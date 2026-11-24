#!/usr/bin/env python3
"""
Leadership Team Finder
----------------------
Give it company domains -> it finds the About / Team / Leadership pages
and extracts leadership names, titles and LinkedIn links into a CSV.

Usage:
    pip install requests beautifulsoup4 lxml
    pip install anthropic          # optional, for AI extraction (much more accurate)
    pip install playwright && playwright install chromium   # optional, for JS-heavy sites

    python leadership_scraper.py tata.com infosys.com
    python leadership_scraper.py -i domains.txt -o leaders.csv
    python leadership_scraper.py -i domains.txt --ai        # needs ANTHROPIC_API_KEY
    python leadership_scraper.py -i domains.txt --js        # render JavaScript pages
"""
import argparse, csv, json, os, re, sys, time
from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup

UA = "Mozilla/5.0 (compatible; LeadershipFinder/1.0)"
TIMEOUT = 15
DELAY = 1.0  # polite delay between requests (seconds)

PAGE_KEYWORDS = ["leadership", "management", "our-team", "team", "about", "board",
                 "directors", "founders", "executive", "people", "who-we-are", "company"]
COMMON_PATHS = ["/about", "/about-us", "/leadership", "/team", "/our-team",
                "/management", "/about/leadership", "/about-us/leadership",
                "/company/leadership", "/board-of-directors", "/who-we-are"]

TITLE_RE = re.compile(
    r"\b(CEO|CTO|CFO|COO|CMO|CIO|CHRO|CRO|CPO|Founder|Co-?Founder|Chairman|Chairperson|"
    r"President|Vice President|VP|Chief [A-Z][a-z]+(?: [A-Z][a-z]+)* Officer|"
    r"Managing Director|Executive Director|Director|Head of [A-Za-z &]+|Partner|"
    r"General Manager|Board Member|Independent Director|Whole-?time Director)\b", re.I)
NAME_RE = re.compile(r"^(?:(?:Mr|Ms|Mrs|Dr|Prof)\.?\s+)?[A-Z][a-zA-Z'\-]+(?:\s+[A-Z]\.?)?(?:\s+[A-Z][a-zA-Z'\-]+){0,3}$")


class Fetcher:
    def __init__(self, use_js=False):
        self.s = requests.Session()
        self.s.headers["User-Agent"] = UA
        self.robots = {}
        self.use_js = use_js
        self._browser = None

    def allowed(self, url):
        root = "{0.scheme}://{0.netloc}".format(urlparse(url))
        if root not in self.robots:
            rp = RobotFileParser()
            try:
                r = self.s.get(root + "/robots.txt", timeout=TIMEOUT)
                rp.parse(r.text.splitlines() if r.ok else [])
            except requests.RequestException:
                rp.parse([])
            self.robots[root] = rp
        return self.robots[root].can_fetch(UA, url)

    def get(self, url):
        if not self.allowed(url):
            print(f"   robots.txt disallows {url}", file=sys.stderr)
            return None
        time.sleep(DELAY)
        try:
            r = self.s.get(url, timeout=TIMEOUT, allow_redirects=True)
            html = r.text if r.ok and "html" in r.headers.get("content-type", "") else None
        except requests.RequestException:
            html = None
        # Fall back to a real browser when the page is empty / JS-rendered
        if self.use_js and (html is None or len(BeautifulSoup(html, "lxml").get_text(strip=True)) < 300):
            html = self.get_js(url) or html
        return html

    def get_js(self, url):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            return None
        if self._browser is None:
            self._pw = sync_playwright().start()
            self._browser = self._pw.chromium.launch()
        page = self._browser.new_page(user_agent=UA)
        try:
            page.goto(url, timeout=30000, wait_until="networkidle")
            return page.content()
        except Exception:
            return None
        finally:
            page.close()


def find_candidate_pages(fetcher, base):
    """Return likely leadership/about URLs, best first."""
    scored = {}
    home = fetcher.get(base)
    if home:
        soup = BeautifulSoup(home, "lxml")
        for a in soup.find_all("a", href=True):
            url = urljoin(base, a["href"]).split("#")[0]
            if urlparse(url).netloc.replace("www.", "") != urlparse(base).netloc.replace("www.", ""):
                continue
            hay = (url + " " + a.get_text(" ", strip=True)).lower()
            score = sum((len(PAGE_KEYWORDS) - i) for i, k in enumerate(PAGE_KEYWORDS) if k in hay)
            if score:
                scored[url] = max(scored.get(url, 0), score)
    for p in COMMON_PATHS:  # guesses, low priority
        scored.setdefault(urljoin(base, p), 1)
    return [u for u, _ in sorted(scored.items(), key=lambda x: -x[1])][:8]


def page_text(soup):
    for t in soup(["script", "style", "noscript", "svg", "header", "footer", "nav"]):
        t.decompose()
    return [l.strip() for l in soup.get_text("\n").split("\n") if l.strip()]


def linkedin_links(soup):
    return [a["href"] for a in soup.find_all("a", href=True) if "linkedin.com/in/" in a["href"]]


def extract_jsonld(soup):
    people = []
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(tag.string or "")
        except (json.JSONDecodeError, TypeError):
            continue
        stack = data if isinstance(data, list) else [data]
        while stack:
            d = stack.pop()
            if isinstance(d, dict):
                if d.get("@type") == "Person" and d.get("name"):
                    people.append({"name": d["name"], "title": d.get("jobTitle", "")})
                stack.extend(v for v in d.values() if isinstance(v, (dict, list)))
            elif isinstance(d, list):
                stack.extend(d)
    return people


def extract_heuristic(lines):
    """Pair a name-looking line with an adjacent title-looking line."""
    people, seen = [], set()
    def is_name(s):
        return 3 < len(s) < 45 and NAME_RE.match(s) and not TITLE_RE.search(s)
    for i, line in enumerate(lines):
        # "Jane Doe, CEO" / "Jane Doe - Chief Financial Officer"
        m = re.match(r"^(.{3,45}?)\s*[,\-–|:]\s*(.{2,80})$", line)
        if m and is_name(m.group(1)) and TITLE_RE.search(m.group(2)):
            cand = (m.group(1), m.group(2))
        elif len(line) <= 90 and TITLE_RE.search(line):
            name = next((lines[j] for j in (i - 1, i + 1) if 0 <= j < len(lines) and is_name(lines[j])), None)
            cand = (name, line) if name else None
        else:
            cand = None
        if cand and cand[0] not in seen:
            seen.add(cand[0])
            people.append({"name": cand[0], "title": cand[1]})
    return people


def extract_ai(lines, url, model):
    import anthropic
    client = anthropic.Anthropic()
    text = "\n".join(lines)[:40000]
    prompt = ("From this company web page, list the leadership team / executives / board members / "
              "founders. Return ONLY a JSON array of objects with keys \"name\" and \"title\". "
              "Return [] if none. Do not invent anyone.\n\nURL: " + url + "\n\nPAGE TEXT:\n" + text)
    msg = client.messages.create(model=model, max_tokens=2000,
                                 messages=[{"role": "user", "content": prompt}])
    raw = "".join(b.text for b in msg.content if b.type == "text")
    raw = re.sub(r"(?:json)?|", "", raw).strip()
    try:
        return [p for p in json.loads(raw) if p.get("name")]
    except json.JSONDecodeError:
        return []


def process_domain(fetcher, domain, use_ai, model):
    domain = re.sub(r"^https?://", "", domain.strip()).strip("/")
    base = "https://" + domain
    print(f"-> {domain}", file=sys.stderr)
    results, seen = [], set()
    for url in find_candidate_pages(fetcher, base):
        html = fetcher.get(url)
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        lis = linkedin_links(soup)
        people = extract_jsonld(soup)
        method = "jsonld"
        if not people:
            lines = page_text(soup)
            if use_ai:
                people, method = extract_ai(lines, url, model), "ai"
            else:
                people, method = extract_heuristic(lines), "heuristic"
        for p in people:
            key = p["name"].lower()
            if key in seen:
                continue
            seen.add(key)
            slug = re.sub(r"[^a-z]", "", p["name"].split()[0].lower()) if p["name"].split() else ""
            li = next((l for l in lis if slug and slug in l.lower()), "")
            results.append({"domain": domain, "name": p["name"], "title": p.get("title", ""),
                            "linkedin": li, "source_url": url, "method": method})
        if len(results) >= 3 and any(k in url.lower() for k in ("leader", "team", "management", "board")):
            break  # found a proper leadership page; stop crawling
    print(f"   found {len(results)} people", file=sys.stderr)
    return results


def main():
    ap = argparse.ArgumentParser(description="Extract leadership teams from company websites")
    ap.add_argument("domains", nargs="*", help="company domains, e.g. example.com")
    ap.add_argument("-i", "--input", help="text/CSV file with one domain per line")
    ap.add_argument("-o", "--output", default="leadership.csv")
    ap.add_argument("--ai", action="store_true", help="use Claude for extraction (needs ANTHROPIC_API_KEY)")
    ap.add_argument("--model", default="claude-sonnet-5-5")
    ap.add_argument("--js", action="store_true", help="render JS-heavy sites with Playwright")
    a = ap.parse_args()

    domains = list(a.domains)
    if a.input:
        with open(a.input, encoding="utf-8") as f:
            domains += [l.split(",")[0].strip() for l in f if l.strip() and not l.lower().startswith("domain")]
    if not domains:
        ap.error("give at least one domain or --input file")

    fetcher = Fetcher(use_js=a.js)
    fields = ["domain", "name", "title", "linkedin", "source_url", "method"]
    with open(a.output, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for d in domains:
            try:
                for row in process_domain(fetcher, d, a.ai, a.model):
                    w.writerow(row)
                f.flush()
            except Exception as e:
                print(f"   error on {d}: {e}", file=sys.stderr)
    print(f"Done -> {a.output}", file=sys.stderr)


if __name__ == "__main__":
    main()