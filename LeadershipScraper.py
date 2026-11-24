#!/usr/bin/env python3

"""
Leadership Team Finder
----------------------

Finds About / Team / Leadership pages on company websites
and extracts leadership names, titles and LinkedIn links.

Usage:

    python main.py tata.com infosys.com

    python main.py -i domains.txt -o leaders.csv

    python main.py -i domains.csv -o leaders.csv

    python main.py -i domains.xlsx -o leaders.csv

    python main.py -i domains.xlsx --ai

    python main.py -i domains.xlsx --js

    python main.py -i domains.xlsx --ai --js

Required:

    pip install requests beautifulsoup4 lxml

For Excel:

    pip install pandas openpyxl

Optional AI:

    pip install anthropic

Optional JavaScript:

    pip install playwright
    playwright install chromium
"""

import argparse
import csv
import json
import os
import re
import sys
import time

from urllib.parse import urljoin, urlparse
from urllib.robotparser import RobotFileParser

import requests
from bs4 import BeautifulSoup


# ============================================================
# CONFIGURATION
# ============================================================

UA = "Mozilla/5.0 (compatible; LeadershipFinder/1.0)"

TIMEOUT = 15

# Delay between website requests
DELAY = 1.0

PAGE_KEYWORDS = [
    "leadership",
    "management",
    "our-team",
    "team",
    "about",
    "board",
    "directors",
    "founders",
    "executive",
    "people",
    "who-we-are",
    "company",
]

COMMON_PATHS = [
    "/about",
    "/about-us",
    "/leadership",
    "/team",
    "/our-team",
    "/management",
    "/about/leadership",
    "/about-us/leadership",
    "/company/leadership",
    "/board-of-directors",
    "/who-we-are",
]

TITLE_RE = re.compile(
    r"\b("
    r"CEO|CTO|CFO|COO|CMO|CIO|CHRO|CRO|CPO|"
    r"Founder|Co-?Founder|"
    r"Chairman|Chairperson|"
    r"President|Vice President|VP|"
    r"Chief [A-Z][a-z]+(?: [A-Z][a-z]+)* Officer|"
    r"Managing Director|Executive Director|"
    r"Director|"
    r"Head of [A-Za-z &]+|"
    r"Partner|"
    r"General Manager|"
    r"Board Member|"
    r"Independent Director|"
    r"Whole-?time Director"
    r")\b",
    re.I
)

NAME_RE = re.compile(
    r"^(?:(?:Mr|Ms|Mrs|Dr|Prof)\.?\s+)?"
    r"[A-Z][a-zA-Z'\-]+"
    r"(?:\s+[A-Z]\.?)?"
    r"(?:\s+[A-Z][a-zA-Z'\-]+){0,3}$"
)


# ============================================================
# DOMAIN NORMALIZATION
# ============================================================

def normalize_domain(domain):
    """
    Converts:

        https://www.example.com
        http://example.com/about
        www.example.com
        example.com

    into:

        example.com
    """

    if not domain:
        return ""

    domain = str(domain).strip()

    if not domain:
        return ""

    if not re.match(r"^https?://", domain, re.I):
        domain = "https://" + domain

    try:
        parsed = urlparse(domain)

        result = parsed.netloc.lower()

        if result.startswith("www."):
            result = result[4:]

        # Remove port
        result = result.split(":")[0]

        return result.strip("/")

    except Exception:
        return ""


# ============================================================
# FETCHER
# ============================================================

class Fetcher:

    def __init__(self, use_js=False):

        self.s = requests.Session()

        self.s.headers["User-Agent"] = UA

        self.robots = {}

        self.use_js = use_js

        self._browser = None
        self._pw = None

    def allowed(self, url):

        root = "{0.scheme}://{0.netloc}".format(
            urlparse(url)
        )

        if root not in self.robots:

            rp = RobotFileParser()

            try:

                r = self.s.get(
                    root + "/robots.txt",
                    timeout=TIMEOUT
                )

                rp.parse(
                    r.text.splitlines()
                    if r.ok
                    else []
                )

            except requests.RequestException:

                rp.parse([])

            self.robots[root] = rp

        return self.robots[root].can_fetch(
            UA,
            url
        )

    def get(self, url):

        if not self.allowed(url):

            print(
                f"   robots.txt disallows {url}",
                file=sys.stderr
            )

            return None

        time.sleep(DELAY)

        try:

            r = self.s.get(
                url,
                timeout=TIMEOUT,
                allow_redirects=True
            )

            content_type = r.headers.get(
                "content-type",
                ""
            ).lower()

            html = (
                r.text
                if r.ok and "html" in content_type
                else None
            )

        except requests.RequestException:

            html = None

        # JavaScript fallback
        if self.use_js:

            text_length = 0

            if html:

                try:

                    text_length = len(
                        BeautifulSoup(
                            html,
                            "lxml"
                        ).get_text(
                            strip=True
                        )
                    )

                except Exception:

                    text_length = 0

            if html is None or text_length < 300:

                js_html = self.get_js(url)

                if js_html:

                    html = js_html

        return html

    def get_js(self, url):

        try:

            from playwright.sync_api import sync_playwright

        except ImportError:

            print(
                "   Playwright is not installed.",
                file=sys.stderr
            )

            print(
                "   Run: pip install playwright",
                file=sys.stderr
            )

            return None

        try:

            if self._browser is None:

                self._pw = sync_playwright().start()

                self._browser = (
                    self._pw.chromium.launch()
                )

            page = self._browser.new_page(
                user_agent=UA
            )

            try:

                page.goto(
                    url,
                    timeout=30000,
                    wait_until="networkidle"
                )

                return page.content()

            except Exception:

                return None

            finally:

                page.close()

        except Exception:

            return None

    def close(self):

        try:

            if self._browser:
                self._browser.close()

            if self._pw:
                self._pw.stop()

        except Exception:

            pass


# ============================================================
# FIND CANDIDATE PAGES
# ============================================================

def find_candidate_pages(fetcher, base):

    """
    Find likely leadership/about/team URLs.
    """

    scored = {}

    home = fetcher.get(base)

    if home:

        soup = BeautifulSoup(
            home,
            "lxml"
        )

        for a in soup.find_all(
            "a",
            href=True
        ):

            url = urljoin(
                base,
                a["href"]
            ).split("#")[0]

            if (
                urlparse(url).netloc.replace("www.", "")
                !=
                urlparse(base).netloc.replace("www.", "")
            ):

                continue

            hay = (
                url
                + " "
                + a.get_text(
                    " ",
                    strip=True
                )
            ).lower()

            score = sum(
                (
                    len(PAGE_KEYWORDS) - i
                )
                for i, k in enumerate(
                    PAGE_KEYWORDS
                )
                if k in hay
            )

            if score:

                scored[url] = max(
                    scored.get(url, 0),
                    score
                )

    # Add common paths
    for p in COMMON_PATHS:

        scored.setdefault(
            urljoin(base, p),
            1
        )

    return [
        u
        for u, _
        in sorted(
            scored.items(),
            key=lambda x: -x[1]
        )
    ][:8]


# ============================================================
# PAGE TEXT
# ============================================================

def page_text(soup):

    for t in soup(
        [
            "script",
            "style",
            "noscript",
            "svg",
            "header",
            "footer",
            "nav"
        ]
    ):

        t.decompose()

    return [
        l.strip()
        for l in soup.get_text(
            "\n"
        ).split("\n")
        if l.strip()
    ]


# ============================================================
# LINKEDIN LINKS
# ============================================================

def linkedin_links(soup):

    return [
        a["href"]
        for a in soup.find_all(
            "a",
            href=True
        )
        if "linkedin.com/in/" in a["href"].lower()
    ]


# ============================================================
# JSON-LD EXTRACTION
# ============================================================

def extract_jsonld(soup):

    people = []

    for tag in soup.find_all(
        "script",
        type="application/ld+json"
    ):

        try:

            data = json.loads(
                tag.string or ""
            )

        except (
            json.JSONDecodeError,
            TypeError
        ):

            continue

        stack = (
            data
            if isinstance(data, list)
            else [data]
        )

        while stack:

            d = stack.pop()

            if isinstance(d, dict):

                if (
                    d.get("@type") == "Person"
                    and d.get("name")
                ):

                    people.append({
                        "name": d["name"],
                        "title": d.get(
                            "jobTitle",
                            ""
                        )
                    })

                stack.extend(
                    v
                    for v in d.values()
                    if isinstance(
                        v,
                        (dict, list)
                    )
                )

            elif isinstance(
                d,
                list
            ):

                stack.extend(d)

    return people


# ============================================================
# HEURISTIC EXTRACTION
# ============================================================

def extract_heuristic(lines):

    """
    Pair a name-looking line with
    an adjacent title-looking line.
    """

    people = []

    seen = set()

    def is_name(s):

        return (
            3 < len(s) < 45
            and NAME_RE.match(s)
            and not TITLE_RE.search(s)
        )

    for i, line in enumerate(lines):

        # Example:
        #
        # Jane Doe, CEO
        # Jane Doe - Chief Financial Officer

        m = re.match(
            r"^(.{3,45}?)\s*[,\-–|:]\s*(.{2,80})$",
            line
        )

        if (
            m
            and is_name(m.group(1))
            and TITLE_RE.search(m.group(2))
        ):

            cand = (
                m.group(1),
                m.group(2)
            )

        elif (
            len(line) <= 90
            and TITLE_RE.search(line)
        ):

            name = next(
                (
                    lines[j]
                    for j in (
                        i - 1,
                        i + 1
                    )
                    if (
                        0 <= j < len(lines)
                        and is_name(lines[j])
                    )
                ),
                None
            )

            cand = (
                name,
                line
            ) if name else None

        else:

            cand = None

        if cand and cand[0] not in seen:

            seen.add(cand[0])

            people.append({
                "name": cand[0],
                "title": cand[1]
            })

    return people


# ============================================================
# AI EXTRACTION
# ============================================================

def extract_ai(lines, url, model):

    try:

        import anthropic

    except ImportError:

        print(
            "Anthropic is not installed.",
            file=sys.stderr
        )

        print(
            "Run: pip install anthropic",
            file=sys.stderr
        )

        return []

    try:

        client = anthropic.Anthropic()

        text = "\n".join(lines)[:40000]

        prompt = (
            "From this company web page, list the "
            "leadership team, executives, board members "
            "and founders.\n\n"

            "Return ONLY a valid JSON array.\n\n"

            "Required format:\n"
            "["
            "{\"name\": \"Full Name\", "
            "\"title\": \"Job Title\"}"
            "]\n\n"

            "Return [] if none.\n"
            "Do not invent anyone.\n\n"

            "URL: "
            + url
            + "\n\nPAGE TEXT:\n"
            + text
        )

        msg = client.messages.create(
            model=model,
            max_tokens=2000,
            messages=[
                {
                    "role": "user",
                    "content": prompt
                }
            ]
        )

        raw = "".join(
            b.text
            for b in msg.content
            if b.type == "text"
        )

        # ----------------------------------------------------
        # FIXED JSON CLEANING
        # ----------------------------------------------------

        raw = raw.strip()

        # Remove opening ```json or ```
        raw = re.sub(
            r"^\s*```(?:json)?\s*",
            "",
            raw,
            flags=re.IGNORECASE
        )

        # Remove closing ```
        raw = re.sub(
            r"\s*```\s*$",
            "",
            raw
        )

        raw = raw.strip()

        # Extract JSON array if AI added text
        # before or after the JSON.
        start = raw.find("[")
        end = raw.rfind("]")

        if start != -1 and end != -1:

            raw = raw[
                start:end + 1
            ]

        try:

            data = json.loads(raw)

        except json.JSONDecodeError:

            return []

        if not isinstance(
            data,
            list
        ):

            return []

        return [
            p
            for p in data
            if isinstance(p, dict)
            and p.get("name")
        ]

    except Exception as e:

        print(
            f"   AI extraction error: {e}",
            file=sys.stderr
        )

        return []


# ============================================================
# PROCESS DOMAIN
# ============================================================

def process_domain(
    fetcher,
    domain,
    use_ai,
    model
):

    domain = normalize_domain(
        domain
    )

    if not domain:

        return []

    base = "https://" + domain

    print(
        f"-> {domain}",
        file=sys.stderr
    )

    results = []

    seen = set()

    for url in find_candidate_pages(
        fetcher,
        base
    ):

        html = fetcher.get(url)

        if not html:
            continue

        soup = BeautifulSoup(
            html,
            "lxml"
        )

        lis = linkedin_links(
            soup
        )

        people = extract_jsonld(
            soup
        )

        method = "jsonld"

        if not people:

            lines = page_text(
                soup
            )

            if use_ai:

                people = extract_ai(
                    lines,
                    url,
                    model
                )

                method = "ai"

            else:

                people = extract_heuristic(
                    lines
                )

                method = "heuristic"

        for p in people:

            name = p.get(
                "name",
                ""
            ).strip()

            if not name:
                continue

            key = name.lower()

            if key in seen:
                continue

            seen.add(key)

            slug = re.sub(
                r"[^a-z]",
                "",
                name.split()[0].lower()
            ) if name.split() else ""

            li = next(
                (
                    l
                    for l in lis
                    if slug
                    and slug in l.lower()
                ),
                ""
            )

            results.append({
                "domain": domain,
                "name": name,
                "title": p.get(
                    "title",
                    ""
                ),
                "linkedin": li,
                "source_url": url,
                "method": method
            })

        if (
            len(results) >= 3
            and any(
                k in url.lower()
                for k in (
                    "leader",
                    "team",
                    "management",
                    "board"
                )
            )
        ):

            break

    print(
        f"   found {len(results)} people",
        file=sys.stderr
    )

    return results


# ============================================================
# INPUT FILE READER
# ============================================================

def load_domains_from_file(
    input_file,
    domain_column=None
):

    extension = os.path.splitext(
        input_file
    )[1].lower()

    # --------------------------------------------------------
    # TXT
    # --------------------------------------------------------

    if extension == ".txt":

        with open(
            input_file,
            encoding="utf-8"
        ) as f:

            domains = []

            for line in f:

                line = line.strip()

                if not line:
                    continue

                if line.lower() in (
                    "domain",
                    "website",
                    "url"
                ):
                    continue

                domains.append(
                    line.split(",")[0].strip()
                )

        return domains

    # --------------------------------------------------------
    # CSV
    # --------------------------------------------------------

    if extension == ".csv":

        with open(
            input_file,
            encoding="utf-8-sig",
            newline=""
        ) as f:

            reader = csv.DictReader(f)

            if not reader.fieldnames:

                raise ValueError(
                    "CSV file has no header row."
                )

            # Explicit column supplied
            if domain_column:

                matching = next(
                    (
                        col
                        for col in reader.fieldnames
                        if col.lower().strip()
                        == domain_column.lower().strip()
                    ),
                    None
                )

                if not matching:

                    raise ValueError(
                        f"Column '{domain_column}' "
                        "was not found in CSV."
                    )

                domain_col = matching

            else:

                domain_col = find_domain_column(
                    reader.fieldnames
                )

                if not domain_col:

                    raise ValueError(
                        "Could not find domain column. "
                        "Use --domain-column."
                    )

            return [
                row.get(
                    domain_col,
                    ""
                ).strip()
                for row in reader
                if row.get(
                    domain_col,
                    ""
                ).strip()
            ]

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    if extension in (
        ".xlsx",
        ".xls"
    ):

        try:

            import pandas as pd

        except ImportError:

            raise ValueError(
                "Excel support requires pandas "
                "and openpyxl.\n"
                "Run: pip install pandas openpyxl"
            )

        df = pd.read_excel(
            input_file
        )

        if df.empty:

            return []

        if domain_column:

            matching = next(
                (
                    col
                    for col in df.columns
                    if str(col).lower().strip()
                    ==
                    domain_column.lower().strip()
                ),
                None
            )

            if not matching:

                raise ValueError(
                    f"Column '{domain_column}' "
                    "was not found in Excel."
                )

            domain_col = matching

        else:

            domain_col = find_domain_column(
                df.columns
            )

            if not domain_col:

                raise ValueError(
                    "Could not find domain column. "
                    "Use --domain-column."
                )

        return [
            str(value).strip()
            for value in df[domain_col]
            if (
                pd.notna(value)
                and str(value).strip()
            )
        ]

    raise ValueError(
        "Unsupported input format. "
        "Use TXT, CSV, XLSX or XLS."
    )


def find_domain_column(columns):

    possible_columns = [
        "domain",
        "website",
        "website/domain",
        "company website",
        "company domain",
        "url",
        "website url",
        "company url"
    ]

    for column in columns:

        normalized = str(
            column
        ).strip().lower()

        if normalized in possible_columns:

            return column

    return None


# ============================================================
# MAIN
# ============================================================

def main():

    parser = argparse.ArgumentParser(
        description=(
            "Extract leadership teams "
            "from company websites."
        )
    )

    parser.add_argument(
        "domains",
        nargs="*",
        help=(
            "Company domains, e.g. "
            "example.com"
        )
    )

    parser.add_argument(
        "-i",
        "--input",
        help=(
            "Input TXT, CSV or Excel file"
        )
    )

    parser.add_argument(
        "--domain-column",
        help=(
            "Column containing domains "
            "in CSV/Excel"
        )
    )

    parser.add_argument(
        "-o",
        "--output",
        default="leadership.csv",
        help=(
            "Output CSV file "
            "(default: leadership.csv)"
        )
    )

    parser.add_argument(
        "--ai",
        action="store_true",
        help=(
            "Use Claude for extraction "
            "(requires ANTHROPIC_API_KEY)"
        )
    )

    parser.add_argument(
        "--model",
        default="claude-sonnet-4-5",
        help="Claude model"
    )

    parser.add_argument(
        "--js",
        action="store_true",
        help=(
            "Render JavaScript-heavy sites "
            "using Playwright"
        )
    )

    args = parser.parse_args()

    # ========================================================
    # GET DOMAINS
    # ========================================================

    domains = list(
        args.domains
    )

    if args.input:

        try:

            file_domains = load_domains_from_file(
                args.input,
                args.domain_column
            )

            domains.extend(
                file_domains
            )

        except Exception as e:

            parser.error(
                str(e)
            )

    # Remove blanks
    domains = [
        d.strip()
        for d in domains
        if d and d.strip()
    ]

    # Remove duplicates
    domains = list(
        dict.fromkeys(
            domains
        )
    )

    if not domains:

        parser.error(
            "give at least one domain "
            "or --input file"
        )

    # ========================================================
    # CREATE OUTPUT DIRECTORY
    # ========================================================

    output_dir = os.path.dirname(
        os.path.abspath(
            args.output
        )
    )

    os.makedirs(
        output_dir,
        exist_ok=True
    )

    # ========================================================
    # OUTPUT
    # ========================================================

    fields = [
        "domain",
        "name",
        "title",
        "linkedin",
        "source_url",
        "method"
    ]

    fetcher = Fetcher(
        use_js=args.js
    )

    try:

        with open(
            args.output,
            "w",
            newline="",
            encoding="utf-8"
        ) as f:

            writer = csv.DictWriter(
                f,
                fieldnames=fields
            )

            writer.writeheader()

            for domain in domains:

                try:

                    rows = process_domain(
                        fetcher,
                        domain,
                        args.ai,
                        args.model
                    )

                    for row in rows:

                        writer.writerow(
                            row
                        )

                    f.flush()

                except Exception as e:

                    print(
                        f"   error on {domain}: {e}",
                        file=sys.stderr
                    )

    finally:

        fetcher.close()

    print(
        f"Done -> {args.output}",
        file=sys.stderr
    )


if __name__ == "__main__":

    main()