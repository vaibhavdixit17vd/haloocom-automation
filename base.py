# backend/app/models/base.py
import uuid
from datetime import datetime, timezone
from sqlalchemy import DateTime, String, ForeignKey, Integer, Boolean, JSON, Enum as SQLEnum
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship
from sqlalchemy.dialects.postgresql import UUID
import enum

class Base(DeclarativeBase):
    pass

class UserRole(str, enum.Enum):
    ADMIN = "ADMIN"
    STANDARD = "STANDARD"

class JobStatus(str, enum.Enum):
    QUEUED = "QUEUED"
    VALIDATING = "VALIDATING"
    PROCESSING = "PROCESSING"
    FINALIZING = "FINALIZING"
    COMPLETED = "COMPLETED"
    COMPLETED_WITH_ERRORS = "COMPLETED_WITH_ERRORS"
    FAILED = "FAILED"
    CANCELLED = "CANCELLED"

class NormalizedStatus(str, enum.Enum):
    VALID = "valid"
    INVALID = "invalid"
    ACCEPT_ALL = "accept_all"
    UNKNOWN = "unknown"
    RISKY = "risky"