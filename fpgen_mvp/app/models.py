"""SQLAlchemy ORM models for FP-GEN master data and proposals.

Mirrors the previous JSON files (personnel, team presets, reference projects,
asset registry) plus a dedicated ``unit_rates`` table that is the source of
truth for per-role rates, and a ``proposals`` table replacing file storage.
"""
from datetime import date, datetime, timezone

from sqlalchemy import (
    Date,
    DateTime,
    Float,
    ForeignKey,
    Integer,
    String,
    Text,
    UniqueConstraint,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.db import Base, JSONType


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


class TimestampMixin:
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=_utcnow)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=_utcnow, onupdate=_utcnow
    )


class Personnel(TimestampMixin, Base):
    __tablename__ = "personnel"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    name: Mapped[str] = mapped_column(String(200), nullable=False)
    title: Mapped[str] = mapped_column(String(200), default="")
    roles: Mapped[list] = mapped_column(JSONType, default=list)
    cv_asset_id: Mapped[str | None] = mapped_column(String(128), nullable=True)


class UnitRate(TimestampMixin, Base):
    """Source of truth for per-role rates. Presets and the fee UI default from here."""

    __tablename__ = "unit_rates"
    __table_args__ = (UniqueConstraint("role", "currency", name="uq_unit_rates_role_currency"),)

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    role: Mapped[str] = mapped_column(String(120), nullable=False)
    rate: Mapped[float] = mapped_column(Float, nullable=False)
    currency: Mapped[str] = mapped_column(String(8), default="USD")
    effective_from: Mapped[date | None] = mapped_column(Date, nullable=True)


class TeamPreset(TimestampMixin, Base):
    __tablename__ = "team_presets"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    name: Mapped[str] = mapped_column(String(200), nullable=False)
    types: Mapped[list] = mapped_column(JSONType, default=list)

    assignments: Mapped[list["TeamAssignment"]] = relationship(
        back_populates="preset",
        cascade="all, delete-orphan",
        order_by="TeamAssignment.id",
    )


class TeamAssignment(TimestampMixin, Base):
    __tablename__ = "team_assignments"

    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    preset_id: Mapped[str] = mapped_column(
        ForeignKey("team_presets.id", ondelete="CASCADE"), nullable=False
    )
    role: Mapped[str] = mapped_column(String(120), nullable=False)
    person_id: Mapped[str | None] = mapped_column(String(64), nullable=True)
    # Optional per-assignment override; when null, defaults from UnitRate for the role.
    rate: Mapped[float | None] = mapped_column(Float, nullable=True)

    preset: Mapped["TeamPreset"] = relationship(back_populates="assignments")


class ReferenceProject(TimestampMixin, Base):
    __tablename__ = "reference_projects"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    name: Mapped[str] = mapped_column(String(300), nullable=False)
    project_type: Mapped[str] = mapped_column(String(120), default="")
    location: Mapped[str] = mapped_column(String(200), default="")
    keywords: Mapped[list] = mapped_column(JSONType, default=list)
    summary: Mapped[str] = mapped_column(Text, default="")


class Asset(TimestampMixin, Base):
    """CV / experience asset metadata. The file itself lives in the File Manager
    (when configured) or a local/volume path, referenced by ``storage_ref``."""

    __tablename__ = "assets"

    id: Mapped[str] = mapped_column(String(128), primary_key=True)
    kind: Mapped[str] = mapped_column(String(32), default="cv")  # cv | experience | template
    role: Mapped[str | None] = mapped_column(String(120), nullable=True)
    reference_project_id: Mapped[str | None] = mapped_column(String(64), nullable=True)
    filename: Mapped[str | None] = mapped_column(String(300), nullable=True)
    storage_ref: Mapped[str | None] = mapped_column(String(500), nullable=True)
    mime_type: Mapped[str | None] = mapped_column(String(120), nullable=True)
    size_bytes: Mapped[int | None] = mapped_column(Integer, nullable=True)
    uploaded_by: Mapped[str | None] = mapped_column(String(200), nullable=True)


class Proposal(TimestampMixin, Base):
    __tablename__ = "proposals"

    proposal_id: Mapped[str] = mapped_column(String(64), primary_key=True)
    payload: Mapped[dict] = mapped_column(JSONType, default=dict)
