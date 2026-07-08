"""initial FP-GEN schema

Revision ID: 0001_initial
Revises:
Create Date: 2026-07-08
"""
from typing import Sequence, Union

import sqlalchemy as sa
from alembic import op

from app.db import DB_SCHEMA as SCHEMA
from app.db import JSONType

revision: str = "0001_initial"
down_revision: Union[str, None] = None
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def _fk(table: str) -> str:
    return f"{SCHEMA}.{table}.id" if SCHEMA else f"{table}.id"


def upgrade() -> None:
    op.create_table(
        "personnel",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("name", sa.String(200), nullable=False),
        sa.Column("title", sa.String(200), nullable=True),
        sa.Column("roles", JSONType, nullable=True),
        sa.Column("cv_asset_id", sa.String(128), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        schema=SCHEMA,
    )
    op.create_table(
        "unit_rates",
        sa.Column("id", sa.Integer, primary_key=True, autoincrement=True),
        sa.Column("role", sa.String(120), nullable=False),
        sa.Column("rate", sa.Float, nullable=False),
        sa.Column("currency", sa.String(8), nullable=True),
        sa.Column("effective_from", sa.Date, nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        sa.UniqueConstraint("role", "currency", name="uq_unit_rates_role_currency"),
        schema=SCHEMA,
    )
    op.create_table(
        "team_presets",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("name", sa.String(200), nullable=False),
        sa.Column("types", JSONType, nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        schema=SCHEMA,
    )
    op.create_table(
        "team_assignments",
        sa.Column("id", sa.Integer, primary_key=True, autoincrement=True),
        sa.Column(
            "preset_id",
            sa.String(64),
            sa.ForeignKey(_fk("team_presets"), ondelete="CASCADE"),
            nullable=False,
        ),
        sa.Column("role", sa.String(120), nullable=False),
        sa.Column("person_id", sa.String(64), nullable=True),
        sa.Column("rate", sa.Float, nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        schema=SCHEMA,
    )
    op.create_table(
        "reference_projects",
        sa.Column("id", sa.String(64), primary_key=True),
        sa.Column("name", sa.String(300), nullable=False),
        sa.Column("project_type", sa.String(120), nullable=True),
        sa.Column("location", sa.String(200), nullable=True),
        sa.Column("keywords", JSONType, nullable=True),
        sa.Column("summary", sa.Text, nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        schema=SCHEMA,
    )
    op.create_table(
        "assets",
        sa.Column("id", sa.String(128), primary_key=True),
        sa.Column("kind", sa.String(32), nullable=True),
        sa.Column("role", sa.String(120), nullable=True),
        sa.Column("reference_project_id", sa.String(64), nullable=True),
        sa.Column("filename", sa.String(300), nullable=True),
        sa.Column("storage_ref", sa.String(500), nullable=True),
        sa.Column("mime_type", sa.String(120), nullable=True),
        sa.Column("size_bytes", sa.Integer, nullable=True),
        sa.Column("uploaded_by", sa.String(200), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        schema=SCHEMA,
    )
    op.create_table(
        "proposals",
        sa.Column("proposal_id", sa.String(64), primary_key=True),
        sa.Column("payload", JSONType, nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=True),
        schema=SCHEMA,
    )


def downgrade() -> None:
    op.drop_table("proposals", schema=SCHEMA)
    op.drop_table("assets", schema=SCHEMA)
    op.drop_table("reference_projects", schema=SCHEMA)
    op.drop_table("team_assignments", schema=SCHEMA)
    op.drop_table("team_presets", schema=SCHEMA)
    op.drop_table("unit_rates", schema=SCHEMA)
    op.drop_table("personnel", schema=SCHEMA)
