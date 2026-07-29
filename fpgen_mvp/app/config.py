from pathlib import Path
import os

from dotenv import load_dotenv


ROOT_DIR = Path(__file__).resolve().parents[1]
load_dotenv(ROOT_DIR / ".env")


class Settings:
    app_name: str = "FP-GEN MVP API"
    llm_provider: str = os.getenv("LLM_PROVIDER", "openai")
    openai_api_key: str = os.getenv("OPENAI_API_KEY", "")
    openai_model: str = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    litellm_url: str = os.getenv("LITELLM_URL", "")
    litellm_master_key: str = os.getenv("LITELLM_MASTER_KEY", "")
    enable_openai_synthesis: bool = os.getenv("ENABLE_OPENAI_SYNTHESIS", "true").lower() == "true"
    enable_indd_export: bool = os.getenv("FPGEN_ENABLE_INDD_EXPORT", "true").lower() == "true"
    storage_dir: Path = Path(os.getenv("FPGEN_STORAGE_DIR", ROOT_DIR / "data" / "proposals"))
    export_dir: Path = Path(os.getenv("FPGEN_EXPORT_DIR", ROOT_DIR / "data" / "exports"))
    indd_commercial_template_path: Path = Path(
        os.getenv(
            "FPGEN_INDD_COMMERCIAL_TEMPLATE",
            ROOT_DIR / "examples" / "INDD" / "cp" / "I. Commercial.indd",
        )
    )
    reference_projects_path: Path = Path(
        os.getenv("FPGEN_REFERENCE_PROJECTS", ROOT_DIR / "data" / "reference_projects.json")
    )
    assets_dir: Path = Path(os.getenv("FPGEN_ASSETS_DIR", ROOT_DIR / "data" / "assets"))
    brand_assets_dir: Path = Path(
        os.getenv("FPGEN_BRAND_ASSETS_DIR", ROOT_DIR / "data" / "brand")
    )
    assets_registry_path: Path = Path(
        os.getenv("FPGEN_ASSETS_REGISTRY", ROOT_DIR / "data" / "assets_registry.json")
    )
    personnel_path: Path = Path(os.getenv("FPGEN_PERSONNEL", ROOT_DIR / "data" / "personnel.json"))
    team_presets_path: Path = Path(os.getenv("FPGEN_TEAM_PRESETS", ROOT_DIR / "data" / "team_presets.json"))
    default_currency: str = os.getenv("DEFAULT_CURRENCY", "USD")
    default_overhead_pct: float = float(os.getenv("DEFAULT_OVERHEAD_PCT", "0.10"))
    travel_unit_cost_usd: float = float(os.getenv("TRAVEL_UNIT_COST_USD", "6000"))
    openai_ssl_verify: bool = os.getenv("OPENAI_SSL_VERIFY", "true").lower() != "false"

    # --- Database (exaBase PostgreSQL; SQLite fallback for local dev) ---
    database_url: str = os.getenv(
        "DATABASE_URL", f"sqlite:///{(ROOT_DIR / 'data' / 'fpgen.db').as_posix()}"
    )
    db_schema: str = os.getenv("FPGEN_DB_SCHEMA", "fpgen")
    db_echo: bool = os.getenv("FPGEN_DB_ECHO", "false").lower() == "true"

    # --- Auth (exaBase oauth2-proxy + Keycloak forwarded headers) ---
    auth_enabled: bool = os.getenv("FPGEN_AUTH_ENABLED", "false").lower() == "true"
    auth_dev_user: str = os.getenv("FPGEN_AUTH_DEV_USER", "dev-admin@example.com")
    auth_dev_groups: str = os.getenv("FPGEN_AUTH_DEV_GROUPS", "fpgen_admin,Finance")
    admin_roles: str = os.getenv("FPGEN_ADMIN_ROLES", "fpgen_admin")
    finance_roles: str = os.getenv("FPGEN_FINANCE_ROLES", "fpgen_admin,Finance")

    # --- Asset files (CVs / experience): File Manager if configured, else local volume ---
    file_manager_url: str = os.getenv("FPGEN_FILE_MANAGER_URL", "")
    asset_storage_dir: Path = Path(
        os.getenv("FPGEN_ASSET_STORAGE_DIR", ROOT_DIR / "data" / "asset_files")
    )


settings = Settings()
