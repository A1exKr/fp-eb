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
    default_currency: str = os.getenv("DEFAULT_CURRENCY", "USD")
    default_overhead_pct: float = float(os.getenv("DEFAULT_OVERHEAD_PCT", "0.10"))
    travel_unit_cost_usd: float = float(os.getenv("TRAVEL_UNIT_COST_USD", "6000"))


settings = Settings()
