from pathlib import Path
import os
import yaml
from pydantic_settings import BaseSettings, SettingsConfigDict

class Settings(BaseSettings):
    database_url: str = "sqlite:///./data/ndr.sqlite3"
    app_timezone: str = "Asia/Kolkata"
    input_dir: str = "2_Input_MIS"
    output_dir: str = "4_Generated_Calling_Sheets"
    backup_dir: str = "0_Backups"
    admin_email: str = "admin@example.com"
    smtp_host: str = "smtp.example.com"
    smtp_port: int = 587
    smtp_user: str = "ndr@example.com"
    smtp_password: str = "CHANGE_ME"
    smtp_enabled: bool = False
    model_config = SettingsConfigDict(env_file=".env", extra="ignore")

settings = Settings()

def load_rules(path: str = "config/default.yaml") -> dict:
    with Path(path).open(encoding="utf-8") as stream:
        return yaml.safe_load(stream)
