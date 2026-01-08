"""
Configuration management for IHADIS Data Comparison Tool.

Supports multiple configuration sources:
1. Environment variables (highest priority for server secrets)
2. Database (for user-configurable settings)
3. YAML file (fallback for initial setup)
"""
import os
from pathlib import Path
from typing import Optional, Dict
from dataclasses import dataclass, field

import yaml
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()


@dataclass
class ProcessingConfig:
    batch_size: int = 5
    max_retries: int = 3
    retry_delay: int = 0
    save_interval: int = 5
    start_row: int = 0


@dataclass
class ApiConfig:
    api_key: str = ''
    model: str = ''
    max_tokens: Optional[int] = None


@dataclass
class FileSettings:
    input_file: str = ''
    output_file: str = ''
    prompts_file: str = 'prompts.txt'
    chunks_directory: str = 'chunks'
    merged_file: str = 'merged_output.xlsx'
    rows_per_chunk: int = 500
    action: str = 'split'


@dataclass
class ExcelSettings:
    sheet_name: str = 'Sheet1'
    columns: Dict[str, str] = field(default_factory=lambda: {
        'primary_text': 'bn',
        'secondary_text': 'step_1_output',
        'ratio': 'ratio',
        'number': 'id',
        'arabic_text': 'ar'
    })


@dataclass
class Config:
    processing: ProcessingConfig
    api_settings: Dict[str, ApiConfig]
    file_settings: FileSettings
    excel_settings: ExcelSettings


def load_config_from_yaml(config_path: str = 'config_flash.yaml') -> Optional[Config]:
    """Load configuration from YAML file."""
    config_path = Path(config_path)

    if not config_path.exists():
        return None

    try:
        with config_path.open('r') as f:
            config_dict = yaml.safe_load(f)

        return Config(
            processing=ProcessingConfig(**config_dict.get('processing', {})),
            api_settings={
                key: ApiConfig(**value)
                for key, value in config_dict.get('api_settings', {}).items()
                if not key.startswith('#')
            },
            file_settings=FileSettings(**config_dict.get('file_settings', {})),
            excel_settings=ExcelSettings(**config_dict.get('excel_settings', {
                'sheet_name': 'Sheet1',
                'columns': {
                    'primary_text': 'bn',
                    'secondary_text': 'step_1_output',
                    'ratio': 'ratio',
                    'number': 'id',
                    'arabic_text': 'ar'
                }
            }))
        )
    except Exception as e:
        print(f"Warning: Failed to load config from YAML: {e}")
        return None


def get_default_config() -> Config:
    """Get default configuration."""
    return Config(
        processing=ProcessingConfig(),
        api_settings={
            'google': ApiConfig(model='gemini-2.0-flash', max_tokens=8192),
            'claude': ApiConfig(model='claude-3-haiku-20240307', max_tokens=4096),
            'openai': ApiConfig(model='gpt-4o', max_tokens=4096),
            'deepseek': ApiConfig(model='deepseek-chat', max_tokens=4096),
            'grok': ApiConfig(model='grok-1', max_tokens=4096),
        },
        file_settings=FileSettings(),
        excel_settings=ExcelSettings()
    )


def load_config(config_path: str = 'config_flash.yaml') -> Config:
    """
    Load configuration with fallback chain:
    1. Try loading from YAML file
    2. Fall back to defaults

    Note: In production, settings will be loaded from database via get_settings_from_db()
    """
    # Try YAML config first
    yaml_config = load_config_from_yaml(config_path)
    if yaml_config:
        return yaml_config

    # Fall back to defaults
    return get_default_config()


# Environment variable getters
def get_env(key: str, default: str = '') -> str:
    """Get environment variable with default."""
    return os.environ.get(key, default)


def get_env_int(key: str, default: int = 0) -> int:
    """Get environment variable as integer."""
    try:
        return int(os.environ.get(key, default))
    except (ValueError, TypeError):
        return default


def get_env_bool(key: str, default: bool = False) -> bool:
    """Get environment variable as boolean."""
    val = os.environ.get(key, str(default)).lower()
    return val in ('true', '1', 'yes', 'on')


# Server configuration from environment
class ServerConfig:
    """Server-side configuration from environment variables."""

    @staticmethod
    def get_secret_key() -> str:
        return get_env('SECRET_KEY', 'dev-secret-key-change-in-production')

    @staticmethod
    def get_database_url() -> str:
        return get_env('DATABASE_URL', 'sqlite:///data/app.db')

    @staticmethod
    def get_upload_folder() -> str:
        return get_env('UPLOAD_FOLDER', 'uploads')

    @staticmethod
    def get_max_upload_size() -> int:
        return get_env_int('MAX_UPLOAD_SIZE_MB', 50) * 1024 * 1024  # Convert to bytes

    @staticmethod
    def get_google_credentials_path() -> str:
        return get_env('GOOGLE_APPLICATION_CREDENTIALS', 'service_account.json')

    @staticmethod
    def is_production() -> bool:
        return get_env('FLASK_ENV', 'development') == 'production'

    @staticmethod
    def get_host() -> str:
        return get_env('HOST', '0.0.0.0')

    @staticmethod
    def get_port() -> int:
        return get_env_int('PORT', 8000)


# Database-backed settings functions (for runtime use)
def get_settings_from_db() -> Dict:
    """
    Get processing settings from database.
    Import here to avoid circular imports.
    """
    try:
        from .models import Settings
        return Settings.get_all()
    except Exception:
        # Return defaults if database not available
        return {
            'batch_size': '5',
            'max_retries': '3',
            'retry_delay': '0',
            'save_interval': '5',
        }


def get_api_key_from_db(provider: str) -> Optional[Dict]:
    """
    Get API key and settings for a provider from database.
    """
    try:
        from .models import ApiKey
        from .database import decrypt_api_key

        api_key = ApiKey.query.filter_by(provider=provider, is_active=True).first()
        if api_key:
            return {
                'api_key': decrypt_api_key(api_key.api_key_encrypted),
                'model': api_key.model_name,
                'max_tokens': api_key.max_tokens,
            }
    except Exception:
        pass
    return None


def get_project_config(project_id: int) -> Optional[Dict]:
    """
    Get project-specific configuration from database.
    """
    try:
        from .models import Project

        project = Project.query.get(project_id)
        if project:
            return {
                'excel_path': project.excel_path,
                'sheet_name': project.sheet_name,
                'columns': {
                    'primary_text': project.col_primary_text,
                    'secondary_text': project.col_secondary_text,
                    'arabic_text': project.col_arabic_text,
                    'number': project.col_id,
                    'ratio': project.col_ratio,
                },
                'rows_per_chunk': project.rows_per_chunk,
            }
    except Exception:
        pass
    return None


# Try to load config at module import (for backwards compatibility)
try:
    config = load_config('config_flash.yaml')
except Exception:
    config = get_default_config()
