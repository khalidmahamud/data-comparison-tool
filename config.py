from pathlib import Path
from typing import Optional, Dict
import yaml
from dataclasses import  dataclass


@dataclass
class ProcessingConfig:
    batch_size: int
    max_retries: int
    retry_delay: int
    save_interval: int
    start_row: int


@dataclass
class ApiConfig:
    api_key: str
    model: str
    max_tokens: int


@dataclass
class FileSettings:
    input_file: str
    output_file: str
    prompts_file: str
    chunks_directory: str
    merged_file: str


@dataclass
class Config:
    processing: ProcessingConfig
    api_settings: Dict[str, ApiConfig]
    file_settings: FileSettings


def load_config(config_path: str) -> Config:
    config_path = Path(config_path)

    if not config_path.exists():
        raise FileNotFoundError(f"Config file '{config_path}' not found")
    
    with config_path.open('r') as f:
        config_dict = yaml.safe_load(f)
    return Config(
        processing=ProcessingConfig(**config_dict['processing']),
        api_settings={
            key: ApiConfig(**value) 
            for key, value in config_dict['api_settings'].items()
            if not key.startswith('#')
        },
        file_settings=FileSettings(**config_dict['file_settings'])
    )


config = load_config('config_flash.yaml')

 
