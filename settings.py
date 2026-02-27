from pydantic_settings import BaseSettings, SettingsConfigDict
from pathlib import Path
from pydantic import ConfigDict
import sys

BASE_DIR = Path(__file__).parent

def get_env_path():
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        base_path = Path(sys._MEIPASS)
    else:
        # Running as script
        base_path = Path(__file__).parent
    return base_path / '.env'


class Settings(BaseSettings):
    openai_api_key: str
    hunter_api_key: str
    initial_prompt: str
    #temperature: int
    max_tokens: int
    initial_prompt_mayor: str
    initial_prompt_assessor: str
    prompt_format_gis: str
    prompt_format_assessor: str
    model_config = ConfigDict(env_file=get_env_path())
settings = Settings()