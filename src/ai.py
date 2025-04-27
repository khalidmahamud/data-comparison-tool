from google import generativeai as genai
from google.generativeai.types import GenerationConfigType
from src.config import config

flash = config.api_settings["gemini_flash_1"]

genai.configure(api_key=flash.api_key)
client = genai.GenerativeModel(flash.model)

def ask(query: str, config: GenerationConfigType | None = {}):
  if config is None:
    config = {}

  return client.generate_content(
      contents=query,
      generation_config={"max_output_tokens": flash.max_tokens, **config}
  )

