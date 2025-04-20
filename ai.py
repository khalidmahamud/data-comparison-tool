from google import generativeai as genai
from google.generativeai.types import GenerationConfigType
import time
from config import config

flash = config.api_settings["gemini_flash_1"]

genai.configure(api_key=flash.api_key)
client = genai.GenerativeModel(flash.model)


class RateLimiter:
  def __init__(self, max_calls: int, time_window: int):
    self.max_calls = max_calls
    self.time_window = time_window
    self.calls = 0
    self.window_start = time.time()
    print(
        f"Rate limiter initialized: {max_calls} calls per {time_window} seconds")

  def __call__(self):
    current_time = time.time()
    time_since_start = current_time - self.window_start

    if time_since_start >= self.time_window:
      print(f"Rate limit window reset after {time_since_start:.1f} seconds")
      self.calls = 0
      self.window_start = current_time

    if self.calls >= self.max_calls:
      sleep_time = self.time_window - time_since_start
      print(
          f"Rate limit reached ({self.calls}/{self.max_calls} calls). Waiting {sleep_time:.1f} seconds...")
      if sleep_time > 0:
        time.sleep(sleep_time)
      self.calls = 0
      self.window_start = time.time()
      print("Rate limit window reset after waiting")

    self.calls += 1
    print(
        f"API call {self.calls}/{self.max_calls} in current window ({time_since_start:.1f}s elapsed)"
    )


rate_limiter = RateLimiter(max_calls=15, time_window=60)


def ask(query: str, config: GenerationConfigType | None = {}):
  rate_limiter()

  if config is None:
    config = {}

  print(query)

  return client.generate_content(contents=query,generation_config={ "max_output_tokens": flash.max_tokens, **config })
