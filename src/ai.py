from typing import Dict, Any, Optional, Union, Literal
import os
from src.config import config
from openai import OpenAI  # Import OpenAI SDK for Deepseek and Grok

# Type for supported AI providers
ProviderType = Literal["google", "claude", "deepseek", "grok", "openai"]

# Default provider and model mappings
DEFAULT_PROVIDER = "google"
PROVIDER_DEFAULT_MODELS = {
    "google": "gemini-2.0-flash",
    "claude": "claude-3-haiku-20240307",
    "deepseek": "deepseek-chat",
    "grok": "grok-3",  # Note: Example uses "grok-3-beta", you may need to adjust this
    "openai": "gpt-4"
}

class AIResponse:
    """Standardized response object for all AI providers"""
    def __init__(self, text: str, raw_response: Any = None):
        self.text = text
        self.raw_response = raw_response
    
    def __str__(self):
        return self.text

class AIProvider:
    """Base class for AI providers"""
    def __init__(self, api_key: str, model: str):
        self.api_key = api_key
        self.model = model
        
    def generate_content(self, query: str, generation_config: Dict[str, Any]) -> AIResponse:
        """Generate content using the AI provider's API"""
        raise NotImplementedError("This method should be implemented by subclasses")

class GoogleAI(AIProvider):
    """Google AI provider implementation"""
    def __init__(self, api_key: str, model: str):
        super().__init__(api_key, model)
        from google import generativeai as genai
        genai.configure(api_key=api_key)
        self.client = genai.GenerativeModel(model)
        
    def generate_content(self, query: str, generation_config: Dict[str, Any]) -> AIResponse:
        # Google expects "max_output_tokens", so we keep the original config key
        response = self.client.generate_content(
            contents=query,
            generation_config=generation_config
        )
        return AIResponse(text=response.text, raw_response=response)

class ClaudeAI(AIProvider):
    """Anthropic Claude implementation"""
    def __init__(self, api_key: str, model: str):
        super().__init__(api_key, model)
        try:
            import anthropic
            self.client = anthropic.Anthropic(api_key=api_key)
        except ImportError:
            raise ImportError("Please install anthropic package: pip install anthropic")
        
    def generate_content(self, query: str, generation_config: Dict[str, Any]) -> AIResponse:
        # Claude uses "max_tokens", adjust key if provided as "max_output_tokens"
        max_tokens = generation_config.get("max_output_tokens", 1024)
        response = self.client.messages.create(
            model=self.model,
            max_tokens=max_tokens,
            messages=[{"role": "user", "content": query}]
        )
        return AIResponse(text=response.content[0].text, raw_response=response)

class DeepseekAI(AIProvider):
    """DeepSeek AI implementation using OpenAI SDK"""
    def __init__(self, api_key: str, model: str):
        super().__init__(api_key, model)
        # Use OpenAI client with Deepseek's base URL
        self.client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
        
    def generate_content(self, query: str, generation_config: Dict[str, Any]) -> AIResponse:
        # OpenAI SDK expects "max_tokens", adjust from "max_output_tokens" if needed
        max_tokens = generation_config.get("max_output_tokens", 1024)
        messages = [{"role": "user", "content": query}]
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            max_tokens=max_tokens,
        )
        response_text = response.choices[0].message.content
        return AIResponse(text=response_text, raw_response=response)

class GrokAI(AIProvider):
    """Grok AI implementation using OpenAI SDK"""
    def __init__(self, api_key: str, model: str):
        super().__init__(api_key, model)
        # Use OpenAI client with Grok's base URL
        self.client = OpenAI(api_key=api_key, base_url="https://api.x.ai/v1")
        
    def generate_content(self, query: str, generation_config: Dict[str, Any]) -> AIResponse:
        # OpenAI SDK expects "max_tokens", adjust from "max_output_tokens" if needed
        max_tokens = generation_config.get("max_output_tokens", 1024)
        messages = [{"role": "user", "content": query}]
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            max_tokens=max_tokens,
        )
        response_text = response.choices[0].message.content
        return AIResponse(text=response_text, raw_response=response)

class OpenAIProvider(AIProvider):
    """OpenAI ChatGPT implementation"""
    def __init__(self, api_key: str, model: str):
        super().__init__(api_key, model)
        # Use standard OpenAI client
        self.client = OpenAI(api_key=api_key)
        
    def generate_content(self, query: str, generation_config: Dict[str, Any]) -> AIResponse:
        # OpenAI SDK expects "max_tokens", adjust from "max_output_tokens" if needed
        max_tokens = generation_config.get("max_output_tokens", 1024)
        messages = [{"role": "user", "content": query}]
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            max_tokens=max_tokens,
        )
        response_text = response.choices[0].message.content
        return AIResponse(text=response_text, raw_response=response)

# Provider factory
def get_provider(provider_name: ProviderType, model_name: str = None) -> AIProvider:
    """Get the appropriate AI provider instance based on provider name and model"""
    provider_config = None
    config_model = None
    
    # Try to get config from config file first - use exact matching
    if provider_name in config.api_settings:
        provider_config = config.api_settings[provider_name]
        config_model = provider_config.model
    
    # If not found in config, check environment variables
    if not provider_config:
        env_api_key = os.environ.get(f"{provider_name.upper()}_API_KEY")
        if not env_api_key:
            raise ValueError(f"API key for {provider_name} not found in config or environment variables")
        
        provider_config = type('ApiConfig', (), {
            'api_key': env_api_key,
            'model': PROVIDER_DEFAULT_MODELS.get(provider_name),
            'max_tokens': 4096  # Default value
        })
    
    # Model priority: function argument > config file > code default
    final_model = model_name or config_model or PROVIDER_DEFAULT_MODELS.get(provider_name)
    
    # Create provider instance based on name
    if provider_name == "google":
        return GoogleAI(provider_config.api_key, final_model)
    elif provider_name == "claude":
        return ClaudeAI(provider_config.api_key, final_model)
    elif provider_name == "deepseek":
        return DeepseekAI(provider_config.api_key, final_model)
    elif provider_name == "grok":
        return GrokAI(provider_config.api_key, final_model)
    elif provider_name == "openai":
        return OpenAIProvider(provider_config.api_key, final_model)
    else:
        raise ValueError(f"Unsupported AI provider: {provider_name}")

def ask(
    query: str, 
    provider: ProviderType = DEFAULT_PROVIDER, 
    model: Optional[str] = None, 
    config: Optional[Dict[str, Any]] = None
):
    """
    Generate content using specified AI provider and model.
    
    Args:
        query (str): The prompt or query to send to the AI
        provider (str): AI provider ("google", "claude", "deepseek", "grok", "openai")
        model (str, optional): Model name for the specified provider
        config (Dict[str, Any], optional): Generation configuration parameters
        
    Returns:
        AIResponse: A standardized response object with a .text property 
                   containing the generated text
    """
    if config is None:
        config = {}
    
    # Get the appropriate provider instance with provider-specific defaults
    ai_provider = get_provider(provider, model)
    
    # Get max tokens from config if available
    max_tokens = 4096  # Default fallback
    from src.config import config as app_config
    if provider in app_config.api_settings:
        max_tokens = app_config.api_settings[provider].max_tokens
    
    # Merge max tokens with provided config
    generation_config = {"max_output_tokens": max_tokens}
    generation_config.update(config)
    
    return ai_provider.generate_content(query, generation_config)
