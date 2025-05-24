# Data Comparison Tool

A powerful web-based application for comparing and analyzing textual data in Excel files, with support for intelligent AI-assisted corrections and contextual analysis.

## Features

- **Interactive Comparison Interface**: Side-by-side visualization of primary and secondary text columns with difference highlighting
- **AI-Powered Text Analysis**: Utilizes Google's Gemini models to analyze and suggest improvements to content
- **Excel File Handling**:
  - Split large Excel files into manageable chunks
  - Merge processed chunks back into a single file
  - Preserve all Excel styles during processing
- **Visual Difference Highlighting**: See exactly what changed between text versions
- **Similarity Ratio Calculation**: Automatic calculation of similarity between text columns
- **Commenting System**: Add and save comments on individual entries
- **Arabic Text Support**: Special handling for Arabic text content
- **Flexible Configuration**: YAML-based configuration for easy customization

## Getting Started

### Prerequisites

- Python 3.7+
- Flask
- Pandas
- OpenPyXL
- PyYAML

### Installation

1. Clone the repository:

```
git clone https://github.com/your-username/data-comparison-tool.git
cd data-comparison-tool
```

2. Install required packages:

```
pip install -r requirements.txt
```

3. Configure your settings in `config_flash.yaml`

### Usage

1. **Start the web application**:

```
python app.py
```

2. **Split large Excel files into chunks**:

```
python sm.py
```

3. Open your browser and navigate to `http://localhost:5000`

## Configuration

The application is configured through `config_flash.yaml`:

```yaml
processing:
  batch_size: 5
  max_retries: 3
  retry_delay: 0
  save_interval: 5
  start_row: 0

api_settings:
  # Google AI models
  google_gemini:
    api_key: "YOUR_GOOGLE_API_KEY"
    model: "gemini-2.0-flash"
    max_tokens: 8192

  # Anthropic Claude models
  claude_haiku:
    api_key: "YOUR_ANTHROPIC_API_KEY"
    model: "claude-3-haiku-20240307"
    max_tokens: 4096

  # DeepSeek models
  deepseek_chat:
    api_key: "YOUR_DEEPSEEK_API_KEY"
    model: "deepseek-chat"
    max_tokens: 4096

  # Grok models
  grok_1:
    api_key: "YOUR_GROK_API_KEY"
    model: "grok-1"
    max_tokens: 4096

  # OpenAI models
  openai_gpt4o:
    api_key: "YOUR_OPENAI_API_KEY"
    model: "gpt-4o"
    max_tokens: 4096

  openai_gpt4o_mini:
    api_key: "YOUR_OPENAI_API_KEY"
    model: "gpt-4o-mini"
    max_tokens: 4096

file_settings:
  input_file: "input.xlsx"
  output_file: "output.xlsx"
  prompts_file: "prompts.txt"
  chunks_directory: "chunks"
  merged_file: "merged_output.xlsx"
  rows_per_chunk: 500
  action: "merge" # Options: "split" or "merge"

excel_settings:
  sheet_name: "data"
  columns:
    primary_text: "hadith_details"
    secondary_text: "analysis-3"
    ratio: "ratio"
    number: "number"
    arabic_text: "arabic_text"
```

### Configuration Details

#### Processing

- **batch_size**: Number of rows to process in a single batch for AI operations
- **max_retries**: Maximum number of retry attempts if an API call fails
- **retry_delay**: Delay in seconds between retry attempts
- **save_interval**: Number of processed rows after which to save progress
- **start_row**: Row number to start processing from (useful for resuming)

#### API Settings

The application supports multiple AI providers, each with their own configuration:

- **google_gemini**: Configuration for Google AI models

  - **api_key**: Your Google AI API key
  - **model**: The specific model to use (e.g., "gemini-2.0-flash")
  - **max_tokens**: Maximum token limit for generated responses

- **claude_haiku**: Configuration for Anthropic Claude models

  - **api_key**: Your Anthropic API key
  - **model**: The specific Claude model to use (e.g., "claude-3-haiku-20240307")
  - **max_tokens**: Maximum token limit for generated responses

- **deepseek_chat**: Configuration for DeepSeek models

  - **api_key**: Your DeepSeek API key
  - **model**: The specific DeepSeek model to use
  - **max_tokens**: Maximum token limit for generated responses

- **grok_1**: Configuration for Grok models
  - **api_key**: Your Grok API key
  - **model**: The specific Grok model to use
  - **max_tokens**: Maximum token limit for generated responses

- **openai_gpt4o**: Configuration for OpenAI GPT-4o models

  - **api_key**: Your OpenAI API key
  - **model**: The specific GPT-4o model to use
  - **max_tokens**: Maximum token limit for generated responses

- **openai_gpt4o_mini**: Configuration for OpenAI GPT-4o mini models

  - **api_key**: Your OpenAI API key
  - **model**: The specific GPT-4o mini model to use
  - **max_tokens**: Maximum token limit for generated responses

You can configure multiple entries for each provider for load balancing.

#### File Settings

- **input_file**: Path to the source Excel file to process
- **output_file**: Path where processed data will be saved
- **prompts_file**: Path to file containing AI prompts
- **chunks_directory**: Directory where file chunks will be stored
- **merged_file**: Name of the final merged output file
- **rows_per_chunk**: Maximum number of rows in each Excel chunk
- **action**: Operation to perform: "split" (divide file into chunks) or "merge" (combine chunks)

#### Excel Settings

- **sheet_name**: Name of the worksheet to process
- **columns**: Mapping of logical column names to actual Excel column headers
  - **primary_text**: Column containing the primary text data
  - **secondary_text**: Column containing the comparison text data
  - **ratio**: Column where similarity ratios will be stored
  - **number**: Column containing row/entry numbers or IDs
  - **arabic_text**: Column containing Arabic text (if applicable)

## Using Multiple AI Providers

The application supports various AI providers:

1. **Usage in Code**:

```python
from src.ai import ask

# Using Google AI (default provider with model from config or fallback to "gemini-2.0-flash")
response = ask("Your prompt text here")
print(response.text)  # Access the text content of the response

# Using Claude AI with model priority: config file > code default
response = ask("Your prompt text here", provider="claude")
result_text = response.text  # All providers return a consistent .text property

# Using DeepSeek AI with a specific model (explicitly overriding config and defaults)
response = ask("Your prompt text here", provider="deepseek", model="deepseek-coder")
print(f"DeepSeek says: {response.text}")

# Using Grok AI with specific parameters but using model from config or default
response = ask(
    "Your prompt text here",
    provider="grok",
    config={"temperature": 0.7}
)
# Access both the text and raw response if needed
print(response.text)
raw_provider_response = response.raw_response  # Access the original provider-specific response

# Using OpenAI with specific model and parameters
response = ask(
    "Your prompt text here",
    provider="openai",
    model="gpt-4o",
    config={
        "temperature": 0.7,
        "top_p": 0.9,
        "frequency_penalty": 0.1,
        "presence_penalty": 0.1
    }
)
print(f"OpenAI says: {response.text}")

# Using OpenAI GPT-4o mini (more cost-effective option)
response = ask("Your prompt text here", provider="openai", model="gpt-4o-mini")
print(response.text)
```

2. **Environment Variables**:

Instead of using the config file, you can set environment variables for API keys:

```
export GOOGLE_API_KEY=your_google_api_key
export CLAUDE_API_KEY=your_anthropic_api_key
export DEEPSEEK_API_KEY=your_deepseek_api_key
export GROK_API_KEY=your_grok_api_key
export OPENAI_API_KEY=your_openai_api_key
```

3. **Model Selection Priority**:

When selecting which model to use, the system follows this priority order:

1.  Explicitly provided model parameter in the `ask()` function
2.  Model defined in the config file for the provider
3.  Default model defined in the code for the provider

4.  **Standardized Response Object**:

All AI providers return a standardized `AIResponse` object with these properties:

- `response.text` - The text content from the AI provider (always available)
- `response.raw_response` - The original, provider-specific response object

This ensures consistent access to the text content regardless of which AI provider is used.

5.  **Adding New Providers**:

To add support for additional AI providers, extend the AIProvider class in `src/ai.py`.

## Project Structure

- `app.py`: Main Flask web application
- `sm.py`: Excel file splitting and merging utility
- `src/`: Source code directory
  - `generate_cell.py`: AI text generation utilities
  - `config.py`: Configuration management
  - `ai.py`: AI integration module
  - `prompt.py`: Prompt handling utilities
- `templates/`: HTML templates for the web interface
- `static/`: Static assets (CSS, JavaScript, images)
- `chunks/`: Directory for storing Excel file chunks
- `prompts/`: Directory for prompt templates

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
