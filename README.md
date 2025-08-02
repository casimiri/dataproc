# Excel Data Processor with OpenAI Enhancement

A Python tool for processing Excel files containing plant/species data with AI-enhanced extraction of countries, variety names, Latin names, and common names.

## Features

- **Smart Country Detection**: Uses OpenAI API to accurately extract countries from address fields
- **Plant Information Extraction**: AI-powered identification of:
  - Latin/scientific names
  - Standardized common names
  - Variety/cultivar names
- **Robust Data Processing**: Handles multiple varieties per row, dose parsing, and address parsing
- **Fallback Support**: Works with or without OpenAI API key (uses hardcoded mappings as fallback)

## Setup

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Configure OpenAI API (optional but recommended):
```bash
cp .env.example .env
# Edit .env and add your OpenAI API key
```

## Usage

```bash
python excel_processor.py input_file.xlsx [output_file.xlsx]
```

## OpenAI Integration

The tool uses OpenAI's GPT-3.5-turbo model to:
- Extract countries from complex address strings
- Identify plant species and provide accurate Latin names
- Standardize common names and variety names
- Handle edge cases that hardcoded mappings might miss

If no OpenAI API key is provided, the tool falls back to hardcoded mappings and basic text parsing.