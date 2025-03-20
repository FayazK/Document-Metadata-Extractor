# Document Metadata Extractor

A powerful utility for extracting metadata from PDF and DOCX documents using Google's Gemini AI.

## Overview

This tool scans documents in a specified directory, extracts their text content, and uses Gemini AI to identify and extract structured metadata. The results are compiled into an Excel spreadsheet for easy review and analysis.

## Features

- Supports PDF and DOCX file formats
- Extracts comprehensive metadata including:
  - General document information (title, type, dates, version)
  - Document-specific metadata based on document type
  - Keywords, authors, and other contextual information
- Exports results to a formatted Excel file
- Includes logging for troubleshooting
- Handles errors gracefully
- Includes progress reporting for large batches

## Requirements

- Python 3.7+
- Dependencies:
  - `python-docx`: For DOCX file processing
  - `PyPDF2`: For PDF file processing
  - `google-generativeai`: For AI metadata extraction
  - `openpyxl`: For Excel output generation
  - `pathlib`: For path handling

## Installation

1. Clone or download this repository
2. Install required dependencies:

```bash
pip install python-docx PyPDF2 google-generativeai openpyxl
```

3. Create a `config.json` file with your Gemini API key:

```json
{
  "api_key": "YOUR_GEMINI_API_KEY"
}
```

## Usage

Run the script from the command line:

```bash
python metadata_extractor.py --dir "path/to/documents" --output "results.xlsx" --output-dir "output" --config "config.json"
```

### Arguments

- `--dir`, `-d`: Directory containing documents to process (required)
- `--output`, `-o`: Output Excel filename (default: "metadata_output.xlsx")
- `--output-dir`: Directory to save JSON responses and other output files (default: "output")
- `--config`, `-c`: Path to config file with API key (default: "config.json")

### Output Files

The script generates the following outputs:

1. An Excel file with all extracted metadata (saved to the output directory)
2. JSON files for each processed document containing the raw Gemini API responses (saved in the output directory with the same base filename as the original document)

## Metadata Schema

The tool extracts metadata according to a comprehensive schema, including:

### General Metadata (for all document types)
- Document title, type, creation date
- Version/revision information
- Source, keywords, and summary
- File details and confidentiality level
- Relevant products and geographic regions

### Document-Specific Metadata
Tailored metadata extraction based on document type:

- **Research Papers**: Authors, journal name, DOI, methodology, findings
- **Test Documents**: Test name, standards, equipment, materials, results
- **EPDs**: Declared unit, GWP, LCA practitioner, validity period
- **Case Studies**: Project details, challenges, benefits, outcomes
- **Technical Product Data**: Product specs, application instructions, references
- **ASTM Standards**: Designation, issue year, title, relevant sections

## Benefits

- **Time Saving**: Automate metadata extraction from large document collections
- **Consistency**: Apply the same extraction criteria across all documents
- **AI-Powered**: Leverage Gemini AI for intelligent content analysis
- **Structured Output**: Get organized results ready for database import
- **Complete Records**: Store both the processed metadata and raw AI responses

## Contributing

Contributions, bug reports, and feature requests are welcome! Feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

---

Made with ❤️ by [Fayaz K](https://fayazk.com)