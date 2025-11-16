---
name: lang-excel2json
description: Convert multilingual Excel files to i18n JSON format. Use for extracting translation data from Excel spreadsheets into internationalization-ready JSON files.
---

# Excel to i18n JSON Converter

A specialized skill for converting multilingual Excel files into i18n JSON format, perfect for internationalization projects.

## When to Use This Skill

This skill should be triggered when:
- Converting Excel translation files to i18n JSON format
- Extracting multilingual content from Excel spreadsheets
- Preparing translation data for internationalization projects
- Working with Excel files containing multiple language columns

## Features

- Extract multilingual data from Excel files
- Output in standard i18n JSON format: `{"en": {"key": "text"}, "zh": {"key": "文本"}}`
- Support for multiple sheets and row ranges
- Automatic language column detection
- Custom key and default language column configuration

## Quick Reference

### Example Usage

**Basic conversion**:
```bash
python scripts/excel_to_i18n_json.py language.xlsx output.json \
  --sheet "apk词条" --start 10 --end 25 \
  --key-col "Picture（optional）" \
  --default-col "英语/en" \
  --default-lang en
```

**With custom parameters**:
```bash
python scripts/excel_to_i18n_json.py input.xlsx i18n.json \
  --sheet "translations" \
  --start 2 \
  --end 100 \
  --key-col "key" \
  --default-col "en" \
  --default-lang en
```

### Command Options

- `input_file` - Path to input Excel file (.xlsx)
- `output_file` - Path to output JSON file
- `--sheet` - Sheet name to process (default: first sheet)
- `--start` - Starting row number (default: 2)
- `--end` - Ending row number (default: last row)
- `--key-col` - Column name for text keys (default: "key")
- `--default-col` - Column name for default language (default: "default")
- `--default-lang` - Language code for default column (default: "en")
- `--no-abbrev` - Keep full language codes like zh_rCN instead of zh-CN

## Output Format

The script generates JSON in standard i18n format:

```json
{
  "en": {
    "welcome": "Welcome",
    "goodbye": "Goodbye"
  },
  "zh-CN": {
    "welcome": "欢迎",
    "goodbye": "再见"
  },
  "ja": {
    "welcome": "ようこそ",
    "goodbye": "さようなら"
  }
}
```

## Notes

- Supports .xlsx Excel format
- Automatically detects language columns
- Skips empty cells and rows
- Provides progress output during conversion
