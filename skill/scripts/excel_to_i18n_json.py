#!/usr/bin/env python3
"""
Excel to i18n JSON converter - specialized for language corpus extraction.
Outputs format: {"en": {"key1": "text1"}, "zh": {"key1": "文本1"}}
"""

import json
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Set
import sys
import os


class ExcelToI18nConverter:
    """Convert Excel language files to i18n JSON format grouped by language."""

    def __init__(self):
        """Initialize the converter."""
        self.ns = {'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

    @staticmethod
    def get_column_index(col_ref: str) -> int:
        """Convert column reference (e.g., 'A', 'AB') to zero-based index."""
        index = 0
        for char in col_ref.upper():
            index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1

    def load_shared_strings(self, zip_ref: zipfile.ZipFile) -> List[str]:
        """Load shared strings table from Excel file."""
        shared_strings = []
        try:
            sst_xml = zip_ref.read('xl/sharedStrings.xml')
            sst_root = ET.fromstring(sst_xml)

            for si in sst_root.findall('.//main:si', self.ns):
                text_parts = []
                for t in si.findall('.//main:t', self.ns):
                    if t.text:
                        text_parts.append(t.text)
                shared_strings.append(''.join(text_parts))
        except KeyError:
            pass
        return shared_strings

    def find_sheet_path(self, zip_ref: zipfile.ZipFile, sheet_name: Optional[str] = None) -> tuple:
        """Find the XML file path for a specific sheet."""
        wb_xml = zip_ref.read('xl/workbook.xml')
        wb_root = ET.fromstring(wb_xml)

        sheets = {}
        for sheet in wb_root.findall('.//main:sheet', self.ns):
            name = sheet.get('name')
            r_id = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            sheets[name] = r_id

        if not sheets:
            raise ValueError("No sheets found in workbook")

        if sheet_name:
            if sheet_name not in sheets:
                available = ", ".join(sheets.keys())
                raise ValueError(f"Sheet '{sheet_name}' not found. Available: {available}")
            target_name = sheet_name
            target_rid = sheets[sheet_name]
        else:
            target_name = list(sheets.keys())[0]
            target_rid = sheets[target_name]

        rels_xml = zip_ref.read('xl/_rels/workbook.xml.rels')
        rels_root = ET.fromstring(rels_xml)

        for rel in rels_root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
            if rel.get('Id') == target_rid:
                sheet_path = 'xl/' + rel.get('Target')
                return sheet_path, target_name

        raise ValueError(f"Could not find sheet file for '{target_name}'")

    def parse_row(self, row_elem, shared_strings: List[str], max_col: int) -> Dict[int, str]:
        """Parse a row element and extract cell values."""
        cells_dict = {}
        for cell in row_elem.findall('.//main:c', self.ns):
            cell_ref = cell.get('r')
            col_letter = ''.join([c for c in cell_ref if c.isalpha()])
            col_idx = self.get_column_index(col_letter)

            if col_idx > max_col:
                continue

            v = cell.find('.//main:v', self.ns)
            if v is not None and v.text:
                t = cell.get('t')
                if t == 's':
                    idx = int(v.text)
                    if idx < len(shared_strings):
                        cells_dict[col_idx] = shared_strings[idx]
                else:
                    cells_dict[col_idx] = v.text

        return cells_dict

    def convert_to_i18n_json(
        self,
        input_path: str,
        output_path: str,
        sheet_name: Optional[str] = None,
        start_row: int = 2,
        end_row: Optional[int] = None,
        key_column: str = 'key',
        default_column: str = 'default',
        default_lang: str = 'en',
        exclude_columns: Optional[Set[str]] = None,
        use_abbrev: bool = True
    ) -> None:
        """
        Convert Excel file to i18n JSON format grouped by language.

        Args:
            input_path: Path to input .xlsx file
            output_path: Path to output .json file
            sheet_name: Name of sheet to process (None for first sheet)
            start_row: Starting row number (1-based, inclusive)
            end_row: Ending row number (1-based, inclusive, None for last row)
            key_column: Column name to use as text key (e.g., 'key')
            default_column: Column name for default language (e.g., 'default')
            default_lang: Language code for default column (e.g., 'en')
            exclude_columns: Set of column names to exclude from output
            use_abbrev: Use abbreviated language codes (zh_rCN -> zh-CN)
        """
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"File not found: {input_path}")

        if exclude_columns is None:
            exclude_columns = {'特殊说明', 'is_android', 'location'}

        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            # Load shared strings
            shared_strings = self.load_shared_strings(zip_ref)
            print(f"Loaded {len(shared_strings)} shared strings")

            # Find sheet
            sheet_path, actual_sheet_name = self.find_sheet_path(zip_ref, sheet_name)
            print(f"Processing sheet: {actual_sheet_name}")

            # Read sheet XML
            sheet_xml = zip_ref.read(sheet_path)
            sheet_root = ET.fromstring(sheet_xml)
            rows = sheet_root.findall('.//main:row', self.ns)
            print(f"Total rows in sheet: {len(rows)}")

            if not rows:
                raise ValueError("Sheet is empty")

            # Parse header row
            first_row_elem = rows[0]
            header_cells = self.parse_row(first_row_elem, shared_strings, 100)

            # Build headers list
            max_col = max(header_cells.keys()) if header_cells else 0
            headers = []
            for i in range(max_col + 1):
                header = header_cells.get(i, '').strip()
                headers.append(header if header else f'col_{i}')

            print(f"Headers ({len(headers)} columns): {headers[:10]}{'...' if len(headers) > 10 else ''}")

            # Identify language columns
            meta_columns = {key_column, default_column} | exclude_columns
            language_columns = {}
            for idx, header in enumerate(headers):
                if header and header not in meta_columns and not header.startswith('col_'):
                    # Abbreviate language code if needed
                    lang_code = header
                    if use_abbrev:
                        lang_code = header.replace('_r', '-').replace('_', '-')
                    language_columns[idx] = lang_code

            print(f"\nDetected {len(language_columns)} language columns:")
            for idx, lang in language_columns.items():
                print(f"  {headers[idx]} -> {lang}")

            # Find key and default columns
            key_idx = headers.index(key_column) if key_column in headers else None
            default_idx = headers.index(default_column) if default_column in headers else None

            if default_idx is None:
                raise ValueError(f"Default column '{default_column}' not found in headers")

            # Extract data grouped by language
            result = {}
            actual_end_row = end_row if end_row else len(rows)

            for row_elem in rows:
                row_num = int(row_elem.get('r'))

                if start_row <= row_num <= actual_end_row:
                    cells_dict = self.parse_row(row_elem, shared_strings, max_col)

                    # Get key and default values
                    key_value = cells_dict.get(key_idx, '').strip() if key_idx is not None else ''
                    default_value = cells_dict.get(default_idx, '').strip()

                    # Skip empty rows
                    if not default_value:
                        continue

                    # Use key if available, otherwise use default value as key
                    text_id = key_value if key_value else default_value

                    # Add default language
                    if default_lang not in result:
                        result[default_lang] = {}
                    result[default_lang][text_id] = default_value

                    # Add other languages
                    for col_idx, lang_code in language_columns.items():
                        text_value = cells_dict.get(col_idx, '').strip()
                        if text_value:
                            if lang_code not in result:
                                result[lang_code] = {}
                            result[lang_code][text_id] = text_value

            print(f"\nExtracted translations for {len(result)} languages")
            for lang, texts in result.items():
                print(f"  {lang}: {len(texts)} texts")

            # Write JSON
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, ensure_ascii=False, indent=2)

            print(f"\nSaved to: {output_path}")

            # Show sample
            if result:
                print(f"\nSample output:")
                sample_langs = list(result.keys())[:2]
                for lang in sample_langs:
                    print(f"\n{lang}:")
                    items = list(result[lang].items())[:3]
                    for key, value in items:
                        print(f"  \"{key}\": \"{value}\"")


def main():
    """Command line interface."""
    import argparse

    parser = argparse.ArgumentParser(
        description="Convert Excel language files to i18n JSON format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Output Format:
  {
    "en": {
      "key1": "English text",
      "key2": "Another text"
    },
    "zh-CN": {
      "key1": "简体中文",
      "key2": "另一个文本"
    }
  }

Examples:
  # Basic usage
  python excel_to_i18n_json.py language.xlsx i18n.json --sheet buff-web --start 275 --end 305

  # Use full language codes (zh_rCN instead of zh-CN)
  python excel_to_i18n_json.py input.xlsx output.json --no-abbrev

  # Custom columns
  python excel_to_i18n_json.py input.xlsx output.json --key-col id --default-col en --default-lang en-US
        """
    )

    parser.add_argument('input_file', help='Path to input Excel file (.xlsx)')
    parser.add_argument('output_file', help='Path to output JSON file')
    parser.add_argument('--sheet', dest='sheet_name', default=None,
                        help='Sheet name to process (default: first sheet)')
    parser.add_argument('--start', dest='start_row', type=int, default=2,
                        help='Starting row number (default: 2)')
    parser.add_argument('--end', dest='end_row', type=int, default=None,
                        help='Ending row number (default: last row)')
    parser.add_argument('--key-col', dest='key_column', default='key',
                        help='Column name for text keys (default: key)')
    parser.add_argument('--default-col', dest='default_column', default='default',
                        help='Column name for default language (default: default)')
    parser.add_argument('--default-lang', dest='default_lang', default='en',
                        help='Language code for default column (default: en)')
    parser.add_argument('--no-abbrev', dest='use_abbrev', action='store_false', default=True,
                        help='Do not abbreviate language codes (keep zh_rCN format)')

    args = parser.parse_args()

    try:
        converter = ExcelToI18nConverter()
        converter.convert_to_i18n_json(
            input_path=args.input_file,
            output_path=args.output_file,
            sheet_name=args.sheet_name,
            start_row=args.start_row,
            end_row=args.end_row,
            key_column=args.key_column,
            default_column=args.default_column,
            default_lang=args.default_lang,
            use_abbrev=args.use_abbrev
        )
        sys.exit(0)

    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
