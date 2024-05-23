import os
import re
import shutil
from datetime import datetime

def standardize_filename(folder_path, add_default_date=True, rename_folders=False):
    # Improved regular expressions for date formats
    existing_date_pattern = re.compile(r'^\d{6}_')
    full_date_pattern = re.compile(r'(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])')
    full_date_pattern_with_separators = re.compile(r'(\d{4})_(0[1-9]|1[0-2])_(0[1-9]|[12][0-9]|3[01])')
    short_date_pattern = re.compile(r'(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])')
    cn_date_pattern = re.compile(r'(\d{1,2})月(\d{1,2})日')
    space_symbol_normalization_pattern = re.compile(r'[\s\-]+')
    multiple_underscores = re.compile(r'_{2,}')
    trailing_symbols_pattern = re.compile(r'[\s_\-]+(?=\.\w+$)|[\s_\-]+$')

    for filename in os.listdir(folder_path):
        original_path = os.path.join(folder_path, filename)

        # Check if the current entry is a directory
        if os.path.isdir(original_path):
            if not rename_folders:
                continue  # Skip renaming directories if not specified

        # Split the filename and its extension
        name, ext = os.path.splitext(filename)

        # Skip files that already have a standardized date prefix
        if existing_date_pattern.match(name):
            continue

        # Normalize spaces and symbols in filenames
        normalized_name = space_symbol_normalization_pattern.sub('_', name)
        normalized_name = multiple_underscores.sub('_', normalized_name)

        # Remove trailing spaces, underscores, and hyphens
        normalized_name = trailing_symbols_pattern.sub('', normalized_name)

        # Use file creation time instead of last modified time
        creation_timestamp = os.path.getctime(original_path)
        creation_date = datetime.fromtimestamp(creation_timestamp)

        # Date extraction and formatting
        date_str = ""
        date_match = full_date_pattern.search(normalized_name)
        if date_match:
            date_str = datetime.strptime(date_match.group(), '%Y%m%d').strftime('%y%m%d')
            normalized_name = full_date_pattern.sub('', normalized_name)
        else:
            date_match = full_date_pattern_with_separators.search(normalized_name)
            if date_match:
                date_str = datetime.strptime("".join(date_match.groups()), '%Y%m%d').strftime('%y%m%d')
                normalized_name = full_date_pattern_with_separators.sub('', normalized_name)
            else:
                date_match = short_date_pattern.search(normalized_name)
                if date_match:
                    date_str = creation_date.strftime('%y') + date_match.group()
                    normalized_name = short_date_pattern.sub('', normalized_name)
                else:
                    cn_date_match = cn_date_pattern.search(normalized_name)
                    if cn_date_match:
                        month_day_str = f"{int(cn_date_match.group(1)):02d}{int(cn_date_match.group(2)):02d}"
                        date_str = creation_date.strftime('%y') + month_day_str
                        normalized_name = cn_date_pattern.sub('', normalized_name)
                    elif add_default_date:
                        date_str = creation_date.strftime('%y%m%d')
                    else:
                        continue  # Skip renaming if no date found and default date not allowed

        # Build new filename and rename
        new_filename = f"{date_str}_{normalized_name.strip('_')}"
        new_filename = trailing_symbols_pattern.sub('', new_filename)  # Remove trailing symbols again
        new_path = os.path.join(folder_path, f"{new_filename}{ext}")
        print(f"Renaming {original_path} to {new_path}")  # Debug output
        if os.path.isdir(original_path):
            os.rename(original_path, new_path)  # Apply renaming for directories
        else:
            shutil.move(original_path, new_path)  # Apply renaming for files


