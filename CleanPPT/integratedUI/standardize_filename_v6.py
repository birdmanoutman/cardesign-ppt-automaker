# standardize_filename_v6.py
import os
import re
import shutil
from datetime import datetime


def standardize_filename_v6(folder_path, add_default_date=True):
    # Improved regular expressions for date formats
    # Match potential YYMMDD already in filename to prevent re-adding
    existing_date_pattern = re.compile(r'^\d{6}_')
    # Match YYYYMMDD format with valid month and day
    full_date_pattern = re.compile(r'(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])')
    # Match MMDD format with valid month and day
    short_date_pattern = re.compile(r'(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])')
    # Chinese date format is already specific
    cn_date_pattern = re.compile(r'(\d{1,2})月(\d{1,2})日')
    # Pattern to replace one or more spaces and undesired characters with a single underscore
    space_symbol_normalization_pattern = re.compile(r'[\s\-]+')
    # Pattern to collapse multiple underscores into a single one
    multiple_underscores = re.compile(r'_{2,}')

    # Read all files in the specified folder
    for filename in os.listdir(folder_path):
        original_path = os.path.join(folder_path, filename)

        # Skip directories
        if os.path.isdir(original_path):
            continue

        # Normalize spaces and symbols in filenames
        normalized_filename = space_symbol_normalization_pattern.sub('_', filename)
        normalized_filename = multiple_underscores.sub('_', normalized_filename)

        # Ensure no existing date
        if existing_date_pattern.match(normalized_filename):
            continue  # Skip renaming if date is already properly prefixed

        # Get the file's last modified time for date fallback
        last_modified_timestamp = os.path.getmtime(original_path)
        last_modified_date = datetime.fromtimestamp(last_modified_timestamp)

        # Try to find full date format first
        date_match = full_date_pattern.search(normalized_filename)
        if date_match:
            date_str = datetime.strptime(date_match.group(), '%Y%m%d').strftime('%y%m%d')
            # Remove date from filename
            normalized_filename = full_date_pattern.sub('', normalized_filename)
        else:
            # Try to find short date format
            date_match = short_date_pattern.search(normalized_filename)
            if date_match:
                # Combine year from last modified date with MMDD
                date_str = last_modified_date.strftime('%y') + date_match.group()
                # Remove date from filename
                normalized_filename = short_date_pattern.sub('', normalized_filename)
            else:
                # Try to find Chinese date format
                cn_date_match = cn_date_pattern.search(normalized_filename)
                if cn_date_match:
                    month_day_str = datetime.strptime(f'{cn_date_match.group(1)}{cn_date_match.group(2)}',
                                                      '%m%d').strftime('%m%d')
                    date_str = last_modified_date.strftime('%y') + month_day_str
                    # Remove date from filename
                    normalized_filename = cn_date_pattern.sub('', normalized_filename)
                elif add_default_date:
                    # If no date is found and default date is allowed
                    date_str = last_modified_date.strftime('%y%m%d')
                else:
                    # If no date is found and default date is not allowed, skip renaming
                    continue

        # Build new filename
        new_filename = f"{date_str}_{normalized_filename.strip('_')}"
        new_path = os.path.join(folder_path, new_filename)

        # Rename file
        shutil.move(original_path, new_path)


# Uncomment this line to use the function with parameters.
# standardize_filename_v6('/path/to/folder')


standardize_filename_v6(r"C:\Users\dell\Desktop", add_default_date=False)
