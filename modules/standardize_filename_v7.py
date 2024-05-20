import os
import re
import shutil
from datetime import datetime
import tempfile

def standardize_filename(folder_path, add_default_date=True, rename_folders=False):
    # Improved regular expressions for date formats
    existing_date_pattern = re.compile(r'^\d{6}_')
    full_date_pattern = re.compile(r'(\d{4})(0[1-9]|1[0-2])(0[1-9]|[12][0-9]|3[01])')
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

        # Normalize spaces and symbols in filenames
        normalized_name = space_symbol_normalization_pattern.sub('_', name)
        normalized_name = multiple_underscores.sub('_', normalized_name)

        # Remove trailing spaces, underscores, and hyphens
        normalized_name = trailing_symbols_pattern.sub('', normalized_name)

        # Ensure no existing date
        if existing_date_pattern.match(normalized_name):
            continue  # Skip renaming if date is already properly prefixed

        last_modified_timestamp = os.path.getmtime(original_path)
        last_modified_date = datetime.fromtimestamp(last_modified_timestamp)

        # Date extraction and formatting
        date_match = full_date_pattern.search(normalized_name)
        if date_match:
            date_str = datetime.strptime(date_match.group(), '%Y%m%d').strftime('%y%m%d')
            normalized_name = full_date_pattern.sub('', normalized_name)
        else:
            date_match = short_date_pattern.search(normalized_name)
            if date_match:
                date_str = last_modified_date.strftime('%y') + date_match.group()
                normalized_name = short_date_pattern.sub('', normalized_name)
            else:
                cn_date_match = cn_date_pattern.search(normalized_name)
                if cn_date_match:
                    month_day_str = datetime.strptime(f'{cn_date_match.group(1)}{cn_date_match.group(2)}', '%m%d').strftime('%m%d')
                    date_str = last_modified_date.strftime('%y') + month_day_str
                    normalized_name = cn_date_pattern.sub('', normalized_name)
                elif add_default_date:
                    date_str = last_modified_date.strftime('%y%m%d')
                else:
                    continue  # Skip renaming if no date found and default date not allowed

        # Build new filename and rename
        new_filename = f"{date_str}_{normalized_name.strip('_')}"
        print(new_filename)
        new_filename = trailing_symbols_pattern.sub('', new_filename)  # Remove trailing symbols again
        new_path = os.path.join(folder_path, f"{new_filename}{ext}")
        if os.path.isdir(original_path):
            os.rename(original_path, new_path)  # Apply renaming for directories
        else:
            shutil.move(original_path, new_path)  # Apply renaming for files

def generate_standardized_name(file_path, add_date, rename_folders):
    folder, filename = os.path.split(file_path)
    temp_folder = tempfile.mkdtemp()
    try:
        if os.path.isdir(file_path):
            # 如果是文件夹，则不复制内容，只是标准化文件夹名称
            temp_path = os.path.join(temp_folder, filename)
            os.mkdir(temp_path)
            standardize_filename(temp_folder, add_date, rename_folders)
            std_name = os.listdir(temp_folder)[0]
        else:
            temp_path = os.path.join(temp_folder, filename)
            shutil.copy(file_path, temp_path)
            standardize_filename(temp_folder, add_date, rename_folders)
            std_name = os.listdir(temp_folder)[0]
    finally:
        def remove_readonly(func, path, excinfo):
            os.chmod(path, 0o777)
            func(path)
        try:
            shutil.rmtree(temp_folder, onerror=remove_readonly)
        except Exception as e:
            print(f"无法删除临时文件夹: {e}")
    return std_name
