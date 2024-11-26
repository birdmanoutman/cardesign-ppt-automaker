import os
import sqlite3


def check_file_paths(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT img_path FROM image_ppt_mapping")
    rows = cursor.fetchall()

    missing_files = []
    for row in rows:
        img_path = row[0]
        if not os.path.exists(img_path):
            missing_files.append(img_path)

    conn.close()
    return missing_files


### 步骤 2：从数据库中删除不存在的文件路径记录

def delete_missing_files_from_db(db_path, missing_files):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    for img_path in missing_files:
        cursor.execute("DELETE FROM image_ppt_mapping WHERE img_path = ?", (img_path,))
        print(f"已从数据库中删除文件路径: {img_path}")
    conn.commit()
    conn.close()


if __name__ == "__main__":
    db_path = "image_gallery.db"  # 替换为你的数据库路径
    missing_files = check_file_paths(db_path)

    if missing_files:
        print("以下文件路径不存在:")
        for path in missing_files:
            print(path)

        # 删除数据库中不存在的文件路径记录
        delete_missing_files_from_db(db_path, missing_files)
        print("已删除所有不存在的文件路径记录。")
    else:
        print("所有文件路径都有效。")
