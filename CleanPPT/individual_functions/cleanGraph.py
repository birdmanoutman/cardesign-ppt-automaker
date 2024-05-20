# cleanGraph.py
import win32com.client
import os
import sys

def list_open_presentations(powerpoint):
    print("Currently open presentations:")
    for i, presentation in enumerate(powerpoint.Presentations, start=1):
        file_path = presentation.FullName
        file_name = os.path.basename(file_path)
        modified_time = os.path.getmtime(file_path)
        print(f"{i}. {file_name} (Last Modified: {modified_time})")
    selection = input("Enter the number of the presentation you want to adjust: ")
    try:
        selected_index = int(selection) - 1
        return powerpoint.Presentations[selected_index]
    except (ValueError, IndexError):
        print("Invalid selection.")
        sys.exit(1)

def adjust_text_boxes(presentation):
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                # 取消勾选“形状中的文字自动换行”
                shape.TextFrame.WordWrap = 0  # 0 表示False

                # 勾选“不自动调整”，这需要通过设置AutoSize属性来实现
                shape.TextFrame.AutoSize = 0  # 0 表示ppAutoSizeNone，不自动调整

                # 在这里更新文本框，如果有需要的话

                # 最后，勾选“根据文字调整形状大小”
                shape.TextFrame.AutoSize = 1  # 1 表示ppAutoSizeShapeToFitText，根据文字调整形状大小

                # 注意: 以上代码可能需要根据实际情况进行调整

                print("Adjustment complete.")

if __name__ == "__main__":
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    if not powerpoint.Presentations.Count:
        print("No presentations are currently open.")
        sys.exit(1)

    presentation = list_open_presentations(powerpoint)
    adjust_text_boxes(presentation)
