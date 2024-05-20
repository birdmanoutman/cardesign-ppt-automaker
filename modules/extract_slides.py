import win32com.client

def extract_text_from_presentation(presentation):
    text_content = []

    # 遍历演示文稿中的每个幻灯片
    for slide in presentation.Slides:
        # 遍历幻灯片中的每个形状
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                # 如果形状有文本框，提取文本
                if shape.TextFrame.HasText:
                    text_frame = shape.TextFrame
                    text_range = text_frame.TextRange
                    text = text_range.Text
                    text_content.append(text)
    return text_content
