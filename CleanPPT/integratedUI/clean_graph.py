def adjust_text_boxes(presentation):
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text_content = shape.TextFrame.TextRange.Text.strip()  # 移除文本前后的空格和换行符
                if text_content:  # 如果文本内容非空
                    # 取消勾选“形状中的文字自动换行”
                    shape.TextFrame.WordWrap = 0  # 0 表示False

                    # 勾选“不自动调整”，这需要通过设置AutoSize属性来实现
                    shape.TextFrame.AutoSize = 0  # 0 表示ppAutoSizeNone，不自动调整

                    # 最后，勾选“根据文字调整形状大小”
                    shape.TextFrame.AutoSize = 1  # 1 表示ppAutoSizeShapeToFitText，根据文字调整形状大小
                else:
                    # 如果文本框是空的或仅包含空格/换行符，可以选择跳过或执行其他逻辑
                    pass

    print(f"Adjustments made to: {presentation.Name}")
