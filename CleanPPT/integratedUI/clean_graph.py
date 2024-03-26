# clean_graph.py

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
                pass

    # 注意：这个脚本版本不关闭演示文稿和PowerPoint
    print(f"Adjustments made to: {presentation.Fullname}")


# 这个函数用于列出当前打开的所有PowerPoint演示文稿

