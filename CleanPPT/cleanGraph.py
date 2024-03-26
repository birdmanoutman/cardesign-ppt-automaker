import win32com.client


def adjust_text_boxes_via_com(pptx_path):
    # 启动PowerPoint
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = True  # 使PowerPoint可见

    # 打开演示文稿
    presentation = powerpoint.Presentations.Open(pptx_path)

    # 遍历每个幻灯片和每个形状
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

    # 保存并关闭演示文稿
    presentation.SaveAs("adjusted_" + pptx_path)
    presentation.Close()

    # 关闭PowerPoint
    powerpoint.Quit()


# 示例使用路径
pptx_path = r"C:\Users\dell\Desktop\演示文稿1.pptx"  # 请替换为实际的文件路径
adjust_text_boxes_via_com(pptx_path)
