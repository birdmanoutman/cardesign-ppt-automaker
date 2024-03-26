def clean_master_in_open_presentation(presentation):
    # 获取幻灯片母版（设计）
    for design in presentation.Designs:
        slide_master = design.SlideMaster

        # 删除幻灯片母版中的所有占位符
        for shape in list(slide_master.Shapes):
            if shape.Type == 14:  # 14 表示占位符
                try:
                    shape.Delete()
                except Exception as e:
                    print("Error deleting shape:", e)

        # 收集已使用布局的名称
        used_layouts_names = {slide.CustomLayout.Name for slide in presentation.Slides}

        # 删除未使用的布局
        for layout in list(slide_master.CustomLayouts):
            if layout.Name not in used_layouts_names:
                try:
                    layout.Delete()
                except Exception as e:
                    print("Error deleting unused layout:", e)

        # 再次遍历所有布局，删除剩余占位符
        for layout in slide_master.CustomLayouts:
            for shape in list(layout.Shapes):
                if shape.Type == 14:
                    try:
                        shape.Delete()
                    except Exception as e:
                        print("Error deleting shape from layout:", e)
