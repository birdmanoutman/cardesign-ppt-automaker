def clean_master_in_open_presentation(presentation):
    # 首先，收集文档中所有幻灯片实际使用的设计名称和布局名称
    used_designs = set(slide.Design.Name for slide in presentation.Slides)
    used_layouts = set(slide.CustomLayout.Name for slide in presentation.Slides)

    # 删除未使用的设计
    for design in list(presentation.Designs):
        design_name = design.Name  # 保存设计的名称
        if design_name not in used_designs:
            try:
                design.Delete()
                print(f"Deleted unused design: {design_name}")
            except Exception as e:
                print(f"Error deleting unused design {design_name}: {e}")

    # 对于每个剩余的设计，检查并删除未使用的布局
    for design in presentation.Designs:
        slide_master = design.SlideMaster
        for layout in list(slide_master.CustomLayouts):
            layout_name = layout.Name  # 保存布局的名称
            if layout_name not in used_layouts:
                try:
                    layout.Delete()
                    print(f"Deleted unused layout: {layout_name}")
                except Exception as e:
                    print(f"Error deleting unused layout {layout_name}: {e}")

        # 删除设计中所有布局的所有占位符
        for layout in slide_master.CustomLayouts:
            for shape in list(layout.Shapes):
                if shape.Type == 14:  # 14 表示占位符
                    try:
                        shape.Delete()
                        print(f"Deleted placeholder from layout: {layout_name}")
                    except Exception as e:
                        print(f"Error deleting shape from layout {layout_name}: {e}")

    # 遍历演示文稿中的所有设计（母版）
    for design in presentation.Designs:
        slide_master = design.SlideMaster

        # 尝试删除母版中的所有占位符
        for shape in list(slide_master.Shapes):
            if shape.Type == 14:  # 14 表示占位符
                try:
                    shape.Delete()
                    print("Deleted placeholder from master")
                except Exception as e:
                    print(f"Error deleting placeholder from master: {e}")

        # 遍历该母版下的所有布局，并尝试删除其中的占位符
        for layout in slide_master.CustomLayouts:
            for shape in list(layout.Shapes):
                if shape.Type == 14:  # 14 表示占位符
                    try:
                        shape.Delete()
                        print("Deleted placeholder from layout")
                    except Exception as e:
                        print(f"Error deleting placeholder from layout: {e}")
