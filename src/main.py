import win32com.client
import os

def ppt_to_pdf_via_com(ppt_path, pdf_path, print_options=None):
    powerpoint = None
    presentation = None
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

        ppt_path = os.path.abspath(ppt_path)
        pdf_path = os.path.abspath(pdf_path)

        presentation = powerpoint.Presentations.Open(ppt_path, ReadOnly=False, Untitled=False, WithWindow=False)

        if print_options:
            print_options_obj = presentation.PrintOptions

            # 设置打印范围
            if "PrintRangeType" in print_options:
                print_options_obj.RangeType = print_options["PrintRangeType"]
                if print_options["PrintRangeType"] == 4:
                    print_ranges = print_options_obj.Ranges
                    print_ranges.ClearAll()
                    print_ranges.Add(print_options["RangeStart"], print_options["RangeEnd"])

            # 设置其他打印选项
            for key, value in print_options.items():
                if key not in ["PrintRangeType", "RangeStart", "RangeEnd"]:
                    try:
                        setattr(print_options_obj, key, value)
                    except Exception as e:
                        print(f"设置打印选项 {key} 失败: {e}")

        # 另存为 PDF
        presentation.SaveAs(pdf_path, 32)

        print(f"成功将 {ppt_path} 转换为 {pdf_path}")

    except Exception as e:
        print(f"转换过程中发生错误: {e}")

    finally:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()

def convert_ppt_files_in_directory(directory, print_options=None):
    pdf_directory = os.path.join(directory, 'pdf')
    if not os.path.exists(pdf_directory):
        os.makedirs(pdf_directory)
    
    for filename in os.listdir(directory):
        if filename.endswith('.ppt') or filename.endswith('.pptx'):
            ppt_path = os.path.join(directory, filename)
            pdf_path = os.path.join(pdf_directory, os.path.splitext(filename)[0] + '.pdf')
            ppt_to_pdf_via_com(ppt_path, pdf_path, print_options)

# 使用示例
directory = r"C:\Curriculum\2024年秋\人工智能导论-复习资料\选择题判断题-人工智能导论-14课件"  # 替换为你的目录路径

print_settings = {
    "PrintRangeType": 1,         # 打印所有幻灯片 (ppPrintAll)
    "OutputType": 14,            # 每页四张幻灯片的讲义 (ppPrintOutputFourSlideHandouts)
    "NumberOfCopies": 1,         # 打印一份
    "Collate": True,             # 逐份打印
    "HandoutOrder": 2,           # 水平排列 (ppPrintHandoutHorizontalFirst)
    "PrintColorType": 1,         # 彩色打印 (ppPrintColor)
    "FrameSlides": False,        # 不加框
    "FitToPage": True,           # 适应页面大小
    "PrintHiddenSlides": False   # 不打印隐藏幻灯片
}

convert_ppt_files_in_directory(directory, print_settings)
