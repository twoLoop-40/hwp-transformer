import os
import sys

from hangul_util import HwpTool, insert_text, StringPosition, ImageRun, insert_one_image
from main_tool import FolderFileName, insert_images, two_cell_before, one_cell_before, count_content_list, \
    TextOperation, insert_math_form
from help_tool import get_home_directory_name, identity, half

from tkinter import Tk
from tkinter.filedialog import askopenfilename

root = Tk()
file_path = askopenfilename(title='txt 파일을 선택해주세요', initialdir=os.path.join(get_home_directory_name(), "Downloads"))
if not file_path or len(file_path) == 0:
    root.destroy()
    sys.exit()

path_handler = FolderFileName(file_path)

root.destroy()

text_operation = TextOperation(folder_path=path_handler.folder_path, file_name=path_handler.in_file_name)
text_operation.load_text()

hwp_tool = HwpTool(True)
for content_line in text_operation.content_lines:
    insert_text(hwp_tool.hwp, content_line)

# 맨 위로 위동
hwp_tool.goto_start()

count_double_dollar = text_operation.extract_content_info(count_content_list, pattern='$$')
count_double = count_double_dollar(half)
change_math_double_dollar = insert_math_form(hwp_tool, r'\$\$')
change_math_double_dollar(two_cell_before, count_double)

# 맨 위로 위동
hwp_tool.goto_start()

count_single_dollar = text_operation.extract_content_info(count_content_list, pattern='$')
count_single = count_single_dollar(lambda total: half(total - 2 * count_double_dollar(identity)))
change_math_single_dollar = insert_math_form(hwp_tool, r'\$')
change_math_single_dollar(one_cell_before, count_single)

# 그림 삽입
hwp_tool.goto_start()

image_pos = StringPosition(hwp=hwp_tool.hwp, pattern=r"<\!-- image -->")
image_run = ImageRun(image_path_list=path_handler.get_image_filenames(), image_handler=insert_one_image,
                     image_option={"sizeoption": 1, "Width": 60, "Height": 60})
insert_images(image_pos, image_run)

hwp_tool.hwp.SaveAs(path_handler.make_download_name())
hwp_tool.hwp.Clear(1)
hwp_tool.hwp.Quit()
