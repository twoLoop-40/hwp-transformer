import os
import re
from typing import Any, Callable, TypeVar

from pydantic import BaseModel, Field

from hangul_util import StringPosition, find_by_exp, ImageRun, HwpTool, ChangePosition, get_text_from_block, hancom_eqn
from help_tool import get_home_directory_name


T = TypeVar("T")
ExtractContentInfo = Callable[[list[str], Any], T]


def content_lines_to_math_string(content_lines: list[str], pattern: str) -> str:
    math_lines = filter(lambda content: content != "",
                        map(lambda content: content.replace('\r\n', ''),
                            map(lambda content: re.sub(pattern, "", content),
                                content_lines)))
    return " #".join(math_lines)


PathTuple = tuple[str, str, str]


def split_folder_file_path(filepath: str) -> PathTuple:
    path_list = filepath.split('/')
    file_name = path_list.pop(-1)
    current_folder = path_list[-1]
    folder_path = "\\".join(path_list)

    return file_name, folder_path, current_folder


class FolderFileName(BaseModel):
    folder_path: str | None = None
    in_file_name: str | None = None
    out_file_name: str | None = None
    home_directory: str = Field(default_factory=get_home_directory_name)

    def __init__(self, src_path: str, **data):
        super().__init__(**data)
        file_name, folder_path, current_folder = split_folder_file_path(src_path)
        self.folder_path = folder_path
        self.in_file_name = file_name
        self.out_file_name = current_folder

    def make_download_name(self) -> str:
        download_path = os.path.join(self.home_directory, 'Downloads', self.out_file_name + '번 유사문항.hwp')
        return download_path

    def get_image_filenames(self):
        filenames = os.listdir(self.folder_path)
        filenames.remove(self.in_file_name)
        image_filenames = [os.path.join(self.folder_path, filename) for filename in filenames]
        return image_filenames


def insert_images(current_image_pos: StringPosition, image_runner: ImageRun):
    hwp = current_image_pos.hwp
    for image_path in image_runner.image_path_list:
        next_pos = current_image_pos.select_and_next(finder=find_by_exp)
        current_image_pos.delete_back()
        image_runner.image_handler(hwp, image_path, image_runner.image_option)

        current_image_pos = next_pos


def two_cell_before(position: tuple[int, int, int]) -> tuple[int, int, int]:
    return position[0], position[1], position[2] - 2


def one_cell_before(position: tuple[int, int, int]) -> tuple[int, int, int]:
    return position[0], position[1], position[2] - 1


def count_content_list(content_list: list[str], pattern: str) -> Callable:
    count_pattern = sum(map(lambda content: content.count(pattern), content_list))

    def run_count_pattern(runner: Callable[[int], Any]) -> Any:
        return runner(count_pattern)

    return run_count_pattern


class TextOperation(BaseModel):
    folder_path: str | None = None
    file_name: str | None = None
    content_lines: list[str] | None = None

    def load_text(self) -> 'TextOperation':

        file_path = os.path.join(self.folder_path, self.file_name)

        # 파일 열기
        try:
            with open(file_path, 'r', encoding='UTF-8') as f:
                lines = filter(None, [None if line.strip() == "" else line.strip() for line in f.readlines()])
                self.content_lines = [line + "\r\n" for line in lines]
                return self
        except FileNotFoundError as e:
            print(f'{file_path} not found: {e}')
            return self

    def content_map(self, transform: Callable[[str], str]) -> 'TextOperation':
        new_content_lines = list(map(lambda line: transform(line), self.content_lines))
        return TextOperation(folder_path=self.folder_path, file_name=self.file_name, content=new_content_lines)

    def extract_content_info(self, extract: ExtractContentInfo, **kwargs) -> T:
        return extract(self.content_lines, **kwargs)


def insert_math_form(hwp_handler: HwpTool, pattern: str):
    pos = StringPosition(hwp=hwp_handler.hwp, pattern=pattern)

    def string_to_math(adjustment: ChangePosition, repeat_count):
        nonlocal pos
        for _ in range(repeat_count):
            next_pos = pos.find_pos_pair(finder=find_by_exp, adjustment=adjustment)

            gen = hwp_handler.select_block(pos, get_text_from_block)

            try:
                while True:
                    text = next(gen)
                    if text[0] == ' ':
                        break
                    math_string = content_lines_to_math_string(text, pattern)
                    pos.delete_back()
                    hancom_eqn(hwp_handler.hwp, math_string)
            except StopIteration:
                pass

            pos = next_pos

        print("iteration end")
        return

    return string_to_math
