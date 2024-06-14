from collections.abc import Callable
from typing import Optional, Any, TypeVar

from pydantic import BaseModel, ConfigDict, Field
from win32com import client as win32

T = TypeVar('T', bound=win32.CDispatch)

FindByPattern = Callable[[T, str], T]

Position = tuple[int, int, int]

Coordinates = tuple[int, int]

ChangePosition = Callable[[Position], Position]


class StringPosition(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    hwp: Any

    start_position: tuple[int, int, int] | None = None
    end_position: tuple[int, int, int] | None = None
    pattern: str
    next_pos: Optional['StringPosition'] = None

    def find_pos_pair(self, finder: FindByPattern, adjustment=lambda position: position) -> 'StringPosition':
        hwp = finder(self.hwp, self.pattern)
        current_pos = hwp.GetPos()
        self.start_position = adjustment(current_pos)
        hwp = finder(hwp, self.pattern)
        current_pos = hwp.GetPos()
        self.end_position = current_pos
        self.next_pos = StringPosition(hwp=self.hwp, pattern=self.pattern)
        return self.next_pos

    def select_and_next(self, finder: FindByPattern):
        finder(self.hwp, self.pattern)
        self.next_pos = StringPosition(hwp=self.hwp, pattern=self.pattern)
        return self.next_pos

    def delete_back(self):
        hwp = self.hwp
        hwp.HAction.Run('DeleteBack')


def find_by_exp(hwp: win32, pattern: str) -> win32.CDispatch:
    hwp.HAction.GetDefault("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    h_find_replace = hwp.HParameterSet.HFindReplace
    h_find_replace.FindString = pattern
    h_find_replace.Direction = hwp.FindDir("Forward")
    h_find_replace.FindRegExp = 1
    h_find_replace.IgnoreMessage = 1
    h_find_replace.FindType = 1
    hwp.HAction.Execute("RepeatFind", hwp.HParameterSet.HFindReplace.HSet)
    return hwp


def get_text_from_block(hwp: win32.CDispatch) -> list[str]:
    hwp.InitScan(Range=0xff)
    state = 2
    total_text = []
    while state >= 2:
        state, text = hwp.GetText()
        total_text.append(text)
    hwp.ReleaseScan()
    return total_text


def hancom_eqn(hwp: win32.CDispatch, hwp_eqn_text: str):
    heq_edit = hwp.HParameterSet.HEqEdit
    hwp.HAction.GetDefault("EquationCreate", heq_edit.HSet)
    heq_edit.EqFontName = "HancomEQN"
    heq_edit.string = hwp_eqn_text
    heq_edit.BaseUnit = hwp.PointToHwpUnit(10.0)
    hwp.HAction.Execute("EquationCreate", heq_edit.HSet)

    hwp.FindCtrl()
    h_shape = hwp.HParameterSet.HShapeObject
    hwp.HAction.GetDefault("EquationPropertyDialog", h_shape.HSet)
    h_shape.HSet.SetItem("ShapeType", 3)
    h_shape.Version = "Equation Version 60"
    h_shape.EqFontName = "HancomEQN"
    h_shape.HSet.SetItem("ApplyTo", 0)
    h_shape.HSet.SetItem("TreatAsChar", 1)
    hwp.HAction.Execute("EquationPropertyDialog", h_shape.HSet)


def insert_text(hwp: win32.CDispatch, hwp_text: str):
    act = hwp.CreateAction("InsertText")
    p_set = act.CreateSet()
    act.GetDefault(p_set)
    p_set.SetItem("Text", hwp_text)
    act.Execute(p_set)


def insert_one_image(hwp: win32.CDispatch, image_path: str, image_option: dict):
    ctrl = hwp.InsertPicture(image_path, **image_option)


class HwpTool(BaseModel):
    hwp: Any | None = None

    def __init__(self, is_visible: bool = True, **data):
        super().__init__(**data)
        hwp_object = "hwpframe.hwpobject"
        hwp_init = win32.gencache.EnsureDispatch(hwp_object)
        hwp_init.XHwpWindows.Item(0).Visible = is_visible
        self.hwp = hwp_init

    def goto_start(self):
        hwp = self.hwp
        hwp.Run("MoveDocBegin")
        hwp.Run("Cancel")

    def select_block(self, pair_pos: StringPosition, process=None, **kwargs):
        hwp = self.hwp

        hwp.SetPos(*pair_pos.start_position)
        hwp.Run("Select")
        hwp.SetPos(*pair_pos.end_position)

        if process is None:
            hwp.Run('Cancel')

        else:
            try:
                yield process(hwp, **kwargs)
                hwp.Run('Cancel')
            except Exception as e:
                print(f"Error in process: {e}")
            finally:
                hwp.Run('Cancel')


class ImageRun(BaseModel):
    model_config = ConfigDict(arbitrary_types_allowed=True)

    image_path_list: list[str]
    image_handler: Callable
    image_option: dict = Field(default={"sizeoption": 0})
