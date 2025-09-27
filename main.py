import pygame, sys, os, datetime
from openpyxl import load_workbook
import xlrd
import pygame.freetype as ft
from pygame import Vector2
from typing import Any

WIDTH, HEIGHT = 1167, 677
window = pygame.display.set_mode((WIDTH, HEIGHT))
clock = pygame.time.Clock()
PATH = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/")
pygame.init()
ft.init()
DEBUGGING = False

BG_COLOR = (255, 255, 190)  # also for school name color
INFO_COLOR = (23, 121, 254)


"""////////////////////////////////DATA TO FIX WHEN WORK WITH (DIFFERENT FOR EACH EXCEL FILE)/////////////////////////////////////"""
# 1-base indexed
start_row = 4
name_col = 1
split_name = True
gender_col = 4
dob_col = 4
class_name = "1A"
file_name = "data"

"""////////////////////////////////DATA TO FIX WHEN WORK WITH (DIFFERENT FOR EACH EXCEL FILE)/////////////////////////////////////"""

"""//////////////CONST POSITION FOR GENERATION (DIFFERENT FOR EACH DEFAULT IMAGE)//////////////////"""
SCHOOL_POS = Vector2(50, 50)
NAME_POS = Vector2(631, 295)
DOB_POS = Vector2(640, 338)
NK_POS = Vector2(640, 385)
GENDER_POS = ()
FONT_SIZE = 47
"""//////////////CONST POSITION FOR GENERATION (DIFFERENT FOR EACH DEFAULT IMAGE)//////////////////"""

"""//////////////CURRENT CAPTURING INFO//////////////////"""
cur_row_id = start_row
student_index = 1
cur_name = ""
cur_dob = ""
cur_gender = ""
"""//////////////CURRENT CAPTURING INFO//////////////////"""

# Load the workbook
try:  # xlsx first
    workbook = load_workbook(
        filename=os.path.join(PATH, "excel_files", f"{file_name}.xlsx"), read_only=True
    )
    if len(workbook.sheetnames) > 1:
        print(
            f"WARNING: There's more than 1 sheet in {file_name}, manual intervention if needed"
        )
    sheet = workbook.active
    if sheet == None:
        raise Exception("No active sheet found")
    data_list = [
        [
            cell.strftime("%d/%m/%Y") if isinstance(cell, datetime.datetime) else str(cell)
            for cell in row
        ]
        for row in sheet.iter_rows(values_only=True)
    ]
except Exception as e:  # try xls if xlsx failed

    print(f"Failed to load .xlsx file due to {e}, trying .xls file...")

    workbook = xlrd.open_workbook(os.path.join(PATH, "excel_files", f"{file_name}.xls"))
    if len(workbook.sheet_names()) > 1:
        print(
            f"WARNING: There's more than 1 sheet in {file_name}, manual intervention if needed"
        )
    if len(workbook.sheet_names()) == 0:
        raise Exception("No sheet found")

    sheet = workbook.sheet_by_index(0)
    data_list = [
        [(sheet.cell_value(r, c)) for c in range(sheet.ncols)]
        for r in range(sheet.nrows)
    ]


"""//////////////CONST////////////////"""
INFO_FONT = ft.Font(os.path.join(PATH, "assets", "TNR.ttf"), FONT_SIZE)
default_img = pygame.image.load(os.path.join(PATH, "assets", "default.png")).convert()
INFO_STYLE = ft.STYLE_STRONG

if DEBUGGING:
    cur_nk = "2023 - 2028"
else:
    cur_nk = input("nhap nien khoa tu x - y:\n")
cur_nk_surf = INFO_FONT.render(cur_nk, INFO_COLOR, style=INFO_STYLE, size=FONT_SIZE)[0]
# cur_school_surf = SCHOOL_FONT.render(school_name, True, BG_COLOR)
# print(default_img.get_size())


def update():  # to load the current student data
    pygame.display.set_caption(str(int(clock.get_fps())))
    global cur_name_surf, cur_dob_surf
    try:
        cur_name = data_list[cur_row_id][name_col] + split_name * (
            " " * (data_list[cur_row_id][name_col][-1] != " ")
            + data_list[cur_row_id][name_col + 1]
        )
    except:
        exit(0)
    cur_dob = data_list[cur_row_id][dob_col]
    cur_name_surf = INFO_FONT.render(
        cur_name, INFO_COLOR, style=INFO_STYLE, size=FONT_SIZE
    )[0]
    cur_dob_surf = INFO_FONT.render(
        cur_dob, INFO_COLOR, style=INFO_STYLE, size=FONT_SIZE
    )[0]


def draw():  # to draw the data on the screen
    window.blit(default_img, (0, 0))
    window.blit(cur_name_surf, NAME_POS - Vector2(0, cur_name_surf.get_height()))
    window.blit(cur_dob_surf, DOB_POS - Vector2(0, cur_dob_surf.get_height()))
    window.blit(cur_nk_surf, NK_POS - Vector2(0, cur_nk_surf.get_height()))
    # window.blit(cur_school_surf,SCHOOL_POS)


def capture():  # to save the image
    global student_index, cur_row_id
    try:
        os.mkdir(os.path.join(PATH, "result", class_name))
    except:
        pass
    pygame.image.save(
        window, os.path.join(PATH, "result", class_name, f"image{student_index}.png")
    )
    if DEBUGGING:
        # pygame.time.wait(10000)
        exit(0)
    else:
        student_index += 1
        cur_row_id += 1


if __name__ == "__main__":
    while True:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
        update()
        draw()
        capture()
        pygame.display.flip()
        clock.tick(60)
