import pygame, sys, os, datetime
from openpyxl import load_workbook
import xlrd
import pygame.freetype as ft
from pygame import Vector2


WIDTH, HEIGHT = 1167, 677
window = pygame.display.set_mode((WIDTH, HEIGHT))
clock = pygame.time.Clock()
PATH = os.path.dirname(os.path.abspath(__file__)).replace("\\", "/")
pygame.init()
ft.init()
DEBUGGING = False

BG_COLOR = (255, 255, 190)  # also for school name color
INFO_COLOR = (23, 121, 254)


"""////////////////////////////////DATA TO FIX WHEN WORK WITH DIFFERENT FILE/////////////////////////////////////"""
start_row = 4
name_col = 1
split_name = True
gender_col = 4
dob_col = 4
class_name = "1A"
file_name = "data"
# Load the workbook
try:
    workbook = load_workbook(filename=PATH + f"/excel_files/{file_name}.xlsx")
    sheet = workbook.active
    data_list = [[cell.strftime('%d/%m/%Y') if isinstance(cell, datetime.datetime) else cell for cell in row] for row in sheet.iter_rows(values_only=True)]
except:
    workbook = xlrd.open_workbook(PATH + f"/excel_files/{file_name}.xls")
    sheet = workbook.sheet_by_index(0)
    data_list = [[sheet.cell_value(r, c).strftime('%d/%m/%Y') if isinstance(sheet.cell_value(r, c), datetime.datetime) else sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
"""////////////////////////////////DATA TO FIX WHEN WORK WITH DIFFERENT FILE/////////////////////////////////////"""
"""//////////////TESTING CONST POSITION//////////////////"""
SCHOOL_POS = Vector2(50,50)
NAME_POS = Vector2(631,295)
DOB_POS =  Vector2(640,338)
NK_POS =   Vector2(640,385)
GENDER_POS = ()
FONT_SIZE = 47
"""//////////////TESTING CONST POSITION//////////////////"""
"""//////////////CURRENT CAPTURING INFO//////////////////"""
cur_row_id = start_row
student_index = 1
cur_name = ""
cur_dob = ""
cur_gender = ""
# INFO_FONT = pygame.font.Font(PATH + "/assets/TNR.ttf", FONT_SIZE)
INFO_FONT = ft.Font(PATH + "/assets/TNR.ttf", FONT_SIZE)
# SCHOOL_FONT = pygame.font.Font(PATH + "/assets/TNR.ttf", FONT_SIZE - 10)
default_img = pygame.image.load(PATH + "/assets/BeTongDefault.png").convert()
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
    global cur_name_surf, cur_dob_surf, cur_row_id, INFO_COLOR,FONT_SIZE
    try:
        cur_name = data_list[cur_row_id][name_col] + split_name*(' '*(data_list[cur_row_id][name_col][-1]!=' ')+data_list[cur_row_id][name_col+1])
    except:
        exit(0)
    cur_dob = data_list[cur_row_id][dob_col]
    cur_name_surf = INFO_FONT.render(cur_name, INFO_COLOR, style=INFO_STYLE, size=FONT_SIZE)[0]
    cur_dob_surf = INFO_FONT.render(cur_dob, INFO_COLOR,style=INFO_STYLE, size=FONT_SIZE)[0]

def draw():  # to draw the data on the screen
    window.blit(default_img, (0, 0))
    window.blit(cur_name_surf,NAME_POS - Vector2(0,cur_name_surf.get_height()))
    window.blit(cur_dob_surf,DOB_POS - Vector2(0,cur_dob_surf.get_height()))
    window.blit(cur_nk_surf,NK_POS - Vector2(0,cur_nk_surf.get_height()))
    # window.blit(cur_school_surf,SCHOOL_POS)

def capture():  # to save the image
    global student_index,cur_row_id
    try:
        os.mkdir(PATH + f"/result/{class_name}")
    except:
        pass
    pygame.image.save(window,PATH + f"/result/{class_name}/image{student_index}.png")
    if DEBUGGING:
        # pygame.time.wait(10000)
        exit(0)
    else:
        student_index+=1
        cur_row_id+=1


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
