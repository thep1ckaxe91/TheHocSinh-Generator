import pygame
import pygame.freetype

# 1. Configuration
pygame.init()
pygame.freetype.init()

# Setup screen and colors
SCREEN_WIDTH = 1920
SCREEN_HEIGHT = 400
screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
pygame.display.set_caption("Vietnamese Baseline Fix (pygame.freetype)")

WHITE = (255, 255, 255)
BLACK = (0, 0, 0)
RED = (255, 0, 0)
BLUE = (0, 0, 255)
GREEN = (0, 255, 0)
GRAY = (150, 150, 150)

# The Vietnamese text that causes issues (due to stacked accents)
VIETNAMESE_TEXT = "Chữ Việt: gy nguyễn A Ă Â E Ê I O Ô Ơ U Ư Y"
FONT_SIZE = 48  # Large size to make the issue clear
TEXT_X = 50

# NOTE: Since Times New Roman is a system font, we use SysFont. 
# For a .ttf file, use pygame.freetype.Font("path/to/font.ttf", FONT_SIZE)
try:
    # Attempt to load a Times New Roman font
    font = pygame.freetype.Font("./assets/TNR.ttf", FONT_SIZE)
except:
    # Fallback if Times New Roman is not found
    font = pygame.freetype.SysFont("dejavusans", FONT_SIZE)

# Set anti-aliasing for better rendering
font.antialiased = True

# --- Baseline Calculation Functions ---

def calculate_corrected_ascender(font_obj, text, size):
    """
    Calculates the actual highest point (ascender) needed for a text string
    by checking the metrics of all individual glyphs.
    """
    # get_metrics returns a list of tuples: 
    # [(min_x, max_x, min_y, max_y, advance), ...]
    metrics = font_obj.get_metrics(text, size=size)
    
    # max_y_offset is the fourth element (index 3). We find the maximum
    # max_y_offset across all characters in the string.
    max_y = 0
    for glyph_metrics in metrics:
        # glyph_metrics can be None if the character isn't supported
        if glyph_metrics is not None:
            max_y = max(max_y, glyph_metrics[3])
            
    return max_y

# --- Main Loop ---

running = True
while running:
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False

    screen.fill(WHITE)
    
    # ----------------------------------------------------
    ## Text 1: The Standard (Incorrect) Approach
    # ----------------------------------------------------
    
    # 1. Get the ascender from the font object (This is often too low for Vietnamese)
    standard_ascender = font.get_sized_ascender(FONT_SIZE)
    
    # 2. Define the desired baseline Y position
    BASE_Y_STANDARD = 150
    
    # 3. Calculate the blit position (top-left corner of the surface)
    # Blit_Y = Baseline_Y - Ascender_Height
    blit_y_standard = BASE_Y_STANDARD - standard_ascender
    
    # 4. Render and Blit the text
    text_surf_standard, text_rect_standard = font.render(
        VIETNAMESE_TEXT, BLACK, size=FONT_SIZE
    )
    screen.blit(text_surf_standard, (TEXT_X, blit_y_standard))

    # Draw the Baseline in Red to show where the text *should* sit
    pygame.draw.line(screen, RED, (0, BASE_Y_STANDARD), (SCREEN_WIDTH, BASE_Y_STANDARD), 1)
    
    # Draw a helpful label
    font.render_to(screen, (TEXT_X, 10), "Standard Ascender (Clipped Accents):", BLUE, size=24)
    
    # ----------------------------------------------------
    
    # ----------------------------------------------------
    ## Text 2: The Corrected Approach (Using get_metrics)
    # ----------------------------------------------------

    # 1. Calculate the *true* ascender needed for the text string
    corrected_ascender = calculate_corrected_ascender(font, VIETNAMESE_TEXT, FONT_SIZE)

    # 2. Define the desired baseline Y position (lower on the screen)
    BASE_Y_CORRECTED = 350

    # 3. Calculate the corrected blit position
    # Blit_Y = Baseline_Y - Corrected_Ascender_Height
    blit_y_corrected = BASE_Y_CORRECTED - corrected_ascender

    # 4. Render and Blit the text
    text_surf_corrected, _ = font.render(
        VIETNAMESE_TEXT, BLACK, size=FONT_SIZE
    )
    screen.blit(text_surf_corrected, (TEXT_X, blit_y_corrected))
    
    # Draw the Baseline in Green to show where the text is now sitting
    pygame.draw.line(screen, GREEN, (0, BASE_Y_CORRECTED), (SCREEN_WIDTH, BASE_Y_CORRECTED), 1)
    
    # Draw a helpful label
    font.render_to(screen, (TEXT_X, 210), "Corrected Ascender (Perfect Alignment):", BLUE, size=24)
    
    # ----------------------------------------------------

    pygame.display.flip()
    
pygame.quit()