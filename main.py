import pygame
import numpy as np
import random
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage


class Grid:
    def __init__(self, screen_width=800, screen_height=600, cell_size=15):
        self.screen_width = screen_width
        self.screen_height = screen_height
        self.cell_size = cell_size

        # Calculate grid dimensions
        self.width = screen_width // cell_size
        self.height = screen_height // cell_size

        # Create grid data structure
        self.grid = np.zeros((self.height, self.width), dtype=int)

        # Set outer edges as black walls
        self._create_outer_walls()

        # Add random walls
        self._create_random_walls()

        # Colors
        self.COLORS = {
            0: (255, 255, 255),  # Empty cell
            1: (0, 0, 0),  # Wall
            2: (0, 255, 0),  # Start point
            3: (255, 0, 0),  # End point
            4: (100, 100, 255)  # Path
        }

    def _create_outer_walls(self):
        # Top and bottom walls
        self.grid[0, :] = 1
        self.grid[-1, :] = 1

        # Left and right walls
        self.grid[:, 0] = 1
        self.grid[:, -1] = 1

    def _create_random_walls(self):
        directions = ['top', 'left', 'bottom', 'right']
        max_attempts = 10  # Maximum attempts to adjust a wall's position and size

        for direction in directions:
            placed = False  # Track if the wall is successfully placed

            for _ in range(max_attempts):
                # Randomize wall length between 10 and 30
                wall_length = random.randint(10, 30)

                if direction == 'top':
                    start_x = random.randint(2, self.width - 3)
                    valid = True
                    for y in range(1, min(wall_length + 1, self.height - 2)):
                        if self.grid[y, start_x] != 0:  # Collision detected
                            valid = False
                            break

                    if valid:
                        for y in range(1, min(wall_length + 1, self.height - 2)):
                            self.grid[y, start_x] = 1
                        placed = True
                        break  # Exit retry loop for this direction

                elif direction == 'left':
                    start_y = random.randint(2, self.height - 3)
                    valid = True
                    for x in range(1, min(wall_length + 1, self.width - 2)):
                        if self.grid[start_y, x] != 0:  # Collision detected
                            valid = False
                            break

                    if valid:
                        for x in range(1, min(wall_length + 1, self.width - 2)):
                            self.grid[start_y, x] = 1
                        placed = True
                        break

                elif direction == 'bottom':
                    start_x = random.randint(2, self.width - 3)
                    valid = True
                    for y in range(self.height - 2, max(self.height - 2 - wall_length, 1), -1):
                        if self.grid[y, start_x] != 0:  # Collision detected
                            valid = False
                            break

                    if valid:
                        for y in range(self.height - 2, max(self.height - 2 - wall_length, 1), -1):
                            self.grid[y, start_x] = 1
                        placed = True
                        break

                elif direction == 'right':
                    start_y = random.randint(2, self.height - 3)
                    valid = True
                    for x in range(self.width - 2, max(self.width - 2 - wall_length, 1), -1):
                        if self.grid[start_y, x] != 0:  # Collision detected
                            valid = False
                            break

                    if valid:
                        for x in range(self.width - 2, max(self.width - 2 - wall_length, 1), -1):
                            self.grid[start_y, x] = 1
                        placed = True
                        break

            # If all attempts fail, log a message (optional, for debugging)
            if not placed:
                print(f"Failed to place a wall for direction: {direction}")

    def draw(self, screen):
        for y in range(self.height):
            for x in range(self.width):
                rect = pygame.Rect(
                    x * self.cell_size,
                    y * self.cell_size,
                    self.cell_size,
                    self.cell_size
                )
                pygame.draw.rect(screen, self.COLORS[self.grid[y][x]], rect)
                pygame.draw.rect(screen, (200, 200, 200), rect, 1)


def save_playground_with_data(screen, playground_name, best_algorithm):
    # Paths for saving data
    base_path = r"C:\Users\tomer\Desktop\Projects\pathfinding"
    excel_file = os.path.join(base_path, "playground_metadata.xlsx")
    image_folder = os.path.join(base_path, "playground_images")

    # Ensure the image folder exists
    if not os.path.exists(image_folder):
        os.makedirs(image_folder)

    # Save the playground image
    image_path = os.path.join(image_folder, f"{playground_name}.png")
    pygame.image.save(screen, image_path)

    # Save metadata in Excel
    if os.path.exists(excel_file):
        wb = load_workbook(excel_file)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(["Playground Name", "Date", "Time", "Best Algorithm", "Image"])

    now = datetime.now()
    date = now.strftime("%Y-%m-%d")
    time = now.strftime("%H:%M:%S")

    # Find the next row to write, ensuring 7 empty rows
    next_row = ws.max_row + 8  # Add 8 to leave 7 empty rows between entries

    # Write metadata in the correct row
    ws.cell(row=next_row, column=1, value=playground_name)
    ws.cell(row=next_row, column=2, value=date)
    ws.cell(row=next_row, column=3, value=time)
    ws.cell(row=next_row, column=4, value=best_algorithm)
    ws.cell(row=next_row, column=5, value="See Image")

    # Place the image at the correct row and resize its display
    img = ExcelImage(image_path)
    img.anchor = f"E{next_row}"
    img.width = 120  # Set desired width in pixels
    img.height = 120  # Set desired height in pixels
    ws.add_image(img)

    wb.save(excel_file)
    print(f"Playground and data saved in '{excel_file}'")


def main():
    pygame.init()

    # Screen dimensions
    SCREEN_WIDTH = 800
    SCREEN_HEIGHT = 600
    screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
    pygame.display.set_caption("Pathfinding Playground")

    # Create grid
    grid = Grid(cell_size=15)

    # Draw the grid once
    screen.fill((255, 255, 255))
    grid.draw(screen)

    # Save the playground data
    save_playground_with_data(screen, "playground_1", "Algorithm Placeholder")

    # Main game loop
    running = True
    while running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False

        screen.fill((255, 255, 255))
        grid.draw(screen)
        pygame.display.flip()

    pygame.quit()


if __name__ == "__main__":
    main()
