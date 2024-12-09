import pygame
import numpy as np


class Grid:
    def __init__(self, width, height, cell_size=40):
        self.width = width
        self.height = height
        self.cell_size = cell_size

        # Create grid data structure
        self.grid = np.zeros((height, width), dtype=int)

        # Colors
        self.COLORS = {
            0: (255, 255, 255),  # Empty cell
            1: (0, 0, 0),  # Wall
            2: (0, 255, 0),  # Start point
            3: (255, 0, 0),  # End point
            4: (100, 100, 255)  # Path
        }

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

    def set_cell(self, x, y, value):
        if 0 <= x < self.width and 0 <= y < self.height:
            self.grid[y][x] = value

    def get_cell(self, x, y):
        if 0 <= x < self.width and 0 <= y < self.height:
            return self.grid[y][x]
        return None


def main():
    pygame.init()

    # Screen dimensions
    SCREEN_WIDTH = 800
    SCREEN_HEIGHT = 600
    screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
    pygame.display.set_caption("Pathfinding Playground")

    # Create grid
    grid = Grid(width=20, height=15)

    # Main game loop
    running = True
    while running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False

            # Handle mouse interactions
            if event.type == pygame.MOUSEBUTTONDOWN:
                x, y = event.pos
                grid_x = x // grid.cell_size
                grid_y = y // grid.cell_size
                grid.set_cell(grid_x, grid_y, 1)  # Set as wall on click

        screen.fill((255, 255, 255))
        grid.draw(screen)
        pygame.display.flip()

    pygame.quit()


if __name__ == "__main__":
    main()