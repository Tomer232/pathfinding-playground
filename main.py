import pygame
import numpy as np
import random
import os
import heapq
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from collections import deque

class Grid:
    def __init__(self, screen_width=800, screen_height=600, cell_size=7):
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

        # Generate maze
        self._generate_maze()

        # Colors
        self.COLORS = {
            0: (255, 255, 255),  # Empty cell
            1: (0, 0, 0),  # Wall
            2: (0, 255, 0),  # Start point
            3: (255, 0, 0),  # End point
            4: (100, 100, 255),  # Path
            5: (173, 216, 230)  # Visited cell
        }

    def _create_outer_walls(self):
        # Top and bottom walls
        self.grid[0, :] = 1
        self.grid[-1, :] = 1

        # Left and right walls
        self.grid[:, 0] = 1
        self.grid[:, -1] = 1

    def _generate_maze(self):
        for y in range(1, self.height - 1):
            for x in range(1, self.width - 1):
                if random.random() < 0.3:  # 30% chance to place a wall
                    self.grid[y, x] = 1

    def place_start_and_end_points(self):
        self.start_point = self._get_random_safe_position()
        self.end_point = self._get_random_safe_position(far_from=self.start_point, min_distance=30)

        # Set start and end points on the grid
        self.grid[self.start_point[1], self.start_point[0]] = 2  # Start point
        self.grid[self.end_point[1], self.end_point[0]] = 3      # End point

        # Ensure there is a path between start and end points
        self._create_path(self.start_point, self.end_point)

    def _create_path(self, start, end):
        """Ensure a valid path exists between start and end points."""
        x, y = start
        ex, ey = end

        while (x, y) != (ex, ey):
            # Move horizontally or vertically towards the endpoint
            if x != ex:
                x += 1 if x < ex else -1
            elif y != ey:
                y += 1 if y < ey else -1

            # Avoid overwriting the start or end points
            if self.grid[y, x] not in [2, 3]:
                self.grid[y, x] = 0

    def _get_random_safe_position(self, far_from=None, min_distance=0):
        while True:
            x = random.randint(1, self.width - 2)
            y = random.randint(1, self.height - 2)

            # Check if the cell is empty
            if self.grid[y, x] == 0:
                if far_from:
                    distance = abs(far_from[0] - x) + abs(far_from[1] - y)
                    if distance >= min_distance:
                        return x, y
                else:
                    return x, y

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

    def reset_visualization(self):
        """Reset the grid to its initial state for visualization."""
        for y in range(self.height):
            for x in range(self.width):
                if self.grid[y, x] in [4, 5]:  # Reset path and visited cells
                    self.grid[y, x] = 0

        # Ensure start and end points remain the same
        self.grid[self.start_point[1], self.start_point[0]] = 2
        self.grid[self.end_point[1], self.end_point[0]] = 3

    def bfs(self, screen):
        start = self.start_point
        end = self.end_point
        directions = [
            (0, 1), (1, 0), (0, -1), (-1, 0),  # Up, Right, Down, Left
            (1, 1), (-1, -1), (1, -1), (-1, 1)  # Diagonal directions
        ]
        visited = set()
        queue = deque([(start, [])])

        while queue:
            (x, y), path = queue.popleft()

            if (x, y) in visited:
                continue

            visited.add((x, y))
            path = path + [(x, y)]

            # Visualize the visited cell
            if self.grid[y][x] not in [2, 3]:
                self.grid[y][x] = 5
            self.draw(screen)
            pygame.display.flip()
            pygame.time.delay(1)

            if (x, y) == end:
                for px, py in path:
                    if self.grid[py][px] not in [2, 3]:
                        self.grid[py][px] = 4
                return True

            for dx, dy in directions:
                nx, ny = x + dx, y + dy
                if 0 <= nx < self.width and 0 <= ny < self.height and self.grid[ny][nx] != 1:
                    queue.append(((nx, ny), path))

        return False

    def dfs(self, screen):
        start = self.start_point
        end = self.end_point
        directions = [
            (0, 1), (1, 0), (0, -1), (-1, 0),  # Up, Right, Down, Left
            (1, 1), (-1, -1), (1, -1), (-1, 1)  # Diagonal directions
        ]
        visited = set()
        stack = [(start, [])]

        while stack:
            (x, y), path = stack.pop()

            if (x, y) in visited:
                continue

            visited.add((x, y))
            path = path + [(x, y)]

            # Visualize the visited cell
            if self.grid[y][x] not in [2, 3]:
                self.grid[y][x] = 5
            self.draw(screen)
            pygame.display.flip()
            pygame.time.delay(1)

            if (x, y) == end:
                for px, py in path:
                    if self.grid[py][px] not in [2, 3]:
                        self.grid[py][px] = 4
                return True

            for dx, dy in directions:
                nx, ny = x + dx, y + dy
                if 0 <= nx < self.width and 0 <= ny < self.height and self.grid[ny][nx] != 1:
                    stack.append(((nx, ny), path))

        return False

    def a_star(self, screen):
        start = self.start_point
        end = self.end_point
        directions = [
            (0, 1), (1, 0), (0, -1), (-1, 0),  # Up, Right, Down, Left
            (1, 1), (-1, -1), (1, -1), (-1, 1)  # Diagonal directions
        ]
        random.shuffle(directions)

        pq = []
        heapq.heappush(pq, (0, 0, start, []))
        visited = set()

        def heuristic(current, goal):
            return ((current[0] - goal[0]) ** 2 + (current[1] - goal[1]) ** 2) ** 0.5

        while pq:
            _, g, (x, y), path = heapq.heappop(pq)

            if (x, y) in visited:
                continue

            visited.add((x, y))
            path = path + [(x, y)]

            if self.grid[y][x] not in [2, 3]:
                self.grid[y][x] = 5
            self.draw(screen)
            pygame.display.flip()
            pygame.time.delay(1)

            if (x, y) == end:
                for px, py in path:
                    if self.grid[py][px] not in [2, 3]:
                        self.grid[py][px] = 4
                return True

            random.shuffle(directions)
            for dx, dy in directions:
                nx, ny = x + dx, y + dy
                if 0 <= nx < self.width and 0 <= ny < self.height and self.grid[ny][nx] != 1:
                    f = g + 1 + heuristic((nx, ny), end)
                    heapq.heappush(pq, (f, g + 1, (nx, ny), path))

        return False


def save_playground_with_data(screen, playground_name, best_algorithm):
    base_path = r"C:\Users\tomer\Desktop\Projects\pathfinding"
    excel_file = os.path.join(base_path, "playground_metadata.xlsx")
    image_folder = os.path.join(base_path, "playground_images")

    if not os.path.exists(image_folder):
        os.makedirs(image_folder)

    image_path = os.path.join(image_folder, f"{playground_name}.png")
    pygame.image.save(screen, image_path)

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

    next_row = ws.max_row + 1
    ws.cell(row=next_row, column=1, value=playground_name)
    ws.cell(row=next_row, column=2, value=date)
    ws.cell(row=next_row, column=3, value=time)
    ws.cell(row=next_row, column=4, value=best_algorithm)

    img = ExcelImage(image_path)
    img.anchor = f"E{next_row}"
    img.width = 160
    img.height = 120
    ws.add_image(img)

    wb.save(excel_file)
    print(f"Playground and data saved in '{excel_file}'")


current_algorithm_running = False


def main():
    global current_algorithm_running

    pygame.init()

    SCREEN_WIDTH = 1200
    SCREEN_HEIGHT = 600

    screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
    pygame.display.set_caption("Pathfinding Playground")

    grid = Grid(cell_size=7)
    grid.place_start_and_end_points()

    buttons = [
        {"name": "BFS", "rect": pygame.Rect(900, 50, 200, 50)},
        {"name": "DFS", "rect": pygame.Rect(900, 120, 200, 50)},
        {"name": "A*", "rect": pygame.Rect(900, 190, 200, 50)},
        {"name": "NEW MAP", "rect": pygame.Rect(900, 260, 200, 50)}
    ]

    def draw_buttons():
        font = pygame.font.SysFont("Verdana", 36)
        button_color = (232, 241, 242)
        outline_color = (0, 0, 0)
        text_color = (0, 0, 0)
        panel_color = (108, 117, 107)

        panel_rect = pygame.Rect(900, 0, 300, SCREEN_HEIGHT)
        pygame.draw.rect(screen, panel_color, panel_rect)

        for button in buttons:
            pygame.draw.rect(screen, outline_color, button["rect"], 2)
            pygame.draw.rect(screen, button_color, button["rect"].inflate(-4, -4))
            text = font.render(button["name"], True, text_color)
            text_rect = text.get_rect(center=button["rect"].center)
            screen.blit(text, text_rect)

    screen.fill((255, 255, 255))
    grid.draw(screen)
    draw_buttons()

    running = True
    while running:
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                running = False
            elif event.type == pygame.MOUSEBUTTONDOWN:
                mouse_pos = event.pos
                for button in buttons:
                    if button["rect"].collidepoint(mouse_pos):
                        if button["name"] == "BFS" and not current_algorithm_running:
                            current_algorithm_running = True
                            grid.reset_visualization()
                            if grid.bfs(screen):
                                save_playground_with_data(screen, "playground_bfs", "BFS")
                            current_algorithm_running = False
                        elif button["name"] == "DFS" and not current_algorithm_running:
                            current_algorithm_running = True
                            grid.reset_visualization()
                            if grid.dfs(screen):
                                save_playground_with_data(screen, "playground_dfs", "DFS")
                            current_algorithm_running = False
                        elif button["name"] == "A*" and not current_algorithm_running:
                            current_algorithm_running = True
                            grid.reset_visualization()
                            if grid.a_star(screen):
                                save_playground_with_data(screen, "playground_astar", "A*")
                            current_algorithm_running = False
                        elif button["name"] == "NEW MAP":
                            grid = Grid(cell_size=7)
                            grid.place_start_and_end_points()

        screen.fill((255, 255, 255))
        grid.draw(screen)
        draw_buttons()
        pygame.display.flip()

    pygame.quit()


if __name__ == "__main__":
    main()
