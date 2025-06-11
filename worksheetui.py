''' Este módulo implementa la interfaz gráfica de usuario (GUI) para la hoja de cálculo.
    Proporciona una ventana principal con una barra de menú, una barra de herramientas y un área de trabajo 
    que consiste en un área de celdas y una barra de estado que permite cambiar a otras hojas de trabajo '
    además de la activa'
'''

import tkinter as tk
from tkinter import ttk



MAX_ROWS = 1000  # Maximum number of rows in the worksheet
MAX_COLS = 100  # Maximum number of columns in the worksheet

CELL_WIDTH = 60  # Default width for cells in the worksheet
CELL_HEIGHT = 20  # Default height for cells in the worksheet

GRID_COLOR = "lightgray"  # Default grid color for the worksheet


class SheetUI(tk.Canvas):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.active_cell = None  # Variable to store the active cell
        self.selected_cells = None  # Variable to store the selected cell
        self.bind("<Configure>", self.redraw_sheet)
        self.bind("<Button-1>", self.mouse_click)
        # bind arrow keys to move the active cell
        self.bind("<Up>", self.arrow_click)
        self.bind("<Down>", self.arrow_click)
        self.bind("<Left>", self.arrow_click)
        self.bind("<Right>", self.arrow_click)
        self.bind("<Return>", self.arrow_click)
        self.focus_set()  # Set focus to the canvas

    def redraw_sheet(self, event):
        "Redraws the sheetui when the window is resized or needs updating."
        active_cell = self.find_withtag("active_cell")
        self.delete("all")
        x0 = 40
        width = event.width  # Use the event width to set the width dynamically
        for x in range(x0, width, CELL_WIDTH):
            self.create_rectangle(x, 0, x + CELL_WIDTH, CELL_HEIGHT, fill="red", outline="black")
            # Draw cell headings
            label = self.create_text(x + CELL_WIDTH // 2, CELL_HEIGHT // 2, text=f"Col {x // CELL_WIDTH + 1}", fill="black")
            self.addtag_withtag("columns", label)  # Add tag for columns
            # Draw vertical lines
            self.create_line(x, 0, x, event.height, fill=GRID_COLOR, tags="grid_lines")
        y0 = CELL_HEIGHT
        height = event.height  # Use the event height to set the height dynamically
        for y in range(y0, height, CELL_HEIGHT):
            self.create_rectangle(0, y, x0, y + CELL_HEIGHT, fill="green", outline="black")
            # draw row headings
            label = self.create_text(x0 // 2, y + CELL_HEIGHT // 2, text=f"Row {y // CELL_HEIGHT + 1}", fill="black")
            self.addtag_withtag("rows", label)  # Add tag for rows
            # Draw horizontal lines
            self.create_line(0, y, width, y, fill=GRID_COLOR, tags="grid_lines")

        x0, y0 = 40, CELL_HEIGHT
        x1, y1 = x0 + CELL_WIDTH, y0 + CELL_HEIGHT
        if self.active_cell is None:
            self.selected_cells = (x0, y0, x1, y1)
        if active_cell:
            # Redraw the active cell if it exists
            x0, y0, x1, y1 = self.coords(active_cell[0])
        self.set_active_cell(x0, y0, x1, y1)

    def arrow_click(self, event):
        """Sets the active cell based on the arrow key pressed."""

        acell_x0, acell_y0, acell_x1, acell_y1 = self.active_cell

        if event.keysym == "Return":
            if self.selected_cells != self.active_cell:
                sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
                acell_x0, acell_y0, acell_x1, acell_y1 = self.active_cell
                if event.state & 0x0001:  # If SHIFT is pressed
                    acell_x0 = acell_x0 if  acell_y0 > sel_y0 else ((acell_x0 - CELL_WIDTH) if acell_x0 > sel_x0 else (sel_x1 - CELL_WIDTH))
                    acell_y0 = (acell_y0 - CELL_HEIGHT) if acell_y0 > sel_y0 else (sel_y1 - CELL_HEIGHT)
                else:  # If SHIFT is not pressed
                    acell_y0 = acell_y1 if acell_y1 < sel_y1 else sel_y0
                    acell_x0 = acell_x0 if acell_y1 < sel_y1 else (acell_x1 if acell_x1 < sel_x1 else sel_x0)
                self.set_active_cell(acell_x0, acell_y0, acell_x0 + CELL_WIDTH, acell_y0 + CELL_HEIGHT)
                return "break"
            else:
                event.keysym = "Up" if event.state & 0x0001 else "Down"  # Treat Return as Down for consistency
                event.state = 0

        dx = dy = 0  # Initialize dx and dy for movement
        if event.keysym == "Up":
            dy = -CELL_HEIGHT
        elif event.keysym == "Down":
            dy = CELL_HEIGHT
        elif event.keysym == "Left":
            dx = -CELL_WIDTH
        elif event.keysym == "Right":
            dx = CELL_WIDTH

        if not event.state & 0x0001:  # Check if SHIFT is pressed
            acell_x0 = max(40, acell_x0 + dx)
            acell_y0 = max(CELL_HEIGHT, acell_y0 + dy)
            sel_x0, sel_y0 = acell_x0, acell_y0
            sel_x1, sel_y1 = acell_x0 + CELL_WIDTH, acell_y0 + CELL_HEIGHT
        else: # If SHIFT is pressed, adjust the selection
            sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
            xset = set([sel_x0, sel_x1]) - set([acell_x0, acell_x0 + CELL_WIDTH])
            if len(xset) == 2:
                if dx < 0:
                    sel_x0 += dx
                else:
                    sel_x1 += dx
            else:
                if len(xset) == 0:
                    sel_x0 = (acell_x1 if dx > 0 else acell_x0)
                else:   # len(xset) == 1:
                    sel_x0 = xset.pop()
                sel_x0 += dx
                sel_x1 = max(sel_x0, acell_x0, acell_x1)
                sel_x0 = min(sel_x0, acell_x0, acell_x1)

            yset = set([sel_y0, sel_y1]) - set([acell_y0, acell_y0 + CELL_HEIGHT])
            if len(yset) == 2:
                if dy < 0:
                    sel_y0 += dy
                else:
                    sel_y1 += dy
            else:
                if len(yset) == 0:
                    sel_y0 = acell_y1 if dy > 0 else acell_y0
                else:   # len(yset) == 1:
                    sel_y0 = yset.pop()
                sel_y0 += dy
                sel_y1 = max(sel_y0, acell_y0, acell_y1)
                sel_y0 = min(sel_y0, acell_y0, acell_y1)

        sel_x0 = max(40, sel_x0)        # Assure canvas boundaries
        sel_y0 = max(CELL_HEIGHT, sel_y0)

        self.selected_cells = (sel_x0, sel_y0, sel_x1, sel_y1)
        self.set_active_cell(acell_x0, acell_y0, acell_x0 + CELL_WIDTH, acell_y0 + CELL_HEIGHT)

        return "break"  # Prevent default behavior of arrow keys

    def mouse_click(self, event):
        """Sets the active cell based on the click position."""
        x = 40 + (event.x - 40) // CELL_WIDTH * CELL_WIDTH
        y = CELL_HEIGHT + (event.y - CELL_HEIGHT) // CELL_HEIGHT * CELL_HEIGHT
        clicked_cell = (x, y, x + CELL_WIDTH, y + CELL_HEIGHT)
        acell_x0, acell_y0, acell_x1, acell_y1 = self.active_cell
        if event.state & 0x0001:  # If SHIFT is pressed
            sel_x0 = min(x, acell_x0)
            sel_y0 = min(y, acell_y0)
            sel_x1 = max(x + CELL_WIDTH, acell_x1)
            sel_y1 = max(y + CELL_HEIGHT, acell_y1)
            self.selected_cells = (sel_x0, sel_y0, sel_x1, sel_y1)
        else:  # If SHIFT is not pressed
            self.selected_cells = clicked_cell
            self.active_cell = clicked_cell
        self.set_active_cell(*self.active_cell)
        self.focus_set()  # Set focus to the canvas

    def set_active_cell(self, x0, y0, x1, y1):
        # Set the tag "selected" for the region in coords (40, CELL_HEIGHT, 40 + 5*CELL_WIDTH, CELL_HEIGHT + 5*CELL_HEIGHT) rectangle
        self.delete("selected_cells")
        self.create_rectangle(*self.selected_cells,
            fill="white", outline="black", tags="selected_cells")
        self.tag_lower("selected_cells", "grid_lines")


        self.delete("active_cell")
        """Draws the active cell rectangle."""
        self.create_rectangle(
            x0, y0, x1, y1, 
            fill="yellow", outline="black", tags="active_cell"
        )
        self.itemconfigure("active_cell", outline="black", width=2)

        sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells

        # change color for col_selected and row_selected
        old_selected = self.find_withtag("row_selected")
        new_selected = [srow for srow in self.find_enclosed(-1, sel_y0 - 1, 40 + 1, sel_y1 + 1) if self.type(srow) == "rectangle"]
        to_remove = set(old_selected) - set(new_selected)
        for row_id in to_remove:
            self.dtag(row_id, "row_selected")
            self.itemconfigure(row_id, fill="green")
        to_add = set(new_selected) - set(old_selected)
        for row_id in to_add:
            self.addtag_withtag("row_selected", row_id)
            self.itemconfigure(row_id, fill="blue")
        # self.addtag_withtag("row_selected", *selected_rows)
        # delete previous col_selected if exists
        # self.itemconfigure("row_selected", fill="red")
        old_selected = self.find_withtag("col_selected")
        new_selected = [scol for scol in self.find_enclosed(sel_x0 - 1, -1, sel_x1 + 1, CELL_HEIGHT + 1) if self.type(scol) == "rectangle"]
        to_remove = set(old_selected) - set(new_selected)
        for col_id in to_remove:
            self.dtag(col_id, "col_selected")
            self.itemconfigure(col_id, fill="red")
        to_add = set(new_selected) - set(old_selected)
        for col_id in to_add:
            self.addtag_withtag("col_selected", col_id)
            self.itemconfigure(col_id, fill="blue")
        
        self.active_cell = (x0, y0, x1, y1)  # Store the active cell coordinates

    def setGUI(self):
        x0 = 40
        width = 400  # Default width if not set
        for x in range(x0, width, CELL_WIDTH):
            self.create_rectangle(x, 0, x + CELL_WIDTH, CELL_HEIGHT, fill="red", outline="black")
        y0 = CELL_HEIGHT
        height = 400  # Default height if not set
        for y in range(y0, height, CELL_HEIGHT):
            self.create_rectangle(0, y, x0, y + CELL_HEIGHT, fill="green", outline="black")


    def setColor(self, color):
        self.color = color
        self.dtag("all", "paletteSelected")
        self.itemconfigure("palette", outline="white", width=5)
        self.addtag("paletteSelected", "withtag", "palette" + color)
        self.itemconfigure("paletteSelected", outline="#999999")
        

if __name__ == "__main__":
    root = tk.Tk()
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    sheetui = SheetUI(root)
    sheetui.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))

    root.mainloop()
