''' Este módulo implementa la interfaz gráfica de usuario (GUI) para la hoja de cálculo.
    Proporciona una ventana principal con una barra de menú, una barra de herramientas y un área de trabajo 
    que consiste en un área de celdas y una barra de estado que permite cambiar a otras hojas de trabajo '
    además de la activa'
'''

import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog

SHIFT_PRESSED = 0x0001
CTRL_PRESSED = 0x0004

MAX_ROWS = 1000  # Maximum number of rows in the worksheet
MAX_COLS = 100  # Maximum number of columns in the worksheet

COL_CELLS_WIDTH = 40  # Default width for column cells in the worksheet
ROW_CELLS_HEIGHT = 20  # Default height for row cells in the worksheet
CELL_WIDTH = 60  # Default width for cells in the worksheet
CELL_HEIGHT = ROW_CELLS_HEIGHT  # Default height for cells in the worksheet

GRID_COLOR = "lightgray"  # Default grid color for the worksheet


class SheetUI(tk.Canvas):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.sheet_viewport = (5, 5, 5, 5)  # Default pivot cell
        self.active_cell = self.sheet_viewport[:2]  # Variable to store the active cell
        self.selected_cells = (*self.active_cell, *self.active_cell)  # Variable to store the selected cell
        self.bind("<Configure>", self.redraw_sheet)
        self.bind("<Button-1>", self.mouse_click)
        # bind arrow keys to move the active cell
        self.bind("<Up>", self.arrow_click)
        self.bind("<Down>", self.arrow_click)
        self.bind("<Left>", self.arrow_click)
        self.bind("<Right>", self.arrow_click)
        self.bind("<Return>", self.arrow_click)
        self.bind("<Prior>", self.arrow_click)
        self.bind("<Next>", self.arrow_click)
        # self.bind("<Key>", self.arrow_click)

        self.focus_set()  # Set focus to the canvas

    def cell_coordinates(self, x, y):
        """Calculates the coordinates of the cell based on the x and y position."""
        x0 = COL_CELLS_WIDTH + (x - self.sheet_viewport[0]) * CELL_WIDTH
        y0 = CELL_HEIGHT + (y - self.sheet_viewport[1]) * CELL_HEIGHT
        return (x0, y0, x0 + CELL_WIDTH, y0 + CELL_HEIGHT)
    
    def area_coordinates(self, x0, y0, x1, y1):
        sel_x0, sel_y0 = self.cell_coordinates(x0, y0)[:2]
        sel_x1, sel_y1 = self.cell_coordinates(x1, y1)[2:]
        return (sel_x0, sel_y0, sel_x1, sel_y1)
    
    def cell_containing_coords(self, x, y):
        """Returns the cell address containing the given x and y coordinates."""
        xcell = (x - COL_CELLS_WIDTH) // CELL_WIDTH + self.sheet_viewport[0]
        ycell = (y - CELL_HEIGHT) // CELL_HEIGHT + self.sheet_viewport[1]
        return (xcell, ycell)
    
    def setGUI(self):
        """Sets the GUI for the worksheet with a specified width and height."""

        # Determine the areas that needs redraw
        items = sorted(self.find_withtag("column"), key=lambda x: self.coords(x)[0])
        if not items:
            xroot = COL_CELLS_WIDTH
            width = self.winfo_width() - xroot
        else:
            first, last = items[0], items[-1]
            xroot = COL_CELLS_WIDTH
            width = self.coords(first)[0] - xroot
            if not width:
                xroot = self.coords(last)[2]
                winfo_width = self.winfo_width()
                if xroot >= winfo_width:
                    xroot = 0
                width = winfo_width - xroot
        
        items = sorted(self.find_withtag("row"), key=lambda x: self.coords(x)[1])
        if not items:
            yroot = ROW_CELLS_HEIGHT
            height = self.winfo_height() - yroot
        else:
            first, last = items[0], items[-1]
            yroot = ROW_CELLS_HEIGHT
            height = self.coords(first)[1] - yroot
            if not height:
                yroot = self.coords(last)[3]
                winfo_height = self.winfo_height()
                if yroot >= winfo_height:
                    yroot = 0
                height = winfo_height - yroot

        # We secure that the parameters are all integers
        xroot, width, yroot, height = map(int, (xroot, width, yroot, height))
        
        viewport_x0, viewport_y0, xcell, ycell = self.sheet_viewport
        if xroot > 0:
            winfo_height = self.winfo_height()
            x1 = xroot
            xcell = linf_x = self.cell_containing_coords(x1, 0)[0]
            while (x1 < xroot + width) and xcell <= MAX_COLS:
                x0, _, x1, _ = self.cell_coordinates(xcell, 0)
                self.create_rectangle(x0, 0, x1, ROW_CELLS_HEIGHT, fill="green", outline="black", tags="column")
                # Draw cell headings
                label = self.create_text((x0 + x1) // 2, ROW_CELLS_HEIGHT // 2, text=f"C{xcell}", fill="white")
                self.addtag_withtag("columns_tag", label)  # Add tag for columns
                # Draw vertical lines
                self.create_line(x0, 0, x0, winfo_height, fill=GRID_COLOR, tags="grid_lines")
                xcell += 1
            lsup_x = xcell - 1
            print(f"printed columns from {linf_x} to {lsup_x}")

        if yroot > 0:
            winfo_width = self.winfo_width()
            y1 = yroot
            ycell = linf_y = self.cell_containing_coords(0, y1)[1]
            while (y1 < yroot + height) and ycell <= MAX_ROWS:
                _, y0, _, y1 = self.cell_coordinates(0, ycell)
                self.create_rectangle(0, y0, COL_CELLS_WIDTH, y1, fill="green", outline="black", tags="row")
                # Draw row headings
                label = self.create_text(COL_CELLS_WIDTH // 2, (y0 + y1) // 2, text=f"R{ycell}", fill="white")
                self.addtag_withtag("rows_tag", label)
                # Draw horizontal lines
                self.create_line(0, y0, winfo_width, y0, fill=GRID_COLOR, tags="grid_lines")
                ycell += 1
            lsup_y = ycell - 1
            print(f"printed rows from {linf_y} to {lsup_y}")


        viewport_x0, viewport_y0 = self.cell_containing_coords(COL_CELLS_WIDTH, ROW_CELLS_HEIGHT)
        viewport_x1, viewport_y1 = self.cell_containing_coords(self.winfo_width(), self.winfo_height())
        # Draw the cells content
        if yroot > 0:
            print(f"redrawing cells from ({viewport_x0}, {linf_y}, {viewport_x1}, {lsup_y})")
            for y in range(linf_y, lsup_y + 1):
                for x in range(viewport_x0, viewport_x1 + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y)
                    # Draw cell content (placeholder text)
                    self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{x}R{y}", fill="black", tags="cell_content")
        if yroot == 0 and xroot > 0:
            print(f"redrawing cells from ({linf_x}, {viewport_y0}, {lsup_x}, {viewport_y1})")
            for x in range(linf_x, lsup_x + 1):
                for y in range(viewport_y0, viewport_y1 + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y)
                    # Draw cell content (placeholder text)
                    self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{x}R{y}", fill="black", tags="cell_content")
        if yroot > 0 and xroot > 0:
            if yroot == ROW_CELLS_HEIGHT:
                linf_y = lsup_y + 1
                lsup_y = viewport_y1
            else:
                assert yroot + height >= self.winfo_height(), "Height must be greater than the current height"
                lsup_y = linf_y - 1
                linf_y = viewport_y0
            print(f"redrawing cells from ({linf_x}, {linf_y}, {lsup_x}, {lsup_y})")
            for y in range(linf_y, lsup_y + 1):
                for x in range(linf_x, lsup_x + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y)
                    # Draw cell content (placeholder text)
                    self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{x}R{y}", fill="black", tags="cell_content")
        # Update the viewport
        self.sheet_viewport = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)

    def redraw_sheet(self, event):
        "Redraws the sheetui when the window is resized or needs updating."
        self.delete("all")
        self.create_rectangle(0, 0, COL_CELLS_WIDTH, ROW_CELLS_HEIGHT, fill="green", outline="black", tags="corner")
        self.setGUI()
        self.set_active_cell()

    def arrow_click(self, event):
        """Sets the active cell based on the arrow key pressed."""
        # print(f'{event.keysym} pressed')

        if event.keysym in ("Next", "Prior"):
            col = self.sheet_viewport[0]
            coord_acell_y0 = self.cell_coordinates(*self.active_cell)[1]
            if event.keysym == "Next":
                drow = self.cell_coordinates(0, self.sheet_viewport[-1])[1] - 1
                row = self.cell_containing_coords(0, ROW_CELLS_HEIGHT + 2 * (drow - ROW_CELLS_HEIGHT))[1]
            else:
                drow = self.winfo_height() - ROW_CELLS_HEIGHT
                row = self.cell_containing_coords(0, 2* ROW_CELLS_HEIGHT - drow)[1]
            self.set_sheet_viewport(col, row)
            acell_y0 = self.cell_containing_coords(0, coord_acell_y0)[1]
            self.active_cell = self.active_cell[0], acell_y0
            self.selected_cells = *self.active_cell, *self.active_cell
            self.set_active_cell()
            return "break"
        elif event.keysym == "Return":
            if self.selected_cells[:2] != self.selected_cells[2:]:
                sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
                acell_x0, acell_y0 = self.active_cell
                if event.state & SHIFT_PRESSED:  # If SHIFT is pressed
                    acell_x0 = acell_x0 if  acell_y0 > sel_y0 else ((acell_x0 - 1) if acell_x0 > sel_x0 else sel_x1)
                    acell_y0 = (acell_y0 - 1) if acell_y0 > sel_y0 else sel_y1
                else:  # If SHIFT is not pressed
                    acell_x0 = acell_x0 if acell_y0 < sel_y1 else ((acell_x0 + 1) if acell_x0 < sel_x1 else sel_x0)
                    acell_y0 = (acell_y0 + 1) if acell_y0 < sel_y1 else sel_y0
                self.active_cell = (acell_x0, acell_y0)
                self.set_active_cell()
                return "break"
            else:
                event.keysym = "Up" if event.state & SHIFT_PRESSED else "Down"  # Treat Return as Down for consistency
                event.state = 0

        acell_x0, acell_y0 = self.active_cell
        dx = dy = 0  # Initialize dx and dy for movement
        if event.keysym == "Up":
            dy = -1
        elif event.keysym == "Down":
            dy = 1
        elif event.keysym == "Left":
            dx = -1
        elif event.keysym == "Right":
            dx = 1
        self.offset_acell(dx, dy, state=event.state)
        return "break"  # Prevent default behavior of arrow keys
    
    def offset_acell(self, dx, dy, state=0):
        isShiftPressed = state & SHIFT_PRESSED
        isCtrlPressed = state & CTRL_PRESSED
        acell_x0, acell_y0 = self.active_cell
        if isShiftPressed: # Check if SHIFT is pressed
            sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
            xset = set([sel_x0, sel_x1]) - set([acell_x0])
            if len(xset) == 2:
                if dx < 0:
                    sel_x0 += dx
                else:
                    sel_x1 += dx
            else:
                dmy = sel_x0
                if len(xset) == 1:
                    dmy = xset.pop()
                dmy += dx
                sel_x1 = max(dmy, acell_x0)
                sel_x0 = min(dmy, acell_x0)
            
            yset = set([sel_y0, sel_y1]) - set([acell_y0])
            if len(yset) == 2:
                if dy < 0:
                    sel_y0 += dy
                else:
                    sel_y1 += dy
            else:
                dmy = sel_y0
                if len(yset) == 1:
                    dmy = yset.pop()
                dmy += dy
                sel_y1 = max(dmy, acell_y0)
                sel_y0 = min(dmy, acell_y0)
        else:
            if state & CTRL_PRESSED:
                dx = dx * ((acell_x0 - 1) if dx < 0 else (MAX_COLS - acell_x0))
                dy = dy * ((acell_y0 - 1) if dy < 0 else (MAX_ROWS - acell_y0))
            acell_x0 = xin = min(MAX_COLS, max(1, acell_x0 + dx))
            acell_y0 = yin  = min(MAX_ROWS, max(1, acell_y0 + dy))
            self.active_cell = (acell_x0, acell_y0)
            sel_x0, sel_y0 = acell_x0, acell_y0
            sel_x1, sel_y1 = sel_x0, sel_y0

        sel_x0, sel_x1 = map(lambda w: min(MAX_COLS, max(1, w)), (sel_x0, sel_x1))
        sel_y0, sel_y1 = map(lambda w: min(MAX_ROWS, max(1, w)), (sel_y0, sel_y1))

        xin = sel_x0 if sel_x0 != self.selected_cells[0] else sel_x1
        yin = sel_y0 if sel_y0 != self.selected_cells[1] else sel_y1
        
        self.selected_cells = (sel_x0, sel_y0, sel_x1, sel_y1)
        self.set_sheet_viewport(xin, yin)  # Adjust the viewport if necessary
    
    def set_sheet_viewport(self, xin, yin):
        """
        Assure that the point (xin, yin) is visible in the viewpoint
        """
        viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.sheet_viewport
        xin = min(MAX_COLS, max(1, xin))
        yin = min(MAX_ROWS, max(1, yin))

        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        dx = dy = 0  # Initialize dx and dy for movement
        x0 = xin
        if viewport_x0 > x0:
            linf_x = self.cell_coordinates(x0, 0)[0]
            if linf_x + winfo_width < COL_CELLS_WIDTH:
                # We need to move the viewport to the left a distance such that any information
                # displayed in the viewport is lost.
                lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
                items = self.find_enclosed(COL_CELLS_WIDTH - 1, -1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_x0 = x0
                dx = 1 # Needed to indicate redraw
            else:
                # Remove items that falls beyond the right edge
                dx = self.cell_coordinates(x0, 1)[0] - COL_CELLS_WIDTH
                lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
                viewport_x1, dmy = self.cell_containing_coords(winfo_width + dx, 0)
                linf_x = self.cell_coordinates(viewport_x1, dmy)[2]

                items = self.find_enclosed(linf_x - 1, -1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)

                # Move the viewport dx pixel to the left
                items = self.find_enclosed(COL_CELLS_WIDTH - 1, -1, linf_x + 1, lsup_y + 1)
                for item in items:
                    self.move(item, -dx, 0)
                viewport_x0 = x0
            self.sheet_viewport = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
        
        x1 = xin
        if viewport_x1 <= x1:
            x = self.cell_coordinates(x1, 0)[2]
            if x - winfo_width > winfo_width:
                # We need to move the viewport to the right a distance such that any information
                # displayed in the viewport is lost.
                lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
                items = self.find_enclosed(COL_CELLS_WIDTH - 1, -1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_x0 = self.cell_containing_coords(x - (winfo_width - COL_CELLS_WIDTH), 0)[0]
                viewport_x1 = x1
                self.sheet_viewport = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
            # Remove items that falls beyond the right edge
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            dx = self.cell_coordinates(x1, 2)[2] - winfo_width  + COL_CELLS_WIDTH
            x1, dmy = self.cell_containing_coords(dx, 0)
            linf_x = self.cell_coordinates(x1, dmy)[2]
            dx = linf_x - COL_CELLS_WIDTH

            items = self.find_enclosed(COL_CELLS_WIDTH - 1, -1, linf_x + 1, lsup_y + 1)
            self.delete(*items)
            # Move the viewport dx pixel to the left
            items = self.find_enclosed(linf_x - 1, -1, lsup_x + 1, lsup_y + 1)
            for item in items:
                self.move(item, -dx, 0)
            if dx < winfo_width -  COL_CELLS_WIDTH:
                viewport_x0 = self.cell_containing_coords(linf_x + 1, 0)[0]
            else:
                viewport_x0 = x1
            self.sheet_viewport = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates

        y0 = yin
        if viewport_y0 > y0:
            linf_y = self.cell_coordinates(0, y0)[1]
            if linf_y + winfo_height < ROW_CELLS_HEIGHT:
                # We need to move the viewport to the top a distance such that any information
                # displayed in the viewport is lost.
                lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
                items = self.find_enclosed(-1, ROW_CELLS_HEIGHT - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_y0 = y0
                dy = 1  # Needed to indicate redraw
            else:
                # Remove items that falls below the bottom edge
                dy = self.cell_coordinates(1, y0)[1] - ROW_CELLS_HEIGHT
                lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
                dmy, viewport_y1 = self.cell_containing_coords(0, winfo_height + dy)
                linf_y = self.cell_coordinates(dmy, viewport_y1)[3]

                items = self.find_enclosed(-1, linf_y - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)

                # Move the viewport dy pixel up
                items = self.find_enclosed(-1, ROW_CELLS_HEIGHT - 1,  lsup_x -dx + 1, linf_y + 1)
                for item in items:
                    self.move(item, 0, -dy)
                viewport_y0 = y0
            self.sheet_viewport = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
        
        y1 = yin
        if viewport_y1 <= y1:
            y = self.cell_coordinates(0, y1)[3]
            if y - winfo_height > winfo_height:
                # We need to move the viewport to the bottom a distance such that any information
                # displayed in the viewport is lost.
                lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
                items = self.find_enclosed(-1, ROW_CELLS_HEIGHT - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_y0 = self.cell_containing_coords(0, y - (winfo_height - ROW_CELLS_HEIGHT))[1]
                viewport_y1 = y1
                self.sheet_viewport = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
            # Remove items that falls below the bottom edge
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            dy = self.cell_coordinates(1, y1)[3] - winfo_height + ROW_CELLS_HEIGHT
            dmy, viewport_y1 = self.cell_containing_coords(0, dy)
            linf_y = self.cell_coordinates(dmy, viewport_y1)[3]
            dy = linf_y - ROW_CELLS_HEIGHT

            items = self.find_enclosed(-1, ROW_CELLS_HEIGHT - 1, lsup_x + 1, linf_y + 1)
            self.delete(*items)
            # Move the viewport dy pixel up
            items = self.find_enclosed(-1, linf_y - 1, lsup_x + 1, lsup_y + 1)
            for item in items:
                self.move(item, 0, -dy)
            viewport_y0 = self.cell_containing_coords(0, linf_y + 1)[1]
            self.sheet_viewport = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
        if dx or dy:
            self.setGUI()  # Redraw the sheet with the new viewport
        self.set_active_cell()
        # print(f"Viewport set to: {self.sheet_viewport}, Active cell: {self.active_cell}, Selected cells: {self.selected_cells}")

    def mouse_click(self, event):
        """Sets the active cell based on the click position."""
        if event.x < COL_CELLS_WIDTH or event.y < ROW_CELLS_HEIGHT:
            if event.x < COL_CELLS_WIDTH and event.y < ROW_CELLS_HEIGHT:
                # display a message box to get  the pivot cell coordinates
                answ = simpledialog.askstring("Set Pivot Cell", "Enter the pivot cell coordinates (col, row):", parent=self)
                if answ:
                    try:
                        x, y = map(lambda w: max(1, int(w)), answ.split(","))
                        self.set_sheet_viewport(x, y)
                    except ValueError:
                        print("Invalid input. Please enter coordinates in the format 'x, y'.")
                else:
                    print("No input provided.")
            elif event.x < COL_CELLS_WIDTH:
                print("Clicked on the column header")
            else:
                print("Clicked on the row header")
            return "break"
        clicked_cell = clk_x, clk_y = self.cell_containing_coords(event.x, event.y)
        acell_x0, acell_y0 = self.active_cell  # Get the top-left corner of the clicked cell
        if event.state & SHIFT_PRESSED:  # If SHIFT is pressed
            sel_x0 = min(clk_x, acell_x0)
            sel_y0 = min(clk_y, acell_y0)
            sel_x1 = max(clk_x, acell_x0)
            sel_y1 = max(clk_y, acell_y0)
            self.selected_cells = (sel_x0, sel_y0, sel_x1, sel_y1)
        else:  # If SHIFT is not pressed
            self.active_cell = clicked_cell
            self.selected_cells = (*clicked_cell, *clicked_cell)
        self.set_active_cell()
        self.focus_set()  # Set focus to the canvas

    def set_active_cell(self):
        # Set the tag "selected" for the region in coords (40, CELL_HEIGHT, 40 + 5*CELL_WIDTH, CELL_HEIGHT + 5*CELL_HEIGHT) rectangle
        def clip_rectangle(x0, y0, x1, y1, clipping_rgn=None):
            if clipping_rgn:
                linf_x, linf_y, lsup_x, lsup_y = clipping_rgn
                x0, x1 = min(lsup_x, max(linf_x, x0)), max(linf_x, min(x1, lsup_x))
                y0, y1 = min(lsup_y, max(linf_y, y0)), max(linf_y, min(y1, lsup_y))
            return x0, y0, x1, y1
        self.delete("selected_cells")
        clipping_rect = (COL_CELLS_WIDTH, ROW_CELLS_HEIGHT, self.winfo_width(), self.winfo_height())

        x0, y0, x1, y1 = self.area_coordinates(*self.selected_cells)
        sel_x0, sel_y0, sel_x1, sel_y1 = clip_rectangle(x0, y0, x1, y1, clipping_rect)
    
        self.create_rectangle(sel_x0, sel_y0, sel_x1, sel_y1,
            fill="lightblue", outline="black", tags="selected_cells")
        self.tag_lower("selected_cells", "grid_lines")


        self.delete("active_cell")
        """Draws the active cell rectangle."""
        x0, y0, x1, y1 = self.cell_coordinates(*self.active_cell)
        x0, y0, x1, y1 = clip_rectangle(x0, y0, x1, y1, clipping_rect)
        self.create_rectangle(
            x0, y0, x1, y1, 
            fill="yellow", outline="black", tags="active_cell"
        )
        self.itemconfigure("active_cell", outline="black", width=2)
        # place cell_content above the other tags
        self.tag_raise("cell_content")


        # change color for col_selected and row_selected
        old_selected = self.find_withtag("row_selected")
        new_selected = [srow for srow in self.find_enclosed(-1, sel_y0 - 1, COL_CELLS_WIDTH + 1, sel_y1 + 1) if self.type(srow) == "rectangle"]
        to_remove = set(old_selected) - set(new_selected)
        for row_id in to_remove:
            self.dtag(row_id, "row_selected")
            self.itemconfigure(row_id, fill="green")
        to_add = set(new_selected) - set(old_selected)
        for row_id in to_add:
            self.addtag_withtag("row_selected", row_id)
            self.itemconfigure(row_id, fill="blue")
        old_selected = self.find_withtag("col_selected")
        new_selected = [scol for scol in self.find_enclosed(sel_x0 - 1, -1, sel_x1 + 1, ROW_CELLS_HEIGHT + 1) if self.type(scol) == "rectangle"]
        to_remove = set(old_selected) - set(new_selected)
        for col_id in to_remove:
            self.dtag(col_id, "col_selected")
            self.itemconfigure(col_id, fill="green")
        to_add = set(new_selected) - set(old_selected)
        for col_id in to_add:
            self.addtag_withtag("col_selected", col_id)
            self.itemconfigure(col_id, fill="blue")
        
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

    # Create a frame to hold the canvas and scrollbars
    frame = ttk.Frame(root)
    frame.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    # Create vertical and horizontal scrollbars
    v_scroll = ttk.Scrollbar(frame, orient="vertical")
    h_scroll = ttk.Scrollbar(frame, orient="horizontal")

    # Create the SheetUI canvas
    sheetui = SheetUI(frame, bg="white", 
                      yscrollcommand=v_scroll.set, 
                      xscrollcommand=h_scroll.set,
                    #   scrollregion=(200, 200, 400, 400)
    )  # Adjust scrollregion as needed

    # Configure scrollbars to control the canvas
    v_scroll.config(command=sheetui.yview)
    h_scroll.config(command=sheetui.xview)

    # Layout
    sheetui.grid(row=0, column=0, sticky="nsew")
    v_scroll.grid(row=0, column=1, sticky="ns")
    h_scroll.grid(row=1, column=0, sticky="ew")

    root.mainloop()