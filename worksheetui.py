''' Este módulo implementa la interfaz gráfica de usuario (GUI) para la hoja de cálculo.
    Proporciona una ventana principal con una barra de menú, una barra de herramientas y un área de trabajo 
    que consiste en un área de celdas y una barra de estado que permite cambiar a otras hojas de trabajo '
    además de la activa'
'''

import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from contextlib import contextmanager
from types import SimpleNamespace


SHIFT_PRESSED = 0x00001
CTRL_PRESSED = 0x00004
ALT_PRESSED = 0x20000

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
        self.coords_vportq3 = (COL_CELLS_WIDTH, ROW_CELLS_HEIGHT)
        self.viewport_q3 = (2, 3, 4, 6)
        self.coords_vportq1 = (COL_CELLS_WIDTH + 2 * CELL_WIDTH, ROW_CELLS_HEIGHT + 3 * CELL_HEIGHT)
        self.viewport_q1 = (4, 6, 4, 6)  # Default pivot cell
        self.active_cell = self.viewport_q1[:2]  # Variable to store the active cell
        self.selected_cells = (*self.active_cell, *self.active_cell)  # Variable to store the selected cell
        self.bind("<Configure>", self.redraw_sheet)
        self.bind("<Button-1>", self.mouse_click)
        self.bind("<B1-Motion>", self.mouse_drag)
        self.bind("<ButtonRelease-1>", self.mouse_release)

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

    @contextmanager
    def pivot_point(self, isActiveCell=False, isUp=0):
        acell_x0, acell_y0 = x, y = self.active_cell
        if not isActiveCell:
            sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
            bx = len(set([sel_x0, sel_x1]) - set([acell_x0])) <= 1
            if bx:
                x = sel_x0 + sel_x1 - acell_x0
            else:
                x = sel_x0 if isUp & 0x1 else sel_x1 

            by = len(set([sel_y0, sel_y1]) - set([acell_y0])) <= 1
            if by:
                y = sel_y0 + sel_y1 - acell_y0
            else:
                y = sel_y0 if isUp & 0x2 else sel_y1

        pt = SimpleNamespace(x=x, y=y)

        yield pt

        if isActiveCell:
            self.active_cell = pt.x, pt.y
            self.selected_cells = (*self.active_cell, *self.active_cell)
        else:
            acell_x0, acell_y0 = self.active_cell

            x = sel_x0 if isUp & 0x1 else sel_x1 

            if bx:
                sel_x0, sel_x1 = min(acell_x0, pt.x), max(acell_x0, pt.x)
            else:
                if isUp & 0x1:
                    sel_x0 = min(acell_x0, pt.x)
                else:
                    sel_x1 = max(acell_x0, pt.x)

            if by:
                sel_y0, sel_y1 = min(acell_y0, pt.y), max(acell_y0, pt.y)
            else:
                if isUp & 0x02:
                    sel_y0 = min(acell_y0, pt.y)
                else:
                    sel_y1 = max(acell_y0, pt.y)
            self.selected_cells = sel_x0, sel_y0, sel_x1, sel_y1

    def cell_coordinates(self, x, y, viewport=None, coords_viewport=None):
        """Calculates the coordinates of the cell based on the x and y position."""
        if viewport is None:
            coords_viewport, viewport = self.coords_vportq1, self.viewport_q1
        x0 = coords_viewport[0] + (x - viewport[0]) * CELL_WIDTH
        y0 = coords_viewport[1] + (y - viewport[1]) * CELL_HEIGHT
        return (x0, y0, x0 + CELL_WIDTH, y0 + CELL_HEIGHT)
    
    def area_coordinates(self, x0, y0, x1, y1):
        sel_x0, sel_y0 = self.cell_coordinates(x0, y0)[:2]
        sel_x1, sel_y1 = self.cell_coordinates(x1, y1)[2:]
        return (sel_x0, sel_y0, sel_x1, sel_y1)
    
    def cell_containing_coords(self, x, y, viewport=None, coords_viewport=None):
        """Returns the cell address containing the given x and y screen coordinates."""
        if viewport is None:
            coords_viewport, viewport = self.coords_vportq1, self.viewport_q1
        xcell = (x - coords_viewport[0]) // CELL_WIDTH + viewport[0]
        ycell = (y - coords_viewport[1]) // CELL_HEIGHT + viewport[1]
        xcell = max(1, min(MAX_COLS, xcell))
        ycell = max(1, min(MAX_ROWS, ycell))
        return (xcell, ycell)
    
    def quadrant_data(self, nquadrant):
        """Returns the (vieport, coords_viewport) for the given quadrant."""
        if nquadrant == 1:
            return self.viewport_q1, self.coords_vportq1
        elif nquadrant == 2:
            orig = self.viewport_q1[0], self.viewport_q3[1], self.viewport_q1[2], self.viewport_q3[3]
            return orig, (self.coords_vportq1[0], self.coords_vportq3[1])
        elif nquadrant == 3:
            return self.viewport_q3, self.coords_vportq3
        else:
            orig = self.viewport_q3[0], self.viewport_q1[1], self.viewport_q3[1], self.viewport_q1[3]
            return orig, (self.coords_vportq3[0], self.coords_vportq1[1])
    
    def setGUI(self):
        """Sets the GUI for the worksheet with a specified width and height."""

        # Determine the areas that needs redraw
        items = sorted(self.find_withtag("column"), key=lambda x: self.coords(x)[0])
        dmy0 = dmy1 = COL_CELLS_WIDTH
        ndx = 0
        while ndx < len(items):
            dmy, _, dmy1, _ = self.coords(items[ndx])
            if dmy != dmy0:
                xroot = dmy0
                width = dmy - dmy0
                break
            ndx += 1
            dmy0 = dmy1
        else:
            xroot = 0 if dmy1 >= self.winfo_width() else dmy1
            width = self.winfo_width() - xroot
        xroot_in = xroot

        items = sorted(self.find_withtag("row"), key=lambda x: self.coords(x)[1])
        dmy0 = dmy1 = ROW_CELLS_HEIGHT
        ndx = 0
        while ndx < len(items):
            _, dmy, _, dmy1 = self.coords(items[ndx])
            if dmy != dmy0:
                yroot = dmy0
                height = dmy - dmy0
                break
            ndx += 1
            dmy0 = dmy1
        else:
            yroot = 0 if dmy1 >= self.winfo_height() else dmy1
            height = self.winfo_height() - yroot
        yroot_in = yroot

        # We secure that the parameters are all integers
        xroot, width, yroot, height = map(int, (xroot, width, yroot, height))
        
        viewport_x0, viewport_y0, xcell, ycell = self.viewport_q1
        if xroot > 0:
            winfo_height = self.winfo_height()
            if xroot == self.coords_vportq3[0]:
                pane_width = self.coords_vportq1[0] - self.coords_vportq3[0]
                task = [
                    (self.coords_vportq1[0], width - pane_width, None, None),
                    (self.coords_vportq3[0], pane_width, self.viewport_q3, self.coords_vportq3),
                ]
            else:
                task = [(xroot, width, None, None)]
            while task:
                xroot, width, orig, coords_orig = task.pop()
                x1 = xroot
                xcell = linf_x = self.cell_containing_coords(x1, 0, orig, coords_orig)[0]
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
            if yroot == self.coords_vportq3[1]:
                pane_height = self.coords_vportq1[1] - self.coords_vportq3[1]
                task = [
                    (self.coords_vportq1[1], height - pane_height, None, None),
                    (self.coords_vportq3[1], pane_height, self.viewport_q3, self.coords_vportq3),
                ]
            else:
                task = [(yroot, height, None, None)]

            while task:
                yroot, height, orig, coords_orig = task.pop()
                y1 = yroot
                ycell = linf_y = self.cell_containing_coords(0, y1, orig, coords_orig)[1]
                while (y1 < yroot + height) and ycell <= MAX_ROWS:
                    _, y0, _, y1 = self.cell_coordinates(0, ycell)
                    self.create_rectangle(0, y0, COL_CELLS_WIDTH, y1, fill="green", outline="black", tags="row")
                    # Draw cell headings
                    label = self.create_text(COL_CELLS_WIDTH // 2, (y0 + y1) // 2, text=f"C{ycell}", fill="white")
                    self.addtag_withtag("rows_tag", label)  # Add tag for columns
                    # Draw Horizontal lines
                    self.create_line(0, y0, winfo_width, y0, fill=GRID_COLOR, tags="grid_lines")
                    ycell += 1
                lsup_y = ycell - 1

        viewport_x0, viewport_y0 = self.cell_containing_coords(*self.coords_vportq1)
        viewport_x1, viewport_y1 = self.cell_containing_coords(self.winfo_width(), self.winfo_height())


        # Draw the cells content
        if xroot > 0:
            # Segundo cuadrante
            # print(f"redrawing cells from 2 scnd cuadrant ({viewport_x0}, {linf_y}, {viewport_x1}, {lsup_y})")
            orig, coords_orig =  self.quadrant_data(2)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            if linf_x <= orig_x1 and lsup_x >= orig_x0:
                xmin, xmax = max(linf_x, orig_x0), min(lsup_x, orig_x1)
                for x in range(xmin, xmax + 1):
                    for y in range(orig_y0, orig_y1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"Q2_C{x}R{y}", fill="black", tags="cell_content")

              # Cuarto cuadrante
                for x in range(xmin, xmax + 1):
                    for y in range(viewport_y0, viewport_y1 + 1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{x}R{y}", fill="black", tags="cell_content")

        if xroot == 0:
            # Cuarto cuadrante
            orig, coords_orig =  self.quadrant_data(4)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            if linf_y <= orig_y1 and lsup_y >= orig_y0:
                ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
                for y in range(ymin, ymax + 1):
                    for x in range(orig_x0, orig_x1 + 1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"Q4_C{x}R{y}", fill="black", tags="cell_content")
                # Primer cuadrante
                for y in range(ymin, ymax + 1):
                    for x in range(viewport_x0, viewport_x1 + 1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{x}R{y}", fill="black", tags="cell_content")

        if xroot_in > 0 and yroot_in > 0:
            if (xroot_in, yroot_in) == (COL_CELLS_WIDTH, ROW_CELLS_HEIGHT):
                # Tercer cuadrante
                orig, coords_orig =  self.quadrant_data(3)
                for y in range(orig[1], orig[3]):
                    for x in range(orig[0], orig[2]):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"Q3_C{x}R{y}", fill="black", tags="cell_content")
            # Cuarto cuadrante
            orig, coords_orig =  self.quadrant_data(4)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
            for y in range(ymin, ymax + 1):
                for x in range(orig_x0, orig_x1 + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                    # Draw cell content (placeholder text)
                    self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"Q4_C{x}R{y}", fill="black", tags="cell_content")

            # Primer cuadrante
            for x in range(viewport_x0, linf_x):
                for y in range(ymin, ymax + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y)
                    # Draw cell content (placeholder text)
                    self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{x}R{y}", fill="black", tags="cell_content")

            for x in range(lsup_x + 1, viewport_x1 + 1):
                for y in range(ymin, ymax + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y)
                    # Draw cell content (placeholder text)
                    self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{x}R{y}", fill="black", tags="cell_content")   

        # Update the viewport
        self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)

    def redraw_sheet(self, event):
        "Redraws the sheetui when the window is resized or needs updating."
        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        viewport_x1, viewport_y1 = self.cell_containing_coords(winfo_width, winfo_height)
        self.viewport_q1 = self.viewport_q1[:2] + (viewport_x1, viewport_y1)
        self.configure(scrollregion=(0, 0, MAX_COLS * CELL_WIDTH, MAX_ROWS * CELL_HEIGHT))
        self.delete("all")
        self.create_rectangle(0, 0, COL_CELLS_WIDTH, ROW_CELLS_HEIGHT, fill="green", outline="black", tags="corner")
        self.setGUI()
        self.set_active_cell()

    def arrow_click(self, event):
        """Sets the active cell based on the arrow key pressed."""
        # print(f'{event.keysym} pressed')
        
        if event.keysym in ("Next", "Prior"):
            viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.viewport_q1
            winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
            acell_x0, acell_y0 = self.active_cell
            with self.pivot_point(isActiveCell=not event.state & SHIFT_PRESSED) as pivot:
                coord_pivot_x, coord_pivot_y = self.cell_coordinates(pivot.x, pivot.y)[:2]
                if event.state & ALT_PRESSED:
                    if event.keysym == "Next":
                        viewport_x0 = viewport_x1 if viewport_x1 < MAX_COLS else viewport_x0
                    else:
                        xright = self.cell_coordinates(viewport_x0, 0)[2]
                        viewport_x0 = self.cell_containing_coords(xright - (winfo_width - COL_CELLS_WIDTH), 0)[0]
                        viewport_x0 = min(MAX_COLS, max(1, viewport_x0))
                else:
                    if event.keysym == "Next":
                        viewport_y0 = viewport_y1 if viewport_x1 < MAX_ROWS else viewport_y0
                    else:
                        ybottom = self.cell_coordinates(0, viewport_y0)[3]
                        viewport_y0 = self.cell_containing_coords(0, ybottom - (winfo_height - ROW_CELLS_HEIGHT))[1]
                        viewport_y0 = min(MAX_ROWS, max(1, viewport_y0))
                self.move_viewport(viewport_x0, viewport_y0)
                pivot.x = self.cell_containing_coords(coord_pivot_x, 0)[0]
                pivot.y = self.cell_containing_coords(0, coord_pivot_y)[1]
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
    
    def offset_acell(self, dx, dy, state):
        isShiftPressed = state & SHIFT_PRESSED
        isCtrlPressed = state & CTRL_PRESSED
        isup = (dx < 0) * 0x1 + (dy < 0) * 0x2
        with self.pivot_point(isActiveCell=not isShiftPressed, isUp=isup) as pivot:
            if isCtrlPressed:
                dx = dx * ((pivot.x - 1) if dx < 0 else (MAX_COLS - pivot.x))
                dy = dy * ((pivot.y - 1) if dy < 0 else (MAX_ROWS - pivot.y))
            pivot.x = max(self.viewport_q3[2], min(MAX_COLS, pivot.x + dx))
            pivot.y = max(self.viewport_q3[3], min(MAX_ROWS, pivot.y + dy))
            xin, yin = pivot.x, pivot.y
        self.show_cell(xin, yin)

    def move_viewport(self, x, y):
        """
        Move the viewport origen (deltax, deltay) pixels
        """
        viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.viewport_q1
        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        x = max(1, min(MAX_COLS, x))
        y = max(1, min(MAX_ROWS, y))
        # x, y = self.cell_containing_coords(deltax + COL_CELLS_WIDTH, deltay + ROW_CELLS_HEIGHT)
        linf_x, linf_y = self.cell_coordinates(x, y)[:2]
        deltax, deltay = linf_x - self.coords_vportq1[0], linf_y - self.coords_vportq1[1]
        dx = dy = 0
        if deltax < 0:
            # left displacement
            x0 = x
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            if deltax + winfo_width < self.coords_vportq1[0]:
                # We need to move the viewport to the left a distance such that any information
                # displayed in the viewport is lost.
                items = self.find_enclosed(self.coords_vportq1[0] - 1, -1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_x0 = x0
                dx = 1 # Needed to indicate redraw
            else:
                # Remove items that falls beyond the right edge
                dx = deltax
                viewport_x1, dmy = self.cell_containing_coords(winfo_width + dx, 0)
                linf_x = self.cell_coordinates(viewport_x1, dmy)[2]

                items = self.find_enclosed(linf_x - 1, -1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)

                # Move the viewport dx pixel to the left
                items = self.find_enclosed(self.coords_vportq1[0] - 1, -1, linf_x + 1, lsup_y + 1)
                for item in items:
                    self.move(item, -dx, 0)
                viewport_x0 = x0
            self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
        elif deltax > 0:
            # Rigth displacement
            x1 = x
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            if deltax - winfo_width > winfo_width:
                # We need to move the viewport to the right a distance such that any information
                # displayed in the viewport is lost.
                items = self.find_enclosed(self.coords_vportq1[0] - 1, -1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_x0 = x1
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
                dx = 1
            else:
                # Remove items that falls beyond the right edge
                dx = deltax
                items = self.find_enclosed(self.coords_vportq1[0] - 1, -1, linf_x + 1, lsup_y + 1)
                self.delete(*items)
                # Move the viewport dx pixel to the left
                items = self.find_enclosed(linf_x - 1, -1, lsup_x + 1, lsup_y + 1)
                for item in items:
                    self.move(item, -dx, 0)
                viewport_x0 = self.cell_containing_coords(linf_x + 1, 0)[0]
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
            pass

        if deltay < 0:
            y0 = y
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            if deltay + winfo_height < self.coords_vportq1[1]:
                # We need to move the viewport to the top a distance such that any information
                # displayed in the viewport is lost.
                items = self.find_enclosed(-1, self.coords_vportq1[1] - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_y0 = y0
                dy = 1  # Needed to indicate redraw
            else:
                # Remove items that falls below the bottom edge
                dy = deltay
                dmy, viewport_y1 = self.cell_containing_coords(0, winfo_height + dy)
                linf_y = self.cell_coordinates(dmy, viewport_y1)[3]

                items = self.find_enclosed(-1, linf_y - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)

                # Move the viewport dy pixel up
                items = self.find_enclosed(-1, self.coords_vportq1[1] - 1,  lsup_x -dx + 1, linf_y + 1)
                for item in items:
                    self.move(item, 0, -dy)
                viewport_y0 = y0
            self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
        elif deltay > 0:
            y1 = y
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            if deltay - winfo_height > winfo_height:
                # We need to move the viewport to the bottom a distance such that any information
                # displayed in the viewport is lost.
                items = self.find_enclosed(-1, self.coords_vportq1[1] - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_y0 = y1
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
                dy = 1
            else:
                # Remove items that falls below the bottom edge
                dy = deltay
                items = self.find_enclosed(-1, self.coords_vportq1[1] - 1, lsup_x + 1, linf_y + 1)
                self.delete(*items)
                # Move the viewport dy pixel up
                items = self.find_enclosed(-1, linf_y - 1, lsup_x + 1, lsup_y + 1)
                for item in items:
                    self.move(item, 0, -dy)
                viewport_y0 = self.cell_containing_coords(0, linf_y + 1)[1]
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates

        if dx or dy:
            if dx:
                viewport_x1 = self.cell_containing_coords(winfo_width, 0)[0]
                pass
            if dy:
                viewport_y1 = self.cell_containing_coords(0, winfo_height)[1]
                pass
            self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
            self.setGUI()  # Redraw the sheet with the new viewport
            pass
        self.set_active_cell()

    def show_cell(self, xin, yin):
        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.viewport_q1
        if xin >= viewport_x1:
            xright = self.cell_coordinates(xin, 0)[2]
            x = self.cell_containing_coords(xright - (winfo_width - self.coords_vportq1[0]), 0)[0]
            xright = self.cell_coordinates(x, 0)[2]
            viewport_x0 = self.cell_containing_coords(xright + 1, 0)[0]
        elif xin < viewport_x0:
            viewport_x0 = xin
        
        if yin >= viewport_y1:
            ybottom = self.cell_coordinates(0, yin)[3]
            y = self.cell_containing_coords(0, ybottom - (winfo_height - self.coords_vportq1[1]))[1]
            ybottom = self.cell_coordinates(0, y)[3]
            viewport_y0 = self.cell_containing_coords(0, ybottom + 1)[1]
        elif yin < viewport_y0:
            viewport_y0 = yin
        self.move_viewport(viewport_x0, viewport_y0)
        pass
    
    def mouse_click(self, event):
        """Sets the active cell based on the click position."""
        if event.x < COL_CELLS_WIDTH or event.y < ROW_CELLS_HEIGHT:
            if event.x < COL_CELLS_WIDTH and event.y < ROW_CELLS_HEIGHT:
                test = 'show_cell'
                if test == 'show_cell':
                    msg1 = "Enter the pivot cell coordinates (col, row):"
                elif test == "move_viewport":
                    msg1 = "Enter the pivot cell coordinates (delta_col, delta_row):"
                fnc = getattr(self, test)
                # display a message box to get  the pivot cell coordinates
                answ = simpledialog.askstring(test, msg1, parent=self)
                if answ:
                    try:
                        x, y = map(lambda w: int(w), answ.split(","))
                        fnc(x, y)
                    except ValueError:
                        print("Invalid input. Please enter parameters as integers separated with commas in the format 'x, y'.")
                else:
                    print("No input provided.")
            elif event.x < COL_CELLS_WIDTH:
                print("Clicked on the column header")
            else:
                print("Clicked on the row header")
            return "break"
        clk_x, clk_y = self.cell_containing_coords(event.x, event.y)
        with self.pivot_point(isActiveCell=not event.state & SHIFT_PRESSED) as pivot:
            pivot.x = clk_x
            pivot.y = clk_y
        self.set_active_cell()
        self.focus_set()  # Set focus to the canvas

    def mouse_drag(self, event):
        if event.x > COL_CELLS_WIDTH and event.y < ROW_CELLS_HEIGHT:
            self.yview('scroll', '-1', 'units')
            clk_x, clk_y = self.cell_containing_coords(event.x, ROW_CELLS_HEIGHT + 1)
            with self.pivot_point(isActiveCell=False) as pivot:
                pivot.x = clk_x
            self.set_active_cell()
        elif event.x > COL_CELLS_WIDTH and event.y > ROW_CELLS_HEIGHT:
            clk_x, clk_y = self.cell_containing_coords(event.x, event.y)
            with self.pivot_point(isActiveCell=False) as pivot:
                pivot.x = clk_x
                pivot.y = clk_y
            self.show_cell(clk_x, clk_y)
            self.set_active_cell()
        elif event.x < COL_CELLS_WIDTH and event.y > ROW_CELLS_HEIGHT:
            self.xview('scroll', '-1', 'units')
            clk_x, clk_y = self.cell_containing_coords(COL_CELLS_WIDTH + 1, event.y)
            with self.pivot_point(isActiveCell=False) as pivot:
                pivot.y = clk_y
            self.set_active_cell()
        else:
            viewport_x0, viewport_y0 = self.viewport_q1[:2]
            self.move_viewport(viewport_x0 - 1, viewport_y0 - 1)
            with self.pivot_point(isActiveCell=False) as pivot:
                clk_x, clk_y = self.cell_containing_coords(COL_CELLS_WIDTH + 1,ROW_CELLS_HEIGHT + 1)
                pivot.x = clk_x
                pivot.y = clk_y
            self.set_active_cell()


    def mouse_release(self, event):
        pass
    
    def set_active_cell(self):
        # Set the tag "selected" for the region in coords (40, CELL_HEIGHT, 40 + 5*CELL_WIDTH, CELL_HEIGHT + 5*CELL_HEIGHT) rectangle
        def clip_rectangle(x0, y0, x1, y1, clipping_rgn=None):
            if clipping_rgn:
                linf_x, linf_y, lsup_x, lsup_y = clipping_rgn
                x0, x1 = min(lsup_x, max(linf_x, x0)), max(linf_x, min(x1, lsup_x))
                y0, y1 = min(lsup_y, max(linf_y, y0)), max(linf_y, min(y1, lsup_y))
            return x0, y0, x1, y1
        self.delete("selected_cells")
        clipping_rect = (self.coords_vportq1[0], self.coords_vportq1[1], self.winfo_width(), self.winfo_height())

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
        new_selected = [srow for srow in self.find_enclosed(-1, sel_y0 - 1, self.coords_vportq1[0] + 1, sel_y1 + 1) if self.type(srow) == "rectangle"]
        to_remove = set(old_selected) - set(new_selected)
        for row_id in to_remove:
            self.dtag(row_id, "row_selected")
            self.itemconfigure(row_id, fill="green")
        to_add = set(new_selected) - set(old_selected)
        for row_id in to_add:
            self.addtag_withtag("row_selected", row_id)
            self.itemconfigure(row_id, fill="blue")
        old_selected = self.find_withtag("col_selected")
        new_selected = [scol for scol in self.find_enclosed(sel_x0 - 1, -1, sel_x1 + 1, self.coords_vportq1[1] + 1) if self.type(scol) == "rectangle"]
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

    def yview(self, *args):
        if not args:
            return super().yview()
        elif args[0] == 'scroll':
            direction = args[2]
            if direction == 'units':
                delta = int(args[1])
                viewport_x0, viewport_y0 = self.viewport_q1[:2]
                self.move_viewport(viewport_x0, viewport_y0 + delta)
            elif direction == 'pages':
                delta = int(args[1])
                keysym = 'Prior' if delta < 0 else 'Next'
                self.event_generate(f'<{keysym}>', keysym=keysym)
        elif args[0] == 'moveto':
            super().yview_moveto(args[1])
        else:
            super().yview(*args)

    def xview(self, *args):
        if not args:
            return super().xview()
        elif args[0] == 'scroll':
            direction = args[2]
            if direction == 'units':
                delta = int(args[1])
                viewport_x0, viewport_y0 = self.viewport_q1[:2]
                self.move_viewport(viewport_x0 + delta, viewport_y0)
            elif direction == 'pages':
                delta = int(args[1])
                keysym = 'Prior' if delta < 0 else 'Next'
                self.event_generate(f'<{keysym}>', keysym=keysym, state=ALT_PRESSED)
        elif args[0] == 'moveto':
            super().xview_moveto(args[1])
        else:
            super().xview(*args)


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
                    #   scrollregion=(0, 0, 10000, 10000)
    )  # Adjust scrollregion as needed

    # Configure scrollbars to control the canvas
    v_scroll.config(command=sheetui.yview)
    h_scroll.config(command=sheetui.xview)

    # Layout
    sheetui.grid(row=0, column=0, sticky="nsew")
    v_scroll.grid(row=0, column=1, sticky="ns")
    h_scroll.grid(row=1, column=0, sticky="ew")

    root.mainloop()