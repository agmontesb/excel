''' Este módulo implementa la interfaz gráfica de usuario (GUI) para la hoja de cálculo.
    Proporciona una ventana principal con una barra de menú, una barra de herramientas y un área de trabajo 
    que consiste en un área de celdas y una barra de estado que permite cambiar a otras hojas de trabajo '
    además de la activa'
'''
import re
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from contextlib import contextmanager
from types import SimpleNamespace

import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Constants for key states
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


def cell_content_gen(nquadrant: int, x: int, y: int) -> str:
    """Generates the content for a cell based on its quadrant and cell coordinates."""
    if nquadrant == 1:
        return f"C{x}R{y}"
    elif nquadrant == 2:
        return f"Q2_C{x}R{y}"
    elif nquadrant == 3:
        return f"Q3_C{x}R{y}"
    else:
        return f"Q4_C{x}R{y}"


class SheetUI(tk.Canvas):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.coords_vportq3 = (COL_CELLS_WIDTH, ROW_CELLS_HEIGHT)
        # self.viewport_q3 = (2, 3, 4, 6)
        self.viewport_q3 = (1, 1, 1, 1)
        # self.coords_vportq1 = (COL_CELLS_WIDTH + 2 * CELL_WIDTH, ROW_CELLS_HEIGHT + 3 * CELL_HEIGHT)
        self.coords_vportq1 = self.coords_vportq3
        # self.viewport_q1 = (4, 6, 4, 6)  # Default pivot cell
        self.viewport_q1 = (1, 1, 1, 1)  # Default pivot cell
        self.active_cell = self.viewport_q1[:2]  # Variable to store the active cell
        self.selected_cells = (*self.active_cell, *self.active_cell)  # Variable to store the selected cell
        self.cell_content = cell_content_gen

        #flags
        self.f_drag = False  # Flag to indicate if a mouse drag is in progress
        self.f_gridlines = True  # Flag to indicate if gridlines are visible
        self.f_headings = True  # Flag to indicate if headings are visible


        self.error_report = ''


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
        self.bind("<Home>", self.arrow_click)
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
            self.event_generate("<<ActiveCellChanged>>")
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
            self.event_generate("<<SelectedCellsChanged>>")

    def draw_cell_content(self, box: tuple[int, int, int, int], cell_content:str, **kwargs):
        x0, y0, x1, y1 = box
        items = [item for item in self.find_enclosed(x0, y0, x1, y1) if self.type(item) == "text"]
        if items:
            old_txt = self.itemcget(items[0], "text")
            print(f"replacing {old_txt} with {cell_content}")
            self.error_report += f" {old_txt}"
        self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=cell_content, **kwargs)


    def cell_coordinates(self, x, y, viewport=None, coords_viewport=None):
        """Calculates the coordinates of the cell based on the x and y position."""
        if viewport is None:
            coords_viewport, viewport = self.coords_vportq1, self.viewport_q1
        x0 = coords_viewport[0] + (x - viewport[0]) * CELL_WIDTH
        y0 = coords_viewport[1] + (y - viewport[1]) * CELL_HEIGHT
        return (x0, y0, x0 + CELL_WIDTH, y0 + CELL_HEIGHT)
    
    def area_coordinates(self, x0, y0, x1, y1):
        nquadrant = self.cell_quadrant(x0, y0, isCoord=False)
        orig, coords_orig = self.quadrant_data(nquadrant)
        sel_x0, sel_y0 = self.cell_coordinates(x0, y0, orig, coords_orig)[:2]
        nquadrant = self.cell_quadrant(x1, y1, isCoord=False)
        orig, coords_orig = self.quadrant_data(nquadrant)
        sel_x1, sel_y1 = self.cell_coordinates(x1, y1, orig, coords_orig)[2:]
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
    
    def cell_quadrant(self, x: int, y:int, isCoord: bool=True) -> int:
        """Returns the quadrant of the cell containing the given x and y screen coordinates."""
        xdiscr = self.coords_vportq1[0] if isCoord else self.viewport_q3[2]
        ydiscr = self.coords_vportq1[1] if isCoord else self.viewport_q3[3]
        if x >= xdiscr and y >= ydiscr:
            return 1
        if x >= xdiscr and y <= ydiscr:
            return 2
        if x < xdiscr and y < ydiscr:
            return 3
        return 4
    
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
            orig = self.viewport_q3[0], self.viewport_q1[1], self.viewport_q3[2], self.viewport_q1[3]
            return orig, (self.coords_vportq3[0], self.coords_vportq1[1])
    
    def setGUI(self):
        """Sets the GUI for the worksheet with a specified width and height."""

        lsup_coordx, lsup_coordy = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        lsup_coordx = min(self.winfo_width(), lsup_coordx)
        lsup_coordy = min(self.winfo_height(), lsup_coordy)
        self.delete("background")
        self.create_rectangle(0, 0, lsup_coordx, lsup_coordy, fill="white", outline=GRID_COLOR, tags="background")
        self.tag_lower("background")  # Ensure the background is at the bottom of the stack
        # Determine the areas that needs redraw
        items = sorted(self.find_withtag("column"), key=lambda x: self.coords(x)[0])
        dmy0 = dmy1 = self.coords_vportq3[0]
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
            xroot = (self.coords_vportq3[0] - COL_CELLS_WIDTH) if dmy1 >= lsup_coordx else dmy1
            width = lsup_coordx - xroot
        xroot_in = xroot

        items = sorted(self.find_withtag("row"), key=lambda x: self.coords(x)[1])
        dmy0 = dmy1 = self.coords_vportq3[1]
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
            yroot = (self.coords_vportq3[1] - ROW_CELLS_HEIGHT) if dmy1 >= lsup_coordy else dmy1
            height = lsup_coordy - yroot
        yroot_in = yroot

        # We secure that the parameters are all integers
        xroot, width, yroot, height = map(int, (xroot, width, yroot, height))
        linf_x, linf_y, lsup_x, lsup_y = self.viewport_q1
        viewport_x0, viewport_y0, xcell, ycell = self.viewport_q1
        if xroot >= self.coords_vportq3[0]:
            # winfo_height = self.winfo_height()
            if xroot == self.coords_vportq3[0]:
                pane_width = self.coords_vportq1[0] - self.coords_vportq3[0]
                task = [
                    (self.coords_vportq1[0], width - pane_width, None, None),
                    (self.coords_vportq3[0], pane_width, self.viewport_q3, self.coords_vportq3),
                ]
            else:
                task = [(xroot, width, None, None)]
            while task:
                root_x, width, orig, coords_orig = task.pop()
                x1 = root_x
                xcell = linf_x = self.cell_containing_coords(x1, 0, orig, coords_orig)[0]
                while (x1 < root_x + width) and xcell <= MAX_COLS:
                    x0, _, x1, _ = self.cell_coordinates(xcell, 0, orig, coords_orig)
                    y0, y1 = self.coords_vportq3[1] - ROW_CELLS_HEIGHT, self.coords_vportq3[1]
                    self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="column")
                    # Draw cell headings
                    label = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{xcell}", fill="white")
                    self.addtag_withtag("columns_tag", label)  # Add tag for columns
                    # Draw vertical lines
                    self.create_line(x1, y1, x1, self.winfo_height(), fill=GRID_COLOR, tags="grid_lines")
                    xcell += 1
                lsup_x = xcell - 1
            print(f"printed columns from {linf_x} to {lsup_x}")

        if yroot >= self.coords_vportq3[1]:
            # winfo_width = self.winfo_width()
            if yroot == self.coords_vportq3[1]:
                pane_height = self.coords_vportq1[1] - self.coords_vportq3[1]
                task = [
                    (self.coords_vportq1[1], height - pane_height, None, None),
                    (self.coords_vportq3[1], pane_height, self.viewport_q3, self.coords_vportq3),
                ]
            else:
                task = [(yroot, height, None, None)]

            while task:
                root_y, height, orig, coords_orig = task.pop()
                y1 = root_y
                ycell = linf_y = self.cell_containing_coords(0, y1, orig, coords_orig)[1]
                while (y1 < root_y + height) and ycell <= MAX_ROWS:
                    x0, x1 = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[0]
                    _, y0, _, y1 = self.cell_coordinates(0, ycell, orig, coords_orig)
                    self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="row")
                    # Draw cell headings
                    label = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"R{ycell}", fill="white")
                    self.addtag_withtag("rows_tag", label)  # Add tag for columns
                    # Draw Horizontal lines
                    self.create_line(x1, y1, self.winfo_width(), y1, fill=GRID_COLOR, tags="grid_lines")
                    ycell += 1
                lsup_y = ycell - 1
            print(f"printed rows from {linf_y} to {lsup_y}")

        viewport_x0, viewport_y0 = self.cell_containing_coords(*self.coords_vportq1)
        viewport_x1, viewport_y1 = self.cell_containing_coords(self.winfo_width(), self.winfo_height())


        # Draw the cells content
        print(f"{xroot_in=}, {yroot_in=}")
        if xroot > self.coords_vportq1[0]:
            # Segundo cuadrante
            if yroot <= self.coords_vportq1[1]:
                orig, coords_orig =  self.quadrant_data(2)
                orig_x0, orig_y0, orig_x1, orig_y1 = orig
                if orig_y0 < orig_y1:
                    xmin, xmax = max(linf_x, orig_x0), min(lsup_x, orig_x1)
                    for x in range(xmin, xmax + 1):
                        for y in range(orig_y0, orig_y1):
                            x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                            # Draw cell content (placeholder text)
                            cell_content = self.cell_content(2, x, y)
                            self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                    print(f"(xroot > {self.coords_vportq1[0]}) printed cells from {(xmin, xmax)} to {(orig_y0, orig_y1)} in quadrant 2")

            # Primer cuadrante
            orig, coords_orig =  self.quadrant_data(1)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            xmin, xmax = max(linf_x, orig_x0), min(lsup_x, orig_x1)
            for x in range(xmin, xmax + 1):
                for y in range(linf_y, lsup_y + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                    # Draw cell content (placeholder text)
                    cell_content = self.cell_content(1, x, y)
                    self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
            print(f"(xroot > {self.coords_vportq1[0]}) printed cells from {(xmin, xmax)} to {(linf_y, lsup_y)} in quadrant 1")
            
        if xroot == self.coords_vportq1[0]:
            if yroot >= self.coords_vportq1[1]:
                # Cuarto cuadrante
                orig, coords_orig =  self.quadrant_data(4)
                orig_x0, orig_y0, orig_x1, orig_y1 = orig
                ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
                if orig_x1 > orig_x0:
                    for y in range(ymin, ymax + 1):
                        for x in range(orig_x0, orig_x1):
                            x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                            # Draw cell content (placeholder text)
                            cell_content = self.cell_content(4, x, y)
                            self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                    print(f"(xroot == {self.coords_vportq1[0]}) printed cells from {(orig_x0, orig_x1 - 1)} to {(ymin, ymax)} in quadrant 4")
            # Primer cuadrante
            orig, coords_orig =  self.quadrant_data(1)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
            for y in range(ymin, ymax + 1):
                for x in range(linf_x, lsup_x + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y)
                    # Draw cell content (placeholder text)
                    cell_content = self.cell_content(1, x, y)
                    self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
            print(f"(xroot == {self.coords_vportq1[0]})printed cells from {(linf_x, lsup_x)} to {(ymin, ymax)} in quadrant 1")

        if xroot_in > self.viewport_q3[0] - COL_CELLS_WIDTH and yroot_in > self.viewport_q3[1] - ROW_CELLS_HEIGHT:
            if (xroot_in, yroot_in) == self.coords_vportq3:
                # Tercer cuadrante
                orig, coords_orig =  self.quadrant_data(3)
                if orig[0] < orig[2] and orig[1] < orig[3]:
                    for y in range(orig[1], orig[3]):
                        for x in range(orig[0], orig[2]):
                            x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                            # Draw cell content (placeholder text)
                            cell_content = self.cell_content(3, x, y)
                            self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                    print(f"(xroot > 0 and yroot > 0)printed cells from {(orig[0], orig[2] - 1)} to {(orig[1], orig[3] - 1)} in quadrant 3")
            # Cuarto cuadrante
            orig, coords_orig =  self.quadrant_data(4)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
            if orig_x1 > orig_x0:
                for y in range(ymin, ymax + 1):
                    for x in range(orig_x0, orig_x1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        cell_content = self.cell_content(4, x, y)
                        self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                print(f"(xroot > 0 and yroot > 0) printed cells from {(orig_x0, orig_x1 - 1)} to {(ymin, ymax)} in quadrant 4")
            # Primer cuadrante
            if linf_x > viewport_x0:
                for x in range(viewport_x0, linf_x):
                    for y in range(ymin, ymax + 1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y)
                        # Draw cell content (placeholder text)
                        cell_content = self.cell_content(1, x, y)
                        self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                print(f"(xroot > 0 and yroot > 0) printed cells from {(viewport_x0, linf_x - 1)} to {(ymin, ymax)} in quadrant 1")
            if lsup_x < viewport_x1:
                for x in range(lsup_x + 1, viewport_x1 + 1):
                    for y in range(ymin, ymax + 1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y)
                        # Draw cell content (placeholder text)
                        cell_content = self.cell_content(1, x, y)
                        self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                print(f"(xroot > 0 and yroot > 0)printed cells from {(lsup_x + 1, viewport_x1)} to {(ymin, ymax)} in quadrant 1")

        # Update the viewport
        self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)
        self.set_freeze_lines()
        if not self.f_gridlines:
            self.tag_lower("grid_lines")
        if self.error_report:
            self.event_generate("<<errorReport>>")

    def set_freeze_lines(self):
        coord_acell_x, coord_acell_y = self.coords_vportq1
        items = self.find_withtag("freeze_line")
        if not items and coord_acell_y != self.coords_vportq3[1]:
            self.create_line(-COL_CELLS_WIDTH, coord_acell_y, self.winfo_width(), coord_acell_y, fill="black", tags="freeze_line")
        
        if not items and coord_acell_x != self.coords_vportq3[0]:
            self.create_line(coord_acell_x, -ROW_CELLS_HEIGHT, coord_acell_x, self.winfo_height(), fill="black", tags="freeze_line")


    def redraw_sheet(self, event):
        "Redraws the sheetui when the window is resized or needs updating."
        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        viewport_x1, viewport_y1 = self.cell_containing_coords(winfo_width, winfo_height)
        self.viewport_q1 = self.viewport_q1[:2] + (viewport_x1, viewport_y1)
        self.configure(scrollregion=(0, 0, MAX_COLS * CELL_WIDTH, MAX_ROWS * CELL_HEIGHT))
        self.delete("all")
        x0, x1 = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[0]
        y0, y1 = self.coords_vportq3[1] - ROW_CELLS_HEIGHT, self.coords_vportq3[1]
        self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="corner")
        self.toggle_headings()

        # self.setGUI()
        self.set_active_cell()

    def arrow_click(self, event):
        """Sets the active cell based on the arrow key pressed."""
        # print(f'{event.keysym} pressed')

        if event.keysym == "Home":
            isCtrlPressed = event.state & CTRL_PRESSED
            isShiftPressed = event.state & SHIFT_PRESSED
            viewport_x0, viewport_y0 = self.viewport_q1[:2]
            with self.pivot_point(isActiveCell=not isShiftPressed) as pivot:
                pivot.x = viewport_x0 = self.viewport_q3[2]
                if isCtrlPressed:
                    viewport_y0 = self.viewport_q3[3]
                    pivot.y = viewport_y0 
            self.move_viewport(viewport_x0, viewport_y0)
            self.set_active_cell()
            return "break"
        elif event.keysym in ("Next", "Prior"):
            viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.viewport_q1
            winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
            acell_x0, acell_y0 = self.active_cell
            with self.pivot_point(isActiveCell=not event.state & SHIFT_PRESSED) as pivot:
                coord_pivot_x, coord_pivot_y = self.cell_coordinates(pivot.x, pivot.y)[:2]
                nquadrant = self.cell_quadrant(pivot.x, pivot.y, isCoord=False)
                if event.state & ALT_PRESSED:
                    if nquadrant in (4, 3):
                        coord_pivot_x = self.coords_vportq1[0]
                    if event.keysym == "Next":
                        viewport_x0 = viewport_x1 if viewport_x1 < MAX_COLS else viewport_x0
                    else:
                        xright = self.cell_coordinates(viewport_x0, 0)[2]
                        viewport_x0 = self.cell_containing_coords(xright - (winfo_width - COL_CELLS_WIDTH), 0)[0]
                        viewport_x0 = min(MAX_COLS, max(1, viewport_x0))
                else:
                    if nquadrant in (2, 3):
                        coord_pivot_y = self.coords_vportq1[1]
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
            nquadrant = self.cell_quadrant(pivot.x, pivot.y, isCoord=False)
            # linf_x, linf_y = (self.viewport_q3[2], self.viewport_q3[3]) if nquadrant == 1 else (self.viewport_q3[0], self.viewport_q3[1])
            linf_x, linf_y = 1, 1
            pivot.x = max(linf_x, min(MAX_COLS, pivot.x + dx))
            pivot.y = max(linf_y, min(MAX_ROWS, pivot.y + dy))
            xin, yin = pivot.x, pivot.y
        orig = self.quadrant_data(3)[0]
        nquadrant = self.cell_quadrant(xin, yin, isCoord=False)
        if (xin >= orig[0] and yin >= orig[1]) and nquadrant != 3:
            if nquadrant == 2:
                yin = self.viewport_q1[1]
            elif nquadrant == 4:
                xin = self.viewport_q1[0]
            self.show_cell(xin, yin)
        self.set_active_cell()

    def move_viewport(self, x, y):
        """
        Move the viewport origen (deltax, deltay) pixels
        """
        viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.viewport_q1
        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        x = max(1, min(MAX_COLS, x))
        y = max(1, min(MAX_ROWS, y))
        linf_x, linf_y = self.cell_coordinates(x, y)[:2]
        deltax, deltay = linf_x - self.coords_vportq1[0], linf_y - self.coords_vportq1[1]
        dx = dy = 0
        clinf_y = self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        clinf_x = self.coords_vportq3[0] - COL_CELLS_WIDTH
        if deltax < 0:
            # left displacement
            x0 = x
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            if deltax + winfo_width < self.coords_vportq1[0]:
                # We need to move the viewport to the left a distance such that any information
                # displayed in the viewport is lost.
                items = self.find_enclosed(self.coords_vportq1[0] - 1, clinf_y - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_x0 = x0
                dx = 1 # Needed to indicate redraw
            else:
                # Remove items that falls beyond the right edge
                dx = deltax
                viewport_x1, dmy = self.cell_containing_coords(winfo_width + dx, 0)
                linf_x = self.cell_coordinates(viewport_x1, dmy)[2]

                items = self.find_enclosed(linf_x - 1, clinf_y - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)

                # Move the viewport dx pixel to the left
                items = self.find_enclosed(self.coords_vportq1[0] - 1, clinf_y - 1, linf_x + 1, lsup_y + 1)
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
                items = self.find_enclosed(self.coords_vportq1[0] - 1, clinf_y - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_x0 = x1
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
                dx = 1
            else:
                # Remove items that falls beyond the right edge
                dx = deltax
                items = self.find_enclosed(self.coords_vportq1[0] - 1, clinf_y - 1, linf_x + 1, lsup_y + 1)
                self.delete(*items)
                # Move the viewport dx pixel to the left
                items = self.find_enclosed(linf_x - 1, clinf_y - 1, lsup_x + 1, lsup_y + 1)
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
                items = self.find_enclosed(clinf_x - 1, self.coords_vportq1[1] - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_y0 = y0
                dy = 1  # Needed to indicate redraw
            else:
                # Remove items that falls below the bottom edge
                dy = deltay
                dmy, viewport_y1 = self.cell_containing_coords(0, winfo_height + dy)
                linf_y = self.cell_coordinates(dmy, viewport_y1)[3]

                items = self.find_enclosed(clinf_x, linf_y - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)

                # Move the viewport dy pixel up
                items = self.find_enclosed(clinf_x, self.coords_vportq1[1] - 1,  lsup_x -dx + 1, linf_y + 1)
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
                items = self.find_enclosed(clinf_x, self.coords_vportq1[1] - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_y0 = y1
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
                dy = 1
            else:
                # Remove items that falls below the bottom edge
                dy = deltay
                items = self.find_enclosed(clinf_x, self.coords_vportq1[1] - 1, lsup_x + 1, linf_y + 1)
                self.delete(*items)
                # Move the viewport dy pixel up
                items = self.find_enclosed(clinf_x - 1, linf_y - 1, lsup_x + 1, lsup_y + 1)
                for item in items:
                    self.move(item, 0, -dy)
                viewport_y0 = self.cell_containing_coords(0, linf_y + 1)[1]
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
        # dx = dy = 0
        if dx or dy:
            if dx:
                viewport_x1 = self.cell_containing_coords(winfo_width, 0)[0]
                pass
            if dy:
                viewport_y1 = self.cell_containing_coords(0, winfo_height)[1]
                pass
            self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
            self.setGUI()  # Redraw the sheet with the new viewport
            self.tag_raise("freeze_line")  # Move freeze_line above all tags
            pass

    def show_cell(self, xin, yin):
        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.viewport_q1
        lsup_coordx, lsup_coordy = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
        if lsup_coordx > winfo_width and xin >= viewport_x1:
            xright = self.cell_coordinates(xin, 0)[2]
            x = self.cell_containing_coords(xright - (winfo_width - self.coords_vportq1[0]), 0)[0]
            xright = self.cell_coordinates(x, 0)[2]
            viewport_x0 = self.cell_containing_coords(xright + 1, 0)[0]
        elif xin < viewport_x0:
            viewport_x0 = xin
        
        if  lsup_coordy > winfo_height and yin >= viewport_y1:
            ybottom = self.cell_coordinates(0, yin)[3]
            y = self.cell_containing_coords(0, ybottom - (winfo_height - self.coords_vportq1[1]))[1]
            ybottom = self.cell_coordinates(0, y)[3]
            viewport_y0 = self.cell_containing_coords(0, ybottom + 1)[1]
        elif yin < viewport_y0:
            viewport_y0 = yin
        self.move_viewport(viewport_x0, viewport_y0)
        pass

    def toggle_headings(self, *args):
        """Toggles the visibility of headings."""
        if self.f_headings:
            dx, dy = linf_x, linf_y = COL_CELLS_WIDTH, ROW_CELLS_HEIGHT
            lsup_x, lsup_y = self.cell_coordinates(self.viewport_q1[2], self.viewport_q1[3])[2:]
            self.coords_vportq1 = (self.coords_vportq1[0] - dx, self.coords_vportq1[1] - dy)
            self.coords_vportq3 = (0, 0)
            viewport_x1, viewport_y1 = self.cell_containing_coords(self.winfo_width(), self.winfo_height())
            self.viewport_q1 = (self.viewport_q1[0], self.viewport_q1[1], viewport_x1, viewport_y1)
            items = self.find_enclosed(-1, -1, lsup_x + 1, lsup_y + 1)
            for item in items:
                # Move items to the left and up
                self.move(item, -dx, -dy)
            self.setGUI()
        else:
            dx, dy = linf_x, linf_y = COL_CELLS_WIDTH, ROW_CELLS_HEIGHT
            lsup_x, lsup_y = self.cell_coordinates(self.viewport_q1[2], self.viewport_q1[3])[2:]
            self.coords_vportq1 = (self.coords_vportq1[0] + dx, self.coords_vportq1[1] + dy)
            self.coords_vportq3 = (dx, dy)
            viewport_x1, viewport_y1 = self.cell_containing_coords(self.winfo_width(), self.winfo_height())
            self.viewport_q1 = (self.coords_vportq1[0], self.coords_vportq1[1], viewport_x1, viewport_y1)
            items = self.find_enclosed(-linf_x - 1, -linf_y - 1, lsup_x + 1, lsup_y + 1)
            for item in items:
                self.move(item, dx, dy)
        self.f_headings = not self.f_headings

    def toggle_gridlines(self, *args):
        """Toggles the visibility of gridlines."""
        if self.f_gridlines:
            self.tag_lower("grid_lines")
        else:
            self.tag_raise("grid_lines")
        self.f_gridlines = not self.f_gridlines
    
    def mouse_click(self, event):
        """Sets the active cell based on the click position."""
        self.f_drag = True
        if event.x < COL_CELLS_WIDTH or event.y < ROW_CELLS_HEIGHT:
            if event.x < COL_CELLS_WIDTH and event.y < ROW_CELLS_HEIGHT:
                test = 'toggle_headings'
                if test == 'toggle_headings':
                    msg1 = "Toggle headings (yes/no):"
                elif test == 'toggle_gridlines':
                    msg1 = "Toggle gridlines (yes/no):"
                elif test == 'freeze_panes':
                    msg1 = "Enter the pivot cell coordinates (col, row):"
                elif test == 'show_cell':
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
        # Check if the click is not on an existing cell
        items = self.find_overlapping(event.x, event.y, event.x, event.y)
        if not items:
            return "break"
        nquadrant = self.cell_quadrant(event.x, event.y)
        orig, coords_orig = self.quadrant_data(nquadrant)
        clk_x, clk_y = self.cell_containing_coords(event.x, event.y, orig, coords_orig)
        with self.pivot_point(isActiveCell=not event.state & SHIFT_PRESSED) as pivot:
            pivot.x = clk_x
            pivot.y = clk_y
        self.set_active_cell()
        self.focus_set()  # Set focus to the canvas

    def mouse_drag(self, event):
        """Handles mouse drag events to set the active cell."""
        if self.f_drag:
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
        self.f_drag = False
        pass

    def freeze_panes(self, *args):
        if self.viewport_q3 == (1, 1, 1, 1):
            self.coords_vportq1 = coord_acell_x, coord_acell_y = self.cell_coordinates(*self.active_cell)[:2]
            self.viewport_q3 = (*self.viewport_q1[:2], *self.active_cell)
            self.viewport_q1 = *self.active_cell, *self.viewport_q1[2:]

            self.set_freeze_lines()
        else:
            self.move_viewport(*self.viewport_q3[2:])
            self.coords_vportq1 = self.coords_vportq3
            self.viewport_q1 = *self.viewport_q3[:2], *self.viewport_q1[2:]
            self.viewport_q3 = 1, 1, 1, 1
            items = self.find_withtag("freeze_line")
            self.delete(*items)

    def set_active_cell(self):
        # Set the tag "selected" for the region in coords (40, CELL_HEIGHT, 40 + 5*CELL_WIDTH, CELL_HEIGHT + 5*CELL_HEIGHT) rectangle
        def clip_rectangle(x0, y0, x1, y1, clipping_rgn=None):
            if clipping_rgn:
                linf_x, linf_y, lsup_x, lsup_y = clipping_rgn
                x0, x1 = min(lsup_x, max(linf_x, x0)), max(linf_x, min(x1, lsup_x))
                y0, y1 = min(lsup_y, max(linf_y, y0)), max(linf_y, min(y1, lsup_y))
            return x0, y0, x1, y1
        self.delete("selected_cells")
        clipping_rect = (self.coords_vportq3[0], self.coords_vportq3[1], self.winfo_width(), self.winfo_height())

        x0, y0, x1, y1 = self.area_coordinates(*self.selected_cells)
        sel_x0, sel_y0, sel_x1, sel_y1 = clip_rectangle(x0, y0, x1, y1, clipping_rect)
    
        self.create_rectangle(sel_x0, sel_y0, sel_x1, sel_y1,
            fill="lightblue", outline="black", tags="selected_cells")
        self.tag_lower("selected_cells", "grid_lines")


        self.delete("active_cell")
        """Draws the active cell rectangle."""
        nquadrant = self.cell_quadrant(*self.active_cell, isCoord=False)
        orig, coords_orig = self.quadrant_data(nquadrant)
        x0, y0, x1, y1 = self.cell_coordinates(*self.active_cell, orig, coords_orig)
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
        x0, x1 = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[0]
        new_selected = [srow for srow in self.find_enclosed(x0 - 1, sel_y0 - 1, x1 + 1, sel_y1 + 1) if self.type(srow) == "rectangle"]
        to_remove = set(old_selected) - set(new_selected)
        for row_id in to_remove:
            self.dtag(row_id, "row_selected")
            self.itemconfigure(row_id, fill="green")
        to_add = set(new_selected) - set(old_selected)
        for row_id in to_add:
            self.addtag_withtag("row_selected", row_id)
            self.itemconfigure(row_id, fill="blue")
        old_selected = self.find_withtag("col_selected")
        y0, y1 = self.coords_vportq3[1] - ROW_CELLS_HEIGHT, self.coords_vportq3[1]
        new_selected = [scol for scol in self.find_enclosed(sel_x0 - 1, y0 - 1, sel_x1 + 1, y1 + 1) if self.type(scol) == "rectangle"]
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
    class SheetViewer(tk.Tk):
        def __init__(self):
            super().__init__()
            self.setGui()
            self.bind("<<ActiveCellChanged>>", self.on_active_cell_changed)
            self.bind("<<SelectedCellsChanged>>", self.on_selected_cells_changed)
            self.bind("<<errorReport>>", self.on_error_report)

        def on_active_cell_changed(self, event):
            active_cell = event.widget.active_cell
            self.activeCell.config(text=f"Active Cell: {active_cell}")

        def on_selected_cells_changed(self, event):
            selected_cells = event.widget.selected_cells
            sel_x0, sel_y0, sel_x1, sel_y1 = selected_cells
            nrows = sel_y1 - sel_y0 + 1
            ncols = sel_x1 - sel_x0 + 1
            self.activeCell.config(text=f"Selected Cells: {nrows}R x {ncols}C ({sel_x0}, {sel_y0}) to ({sel_x1}, {sel_y1})")

        def on_error_report(self, event):
            widget: SheetUI = event.widget
            error_message = widget.error_report
            self.errorReport.config(text=error_message)
            event.widget.errorReport = ""  # Clear the error message after displaying it

        def setGui(self):
            self.columnconfigure(0, weight=1)
            self.rowconfigure(2, weight=1)  # Change to row 2 for the main frame

            # --- Add labels at the top ---
            self.activeCell = ttk.Label(self, text="Active Cell: ", background="magenta", font=("Arial", 10))
            self.activeCell.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))

            self.errorReport = ttk.Label(self, text="", background="grey", font=("Arial", 14), foreground="red", anchor="w")
            self.errorReport.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))

            # Create a frame to hold the canvas and scrollbars
            frame = ttk.Frame(self)
            frame.grid(column=0, row=2, sticky=(tk.N, tk.W, tk.E, tk.S))
            frame.columnconfigure(0, weight=1)
            frame.rowconfigure(0, weight=1)

            # Create vertical and horizontal scrollbars
            v_scroll = ttk.Scrollbar(frame, orient="vertical")
            h_scroll = ttk.Scrollbar(frame, orient="horizontal")

            # Create the SheetUI canvas
            self.sheetui = sheetui = SheetUI(frame, bg=GRID_COLOR, 
                                    yscrollcommand=v_scroll.set, 
                                    xscrollcommand=h_scroll.set,
                                    # scrollregion=(0, 0, 2000, 2000)
            )  # Adjust scrollregion as needed

                # Configure scrollbars to control the canvas
            v_scroll.config(command=sheetui.yview)
            h_scroll.config(command=sheetui.xview)

            # Layout
            sheetui.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")


    root = SheetViewer()
    root.mainloop()