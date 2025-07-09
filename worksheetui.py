''' Este módulo implementa la interfaz gráfica de usuario (GUI) para la hoja de cálculo.
    Proporciona una ventana principal con una barra de menú, una barra de herramientas y un área de trabajo 
    que consiste en un área de celdas y una barra de estado que permite cambiar a otras hojas de trabajo '
    además de la activa'
'''
import os
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import filedialog
from contextlib import contextmanager
from types import SimpleNamespace

import logging

logging.basicConfig(level=logging.DEBUG)
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
        self.viewport_q3 = (1, 1, 1, 1)
        self.coords_vportq1 = self.coords_vportq3
        self.viewport_q1 = (1, 1, 1, 1)                                 # Default pivot cell
        self.active_cell = self.viewport_q1[:2]                         # Variable to store the active cell     
        self.selected_cells = (*self.active_cell, *self.active_cell)    # Variable to store the selected cell
        self.cell_content = cell_content_gen

        #flags
        self.f_drag = False  # Flag to indicate if a mouse drag is in progress
        self.f_gridlines = True  # Flag to indicate if gridlines are visible
        self.f_headings = True  # Flag to indicate if headings are visible
        self.f_freeze = False  # Flag to indicate if freeze panes are visible


        self.error_report = ''


        self.bind("<Configure>", self.redraw_sheet)
        self.bind("<Button-1>", self.mouse_click)
        self.bind("<B1-Motion>", self.mouse_drag)
        self.bind("<ButtonRelease-1>", self.mouse_release)
        self.bind("<MouseWheel>", self.on_mouse_wheel)  # Windows/macOS
        self.bind("<Button-4>", self.on_mouse_wheel)    # Linux scroll up
        self.bind("<Button-5>", self.on_mouse_wheel)    # Linux scroll down

        # bind arrow keys to move the active cell
        self.bind("<Up>", self.arrow_click)
        self.bind("<Down>", self.arrow_click)
        self.bind("<Left>", self.arrow_click)
        self.bind("<Right>", self.arrow_click)
        self.bind("<Return>", self.arrow_click)
        self.bind("<Tab>", self.arrow_click)
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
            logging.debug(f"replacing {old_txt} with {cell_content}")
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
    
    def area_cells(self, sel_x0, sel_y0, sel_x1, sel_y1):
        nquadrant = self.cell_quadrant(sel_x0, sel_y0, isCoord=True)
        orig, coords_orig = self.quadrant_data(nquadrant)
        x0, y0 = self.cell_containing_coords(sel_x0, sel_y0, orig, coords_orig)
        nquadrant = self.cell_quadrant(sel_x1 - 1, sel_y1 - 1, isCoord=True)
        orig, coords_orig = self.quadrant_data(nquadrant)
        x1, y1 = self.cell_containing_coords(sel_x1 - 1, sel_y1 - 1, orig, coords_orig)
        return (x0, y0, x1, y1)
    
    def cell_containing_coords(self, x, y, viewport=None, coords_viewport=None):
        """Returns the cell address containing the given x and y screen coordinates."""
        if viewport is None:
            coords_viewport, viewport = self.coords_vportq1, self.viewport_q1
        xcell = int((x - coords_viewport[0]) // CELL_WIDTH + viewport[0])
        ycell = int((y - coords_viewport[1]) // CELL_HEIGHT + viewport[1])
        xcell = max(1, min(MAX_COLS, xcell))
        ycell = max(1, min(MAX_ROWS, ycell))
        return (xcell, ycell)
    
    def cell_quadrant(self, x: int, y:int, isCoord: bool=True) -> int:
        """Returns the quadrant of the cell containing the given x and y screen coordinates."""
        if self.coords_vportq1 == self.viewport_q3:
            return 1
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
            orig = self.viewport_q1[0], self.viewport_q3[1], self.viewport_q1[2], self.viewport_q3[3] - 1
            return orig, (self.coords_vportq1[0], self.coords_vportq3[1])
        elif nquadrant == 3:
            return (*self.viewport_q3[:2], self.viewport_q3[2] - 1, self.viewport_q3[3] - 1), self.coords_vportq3
        else: # nquadrant == 4
            orig = self.viewport_q3[0], self.viewport_q1[1], self.viewport_q3[2] - 1, self.viewport_q1[3]
            return orig, (self.coords_vportq3[0], self.coords_vportq1[1])
        
    def setGUI(self):
        winfo_width, winfo_height = self.winfo_width() + int(self.f_headings) * COL_CELLS_WIDTH, self.winfo_height() + int(self.f_headings) * ROW_CELLS_HEIGHT
        self.validate_areas()

        # Draw the background
        self.delete("background")
        linf_coordx, linf_coordy = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        lsup_coordx, lsup_coordy = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        self.create_rectangle(linf_coordx, linf_coordy, lsup_coordx, lsup_coordy, fill="white", outline=GRID_COLOR, tags="background")
        self.tag_lower("background")  # Ensure the background is at the bottom of the stack

        # Draw column headings
        for item in self.find_withtag("cols_to_draw"):
            cx0, cy0, cx1, cy1 = self.coords(item)
            xcell = 1
            while cx0 < cx1 and xcell < MAX_COLS:
                nquadrant = self.cell_quadrant(cx0, cy1, isCoord=True)
                orig, coords_orig = self.quadrant_data(nquadrant)
                xcell = self.cell_containing_coords(cx0, cy0, orig, coords_orig)[0]
                y0, y1 = cy0, cy1
                x0, x1 = self.cell_coordinates(xcell, 0, orig, coords_orig)[::2]
                self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="column")
                # Draw cell headings
                label = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{xcell}", fill="white")
                self.addtag_withtag("columns_tag", label)  # Add tag for columns
                # Draw vertical lines
                self.create_line(x0, y1, x0, winfo_height, fill=GRID_COLOR, tags="grid_lines")
                cx0 = x1
            logging.debug(f"Last column draw {xcell}")

        # Draw row headings
        for item in self.find_withtag("rows_to_draw"):
            cx0, cy0, cx1, cy1 = self.coords(item)
            ycell = 1
            while cy0 < cy1 and ycell < MAX_ROWS:
                nquadrant = self.cell_quadrant(cx1, cy0, isCoord=True)
                orig, coords_orig = self.quadrant_data(nquadrant)
                ycell = self.cell_containing_coords(cx0, cy0, orig, coords_orig)[1]
                x0, x1 = cx0, cx1
                y0, y1 = self.cell_coordinates(0, ycell, orig, coords_orig)[1::2]
                self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="row")
                # Draw cell headings
                label = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"R{ycell}", fill="white")
                self.addtag_withtag("rows_tag", label)  # Add tag for columns
                # Draw vertical lines
                self.create_line(x0, y1, winfo_width, y1, fill=GRID_COLOR, tags="grid_lines")
                cy0 = y1
            logging.debug(f"Last row draw {ycell}")

        # Draw cells content
        quadrants = [1, 2, 3, 4] if self.coords_vportq1 != self.coords_vportq3 else [1]
        for item in self.find_withtag("cells_to_draw"):
            ix0, iy0, ix1, iy1 = self.coords(item)
            for nquadrant in quadrants:
                orig, coords_orig = self.quadrant_data(nquadrant)
                ax0, ay0, ax1, ay1 = self.area_coordinates(*orig)
                # Overlaping area
                cx0, cy0 = max(ix0, ax0), max(iy0, ay0)
                cx1, cy1 = min(ix1, ax1), min(iy1, ay1)
                if not (cx1 > cx0 and cy1 > cy0):
                    continue
                while cx0 < cx1:
                    y0 = cy0
                    while y0 < cy1:
                        xcell, ycell = self.cell_containing_coords(cx0, y0, orig, coords_orig)
                        x0, y0, x1, y1 = self.cell_coordinates(xcell, ycell, orig, coords_orig)
                        cell_content = self.cell_content(nquadrant, xcell, ycell)
                        self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                        y0 = y1
                    cx0 = x1
        [self.tag_lower(tag) for tag in ("cols_to_draw", "rows_to_draw", "cells_to_draw")]
    
    def old_setGUI(self):
        """Sets the GUI for the worksheet with a specified width and height."""

        # Draw the background
        self.delete("background")
        linf_coordx, linf_coordy = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        lsup_coordx, lsup_coordy = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        self.create_rectangle(linf_coordx, linf_coordy, lsup_coordx, lsup_coordy, fill="white", outline=GRID_COLOR, tags="background")
        self.tag_lower("background")  # Ensure the background is at the bottom of the stack
        lsup_coordx = max(lsup_coordx, self.winfo_width() + int(self.f_headings) * COL_CELLS_WIDTH)
        lsup_coordy = max(lsup_coordy, self.winfo_height()+ int(self.f_headings) * ROW_CELLS_HEIGHT)

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
            xroot = linf_coordx if dmy1 >= lsup_coordx else dmy1
            width = lsup_coordx - xroot

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
            yroot = linf_coordy if dmy1 >= lsup_coordy else dmy1
            height = lsup_coordy - yroot

        # We secure that the parameters are all integers
        xroot, width, yroot, height = map(int, (xroot, width, yroot, height))
        # Se verifica que el xroot no exceda inicio de la última celda
        xflag = xroot <= self.cell_coordinates(MAX_COLS, 0)[0]
        # Se verifica que el yroot no exceda inicio de la última celda
        yflag = yroot <= self.cell_coordinates(0, MAX_ROWS)[1]

        if not xflag and not yflag:
            return

        # Draw the headings (rows, columns)
        viewport_x0, viewport_y0, xcell, ycell = self.viewport_q1
        if xroot >= self.coords_vportq3[0]:
            pane_width = self.coords_vportq1[0] - self.coords_vportq3[0]
            if xroot == self.coords_vportq3[0] and pane_width > 0:
                task = [
                    (self.coords_vportq1[0], width - pane_width, None, None),
                    (self.coords_vportq3[0], pane_width, self.viewport_q3, self.coords_vportq3),
                ]
                linf_x = self.viewport_q3[0]
            else:
                task = [(xroot, width, None, None)]
                linf_x = self.cell_containing_coords(xroot, 0)[0] if xflag else self.viewport_q1[0]
            while task:
                root_x, width, orig, coords_orig = task.pop()
                x1 = root_x
                xcell = self.cell_containing_coords(x1, 0, orig, coords_orig)[0]
                while (x1 < root_x + width) and xcell <= MAX_COLS:
                    x0, _, x1, _ = self.cell_coordinates(xcell, 0, orig, coords_orig)
                    y0, y1 = self.coords_vportq3[1] - ROW_CELLS_HEIGHT, self.coords_vportq3[1]
                    self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="column")
                    # Draw cell headings
                    label = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"C{xcell}", fill="white")
                    self.addtag_withtag("columns_tag", label)  # Add tag for columns
                    # Draw vertical lines
                    winfo_height = self.winfo_height() + int(self.f_headings) * ROW_CELLS_HEIGHT
                    self.create_line(x0, y1, x0, winfo_height, fill=GRID_COLOR, tags="grid_lines")
                    xcell += 1
                lsup_x = xcell - 1
            logging.debug(f"printed columns from {linf_x} to {lsup_x}")

        if yroot >= self.coords_vportq3[1]:
            pane_height = self.coords_vportq1[1] - self.coords_vportq3[1]
            if yroot == self.coords_vportq3[1] and pane_height > 0:
                task = [
                    (self.coords_vportq1[1], height - pane_height, None, None),
                    (self.coords_vportq3[1], pane_height, self.viewport_q3, self.coords_vportq3),
                ]
                linf_y = self.viewport_q3[1]
            else:
                task = [(yroot, height, None, None)]
                linf_y = self.cell_containing_coords(0, yroot)[1] if yflag else self.viewport_q1[1]
            while task:
                root_y, height, orig, coords_orig = task.pop()
                y1 = root_y
                ycell = self.cell_containing_coords(0, y1, orig, coords_orig)[1]
                while (y1 < root_y + height) and ycell <= MAX_ROWS:
                    x0, x1 = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[0]
                    _, y0, _, y1 = self.cell_coordinates(0, ycell, orig, coords_orig)
                    self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="row")
                    # Draw cell headings
                    label = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=f"R{ycell}", fill="white")
                    self.addtag_withtag("rows_tag", label)  # Add tag for columns
                    # Draw Horizontal lines
                    winfo_width = self.winfo_width() + int(self.f_headings) * COL_CELLS_WIDTH
                    self.create_line(x1, y0, winfo_width, y0, fill=GRID_COLOR, tags="grid_lines")
                    ycell += 1
                lsup_y = ycell - 1
            logging.debug(f"printed rows from {linf_y} to {lsup_y}")

        # Draw the cells content
        try:
            linf_x, lsup_x = max(linf_x, 1), min(lsup_x, MAX_COLS)
        except NameError:
            linf_x, lsup_x = self.viewport_q3[0], self.viewport_q1[2]

        try:
            linf_y, lsup_y = max(linf_y, 1), min(lsup_y, MAX_ROWS)
        except NameError:
            linf_y, lsup_y = self.viewport_q3[1], self.viewport_q1[3]

        # Segundo cuadrante
        if scnFlag := self.coords_vportq3[1] < self.coords_vportq1[1]:
            orig, coords_orig =  self.quadrant_data(2)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            xmin, xmax = max(linf_x, orig_x0), min(lsup_x, orig_x1)
            if (linf_x, lsup_x) != (orig_x0, orig_x1):
                ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
            else:
                ymin, ymax = orig_y0, orig_y1
            if scnFlag := ymin < ymax:
                for x in range(xmin, xmax + 1):
                    for y in range(ymin, ymax):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        cell_content = self.cell_content(2, x, y)
                        self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                logging.debug(f"(2cnd Quadrant) printed cells from {(xmin, xmax)} to {(ymin, ymax)} in quadrant 2")

        # Cuarto cuadrante
        if fthFlag := self.coords_vportq3[0] < self.coords_vportq1[0]:
            orig, coords_orig =  self.quadrant_data(4)
            orig_x0, orig_y0, orig_x1, orig_y1 = orig
            ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
            if (linf_y, lsup_y) != (orig_y0, orig_y1):
                xmin, xmax = max(linf_x, orig_x0), min(lsup_x, orig_x1)
            else:
                xmin, xmax = orig_x0, orig_x1
            if fthFlag := xmin < xmax:
                for x in range(xmin, xmax):
                    for y in range(ymin, ymax + 1):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        cell_content = self.cell_content(4, x, y)
                        self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                logging.debug(f"(4th Quadrant) printed cells from {(orig_x0, orig_x1 - 1)} to {(ymin, ymax)} in quadrant 4")

        # Tercer cuadrante
        if scnFlag and fthFlag:
            orig, coords_orig =  self.quadrant_data(3)
            xmin, xmax = max(linf_x, orig_x0), min(lsup_x, orig_x1)
            ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)
            if ymin < ymax:
                for y in range(ymin, ymax):
                    for x in range(xmin, xmax):
                        x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                        # Draw cell content (placeholder text)
                        cell_content = self.cell_content(3, x, y)
                        self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
                logging.debug(f"(3rd Quadrant) printed cells from {(orig[0], orig[2] - 1)} to {(orig[1], orig[3] - 1)} in quadrant 3")

        # Primer cuadrante
        orig, coords_orig =  self.quadrant_data(1)
        orig_x0, orig_y0, orig_x1, orig_y1 = orig
        xmin, xmax = max(linf_x, orig_x0), min(lsup_x, orig_x1)
        ymin, ymax = max(linf_y, orig_y0), min(lsup_y, orig_y1)

        if (xmin, xmax) != (orig_x0, orig_x1) or (xmin, ymin, xmax, ymax) == orig:
            for x in range(xmin, xmax + 1):
                for y in range(orig_y0, orig_y1 + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                    # Draw cell content (placeholder text)
                    cell_content = self.cell_content(1, x, y)
                    self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
            logging.debug(f"(Columns Only) printed cells from {(xmin, xmax)} to {(orig_y0, orig_y1)} in quadrant 1")

        if (xmin, xmax) == (orig_x0, orig_x1) and (ymin, ymax) != (orig_y0, orig_y1):
            for y in range(ymin, ymax + 1):
                for x in range(orig_x0, orig_x1 + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                    # Draw cell content (placeholder text)
                    cell_content = self.cell_content(1, x, y)
                    self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
            logging.debug(f"(Rows Only) printed cells from {(orig_x0, orig_x1)} to {(ymin, ymax)} in quadrant 1")

        if (ymin, ymax) != (orig_y0, orig_y1) and orig_x0 < linf_x:
            for x in range(orig_x0, linf_x):
                for y in range(ymin, ymax + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                    # Draw cell content (placeholder text)
                    cell_content = self.cell_content(1, x, y)
                    self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
            logging.debug(f"(left cells) printed cells from {(orig_x0, linf_x - 1)} to {(ymin, ymax)} in quadrant 1")

        if (ymin, ymax) != (orig_y0, orig_y1) and orig_x1 > lsup_x:
            for x in range(lsup_x + 1, orig_x1 + 1):
                for y in range(ymin, ymax + 1):
                    x0, y0, x1, y1 = self.cell_coordinates(x, y, orig, coords_orig)
                    # Draw cell content (placeholder text)
                    cell_content = self.cell_content(1, x, y)
                    self.draw_cell_content((x0, y0, x1, y1), cell_content, fill="black", tags="cell_content")
            logging.debug(f"(Right cells) printed cells from {(lsup_x + 1, orig_x1)} to {(ymin, ymax)} in quadrant 1")

        self.set_freeze_lines()
        if not self.f_gridlines:
            self.tag_lower("grid_lines")
        if self.error_report:
            self.event_generate("<<errorReport>>")

    def set_freeze_lines(self):
        coord_acell_x, coord_acell_y = self.coords_vportq1
        winfo_width, winfo_height = self.winfo_width() + int(self.f_headings) * COL_CELLS_WIDTH, self.winfo_height() + int(self.f_headings) * ROW_CELLS_HEIGHT
        items = self.find_withtag("freeze_line")
        if not items and coord_acell_y != self.coords_vportq3[1]:
            linf_x = self.coords_vportq3[0] - 2 * COL_CELLS_WIDTH
            self.create_line(linf_x, coord_acell_y, winfo_width, coord_acell_y, fill="black", tags="freeze_line")
        
        if not items and coord_acell_x != self.coords_vportq3[0]:
            linf_y = self.coords_vportq3[1] - 2 * ROW_CELLS_HEIGHT
            self.create_line(coord_acell_x, linf_y, coord_acell_x, winfo_height, fill="black", tags="freeze_line")

    def validate_areas(self):
        iareas = self.find_withtag("invalid_area")
        while iareas:
            item, *iareas = iareas
            cx0, cy0, cx1, cy1 = self.coords(item)
            self.delete(item)
            if cx0 < self.coords_vportq3[0]:
                # Rows to draw
                area = cx0, cy0, self.coords_vportq3[0], cy1
                self.tag_area(*area, tag="rows_to_draw")
                cx0 = self.coords_vportq3[0]
            if cy0 < self.coords_vportq3[1]:
                # Columns to draw
                area = cx0, cy0, cx1, self.coords_vportq3[1]
                self.tag_area(*area, tag="cols_to_draw")
                cy0 = self.coords_vportq3[1]
            if cx0 == cx1 or cy0 == cy1 or (cx0 == cy0 and cx1 == cy1):
                continue
            items = [item for item in self.find_overlapping(cx0, cy0, cx1, cy1) if "cell_to_draw" in self.gettags(item)]
            to_draw = [(cx0, cy0, cx1, cy1)]
            for item in items:
                ix0, iy0, ix1, iy1 = self.coords(item)
                n = len(to_draw)
                for i in range(n):
                    cx0, cy0, cx1, cy1 = to_draw[i]
                    # Overlaping area
                    x0 = max(cx0, ix0)
                    y0 = max(cy0, iy0)
                    x1 = min(cx1, ix1)
                    y1 = min(cy1, iy1)
                    if not (x1 > x0 and y1 > y0):
                        to_draw.append((cx0, cy0, cx1, cy1))
                    else:
                        if cx0 < x0:
                            to_draw.append((cx0, cy0, x0, cy1))
                        if cx1 > x1:
                            to_draw.append((x1, cy0, cx1, cy1))
                        if cy0 < y0:
                            to_draw.append((cx0, cy0, cx1, y0))
                        if cy1 > y1:
                            to_draw.append((cx0, y1, cx1, cy1))
                to_draw = to_draw[n:]
            for area in to_draw:
                self.tag_area(*area, tag="cells_to_draw")

    def tag_area(self, *area, tag, cnfg=None):
        kwargs = {"fill": "lightblue", "outline": "black", "width": 4, "stipple": "gray50", "tags": tag}
        if cnfg:
            kwargs.update(cnfg)
        if tag == "invalid_area" and not self.find_withtag(tag):
            # Los tags "invalid_area" y ("cell_to_draw", "cols_to_draw", "rows_to_draw") no coexisten
            [self.delete(item) for atag in ("cols_to_draw", "rows_to_draw", "cells_to_draw") for item in self.find_withtag(atag)]
        return self.create_rectangle(*area, **kwargs)

    def redraw_sheet(self, event=None, width=None, height=None):
        "Redraws the sheetui when the window is resized or needs updating."
        winfo_width = (width or self.winfo_width()) + int(self.f_headings) * COL_CELLS_WIDTH
        winfo_height = (height or self.winfo_height()) + int(self.f_headings) * ROW_CELLS_HEIGHT
        viewport_x1, viewport_y1 = self.cell_containing_coords(winfo_width, winfo_height)
        self.viewport_q1 = self.viewport_q1[:2] + (viewport_x1, viewport_y1)

        clsup_x, clsup_y =self.cell_coordinates(viewport_x1, viewport_y1)[2:]

        # self.delete("all")
        if self.find_withtag("background"):
            bg_coords = self.coords("background")
        else:
            x0, x1 = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[0]
            y0, y1 = self.coords_vportq3[1] - ROW_CELLS_HEIGHT, self.coords_vportq3[1]
            bg_coords = (x0, y0, x1, y1)
            self.create_rectangle(*bg_coords, fill="green", outline="black", tags="corner")
        
        for tag in ("invalid_area", ):
            self.delete(*self.find_withtag(tag))
        x0, y0, x1, y1 = bg_coords

        # Adjust the gridlines to the new viewport
        items = self.find_withtag("grid_lines")
        for item in items:
            gx0, gy0, gx1, gy1 = self.coords(item)
            if gy0 == gy1:
                # Horizontal gridlines
                self.coords(item, gx0, gy0, clsup_x, gy1)
            else:
                # Vertical gridlines
                self.coords(item, gx0, gy0, gx1, clsup_y)

        # Columns to draw
        area = (x1, self.coords_vportq3[1] - ROW_CELLS_HEIGHT, winfo_width, self.coords_vportq3[1])
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")
        # Rows to draw
        area = (self.coords_vportq3[0] - COL_CELLS_WIDTH, y1, self.coords_vportq3[0], winfo_height)
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")
        # Cells to draw
        areas = [(x1, y1, winfo_width, winfo_height)]
        if y1 - y0 > ROW_CELLS_HEIGHT:
            areas.append((x1, self.coords_vportq3[1], winfo_width, y1))
        if x1 - x0 > COL_CELLS_WIDTH:
            areas.append((self.coords_vportq3[0], y1, x1, winfo_height))
        for area in areas:
            self.tag_area(*area, tag="invalid_area")
            logging.debug(f"Invalidated area: {self.area_cells(*area)}")
        self.tag_raise("invalid_area")
        self.setGUI()

        self.xview('scroll', '-1', 'units')
        self.yview('scroll', '-1', 'units')
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
                    self.xview_moveto(viewport_x0)
                else:
                    if nquadrant in (2, 3):
                        coord_pivot_y = self.coords_vportq1[1]
                    if event.keysym == "Next":
                        viewport_y0 = viewport_y1 if viewport_x1 < MAX_ROWS else viewport_y0
                    else:
                        ybottom = self.cell_coordinates(0, viewport_y0)[3]
                        viewport_y0 = self.cell_containing_coords(0, ybottom - (winfo_height - ROW_CELLS_HEIGHT))[1]
                        viewport_y0 = min(MAX_ROWS, max(1, viewport_y0))
                    self.yview_moveto(viewport_y0)
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
                self.show_cell(acell_x0, acell_y0)
                self.set_active_cell()
                return "break"
            else:
                event.keysym = "Up" if event.state & SHIFT_PRESSED else "Down"  # Treat Return as Down for consistency
                event.state = 0
        elif event.keysym == "Tab":
            if self.selected_cells[:2] != self.selected_cells[2:]:
                sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
                acell_x0, acell_y0 = self.active_cell
                if event.state & SHIFT_PRESSED:  # If SHIFT is pressed
                    acell_y0 = acell_y0 if  acell_x0 > sel_x0 else ((acell_y0 - 1) if acell_y0 > sel_y0 else sel_y1)
                    acell_x0 = (acell_x0 - 1) if acell_x0 > sel_x0 else sel_x1
                else:  # If SHIFT is not pressed
                    acell_y0 = acell_y0 if acell_x0 < sel_x1 else ((acell_y0 + 1) if acell_y0 < sel_y1 else sel_y0)
                    acell_x0 = (acell_x0 + 1) if acell_x0 < sel_x1 else sel_x0
                self.active_cell = (acell_x0, acell_y0)
                self.set_active_cell()
                return "break"
            else:
                event.keysym = "Left" if event.state & SHIFT_PRESSED else "Right"  # Treat Return as Down for consistency
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
            linf_x, linf_y = 1, 1
            pivot.x = max(linf_x, min(MAX_COLS, pivot.x + dx))
            pivot.y = max(linf_y, min(MAX_ROWS, pivot.y + dy))
            xin, yin = pivot.x, pivot.y
        orig = self.quadrant_data(3)[0]
        if not self.f_freeze:
            if self.selected_cells[::2] == (1, MAX_COLS):
               xin, yin = self.viewport_q1[0], self.selected_cells[1::2][int(dy > 0)]
            elif self.selected_cells[1::2] == (1, MAX_ROWS):
               xin, yin = self.selected_cells[::2][int(dx > 0)],self.viewport_q1[1]
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
        winfo_width, winfo_height = self.winfo_width() + int(self.f_headings) * COL_CELLS_WIDTH, self.winfo_height() + int(self.f_headings) * ROW_CELLS_HEIGHT
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
                area =(self.coords_vportq3[0], self.coords_vportq3[1] - ROW_CELLS_HEIGHT, winfo_width, winfo_height)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
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

            area =(self.coords_vportq1[0], clinf_y, self.coords_vportq1[0] - dx, lsup_y)
            self.tag_area(*area, tag="invalid_area")
            logging.debug(f"Invalidated area: {self.area_cells(*area)}")
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

                area =(self.coords_vportq3[0], self.coords_vportq3[1] - ROW_CELLS_HEIGHT, winfo_width, winfo_height)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
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

                area =(lsup_x - dx, clinf_y, lsup_x, lsup_y)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
            pass
            if dx:
                viewport_x1 = self.cell_containing_coords(winfo_width, 0)[0]
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
                area =(clinf_x, self.coords_vportq1[1], lsup_x, lsup_y)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
            else:
                # Remove items that falls below the bottom edge
                dy = deltay
                dmy, viewport_y1 = self.cell_containing_coords(0, winfo_height + dy)
                linf_y = self.cell_coordinates(dmy, viewport_y1)[3]

                items = self.find_enclosed(clinf_x - 1, linf_y - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)

                # Move the viewport dy pixel up
                items = self.find_enclosed(clinf_x - 1, self.coords_vportq1[1] - 1,  lsup_x + 1, linf_y + 1)
                for item in items:
                    self.move(item, 0, -dy)
                viewport_y0 = y0
                area =(clinf_x, self.coords_vportq1[1], lsup_x, self.coords_vportq1[1] - dy)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
            self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
        elif deltay > 0:
            y1 = y
            lsup_x, lsup_y = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
            if deltay - winfo_height > winfo_height:
                # We need to move the viewport to the bottom a distance such that any information
                # displayed in the viewport is lost.
                items = self.find_enclosed(clinf_x - 1, self.coords_vportq1[1] - 1, lsup_x + 1, lsup_y + 1)
                self.delete(*items)
                viewport_y0 = y1
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
                dy = 1
                area =(clinf_x, self.coords_vportq1[1], lsup_x, lsup_y)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
            else:
                # Remove items that falls below the bottom edge
                dy = deltay
                items = self.find_enclosed(clinf_x - 1, self.coords_vportq1[1] - 1, lsup_x + 1, linf_y + 1)
                self.delete(*items)
                # Move the viewport dy pixel up
                items = self.find_enclosed(clinf_x - 1, linf_y - 1, lsup_x + 1, lsup_y + 1)
                for item in items:
                    self.move(item, 0, -dy)
                viewport_y0 = self.cell_containing_coords(0, linf_y + 1)[1]
                self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates

                area =(clinf_x, lsup_y - dy, lsup_x, lsup_y)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
        if dy:
            viewport_y1 = self.cell_containing_coords(0, winfo_height)[1]
            self.viewport_q1 = (viewport_x0, viewport_y0, viewport_x1, viewport_y1)  # Update the viewport coordinates
            pass
        # dx = dy = 0
        if dx or dy:
            # self.error_report = f'{self.viewport_q1}'
            # self.event_generate("<<errorReport>>")
            self.setGUI()  # Redraw the sheet with the new viewport
            self.tag_raise("freeze_line")  # Move freeze_line above all tags
            pass

    def show_cell(self, xin, yin):
        winfo_width, winfo_height = self.winfo_width(), self.winfo_height()
        viewport_x0, viewport_y0, viewport_x1, viewport_y1 = self.viewport_q1
        # lsup_coordx, lsup_coordy = self.cell_coordinates(viewport_x1, viewport_y1)[2:]
        # if lsup_coordx > winfo_width and xin >= viewport_x1:
        lsup_x = self.cell_containing_coords(winfo_width, 0)[0]
        if xin >= lsup_x:
            xright = self.cell_coordinates(xin, 0)[2]
            x = self.cell_containing_coords(xright - (winfo_width - self.coords_vportq1[0]), 0)[0]
            xright = self.cell_coordinates(x, 0)[2]
            viewport_x0 = self.cell_containing_coords(xright + 1, 0)[0]
        elif xin < viewport_x0:
            viewport_x0 = xin
        
        # if  lsup_coordy > winfo_height and yin >= viewport_y1:
        lsup_y = self.cell_containing_coords(0, winfo_height)[1]
        if  yin >= lsup_y:
            ybottom = self.cell_coordinates(0, yin)[3]
            y = self.cell_containing_coords(0, ybottom - (winfo_height - self.coords_vportq1[1]))[1]
            ybottom = self.cell_coordinates(0, y)[3]
            viewport_y0 = self.cell_containing_coords(0, ybottom + 1)[1]
        elif yin < viewport_y0:
            viewport_y0 = yin
        self.move_viewport(viewport_x0, viewport_y0)
        self.xview_moveto(viewport_x0)
        self.yview_moveto(viewport_y0)
        pass

    def toggle_headings(self, *args):
        """Toggles the visibility of headings."""
        if self.f_headings:
            dx, dy = -COL_CELLS_WIDTH, -ROW_CELLS_HEIGHT
        else:
            dx, dy = COL_CELLS_WIDTH, ROW_CELLS_HEIGHT
        lsup_x, lsup_y = self.cell_coordinates(self.viewport_q1[2], self.viewport_q1[3])[2:]
        self.coords_vportq1 = (self.coords_vportq1[0] + dx, self.coords_vportq1[1] + dy)
        self.coords_vportq3 = linf_x, linf_y = (self.coords_vportq3[0] + dx, self.coords_vportq3[1] + dy)
        items = self.find_enclosed(-linf_x - 1, -linf_y - 1, lsup_x + 1, lsup_y + 1)
        if self.f_freeze:
            items += self.find_withtag("freeze_line")
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
        # Check if the click is not on an existing cell
        items = self.find_overlapping(event.x, event.y, event.x, event.y)
        if not items:
            return "break"
        if event.x < COL_CELLS_WIDTH and event.y < ROW_CELLS_HEIGHT:
            self.selected_cells = (1, 1, MAX_COLS, MAX_ROWS)
            self.active_cell = self.viewport_q1[:2]
            self.set_active_cell()
            return "break"
        event_x, event_y = max(event.x, COL_CELLS_WIDTH), max(event.y, ROW_CELLS_HEIGHT)
        nquadrant = self.cell_quadrant(event_x, event_y)
        orig, coords_orig = self.quadrant_data(nquadrant)
        clk_x, clk_y = self.cell_containing_coords(event_x, event_y, orig, coords_orig)
        with self.pivot_point(isActiveCell=not event.state & SHIFT_PRESSED) as pivot:
            pivot.x = clk_x
            pivot.y = clk_y
        if row_clk := event.x < COL_CELLS_WIDTH: # and event.y >= ROW_CELLS_HEIGHT:
            sel_y0, sel_y1 = self.selected_cells[1::2]
            self.selected_cells = 1, sel_y0, MAX_COLS, sel_y1
            if not event.state & SHIFT_PRESSED:
                self.active_cell = (self.viewport_q1[0], clk_y)
        elif col_clk := event.y < ROW_CELLS_HEIGHT: # and event.x >= COL_CELLS_WIDTH:
            sel_x0, sel_x1 = self.selected_cells[::2]
            self.selected_cells = sel_x0, 1, sel_x1, MAX_ROWS
            if not event.state & SHIFT_PRESSED:
                self.active_cell = (clk_x, self.viewport_q1[1])
        # if (row_clk or col_clk) and not event.state & SHIFT_PRESSED:
        #     self.active_cell = (clk_x, self.viewport_q1[1])
        self.set_active_cell()
        self.focus_set()  # Set focus to the canvas

    def mouse_drag(self, event):
        """Handles mouse drag events to set the active cell."""
        if self.f_drag:
            # Update the mouse pointer coordinates in screen coordinates
            event_x = self.winfo_pointerx() - self.winfo_rootx()
            event_y = self.winfo_pointery() - self.winfo_rooty()
            if event is None:
                logging.debug("Mouse drag event triggered for col o row selection")
            if event_x >= COL_CELLS_WIDTH and event_y < ROW_CELLS_HEIGHT:
                # mouse over column headings
                if self.selected_cells[1::2] != (1, MAX_ROWS):
                    # Not column selection
                    self.yview('scroll', '-1', 'units')
                    clk_x, clk_y = self.cell_containing_coords(event_x, ROW_CELLS_HEIGHT + 1)
                    with self.pivot_point(isActiveCell=False) as pivot:
                        pivot.y = clk_y
                    self.set_active_cell()
                    self.after(1000, self.mouse_drag, event)  # Repeat the drag event after a delay
                else:
                    # Column selection
                    nquadrant = self.cell_quadrant(event_x, ROW_CELLS_HEIGHT)
                    orig, coords_orig = self.quadrant_data(nquadrant)
                    clk_x, clk_y = self.cell_containing_coords(event_x, ROW_CELLS_HEIGHT, orig, coords_orig)
                    logging.debug(f"{event_x=}, {self.winfo_width()=}")
                    if event_x > self.winfo_width():
                        self.show_cell(clk_x, clk_y)
                        self.after(1000, self.mouse_drag, None)
                    event_y = ROW_CELLS_HEIGHT - 1
                    event = tk.Event()
                    event.x = event_x
                    event.y = event_y
                    event.state = SHIFT_PRESSED
                    self.mouse_click(event)
            elif event_x >= COL_CELLS_WIDTH and event_y >= ROW_CELLS_HEIGHT:
                # mouse over cells
                clk_x, clk_y = self.cell_containing_coords(event_x, event_y)
                with self.pivot_point(isActiveCell=False) as pivot:
                    pivot.x = clk_x
                    pivot.y = clk_y
                self.show_cell(clk_x, clk_y)
                self.set_active_cell()
            elif event_x < COL_CELLS_WIDTH and event_y >= ROW_CELLS_HEIGHT:
                # mouse over row headings
                if self.selected_cells[::2] != (1, MAX_COLS):
                    # Not row selection
                    self.xview('scroll', '-1', 'units')
                    clk_x, clk_y = self.cell_containing_coords(COL_CELLS_WIDTH + 1, event_y)
                    with self.pivot_point(isActiveCell=False) as pivot:
                        pivot.x = clk_x
                    self.set_active_cell()
                    self.after(1000, self.mouse_drag, event)
                else:
                    # Row selection
                    nquadrant = self.cell_quadrant(COL_CELLS_WIDTH, event_y)
                    orig, coords_orig = self.quadrant_data(nquadrant)
                    clk_x, clk_y = self.cell_containing_coords(COL_CELLS_WIDTH, event_y, orig, coords_orig)
                    logging.debug(f"{event_y=}, {self.winfo_height()=}")
                    if event_y > self.winfo_height():
                        self.show_cell(clk_x, clk_y)
                        self.after(1000, self.mouse_drag, None)
                    event_x = COL_CELLS_WIDTH - 1
                    event = tk.Event()
                    event.x = event_x
                    event.y = event_y
                    event.state = SHIFT_PRESSED
                    self.mouse_click(event)
            else:
                # mouse over corners
                viewport_x0, viewport_y0 = self.viewport_q1[:2]
                event = tk.Event()
                if self.selected_cells[::2] == (1, MAX_COLS):
                    # Column selection
                    self.move_viewport(viewport_x0, viewport_y0 - 1)
                    event.x, event.y = COL_CELLS_WIDTH - 1, ROW_CELLS_HEIGHT
                    self.mouse_click(event)
                elif self.selected_cells[1::2] == (1, MAX_ROWS):
                    # Row selection
                    self.move_viewport(viewport_x0 - 1, viewport_y0)
                    event.x, event.y = COL_CELLS_WIDTH, ROW_CELLS_HEIGHT - 1
                    self.mouse_click(event)
                else:
                    self.move_viewport(viewport_x0 - 1, viewport_y0 - 1)
                    with self.pivot_point(isActiveCell=False) as pivot:
                        clk_x, clk_y = self.cell_containing_coords(COL_CELLS_WIDTH + 1,ROW_CELLS_HEIGHT + 1)
                        pivot.x = clk_x
                        pivot.y = clk_y
                self.set_active_cell()
            return "break"  # Prevent default behavior of mouse drag
        else:
            logging.debug("Mouse drag event ignored, not in drag mode.")


    def mouse_release(self, event):
        self.f_drag = False
        pass

    def on_mouse_wheel(self, event):
        logging.debug(f"Mouse wheel:{event=}, {event.delta=}")
        delta = -1 if event.delta > 0 else 1
        fnc = self.xview if event.state & SHIFT_PRESSED else self.yview
        fnc("scroll", delta, 'units')

    def toggle_freeze_panes(self, *args):
        if not self.f_freeze:
            x0, y0, x1, y1 = self.viewport_q3
            coord_acx, coord_acy = self.cell_coordinates(*self.active_cell)[:2]
            if self.active_cell[0] != self.viewport_q1[0]:
                x0, x1 = self.viewport_q1[0], self.active_cell[0]
                self.coords_vportq1 = coord_acx, self.coords_vportq1[1]
            if self.active_cell[1] != self.viewport_q1[1]:
                y0, y1 = self.viewport_q1[1], self.active_cell[1]
                self.coords_vportq1 = self.coords_vportq1[0], coord_acy
            self.viewport_q3 = (x0, y0, x1, y1)
            self.viewport_q1 = *self.active_cell, *self.viewport_q1[2:]
            self.set_freeze_lines()
            self.xview_moveto(0.0)
            self.yview_moveto(0.0)
        else:
            # If freeze is active, unfreeze the panes
            self.move_viewport(*self.viewport_q3[2:])
            self.coords_vportq1 = self.coords_vportq3
            self.viewport_q1 = *self.viewport_q3[:2], *self.viewport_q1[2:]
            self.viewport_q3 = 1, 1, 1, 1
            items = self.find_withtag("freeze_line")
            self.delete(*items)
        self.f_freeze = not self.f_freeze

    def set_active_cell(self):
        # Set the tag "selected" for the region in coords (40, CELL_HEIGHT, 40 + 5*CELL_WIDTH, CELL_HEIGHT + 5*CELL_HEIGHT) rectangle
        def clip_rectangle(x0, y0, x1, y1, clipping_rgn=None):
            if clipping_rgn:
                linf_x, linf_y, lsup_x, lsup_y = clipping_rgn
                x0, x1 = min(lsup_x, max(linf_x, x0)), max(linf_x, min(x1, lsup_x))
                y0, y1 = min(lsup_y, max(linf_y, y0)), max(linf_y, min(y1, lsup_y))
            return x0, y0, x1, y1
        self.delete("selected_cells")
        winfo_width, winfo_height = self.winfo_width() + int(self.f_headings) * COL_CELLS_WIDTH, self.winfo_height() + int(self.f_headings) * ROW_CELLS_HEIGHT
        clipping_rect = (self.coords_vportq3[0], self.coords_vportq3[1], winfo_width, winfo_height)

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

    def ymin_fraction(self):
            y1 = MAX_ROWS
            y0 = int(y1 - (self.winfo_height() - ROW_CELLS_HEIGHT) // CELL_HEIGHT)
            min_fraction = 1 - (y1 - y0) / (MAX_ROWS - self.viewport_q3[3])
            return min_fraction
    
    def yview(self, *args):
        if not args:
            min_fraction = self.ymin_fraction()
            viewport_y0, viewport_y1 = self.viewport_q1[1::2]
            denom = MAX_ROWS - self.viewport_q3[3]
            first = (viewport_y0 - self.viewport_q3[3]) / denom
            first = min(first, min_fraction)
            last = (viewport_y1 - self.viewport_q3[3]) / denom if first < min_fraction else 1.0
            return first, last
        
        elif args[0] == 'scroll':
            direction = args[2]
            if direction == 'units':
                delta = int(args[1])
                viewport_y0 = self.viewport_q1[1]
                viewport_y0 += delta
                self.yview_moveto(viewport_y0)
            elif direction == 'pages':
                delta = int(args[1])
                keysym = 'Prior' if delta < 0 else 'Next'
                viewport_y0 = self.viewport_q1[1]
                ytop, ybottom = self.cell_coordinates(0, viewport_y0)[1::2]
                if keysym == "Next":
                    viewport_y0 = self.cell_containing_coords(0, ytop + (self.winfo_height() - ROW_CELLS_HEIGHT))[1]
                    viewport_y0 = min(MAX_ROWS, max(1, viewport_y0))
                else:
                    viewport_y0 = self.cell_containing_coords(0, ybottom - (self.winfo_height() - ROW_CELLS_HEIGHT))[1]
                    viewport_y0 = min(MAX_ROWS, max(1, viewport_y0))
                fraction = (viewport_y0 - self.viewport_q3[3]) / (MAX_ROWS - self.viewport_q3[3])
                self.yview_moveto(fraction)
            self.set_active_cell()
        elif args[0] == 'moveto':
            self.yview_moveto(args[1])
        else:
            logging.warning(f"Unknown yview command: {args[0]}")
            return super().yview(*args)

    def yview_moveto(self, fraction):
        match fraction:
            case int() as cell_y:
                viewport_y0 = cell_y
            case _:
                fraction = min(self.ymin_fraction(), float(fraction))
                viewport_y0 = int(fraction * (MAX_ROWS - self.viewport_q3[3]) + self.viewport_q3[3])

        viewport_x0 = self.viewport_q1[0]
        viewport_y0 = max(self.viewport_q3[3], min(MAX_ROWS, viewport_y0))
        self.move_viewport(viewport_x0, viewport_y0)

        if scb_get := self.cget("yscrollcommand"):  #vertical scrollbar (scb) get command
            _tk = self._root().tk
            return _tk.call(scb_get, *self.yview())

    def xmin_fraction(self):
        x1 = MAX_COLS
        x0 = int(x1 - (self.winfo_width() - COL_CELLS_WIDTH) // CELL_WIDTH)
        min_fraction = 1 - (x1 - x0) / (MAX_COLS - self.viewport_q3[2])
        return min_fraction
    
    def xview(self, *args):
        if not args:
            min_fraction = self.xmin_fraction()
            viewport_x0, viewport_x1 = self.viewport_q1[::2]
            denom = MAX_COLS - self.viewport_q3[2]
            first = (viewport_x0 - self.viewport_q3[2]) / denom
            first = min(first, min_fraction)
            last = (viewport_x1 - self.viewport_q3[2]) / denom if first < min_fraction else 1.0
            return first, last
        
        elif args[0] == 'scroll':
            direction = args[2]
            if direction == 'units':
                delta = int(args[1])
                viewport_x0 = self.viewport_q1[0]
                viewport_x0 += delta
                self.xview_moveto(viewport_x0)
            elif direction == 'pages':
                delta = int(args[1])
                keysym = 'Prior' if delta < 0 else 'Next'
                viewport_x0 = self.viewport_q1[0]
                xtop, xbottom = self.cell_coordinates(viewport_x0, 0)[::2]
                if keysym == "Next":
                    viewport_x0 = self.cell_containing_coords(xtop + (self.winfo_width() - COL_CELLS_WIDTH), 0)[0]
                    viewport_x0 = min(MAX_COLS, max(1, viewport_x0))
                else:
                    viewport_x0 = self.cell_containing_coords(xbottom - (self.winfo_width() - COL_CELLS_WIDTH), 0)[0]
                    viewport_x0 = min(MAX_COLS, max(1, viewport_x0))
                fraction = (viewport_x0 - self.viewport_q3[2]) / (MAX_COLS - self.viewport_q3[2])
                self.xview_moveto(fraction)
            self.set_active_cell()
        elif args[0] == 'moveto':
            self.xview_moveto(args[1])
        else:
            logging.warning(f"Unknown xview command: {args[0]}")
            return super().xview(*args)
        
    def xview_moveto(self, fraction):
        match fraction:
            case int() as cell_y:
                viewport_x0 = cell_y
            case _:
                fraction = min(self.xmin_fraction(), float(fraction))
                viewport_x0 = int(fraction * (MAX_COLS - self.viewport_q3[2]) + self.viewport_q3[2])

        viewport_y0 = self.viewport_q1[1]
        viewport_x0 = max(self.viewport_q3[2], min(MAX_COLS, viewport_x0))
        self.move_viewport(viewport_x0, viewport_y0)

        if scb_get := self.cget("xscrollcommand"):  #vertical scrollbar (scb) get command
            _tk = self._root().tk
            return _tk.call(scb_get, *self.xview())


if __name__ == "__main__":
    import inspect
    class SheetViewer(tk.Tk):
        def __init__(self):
            super().__init__()
            self.f_rec = False
            self.action_stack = []
            self.setGui()
            self.bind("<<ActiveCellChanged>>", self.on_active_cell_changed)
            self.bind("<<SelectedCellsChanged>>", self.on_selected_cells_changed)
            self.bind("<<errorReport>>", self.on_error_report)
            self.bind("<<EventMonitor>>", self.event_monitor)
            self.geometry("600x400")
            self.monitor = self.bind("<<EventMonitor>>")

        def event_monitor(self, event):
            wdg = event.widget
            wname = f"{wdg.winfo_parent()}.{wdg.winfo_name()}"
            sevent = str(event)
            logging.debug(f"****** Event: {sevent} *****")

            # Action string
            sevent = sevent.strip('<>').replace(' event ', ' ')
            eseq, *kwargs = sevent.split()
            kwargs = dict(item.split('=') for item in kwargs)
            if eseq == 'Configure':
                # Why?
                # <Configure> is a system event (not a user event like <Button-1> or <KeyPress>).
                # Tkinter/Tk will ignore attempts to generate it manually.
                saction = f'self.geometry("{kwargs["width"]}x{kwargs["height"]}+0+0")'
            else:
                for key in set(['keysym', 'state']).intersection(kwargs.keys()):
                    kwargs[key] = f"'{kwargs[key]}'"

                if 'send_event' in kwargs:
                    kwargs['sendevent'] = kwargs.pop('send_event')
                    
                kwargs = ', '.join(f'{k}={v}' for k, v in kwargs.items())
                saction = f"self.nametowidget('{wname}').event_generate('<{eseq}>', {kwargs})"
            logging.debug(f"****** Action: {saction} *****")
            self.action_stack.append(saction)
            # Get widget wit the name 'txt'
            wdg = self.nametowidget('.errorfrm.txt')
            wdg['text'] = saction.rsplit('.', 1)[-1]

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
            self.rowconfigure(3, weight=1)  # Change to row 2 for the main frame

            # --- Add labels at the top ---

            frame = ttk.Frame(self)
            frame.grid(row=0, column=0, sticky="ew", padx=4, pady=(0, 4))
            vals = ["choose an action", "validate_areas", "toggle_headings", "toggle_gridlines", "toggle_freeze_panes", "show_cell", "move_viewport"]
            self.cbox = cbox = ttk.Combobox(frame, name="cbox", values=vals, state="readonly")
            cbox.set(vals[0])  # Set default value
            cbox.pack(side="left")
            cbox.bind("<<ComboboxSelected>>", self.on_combobox_change)  # <-- Bind the event here
            self.activeCell = ttk.Label(frame, text="Active Cell: ", background="magenta", font=("Arial", 10))
            self.activeCell.pack(side="left",expand=True, fill="x", padx=4)
            
            frame = ttk.Frame(self, name='actionfrm')
            frame.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))
            chkbtn = ttk.Checkbutton(frame, name="rec", text="Rec", style="Toolbutton", command=lambda: self.action_cmds('rec'))
            chkbtn.pack(side="left")
            # txt = ttk.Label(frame, name="txt", width=60, text='.....')
            # txt.pack(side="left")
            btn = ttk.Button(frame, text="run", command=lambda: self.action_cmds('step'))
            btn.pack(side="left")
            btn = ttk.Button(frame, text="all", command=lambda: self.action_cmds('run'))
            btn.pack(side="left")
            btn = ttk.Button(frame, text="save", command=lambda: self.action_cmds('save'))
            btn.pack(side="right")
            btn = ttk.Button(frame, text="load", command=lambda: self.action_cmds('load'))
            btn.pack(side="right")

            frame = ttk.Frame(self, name='errorfrm')
            frame.grid(row=2, column=0, sticky="ew", padx=4, pady=(0, 4))
            lbl = ttk.Label(frame, text="Last action:")
            lbl.pack(side="left")
            self.errorReport = ttk.Label(frame, name="txt", text="....", background="grey", font=("Arial", 10), foreground="white", anchor="w")
            self.errorReport.pack(side="left", expand=tk.YES, fill=tk.X)

            # Create a frame to hold the canvas and scrollbars
            frame = ttk.Frame(self)
            frame.grid(column=0, row=3, sticky=(tk.N, tk.W, tk.E, tk.S))
            frame.columnconfigure(0, weight=1)
            frame.rowconfigure(0, weight=1)

            # Create vertical and horizontal scrollbars
            v_scroll = ttk.Scrollbar(frame, name='vscroll', orient="vertical")
            h_scroll = ttk.Scrollbar(frame, name='hscroll', orient="horizontal")

            # Create the SheetUI canvas
            self.sheetui = sheetui = SheetUI(frame, name='sheetui', bg=GRID_COLOR, 
                                    yscrollcommand=v_scroll.set, 
                                    xscrollcommand=h_scroll.set,
                                    scrollregion=(1, 1, MAX_COLS, MAX_ROWS)
            )  # Adjust scrollregion as needed

                # Configure scrollbars to control the canvas
            v_scroll.config(command=sheetui.yview)
            h_scroll.config(command=sheetui.xview)

            # Layout
            sheetui.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")

        def clean_slate(self):
            # Put the canvas in a clean slate
            x, y = self.sheetui.viewport_q3[2:]
            self.sheetui.show_cell(x, y)
            self.sheetui.validate_areas()
            self.sheetui.validate_areas()
        
        def action_cmds(self, cmd):
            self.sheetui.focus_set()
            if cmd == 'save':
                idir = os.path.dirname(os.path.abspath(__file__))
                fname = filedialog.asksaveasfilename(
                    parent=self,
                    title="Save As",
                    defaultextension=".tx",
                    filetypes=[("Macro Files", "*.txt"), ("All Files", "*.*")],
                    initialdir=os.path.join(os.getcwd(), "macros"),

                )
                if fname:
                    if self.f_rec:
                        wdg = self.nametowidget('.actionfrm.rec')
                        wdg.click()
                    # Save your data to 'filename'
                    logging.debug(f"Saving to:{fname}")
                    content = '\n'.join(self.action_stack)
                    with open(fname, "w") as f:
                        f.write(content)
            elif cmd == 'load':
                idir = os.path.dirname(os.path.abspath(__file__))
                fname = filedialog.askopenfilename(
                    parent=self,
                    title="Open",
                    defaultextension=".txt",
                    filetypes=[("Macro Files", "*.txt"), ("All Files", "*.*")],
                    initialdir=os.path.join(idir, "macros"),
                )
                if fname:
                    # Load your data from 'filename'
                    logging.debug(f"Loading from:{fname}")
                    with open(fname, "r") as f:
                        content = f.readlines()
                    self.action_stack = content
                    self.nametowidget('.errorfrm.txt')['text'] = content[-1]
            elif cmd == 'rec':
                wdg = self.nametowidget('.actionfrm.rec')
                self.f_rec = not self.f_rec
                binds = self.sheetui.bind()
                if self.f_rec:
                    for bind in binds:
                        bnd_cb = self.sheetui.bind(bind)
                        bnd_cb = '\n'.join([self.monitor, bnd_cb])
                        self.sheetui.bind(bind, bnd_cb)
                    wdg['text'] = "Stop"
                else:
                    for bind in binds:
                        bnd_cb = self.sheetui.bind(bind)
                        bnd_cb = bnd_cb.split('\n\n')[1]
                        self.sheetui.bind(bind, bnd_cb)
                    wdg['text'] = "Rec"
                pass

            elif cmd == 'run':
                for action in self.action_stack[:-1]:
                    exec(action)
            elif cmd == 'step':
                exec(self.action_stack[-1])


        def on_combobox_change(self, event):
                fname = self.cbox.get()
                if fname == 'validate_areas':
                    msg1 = "Delete invalid areas (yes/no):"
                elif fname == 'toggle_headings':
                    msg1 = "Toggle headings (yes/no):"
                elif fname == 'toggle_gridlines':
                    msg1 = "Toggle gridlines (yes/no):"
                elif fname == 'toggle_freeze_panes':
                    msg1 = "Toggle freeze panes (yes/no):"
                elif fname == 'show_cell':
                    msg1 = "Enter the pivot cell coordinates (col, row):"
                elif fname == "move_viewport":
                    msg1 = "Enter the pivot cell coordinates (delta_col, delta_row):"
                else:
                    return
                fnc = getattr(self.sheetui, fname)
                args = []
                if inspect.signature(fnc).parameters:
                    # display a message box to get  the pivot cell coordinates
                    answ = simpledialog.askstring(fname, msg1, parent=self)
                    if answ:
                        try:
                            args = list(map(lambda w: int(w), answ.split(",")))
                        except ValueError:
                            fnc = lambda *args: 1
                            print("Invalid input. Please enter parameters as integers separated with commas in the format 'x, y'.")
                    else:
                        fnc = lambda *args: 2
                        print("No input provided.")
                sargs = ", ".join(map(str, args))
                logging.debug(f"self.sheetui.{fname}({sargs})")
                if self.f_rec:
                    saction = f"self.sheetui.{fname}({sargs})"
                    self.action_stack.append(saction)
                    # Get widget wit the name 'txt'
                    wdg = self.nametowidget('.errorfrm.txt')
                    wdg['text'] = f"{fname}({sargs})"
                fnc(*args)
                self.cbox.set("choose an action")
                self.sheetui.focus_set()



    root = SheetViewer()
    root.mainloop()