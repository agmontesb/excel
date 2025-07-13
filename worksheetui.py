''' Este módulo implementa la interfaz gráfica de usuario (GUI) para la hoja de cálculo.
    Proporciona una ventana principal con una barra de menú, una barra de herramientas y un área de trabajo 
    que consiste en un área de celdas y una barra de estado que permite cambiar a otras hojas de trabajo '
    además de la activa'
'''
import os
import inspect
import tkinter as tk
from tkinter import ttk
from tkinter import simpledialog
from tkinter import filedialog
from contextlib import contextmanager
from types import SimpleNamespace
from collections import Counter

import logging
from typing import Callable, Literal

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


class SheetLook:
    def __init__(self, canvas: 'SheetUI', cell_content_gen: Callable[[int, int], str]):
        self.canvas = canvas
        self.headings_dim = {}
        self.headings_hided = {}

        self.coords_vportq3 = (COL_CELLS_WIDTH, ROW_CELLS_HEIGHT)
        self.viewport_q3 = (1, 1, 1, 1)
        self.coords_vportq1 = self.coords_vportq3
        self.viewport_q1 = (1, 1, 1, 1)                                 # Default pivot cell
        self.active_cell = self.viewport_q1[:2]                         # Variable to store the active cell     
        self.selected_cells = (*self.active_cell, *self.active_cell)    # Variable to store the selected cell
        self.cell_content = cell_content_gen
        pass

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
            self.canvas.event_generate("<<ActiveCellChanged>>")
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
            self.canvas.event_generate("<<SelectedCellsChanged>>")
        pass

    def cell_coordinates(self, x: int, y:int, viewport:tuple[int, ...]=None, coords_viewport: tuple[int, int]=None) -> tuple[int, int, int, int]:
        """Calculates the coordinates of the cell based on the x and y position."""
        x, y =  map(int, (x, y))
        if viewport is None:
            viewport = self.viewport_q1
        if coords_viewport is None:
            coords_viewport = self.coords_vportq1
        col_width = [self.headings_dim[f"C{ikey}"] for ikey in range(min(viewport[0], x), max(viewport[0], x)) if f"C{ikey}" in self.headings_dim]
        x0 = coords_viewport[0] + ((-1) ** int(x < viewport[0]))*((abs(x - viewport[0]) - len(col_width)) * CELL_WIDTH + sum(col_width))
        row_height = [self.headings_dim[f"R{ikey}"] for ikey in range(min(viewport[1], y), max(viewport[1], y)) if f"R{ikey}" in self.headings_dim]
        y0 = coords_viewport[1] + ((-1) ** int(y < viewport[1]))*((abs(y - viewport[1]) - len(row_height)) * CELL_HEIGHT + sum(row_height))
        return (x0, y0, x0 + self.headings_dim.get(f"C{x}", CELL_WIDTH), y0 + self.headings_dim.get(f"R{y}", CELL_HEIGHT))
    
    def area_coordinates(self, x0:int, y0:int, x1:int, y1:int) -> tuple[int, int, int, int]:
        nquadrant = self.cell_quadrant(x0, y0, isCoord=False)
        orig, coords_orig = self.quadrant_data(nquadrant)
        sel_x0, sel_y0 = self.cell_coordinates(x0, y0, orig, coords_orig)[:2]
        nquadrant = self.cell_quadrant(x1, y1, isCoord=False)
        orig, coords_orig = self.quadrant_data(nquadrant)
        sel_x1, sel_y1 = self.cell_coordinates(x1, y1, orig, coords_orig)[2:]
        return (sel_x0, sel_y0, sel_x1, sel_y1)
    
    def area_cells(self, sel_x0:int, sel_y0:int, sel_x1:int, sel_y1:int) -> tuple[int, int, int, int]:
        nquadrant = self.cell_quadrant(sel_x0, sel_y0, isCoord=True)
        orig, coords_orig = self.quadrant_data(nquadrant)
        x0, y0 = self.cell_containing_coords(sel_x0, sel_y0, orig, coords_orig)
        nquadrant = self.cell_quadrant(sel_x1 - 1, sel_y1 - 1, isCoord=True)
        orig, coords_orig = self.quadrant_data(nquadrant)
        x1, y1 = self.cell_containing_coords(sel_x1 - 1, sel_y1 - 1, orig, coords_orig)
        return (x0, y0, x1, y1)
    
    def cell_containing_coords(self, ptx:int, pty:int, viewport:tuple[int, ...]=None, coords_viewport:tuple[int, int]=None) -> tuple[int, int]:
        """Returns the cell address containing the given x and y screen coordinates."""
        if viewport is None:
            viewport = self.viewport_q1
        if coords_viewport is None:
            coords_viewport = self.coords_vportq1
        xcell = viewport[0]
        while True:
            ptx0, ptx1 = self.cell_coordinates(xcell, 0, viewport, coords_viewport)[::2]
            if ptx0 <= ptx < ptx1:
                break
            if ptx < ptx0:
                delta = max((ptx0 - ptx) // CELL_WIDTH, 1)
                while f"C{xcell - delta}" in self.headings_hided:
                    delta += 1
                xcell -= delta
            elif ptx >= ptx1:
                delta = max((ptx - ptx1) // CELL_WIDTH, 1)
                while f"C{xcell + delta}" in self.headings_hided:
                    delta += 1
                xcell += delta

        ycell = viewport[1]
        while True:
            pty0, pty1 = self.cell_coordinates(0, ycell, viewport, coords_viewport)[1::2]
            if pty0 <= pty < pty1:
                break
            if pty < pty0:
                delta = max((pty0 - pty) // CELL_HEIGHT, 1)
                while f"R{ycell - delta}" in self.headings_hided:
                    delta += 1
                ycell -= delta
            elif pty >= pty1:
                delta = max((pty - pty1) // CELL_HEIGHT, 1)
                while f"R{ycell + delta}" in self.headings_hided:
                    delta += 1
                ycell += delta
        xcell = max(1, min(MAX_COLS, int(xcell)))
        ycell = max(1, min(MAX_ROWS, int(ycell)))
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
    
    def quadrant_data(self, nquadrant:int) -> tuple[tuple[int, ...], tuple[int, int]]:
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

    def map_cell_to_coords(self, x, y, coords=None, coords_viewport=None):
        """Link a cell coordinates to the canvas coordinates."""
        if isViewportOrig := coords is None:
            coords_viewport = self.coords_vportq1
            viewport = self.viewport_q1[:2]
        else:
            pass
        viewport_x0, viewport_y0 = viewport
        winfo_width, winfo_height = self.canvas.efective_width(), self.canvas.efective_height()
        x = max(1, min(MAX_COLS, x))
        y = max(1, min(MAX_ROWS, y))
        linf_x, linf_y = self.cell_coordinates(x, y, viewport, coords_viewport)[:2]
        deltax, deltay = linf_x - coords_viewport[0], linf_y - coords_viewport[1]
        clinf_y = self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        clinf_x = self.coords_vportq3[0] - COL_CELLS_WIDTH

        ptx0 =  ptx1 = pty0 = pty1 = None
        lsup_x, lsup_y = self.cell_coordinates(*self.viewport_q1[2:], viewport, coords_viewport)[2:]
        gx1, gy1 = map(int, self.canvas.coords("background")[2:])

        if deltax and ((deltax + winfo_width < coords_viewport[0]) or (deltax - winfo_width > winfo_width)):
            # All the cell information and column headings need to be updated
            items = self.canvas.find_enclosed(coords_viewport[0] - 1, clinf_y - 1, gx1 + 1, gy1 + 1)
            self.canvas.delete(*items)
            viewport_x0 = x
            ptx0 = self.coords_vportq3[0], clinf_y
        elif deltax < 0:
            # left displacement
            x0 = x
            dx = deltax
            viewport_x1, dmy = self.cell_containing_coords(winfo_width + dx, 0, viewport, coords_viewport)
            linf_x = self.cell_coordinates(viewport_x1, dmy, viewport, coords_viewport)[2]

            items = self.canvas.find_enclosed(linf_x - 1, clinf_y - 1, gx1 + 1, gy1 + 1)
            self.canvas.delete(*items)

            # Move the viewport dx pixel to the left
            items = self.canvas.find_enclosed(coords_viewport[0] - 1, clinf_y - 1, linf_x + 1, gy1 + 1)
            for item in items:
                self.canvas.move(item, -dx, 0)
            
            # Resize the horizontal gridlines
            if gx1 != lsup_x:
                for item in self.canvas.find_withtag("hgrid_lines"):
                    y0, y1 = self.canvas.coords(item)[::2]
                    self.canvas.coords(item, coords_viewport[0], y0, lsup_x, y1)

            viewport_x0 = x0
            ptx0, ptx1 =(coords_viewport[0], clinf_y), (coords_viewport[0] - dx, lsup_y)
            # self.canvas.tag_area(*area, tag="invalid_area")
        elif deltax > 0:
            # Rigth displacement
            dx = deltax
            items = self.canvas.find_enclosed(coords_viewport[0] - 1, clinf_y - 1, linf_x + 1, gy1 + 1)
            self.canvas.delete(*items)
            # Move the viewport dx pixel to the left
            items = self.canvas.find_enclosed(linf_x - 1, clinf_y - 1, gx1 + 1, gy1 + 1)
            for item in items:
                self.canvas.move(item, -dx, 0)
            viewport_x0 = self.cell_containing_coords(linf_x + 1, 0, viewport, coords_viewport)[0]

            # Resize the horizontal gridlines
            if gx1 != lsup_x:
                for item in self.canvas.find_withtag("hgrid_lines"):
                    y0, y1 = self.canvas.coords(item)[::2]
                    self.canvas.coords(item, coords_viewport[0], y0, lsup_x, y1)

            ptx0 = lsup_x - dx, clinf_y
            pass
        if ptx0:
            viewport_x1 = self.cell_containing_coords(winfo_width, 0, (viewport_x0, viewport[1]), coords_viewport)[0]
            viewport_x0 = viewport_x0 if isViewportOrig else self.viewport_q1[0]
            self.viewport_q1 = (viewport_x0, self.viewport_q1[1],viewport_x1, self.viewport_q1[3])
            pass

        if deltay and ((deltay + winfo_height < coords_viewport[1]) or (deltay - winfo_height > winfo_height)):
            # All the cell information and row headings need to be updated
            items = self.canvas.find_enclosed(clinf_x - 1, coords_viewport[1] - 1, gx1 + 1, gy1 + 1)
            self.canvas.delete(*items)
            viewport_y0 = y
            pty0 = clinf_x, coords_viewport[1]
        elif deltay < 0:
            # top displacement
            y0 = y
            dy = deltay
            dmy, viewport_y1 = self.cell_containing_coords(0, winfo_height + dy, viewport, coords_viewport)
            linf_y = self.cell_coordinates(dmy, viewport_y1, viewport, coords_viewport)[3]

            items = self.canvas.find_enclosed(clinf_x - 1, linf_y - 1, gx1 + 1, gy1 + 1)
            self.canvas.delete(*items)

            # Move the viewport dy pixel up
            items = self.canvas.find_enclosed(clinf_x - 1, coords_viewport[1] - 1,  gx1 + 1, linf_y + 1)
            for item in items:
                self.canvas.move(item, 0, -dy)

            # Resize the vertical gridlines
            if gy1 != lsup_y:
                for item in self.canvas.find_withtag("vgrid_lines"):
                    x0, x1 = self.canvas.coords(item)[1::2]
                    self.canvas.coords(item, x0, clinf_y, x1, lsup_y)

            viewport_y0 = y0
            pty0, pty1 =(clinf_x, coords_viewport[1]), (lsup_x, coords_viewport[1] - dy)
            pass
        elif deltay > 0:
            # bottom displacement
            dy = deltay
            items = self.canvas.find_enclosed(clinf_x - 1, coords_viewport[1] - 1, gx1 + 1, linf_y + 1)
            self.canvas.delete(*items)
            # Move the viewport dy pixel up
            items = self.canvas.find_enclosed(clinf_x - 1, linf_y - 1, gx1 + 1, gy1 + 1)
            for item in items:
                self.canvas.move(item, 0, -dy)
            
            # Resize the vertical gridlines
            if gy1 != lsup_y:
                for item in self.canvas.find_withtag("vgrid_lines"):
                    x0, x1 = self.canvas.coords(item)[1::2]
                    self.canvas.coords(item, x0, clinf_y, x1, lsup_y)

            viewport_y0 = self.cell_containing_coords(0, linf_y + 1, viewport, coords_viewport)[1]
            pty0 = clinf_x, lsup_y - dy
        if pty0:
            viewport_y1 = self.cell_containing_coords(0, winfo_height, (viewport_x0, viewport_y0), coords_viewport)[1]
            viewport_y0 = viewport_y0 if isViewportOrig else self.viewport_q1[1]
            self.viewport_q1 = (self.viewport_q1[0], viewport_y0, self.viewport_q1[2], viewport_y1)
        if ptx0:
            if ptx1 is None:
                ptx1 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
            area = ptx0 + ptx1
            self.canvas.tag_area(*area, tag="invalid_area")
        if pty0:
            if pty1 is None:
                pty1 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
            area = pty0 + pty1
            self.canvas.tag_area(*area, tag="invalid_area")
        return ptx0 or pty0

    def set_dimension(self, x0:int, x1:int, width:int, axis:Literal[0, 1]=0) -> int:
        """Sets the width of the columns in the range x0:x1 and returns the change in width."""
        prefix = 'C' if axis == 0 else 'R'
        default = (CELL_WIDTH, ROW_CELLS_HEIGHT)[axis]
        if width > 0:   # Set headings
            twidth1 = sum(self.headings_dim.get(f"{prefix}{x}", default) for x in range(x0, x1 + 1))
            to_update = {f"{prefix}{x}": width for x in range(x0, x1 + 1)}
            [self.headings_hided.pop(key, None) for key in to_update]
            self.headings_dim.update(to_update)
            twidth2 = sum(self.headings_dim.get(f"{prefix}{x}", default) for x in range(x0, x1 + 1))
            delta = twidth2 - twidth1
        elif width == 0:  # Hide headings
            to_update = {f"{prefix}{x}": self.headings_dim.pop(f"{prefix}{x}", default) for x in range(x0, x1 + 1)}
            self.headings_hided.update(to_update)
            self.headings_dim.update((key, 0) for key in to_update)
            delta = -sum(to_update.values())
        else:
            # Unhide headings
            to_update = {f"{prefix}{x}": self.headings_hided.pop(f"{prefix}{x}", default) for x in range(x0, x1 + 1)}
            [self.headings_dim.pop(key, None) for key in to_update]
            self.headings_dim.update([(key, value) for key, value in to_update.items() if value != default])        
            delta = sum(to_update.values())
        # Viewport rbcorner is updated
        rbcorner_vp = self.cell_containing_coords(self.canvas.efective_width(), self.canvas.efective_height())
        self.viewport_q1 = (*self.viewport_q1[:2], *rbcorner_vp)
        return delta
    
    def insert(self, x0:int, x1:int, axis:Literal[0, 1]=0) -> int:
        """Inserts (x1 - x0) headings with default dimension before heading x0."""
        prefix = 'C' if axis == 0 else 'R'
        default = (CELL_WIDTH, ROW_CELLS_HEIGHT)[axis]
        nheadings = x1 - x0 + 1
        to_update = {f"{prefix}{x + nheadings}": self.headings_dim.pop(key) for key in list(self.headings_dim) if key[0] == prefix and (x := int(key[1:])) >= x0}
        self.headings_dim.update(to_update)
        to_update = {f"{prefix}{x + nheadings}": self.headings_hided.pop(key) for key in list(self.headings_hided) if key[0] == prefix and (x := int(key[1:])) >= x0}
        self.headings_hided.update(to_update)
        return nheadings * default
    
    def delete(self, x0:int, x1:int, axis:Literal[0, 1]=0) -> int:
        """Deletes (x1 - x0) headings with default dimension before heading x0."""
        prefix = 'C' if axis == 0 else 'R'
        default = (CELL_WIDTH, ROW_CELLS_HEIGHT)[axis]
        nheadings = x1 - x0 + 1
        delta = sum(self.headings_dim.pop(f"{prefix}{x}", default) for x in range(x0, x1 + 1))
        to_update = {f"{prefix}{x - nheadings}": self.headings_dim.pop(key) for key in list(self.headings_dim) if key[0] == prefix and (x := int(key[1:])) > x1}
        self.headings_dim.update(to_update)
        delta += sum(self.headings_hided.pop(f"{prefix}{x}") for key in list(self.headings_hided) if key[0] == prefix and x0 <= (x:= int(key[1:])) <= x1)
        to_update = {f"{prefix}{x - nheadings}": self.headings_hided.pop(key) for key in list(self.headings_hided) if key[0] == prefix and (x := int(key[1:])) > x1}
        self.headings_hided.update(to_update)
        return -delta


class SheetUI(tk.Canvas):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.look = SheetLook(self, cell_content_gen)
        #flags
        self.f_drag = False  # Flag to indicate if a mouse drag is in progress
        self.f_gridlines = True  # Flag to indicate if gridlines are visible
        self.f_headings = True  # Flag to indicate if headings are visible
        self.f_freeze = False  # Flag to indicate if freeze panes are visible
        self.error_report = ""

        self.bind("<Configure>", self.redraw_sheet)
        self.bind("<Button-1>", self.on_mouse_click)
        self.bind("<B1-Motion>", self.on_mouse_drag)
        self.bind("<ButtonRelease-1>", self.on_mouse_release)
        self.bind("<MouseWheel>", self.on_mouse_wheel)  # Windows/macOS
        self.bind("<Button-4>", self.on_mouse_wheel)    # Linux scroll up
        self.bind("<Button-5>", self.on_mouse_wheel)    # Linux scroll down

        # bind arrow keys to move the active cell
        self.bind("<Up>", self.on_key_press)
        self.bind("<Down>", self.on_key_press)
        self.bind("<Left>", self.on_key_press)
        self.bind("<Right>", self.on_key_press)
        self.bind("<Return>", self.on_key_press)
        self.bind("<Tab>", self.on_key_press)
        self.bind("<Home>", self.on_key_press)
        self.bind("<Prior>", self.on_key_press)
        self.bind("<Next>", self.on_key_press)
        # self.bind("<Key>", self.on_key_press)
        self.focus_set()  # Set focus to the canvas

    def __getattr__(self, attr):
        "Delegate attribute access to the look object"
        if attr in self.look.__dir__():
            return getattr(self.look, attr)
        return super().__getattr__(attr)
    
    def reset_sheet(self):
        self.look = SheetLook(self, cell_content_gen)
        #flags
        self.f_drag = False  # Flag to indicate if a mouse drag is in progress
        self.f_gridlines = True  # Flag to indicate if gridlines are visible
        self.f_headings = True  # Flag to indicate if headings are visible
        self.f_freeze = False  # Flag to indicate if freeze panes are visible
        self.delete("all")
        self.redraw_sheet()

    def move_viewport(self, x, y):
        # if self.look.move_viewport(x, y):
        if self.look.map_cell_to_coords(x, y):
            items = self.find_withtag("invalid_area")
            areas = [self.coords(item) for item in items]
            logging.debug(f"Invalidated area: {[self.look.area_cells(*area) for area in areas]}")
            self.setGUI()  # Redraw the sheet with the new viewport
            self.tag_raise("freeze_line")  # Move freeze_line above all tags
    
    def efective_width(self):
        return self.winfo_width() + int(self.f_headings) * COL_CELLS_WIDTH
    
    def efective_height(self):
        return self.winfo_height() + int(self.f_headings) * ROW_CELLS_HEIGHT

    def draw_cell_content(self, box: tuple[int, int, int, int], cell_content:str, **kwargs):
        x0, y0, x1, y1 = box
        items = [item for item in self.find_enclosed(x0, y0, x1, y1) if self.type(item) == "text"]
        if items:
            old_txt = self.itemcget(items[0], "text")
            logging.debug(f"replacing {old_txt} with {cell_content}")
            self.error_report += f" {old_txt}"
        tid = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text=cell_content, anchor="center", **kwargs)
        tx0, tx1 = self.bbox(tid)[::2]
        if (tx1 - tx0) > (x1 - x0):
            self.delete(tid)
            tid = self.create_text((x0 + x1) // 2, (y0 + y1) // 2, text="*", **kwargs)

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
            items = [item for item in self.find_overlapping(cx0, cy0, cx1, cy1) if "cells_to_draw" in self.gettags(item)]
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
            [self.delete(item) for atag in ("cols_drawn", "rows_drawn", "cells_drawn") for item in self.find_withtag(atag)]
        return self.create_rectangle(*area, **kwargs)

    def setGUI(self):
        winfo_width, winfo_height = self.efective_width(), self.efective_height()
        self.validate_areas()

        # Draw the background
        linf_coordx, linf_coordy = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        lsup_coordx, lsup_coordy = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        self.coords("background", linf_coordx, linf_coordy, lsup_coordx, lsup_coordy)
        self.tag_lower("background")  # Ensure the background is at the bottom of the stack
        # Draw column headings
        for item in self.find_withtag("cols_to_draw"):
            cx0, cy0, cx1, cy1 = map(int, self.coords(item))
            xcell = 1
            while cx0 < cx1 and xcell < MAX_COLS:
                nquadrant = self.cell_quadrant(cx0, cy1, isCoord=True)
                orig, coords_orig = self.quadrant_data(nquadrant)
                xcell = self.cell_containing_coords(cx0, cy0, orig, coords_orig)[0]
                y0, y1 = cy0, cy1
                x0, x1 = self.cell_coordinates(xcell, 0, orig, coords_orig)[::2]
                self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="column")
                # Draw cell headings
                self.draw_cell_content((x0, y0, x1, y1), f"C{xcell}", fill="white", tags="columns_tag")
                # Draw vertical lines
                self.create_line(x0, y0, x0, winfo_height, fill=GRID_COLOR, tags="vgrid_lines")
                cx0 = x1
            assert xcell >= MAX_COLS or cx0 == cx1
            logging.debug(f"Last column draw {xcell}")
            self.itemconfig(item, tags="cols_drawn", state="hidden")
            pass

        # Draw row headings
        for item in self.find_withtag("rows_to_draw"):
            cx0, cy0, cx1, cy1 = map(int, self.coords(item))
            ycell = 1
            while cy0 < cy1 and ycell < MAX_ROWS:
                nquadrant = self.cell_quadrant(cx1, cy0, isCoord=True)
                orig, coords_orig = self.quadrant_data(nquadrant)
                ycell = self.cell_containing_coords(cx0, cy0, orig, coords_orig)[1]
                x0, x1 = cx0, cx1
                y0, y1 = self.cell_coordinates(0, ycell, orig, coords_orig)[1::2]
                self.create_rectangle(x0, y0, x1, y1, fill="green", outline="black", tags="row")
                # Draw cell headings
                self.draw_cell_content((x0, y0, x1, y1), f"C{ycell}", fill="white", tags="rows_tag")
                # Draw horizontal lines
                self.create_line(x0, y1, winfo_width, y1, fill=GRID_COLOR, tags="hgrid_lines")
                cy0 = y1
            assert ycell >= MAX_ROWS or cy0 == cy1
            logging.debug(f"Last row draw {ycell}")
            self.itemconfig(item, tags="rows_drawn", state="hidden")
            pass

        # Draw cells content
        quadrants = [1, 2, 3, 4] if self.coords_vportq1 != self.coords_vportq3 else [1]
        for item in self.find_withtag("cells_to_draw"):
            ix0, iy0, ix1, iy1 = map(int, self.coords(item))
            assert tuple(map(min, zip((ix1, iy1), self.cell_coordinates(MAX_COLS, MAX_ROWS)[2:]))) == self.area_coordinates(*self.area_cells(ix0, iy0, ix1, iy1))[2:]
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
            self.itemconfig(item, tags="cells_drawn", state="hidden")
            pass
        # [self.tag_lower(tag) for tag in ("cols_drawn", "rows_drawn", "cells_drawn")]
        if logger.isEnabledFor(logging.DEBUG):
            logging.debug(sorted(Counter([self.itemcget(item, 'tags') for item in self.find_all()]).items()))
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
        clipping_rect = self.look.area_coordinates(*self.viewport_q3[:2], *self.viewport_q1[2:])
        # winfo_width, winfo_height = self.efective_width(), self.efective_height()
        # clipping_rect = (self.coords_vportq3[0], self.coords_vportq3[1], winfo_width, winfo_height)

        x0, y0, x1, y1 = self.area_coordinates(*self.selected_cells)
        sel_x0, sel_y0, sel_x1, sel_y1 = clip_rectangle(x0, y0, x1, y1, clipping_rect)
    
        self.create_rectangle(sel_x0, sel_y0, sel_x1, sel_y1,
            fill="lightblue", outline="black", tags="selected_cells")
        self.tag_lower("selected_cells", "vgrid_lines")
        self.tag_lower("selected_cells", "hgrid_lines")


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
        
    def set_freeze_lines(self):
        coord_acell_x, coord_acell_y = self.coords_vportq1
        winfo_width, winfo_height = self.efective_width(), self.efective_height()
        items = self.find_withtag("freeze_line")
        if not items and coord_acell_y != self.coords_vportq3[1]:
            linf_x = self.coords_vportq3[0] - 2 * COL_CELLS_WIDTH
            self.create_line(linf_x, coord_acell_y, winfo_width, coord_acell_y, fill="black", tags="freeze_line")
        
        if not items and coord_acell_x != self.coords_vportq3[0]:
            linf_y = self.coords_vportq3[1] - 2 * ROW_CELLS_HEIGHT
            self.create_line(coord_acell_x, linf_y, coord_acell_x, winfo_height, fill="black", tags="freeze_line")

    def set_rows_height(self, height):
        """Sets the height of the rows in the range y0:y1 and returns the change in height"""
        height = max(-1, height)
        sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
        if (sel_x0, sel_x1) != (1, MAX_COLS):
            return
        linf_y, lsup_y = self.look.area_coordinates(sel_x0, sel_y0, sel_x1, sel_y1)[1::2]
        clinf_x = self.coords_vportq3[0] - COL_CELLS_WIDTH
        vplsup_x0, vplsup_y0 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        to_delete = (clinf_x - 1, linf_y - 1, vplsup_x0 + 1, lsup_y + 1)
        to_move = (clinf_x - 1, linf_y - 1, vplsup_x0 + 1, vplsup_y0 + 1)

        delta = self.look.set_dimension(sel_y0, sel_y1, height, axis=1)

        self.delete(*self.find_enclosed(*to_delete))
        for item in self.find_enclosed(*to_move):
            self.move(item, 0, delta)

        area = clinf_x, linf_y, vplsup_x0, lsup_y + delta
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        vplsup_x1, vplsup_y1 = self.look.cell_coordinates(*self.viewport_q1[2:])[2:]
        assert vplsup_x0 == vplsup_x1
        if vplsup_y0 + delta > vplsup_y1:
            to_delete = (clinf_x - 1, vplsup_y1 - 1, vplsup_x1 + 1, vplsup_y0 + delta + 1)
            self.delete(*self.find_enclosed(*to_delete))
        else:
            area = (clinf_x, vplsup_y0 + delta, vplsup_x1, vplsup_y1)
            self.tag_area(*area, tag="invalid_area")
            logging.debug(f"Invalidated area: {self.area_cells(*area)}")
        self.setGUI()
        self.set_active_cell()

    def insert_rows(self):
        """Inserts (y1 - y0) headings with default dimension before heading y0."""
        sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
        if (sel_x0, sel_x1) != (1, MAX_COLS):
            return
        linf_y, lsup_y = self.look.area_coordinates(sel_x0, sel_y0, sel_x1, sel_y1)[1::2]
        clinf_x = self.coords_vportq3[0] - COL_CELLS_WIDTH
        vplsup_x0, vplsup_y0 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        to_delete = (clinf_x - 1, linf_y - 1, vplsup_x0 + 1, lsup_y + 1)
        to_move = (clinf_x - 1, linf_y - 1, vplsup_x0 + 1, vplsup_y0 + 1)

        delta = self.look.insert(sel_y0, sel_y1, axis=1)

        self.delete(*self.find_enclosed(*to_delete))
        for item in self.find_enclosed(*to_move):
            self.move(item, 0, delta)

        area = clinf_x, linf_y, vplsup_x0, lsup_y + delta
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        vplsup_x1, vplsup_y1 = self.look.cell_coordinates(*self.viewport_q1[2:])[2:]
        assert vplsup_x0 == vplsup_x1
        to_delete = (clinf_x - 1, vplsup_y1 - 1, vplsup_x1 + 1, vplsup_y0 + delta + 1)
        self.delete(*self.find_enclosed(*to_delete))

        to_delete = (clinf_x - 1, linf_y - 1, self.coords_vportq3[0] + 1, vplsup_y1 + 1)
        self.delete(*self.find_enclosed(*to_delete))

        area = (clinf_x, linf_y, self.coords_vportq3[0], vplsup_y1)
        self.tag_area(*area, tag="invalid_area")

        self.setGUI()
        self.set_active_cell()

    def delete_rows(self):
        """Deletes the rows in the range y0:y1 and returns the change in height."""
        sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
        if (sel_x0, sel_x1) != (1, MAX_COLS):
            return
        linf_y, lsup_y = self.look.area_coordinates(sel_x0, sel_y0, sel_x1, sel_y1)[1::2]
        clinf_x = self.coords_vportq3[0] - COL_CELLS_WIDTH
        vplsup_x0, vplsup_y0 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        # Marks for deletion the rows from sel_y0 to sel_y1
        to_delete = (clinf_x - 1, linf_y - 1, vplsup_x0 + 1, lsup_y + 1)

        # MArks for movement the rows from sel_y1 < y < viewport_y1
        to_move = (clinf_x - 1, lsup_y - 1, vplsup_x0 + 1, vplsup_y0 + 1)
        delta = self.look.delete(sel_y0, sel_y1, axis=1)

        self.delete(*self.find_enclosed(*to_delete))
        for item in self.find_enclosed(*to_move):
            self.move(item, 0, delta)

        # Updates the coordinates for the viewport brcorner 
        vplsup_x1, vplsup_y1 = self.look.cell_coordinates(*self.viewport_q1[2:])[2:]
        assert vplsup_x0 == vplsup_x1

        # Deletes the row headings from linf_y to vplsup_y1 and invalidates the area
        area = (clinf_x, linf_y, self.coords_vportq1[0] + 1, vplsup_y1 + 1)
        self.delete(*self.find_enclosed(*area))
        area = (clinf_x, linf_y, self.coords_vportq1[0], vplsup_y1)
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        # Invalidates the area leave blank by the move of cell content move previously.
        area = (clinf_x, vplsup_y0 + delta, vplsup_x1, vplsup_y1)        
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        self.setGUI()
        self.set_active_cell()

    def set_cols_width(self, width):
        """Sets the width of the columns in the range x0:x1 and returns the change in width."""
        width = max(-1, width)
        sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
        if (sel_y0, sel_y1) != (1, MAX_ROWS):
            return
        linf_x, lsup_x = self.look.area_coordinates(sel_x0, sel_y0, sel_x1, sel_y1)[::2]
        clinf_y = self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        vplsup_x0, vplsup_y0 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        to_delete = (linf_x - 1, clinf_y - 1, lsup_x + 1, vplsup_y0 + 1)
        to_move = (lsup_x - 1, clinf_y - 1, vplsup_x0 + 1, vplsup_y0 + 1)

        delta = self.look.set_dimension(sel_x0, sel_x1, width)

        self.delete(*self.find_enclosed(*to_delete))
        for item in self.find_enclosed(*to_move):
            self.move(item, delta, 0)

        area = linf_x, clinf_y, lsup_x + delta, vplsup_y0
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        vplsup_x1, vplsup_y1 = self.look.cell_coordinates(*self.viewport_q1[2:])[2:]
        assert vplsup_y0 == vplsup_y1
        if vplsup_x0 + delta > vplsup_x1:
            to_delete = (vplsup_x1 - 1, clinf_y - 1, vplsup_x0 + delta + 1, vplsup_y1 + 1)
            self.delete(*self.find_enclosed(*to_delete))
        else:
            area = (vplsup_x0 + delta, clinf_y, vplsup_x1, vplsup_y1)
            self.tag_area(*area, tag="invalid_area")
            logging.debug(f"Invalidated area: {self.area_cells(*area)}")
        self.setGUI()
        self.set_active_cell()

    def insert_columns(self):
        """Inserts (x1 - x0) headings with default dimension before heading x0."""
        sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
        if (sel_y0, sel_y1) != (1, MAX_ROWS):
            return
        linf_x, lsup_x = self.look.area_coordinates(sel_x0, sel_y0, sel_x1, sel_y1)[::2]
        clinf_y = self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        vplsup_x0, vplsup_y0 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        to_move = (linf_x - 1, clinf_y - 1, vplsup_x0 + 1, vplsup_y0 + 1)

        delta = self.look.insert(sel_x0, sel_x1)

        for item in self.find_enclosed(*to_move):
            self.move(item, delta, 0)
        area = linf_x, clinf_y, linf_x + delta, vplsup_y0
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        vplsup_x1, vplsup_y1 = self.look.cell_coordinates(*self.viewport_q1[2:])[2:]
        assert vplsup_y0 == vplsup_y1
        to_delete = (vplsup_x1 - 1, clinf_y - 1, vplsup_x0 + delta + 1, vplsup_y1 + 1)
        self.delete(*self.find_enclosed(*to_delete))

        to_delete = (linf_x - 1, clinf_y - 1, vplsup_x1 + 1, self.coords_vportq3[1] + 1)
        self.delete(*self.find_enclosed(*to_delete))

        area = linf_x, clinf_y, vplsup_x1, self.coords_vportq3[1]
        self.tag_area(*area, tag="invalid_area")

        self.setGUI()
        self.set_active_cell()

    def delete_columns(self):
        """Deletes the columns in the range x0:x1 and returns the change in width."""
        sel_x0, sel_y0, sel_x1, sel_y1 = self.selected_cells
        if (sel_y0, sel_y1) != (1, MAX_ROWS):
            return
        linf_x, lsup_x = self.look.area_coordinates(sel_x0, sel_y0, sel_x1, sel_y1)[::2]
        clinf_y = self.coords_vportq3[1] - ROW_CELLS_HEIGHT
        vplsup_x0, vplsup_y0 = self.cell_coordinates(*self.viewport_q1[2:])[2:]
        # Marks for deletion the columns from sel_x0 to sel_x1
        to_delete = (linf_x - 1, clinf_y - 1, lsup_x + 1, vplsup_y0 + 1)

        # MArks for movement the columns from sel_x1 < x < viewport_x1
        to_move = (lsup_x - 1, clinf_y - 1, vplsup_x0 + 1, vplsup_y0 + 1)
        delta = self.look.delete(sel_x0, sel_x1)

        self.delete(*self.find_enclosed(*to_delete))
        for item in self.find_enclosed(*to_move):
            self.move(item, delta, 0)

        # Updates the coordinates for the viewport brcorner 
        vplsup_x1, vplsup_y1 = self.look.cell_coordinates(*self.viewport_q1[2:])[2:]
        assert vplsup_y0 == vplsup_y1

        # Deletes the column headings from linf_x to vplsup_x1 and invalidates the area
        area = (linf_x - 1, clinf_y - 1, vplsup_x1 + 1, self.coords_vportq3[1] + 1)
        self.delete(*self.find_enclosed(*area))
        area = (linf_x, clinf_y, vplsup_x1, self.coords_vportq3[1])
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        # Invalidates the area leave blank by the move of cell content move previously.
        area = (vplsup_x0 + delta, clinf_y, vplsup_x1, vplsup_y1)
        self.tag_area(*area, tag="invalid_area")
        logging.debug(f"Invalidated area: {self.area_cells(*area)}")

        self.setGUI()
        self.set_active_cell()

    def redraw_sheet(self, event=None, width=None, height=None):
        "Redraws the sheetui when the window is resized or needs updating."

        if self.find_withtag("background"):
            bg_coords = self.coords("background")
        else:
            bgc_x0, bgc_x1 = self.coords_vportq3[0] - COL_CELLS_WIDTH, self.coords_vportq3[0]
            bgc_y0, bgc_y1 = self.coords_vportq3[1] - ROW_CELLS_HEIGHT, self.coords_vportq3[1]
            bg_coords = (bgc_x0, bgc_y0, bgc_x1, bgc_y1)
            self.create_rectangle(*bg_coords, fill="green", outline="black", tags="corner")
            self.create_rectangle(*bg_coords, fill="white", outline="black", tags="background")
        
        bgc_x0, bgc_y0, bgc_x1, bgc_y1 = bg_coords

        winfo_width = self.efective_width()
        winfo_height = self.efective_height()
        viewport_x1, viewport_y1 = self.cell_containing_coords(winfo_width, winfo_height)
        self.look.viewport_q1 = self.viewport_q1[:2] + (viewport_x1, viewport_y1)

        clsup_x, clsup_y =self.cell_coordinates(viewport_x1, viewport_y1)[2:]
        
        # Adjust the gridlines to the new viewport
        items = self.find_withtag("hgrid_lines") + self.find_withtag("vgrid_lines") + self.find_withtag("freeze_line")
        for item in items:
            gx0, gy0, gx1, gy1 = self.coords(item)
            if gy0 == gy1:
                # Horizontal gridlines
                self.coords(item, gx0, gy0, clsup_x, gy1)
            else:
                # Vertical gridlines
                self.coords(item, gx0, gy0, gx1, clsup_y)

        items = self.find_withtag("cells_drawn")
        bflag = bool(items) and self.itemcget(items[0], "state") == "normal"
        if not bflag:
            # Columns to draw
            if bgc_x1 < clsup_x:
                area = (bgc_x1, self.coords_vportq3[1] - ROW_CELLS_HEIGHT, clsup_x, self.coords_vportq3[1])
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
            else:
                area = (clsup_x, self.coords_vportq3[1] - ROW_CELLS_HEIGHT, bgc_x1, self.coords_vportq3[1])
                self.delete(*self.find_enclosed(*area))
            # Rows to draw
            if bgc_y1 < clsup_y:
                area = (self.coords_vportq3[0] - COL_CELLS_WIDTH, bgc_y1, self.coords_vportq3[0], clsup_y)
                self.tag_area(*area, tag="invalid_area")
                logging.debug(f"Invalidated area: {self.area_cells(*area)}")
            else:
                area = (self.coords_vportq3[0] - COL_CELLS_WIDTH, clsup_y, self.coords_vportq3[0], bgc_y1)
                self.delete(*self.find_enclosed(*area))

            # Cells to draw
            areas = []
            if (bgc_x1, bgc_y1) < (clsup_x, clsup_y):
                areas.append((bgc_x1, bgc_y1, clsup_x, clsup_y))
            else:
                area = (clsup_x - 1, clsup_y - 1, bgc_x1 + 1, bgc_y1 + 1)
                self.delete(*self.find_enclosed(*area))

            if bgc_x1 < clsup_x:
                if bgc_y1 - bgc_y0 > ROW_CELLS_HEIGHT:  
                    areas.append((bgc_x1, self.coords_vportq3[1], clsup_x, bgc_y1))
            else:
                area = (clsup_x - 1, self.coords_vportq3[1] -1, bgc_x1 + 1, bgc_y1 + 1)
                self.delete(*self.find_enclosed(*area))
                
            if bgc_y1 < clsup_y:
                if bgc_x1 - bgc_x0 > COL_CELLS_WIDTH:
                    areas.append((self.coords_vportq3[0], bgc_y1, bgc_x1, clsup_y))
            else:
                area = (self.coords_vportq3[0] - 1, clsup_y - 1, bgc_x1 + 1, bgc_y1 + 1)
                self.delete(*self.find_enclosed(*area))
            if areas:
                for area in areas:
                    self.tag_area(*area, tag="invalid_area")
                    logging.debug(f"Invalidated area: {self.area_cells(*area)}")
                self.tag_raise("invalid_area")
                self.setGUI()
            else:
                self.coords("background", bgc_x0, bgc_y0, clsup_x, clsup_y)

            self.xview('scroll', '-1', 'units')
            self.yview('scroll', '-1', 'units')
            self.set_active_cell()

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

    def show_cell(self, xin, yin):
        winfo_width, winfo_height = map(int, (self.winfo_width(), self.winfo_height()))
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

    def toggle_areas_drawn(self):
        drawn = list(map(self.find_withtag, ("cells_drawn", "cols_drawn", "rows_drawn")))
        for items in drawn:
            if items:
                ndx = int(self.itemcget(items[0], "state") == "normal")
                state = ("normal", "hidden")[ndx]
                for item in items:
                    self.itemconfig(item, state=state)
        pass

    def toggle_headings(self):
        """Toggles the visibility of headings."""
        if self.f_headings:
            dx, dy = -COL_CELLS_WIDTH, -ROW_CELLS_HEIGHT
        else:
            dx, dy = COL_CELLS_WIDTH, ROW_CELLS_HEIGHT
        lsup_x, lsup_y = self.cell_coordinates(self.viewport_q1[2], self.viewport_q1[3])[2:]
        self.look.coords_vportq1 = (self.coords_vportq1[0] + dx, self.coords_vportq1[1] + dy)
        self.look.coords_vportq3 = linf_x, linf_y = (self.coords_vportq3[0] + dx, self.coords_vportq3[1] + dy)
        items = self.find_enclosed(-linf_x - 1, -linf_y - 1, lsup_x + 1, lsup_y + 1)
        if self.f_freeze:
            items += self.find_withtag("freeze_line")
        for item in items:
            self.move(item, dx, dy)
        self.f_headings = not self.f_headings

    def toggle_gridlines(self):
        """Toggles the visibility of gridlines."""
        if self.f_gridlines:
            [self.itemconfig(item, state="hidden") for item in self.find_withtag("vgrid_lines")]
            [self.itemconfig(item, state="hidden") for item in self.find_withtag("hgrid_lines")]
        else:
            [self.itemconfig(item, state="normal") for item in self.find_withtag("vgrid_lines")]
            [self.itemconfig(item, state="normal") for item in self.find_withtag("hgrid_lines")]
        self.f_gridlines = not self.f_gridlines
    
    def toggle_freeze_panes(self, *args):
        if not self.f_freeze:
            x0, y0, x1, y1 = self.viewport_q3
            coord_acx, coord_acy = self.cell_coordinates(*self.active_cell)[:2]
            if self.active_cell[0] != self.viewport_q1[0]:
                x0, x1 = self.viewport_q1[0], self.active_cell[0]
                self.look.coords_vportq1 = coord_acx, self.coords_vportq1[1]
            if self.active_cell[1] != self.viewport_q1[1]:
                y0, y1 = self.viewport_q1[1], self.active_cell[1]
                self.look.coords_vportq1 = self.coords_vportq1[0], coord_acy
            self.look.viewport_q3 = (x0, y0, x1, y1)
            self.look.viewport_q1 = *self.active_cell, *self.viewport_q1[2:]
            self.set_freeze_lines()
            self.xview_moveto(0.0)
            self.yview_moveto(0.0)
        else:
            # If freeze is active, unfreeze the panes
            self.move_viewport(*self.viewport_q3[2:])
            self.look.coords_vportq1 = self.coords_vportq3
            self.look.viewport_q1 = *self.viewport_q3[:2], *self.viewport_q1[2:]
            self.look.viewport_q3 = 1, 1, 1, 1
            items = self.find_withtag("freeze_line")
            self.delete(*items)
        self.f_freeze = not self.f_freeze

    def on_key_press(self, event):
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
            self.xview_moveto(viewport_x0)
            self.yview_moveto(viewport_y0)
            self.set_active_cell()
            return "break"
        elif event.keysym in ("Next", "Prior"):
            viewport_x0, viewport_y0 = self.viewport_q1[:2]
            acell_x0, acell_y0 = self.active_cell
            with self.pivot_point(isActiveCell=not event.state & SHIFT_PRESSED) as pivot:
                coord_pivot_x, coord_pivot_y = self.cell_coordinates(pivot.x, pivot.y)[:2]
                nquadrant = self.cell_quadrant(pivot.x, pivot.y, isCoord=False)
                if event.state & ALT_PRESSED:
                    if nquadrant in (4, 3):
                        coord_pivot_x = self.coords_vportq1[0]
                    self.xview('scroll', -1 if event.keysym == 'Prior' else 1, 'pages')
                else:
                    if nquadrant in (2, 3):
                        coord_pivot_y = self.coords_vportq1[1]
                    self.yview('scroll', -1 if event.keysym == 'Prior' else 1, 'pages')
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
                self.look.active_cell = (acell_x0, acell_y0)
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
                self.look.active_cell = (acell_x0, acell_y0)
                self.set_active_cell()
                return "break"
            else:
                event.keysym = "Left" if event.state & SHIFT_PRESSED else "Right"  # Treat Return as Down for consistency
                event.state = 0

        n = ("Down", "Up", "Right", "Left").index(event.keysym)
        m = (-1) ** n
        dx, dy = (0, m) if n in (0, 1) else (m, 0)
        self.offset_acell(dx, dy, state=event.state)
        return "break"  # Prevent default behavior of arrow keys
    
    def on_mouse_click(self, event):
        """Sets the active cell based on the click position."""
        self.f_drag = True
        # Check if the click is not on an existing cell
        items = self.find_overlapping(event.x, event.y, event.x, event.y)
        if not items:
            return "break"
        if event.x < COL_CELLS_WIDTH and event.y < ROW_CELLS_HEIGHT:
            self.selected_cells = (1, 1, MAX_COLS, MAX_ROWS)
            self.look.active_cell = self.viewport_q1[:2]
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
            self.look.selected_cells = 1, sel_y0, MAX_COLS, sel_y1
            if not event.state & SHIFT_PRESSED:
                self.look.active_cell = (self.viewport_q1[0], clk_y)
        elif col_clk := event.y < ROW_CELLS_HEIGHT: # and event.x >= COL_CELLS_WIDTH:
            sel_x0, sel_x1 = self.selected_cells[::2]
            self.look.selected_cells = sel_x0, 1, sel_x1, MAX_ROWS
            if not event.state & SHIFT_PRESSED:
                self.look.active_cell = (clk_x, self.viewport_q1[1])
        # if (row_clk or col_clk) and not event.state & SHIFT_PRESSED:
        #     self.active_cell = (clk_x, self.viewport_q1[1])
        self.set_active_cell()
        self.focus_set()  # Set focus to the canvas

    def on_mouse_drag(self, event):
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
                    self.after(1000, self.on_mouse_drag, event)  # Repeat the drag event after a delay
                else:
                    # Column selection
                    nquadrant = self.cell_quadrant(event_x, ROW_CELLS_HEIGHT)
                    orig, coords_orig = self.quadrant_data(nquadrant)
                    clk_x, clk_y = self.cell_containing_coords(event_x, ROW_CELLS_HEIGHT, orig, coords_orig)
                    logging.debug(f"{event_x=}, {self.winfo_width()=}")
                    if event_x > self.winfo_width():
                        self.show_cell(clk_x, clk_y)
                        self.after(1000, self.on_mouse_drag, None)
                    event_y = ROW_CELLS_HEIGHT - 1
                    event = tk.Event()
                    event.x = event_x
                    event.y = event_y
                    event.state = SHIFT_PRESSED
                    self.on_mouse_click(event)
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
                    self.after(1000, self.on_mouse_drag, event)
                else:
                    # Row selection
                    nquadrant = self.cell_quadrant(COL_CELLS_WIDTH, event_y)
                    orig, coords_orig = self.quadrant_data(nquadrant)
                    clk_x, clk_y = self.cell_containing_coords(COL_CELLS_WIDTH, event_y, orig, coords_orig)
                    logging.debug(f"{event_y=}, {self.winfo_height()=}")
                    if event_y > self.winfo_height():
                        self.show_cell(clk_x, clk_y)
                        self.after(1000, self.on_mouse_drag, None)
                    event_x = COL_CELLS_WIDTH - 1
                    event = tk.Event()
                    event.x = event_x
                    event.y = event_y
                    event.state = SHIFT_PRESSED
                    self.on_mouse_click(event)
            else:
                # mouse over corners
                viewport_x0, viewport_y0 = self.viewport_q1[:2]
                event = tk.Event()
                if self.selected_cells[::2] == (1, MAX_COLS):
                    # Column selection
                    self.move_viewport(viewport_x0, viewport_y0 - 1)
                    event.x, event.y = COL_CELLS_WIDTH - 1, ROW_CELLS_HEIGHT
                    self.on_mouse_click(event)
                elif self.selected_cells[1::2] == (1, MAX_ROWS):
                    # Row selection
                    self.move_viewport(viewport_x0 - 1, viewport_y0)
                    event.x, event.y = COL_CELLS_WIDTH, ROW_CELLS_HEIGHT - 1
                    self.on_mouse_click(event)
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

    def on_mouse_release(self, event):
        self.f_drag = False
        pass

    def on_mouse_wheel(self, event):
        logging.debug(f"Mouse wheel:{event=}, {event.delta=}")
        delta = -1 if event.delta > 0 else 1
        fnc = self.xview if event.state & SHIFT_PRESSED else self.yview
        fnc("scroll", delta, 'units')

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


class SheetViewer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.f_rec = False
        self.action_stack = []
        self.fnc_to_test = [
            "choose an action", 
            "delete_rows", "insert_rows", "set_rows_height", 
            "delete_columns", "insert_columns", "set_cols_width", 
            "toggle_areas_drawn", "toggle_headings", 
            "toggle_gridlines", "toggle_freeze_panes", 
            "show_cell", "move_viewport"
        ]
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
        tk.Event
        # Action string
        sevent = sevent.strip('<>').replace(' event ', ' ')
        eseq, *kwargs = sevent.split()
        kwargs = dict(item.split('=') for item in kwargs)
        if eseq == 'Configure':
            # Why?
            # <Configure> is a system event (not a user event like <Button-1> or <KeyPress>).
            # Tkinter/Tk will ignore attempts to generate it manually.
            if self.state() == "zoomed":
                saction = f'self.state("zoomed")'
            else:
                saction = f'self.state("normal")\nself.geometry("{self.geometry()}")'
        else:
            if eseq in ('ButtonPress', 'ButtonRelease',):
                eseq = f"{eseq}-{kwargs.pop('num')}"
            if 'state' in kwargs:
                mods = ('Shift', 'Lock', 'Control', 
                        'Mod1', 'Mod2', 'Mod3', 'Mod4', 'Mod5',
                        'Button1', 'Button2', 'Button3', 'Button4', 'Button5'
                    )
                val = []
                for mod in  kwargs['state'].split('|'):
                    try:
                        n = mods.index(mod)
                        val.append(f"0x{1 << n:05x}")
                    except ValueError:
                        val.append(mod)
                kwargs['state'] = '|'.join(val)

            for key in set(['keysym',]).intersection(kwargs.keys()):
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
        vals = self.fnc_to_test
        self.cbox = cbox = ttk.Combobox(frame, name="cbox", values=vals, state="readonly")
        cbox.set(vals[0])  # Set default value
        cbox.pack(side="left")
        cbox.bind("<<ComboboxSelected>>", self.on_combobox_change)  # <-- Bind the event here
        self.activeCell = ttk.Label(frame, text="Active Cell: ", background="magenta", font=("Arial", 10))
        self.activeCell.pack(side="left",expand=True, fill="x", padx=4)
        
        frame = ttk.Frame(self, name='actionfrm')
        frame.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 4))
        chkbtn = ttk.Checkbutton(frame, name="rec", text="Rec", command=lambda: self.action_cmds('rec'))
        chkbtn.pack(side="left")
        # txt = ttk.Label(frame, name="txt", width=60, text='.....')
        # txt.pack(side="left")
        btn = ttk.Button(frame, text="run", command=lambda: self.action_cmds('step'))
        btn.pack(side="left")
        btn = ttk.Button(frame, text="all", command=lambda: self.action_cmds('run'))
        btn.pack(side="left")
        btn = ttk.Button(frame, text="Reset", command=lambda: self.action_cmds('reset'))
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
                initialfile="current_bug.txt"
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
                logging.debug(f"Executing action: {action}")
                exec(action)
        elif cmd == 'step':
            exec(self.action_stack[-1])
        elif cmd == 'reset':
            # Put the canvas in a clean slate
            self.action_stack = []
            self.nametowidget('.errorfrm.txt')['text'] = "...."
            self.sheetui.reset_sheet()

    def on_combobox_change(self, event):
            fname = self.cbox.get()
            fnc = getattr(self.sheetui, fname)
            args = []
            if (sig:=inspect.signature(fnc)).parameters:
                # display a message box to get  the pivot cell coordinates
                msg1 = f"{fname} requires parameters.\n {fname}({', '.join(p.name for p in sig.parameters.values())}) \nPlease enter parameters as integers separated with commas in the format 'x, y'."
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


def main():
    root = SheetViewer()
    root.mainloop()

def look():
    look = SheetLook(None, None)
    for x in range(1, 10):
        x0, x1 = look.cell_coordinates(x, 1)[::2]
        xmed = (x0 + x1) // 2
        col, row = look.cell_containing_coords(xmed, 30)
        print(f'{x=}, ({x0=}, {x1=}), width={x1-x0}, {col=}')

if __name__ == "__main__":
    main()
    # look()
