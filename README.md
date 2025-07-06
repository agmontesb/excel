Ideas:
- It can be posible to determine the screen areas that needs redrawing, by using the background rectangle in which the area present are drawn leaving the rest to be redrawn.

Known issues:
- state:
    f_freeze_panes = False, f_headings = True, f_gridlines = True, active_cell = (1, 1).
  Action: <Control>+<Right> and <Control>+<Down> then <Control>+<Up>
  Bug: The content in the MAX_COLS is erased.
