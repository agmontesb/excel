1. SheetUI:

1.1 Ideas:
    - It can be posible to determine the screen areas that needs redrawing, by using the background rectangle in which the area present are drawn leaving the rest to be redrawn.

1.2 Known issues:
    - state: f_freeze_panes = False, f_headings = True, f_gridlines = True, active_cell = (1, 1).
    Action: <Control>+<Right> and <Control>+<Down> then <Control>+<Up>
    Bug: The content in the MAX_COLS is erased.

    - state: f_freeze_panes = False, f_headings = True, f_gridlines = True, active_cell = (1, 1).
    Action: <Alt>+<Next> and <Alt>+<Next> then <Control>+<Left>
    Bug: The horizontal gridlines disappear.

    - state: f_freeze_panes = False, f_headings = True, f_gridlines = True, active_cell = (3, 4).
    Action: toggle_freeze_panes then maximize the window
    Bug: The content in Q3 is erased and the content in Q4 is rewritten.

