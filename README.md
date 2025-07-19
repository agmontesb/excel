1. SheetUI:

1.1 Ideas:
    - Rethink the redraw method to be call whenever a change is made int the viewport coordinates (active_cell, selected_cells, viewport_q1, viewport_q3).

1.2 Known bugs:

1.3 Implemented ideas:
    - The invalidated zones, for areas that need to be redrawn.

1.3 Solved bugs:
    bug001 - state: f_freeze_panes = False, f_headings = True, f_gridlines = True, active_cell = (1, 1).
    Action: <Control>+<Right> and <Control>+<Down> then <Control>+<Up>
    Bug: The content in the MAX_COLS is erased.

    bug002- state: f_freeze_panes = False, f_headings = True, f_gridlines = True, active_cell = (1, 1).
    Action: <Alt>+<Next> and <Alt>+<Next> then <Control>+<Left>
    Bug: The horizontal gridlines disappear.

    bug003 - state: f_freeze_panes = False, f_headings = True, f_gridlines = True, active_cell = (3, 4).
    Action: toggle_freeze_panes then maximize the window
    Bug: The content in Q3 is erased and the content in Q4 is rewritten.




(variable) def tag_config(
    tagName: str,
    cnf: dict[str, Any] | None = None,
    *,
    background: str = ...,
    bgstipple: str = ...,
    borderwidth: _ScreenUnits = ...,
    border: _ScreenUnits = ...,
    elide: bool = ...,
    fgstipple: str = ...,
    font: _FontDescription = ...,
    foreground: str = ...,
    justify: Literal['left', 'right', 'center'] = ...,
    lmargin1: _ScreenUnits = ...,
    lmargin2: _ScreenUnits = ...,
    lmargincolor: str = ...,
    offset: _ScreenUnits = ...,
    overstrike: bool = ...,
    overstrikefg: str = ...,
    relief: _Relief = ...,
    rmargin: _ScreenUnits = ...,
    rmargincolor: str = ...,
    selectbackground: str = ...,
    selectforeground: str = ...,
    spacing1: _ScreenUnits = ...,
    spacing2: _ScreenUnits = ...,
    spacing3: _ScreenUnits = ...,
    tabs: Any = ...,
    tabstyle: Literal['tabular', 'wordprocessor'] = ...,
    underline: bool = ...,
    underlinefg: str = ...,
    wrap: Literal['none', 'char', 'word'] = ...
) -> (dict[str, tuple[str, str, str, Any, Any]] | None)
Configure a tag TAGNAME.

