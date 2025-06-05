''' Este módulo implementa la interfaz gráfica de usuario (GUI) para la hoja de cálculo.
    Proporciona una ventana principal con una barra de menú, una barra de herramientas y un área de trabajo 
    que consiste en un área de celdas y una barra de estado que permite cambiar a otras hojas de trabajo '
    además de la activa'
'''

import tkinter as tk
from tkinter import ttk

class WorksheetUI(tk.Frame):
    def __init__(self, root):
        super().__init__(root)  # Properly initialize the tk.Frame class
        self.root = root
        self.root.title("Hoja de Cálculo")
        
        # Crear barra de menú
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        self._create_menu()

        # Crear headings
        self.headings = WsHeadindg(self.root, height=10)
        self.headings.pack(side=tk.TOP, fill=tk.X)
        self.headings.propagate = False

        # Crear barra de herramientas
        self.toolbar = ttk.Frame(self.root, relief=tk.RAISED, borderwidth=1)
        self.toolbar.pack(side=tk.TOP, fill=tk.X)
        self._create_toolbar()

        # Crear área de trabajo
        self.workspace = ttk.Frame(self.root)
        self.workspace.pack(fill=tk.BOTH, expand=True)
        self._create_workspace()

        # Crear barra de estado
        self.status_bar = ttk.Label(self.root, text="Listo", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _create_menu(self):
        # Menú Archivo
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        file_menu.add_command(label="Nuevo")
        file_menu.add_command(label="Abrir")
        file_menu.add_command(label="Guardar")
        file_menu.add_separator()
        file_menu.add_command(label="Salir", command=self.root.quit)
        self.menu_bar.add_cascade(label="Archivo", menu=file_menu)

        # Menú Edición
        edit_menu = tk.Menu(self.menu_bar, tearoff=0)
        edit_menu.add_command(label="Copiar")
        edit_menu.add_command(label="Pegar")
        self.menu_bar.add_cascade(label="Edición", menu=edit_menu)

    def _create_toolbar(self):
        ttk.Button(self.toolbar, text="Nuevo").pack(side=tk.LEFT, padx=2, pady=2)
        ttk.Button(self.toolbar, text="Guardar").pack(side=tk.LEFT, padx=2, pady=2)

    def _create_workspace(self):
        # Crear área de celdas con scrollbars
        self.cell_frame = ttk.Frame(self.workspace)
        self.cell_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbars
        self.v_scroll = ttk.Scrollbar(self.cell_frame, orient=tk.VERTICAL)
        self.h_scroll = ttk.Scrollbar(self.cell_frame, orient=tk.HORIZONTAL)
        self.v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        # Crear área de celdas con columna de etiquetas de fila
        self.cell_area = ttk.Treeview(
            self.cell_frame,
            columns=("Row", "A", "B", "C"),
            show="headings",
            yscrollcommand=self.v_scroll.set,
            xscrollcommand=self.h_scroll.set
        )
        self.cell_area.heading("Row", text="#")
        self.cell_area.heading("A", text="A")
        self.cell_area.heading("B", text="B")
        self.cell_area.heading("C", text="C")
        self.cell_area.column("Row", width=50, anchor=tk.CENTER)  # Set width for row labels
        self.cell_area.pack(fill=tk.BOTH, expand=True)

        # Configure scrollbars
        self.v_scroll.config(command=self.cell_area.yview)
        self.h_scroll.config(command=self.cell_area.xview)

        # Populate rows for demonstration
        for i in range(1, 101):  # Add 100 rows for testing
            self.cell_area.insert("", "end", values=(i, "", "", ""))

        # Bind click event to handle cell selection
        self.cell_area.bind("<Button-1>", self._on_cell_click)

        # Crear barra de estado para cambiar hojas
        self.sheet_selector = ttk.Combobox(self.workspace, values=["Hoja1", "Hoja2", "Hoja3"])
        self.sheet_selector.pack(side=tk.BOTTOM, fill=tk.X)
        self.sheet_selector.set("Hoja1")

    def _on_cell_click(self, event):
        # Clear the default selection
        self.cell_area.selection_remove(self.cell_area.selection())

        # Identify the clicked cell
        region = self.cell_area.identify("region", event.x, event.y)
        if region == "cell":
            row_id = self.cell_area.identify_row(event.y)
            column_id = self.cell_area.identify_column(event.x)
            cell_value = self.cell_area.set(row_id, column_id)

            # Highlight the selected cell
            self._highlight_cell(row_id, column_id)

            print(f"Selected cell: Row {row_id}, Column {column_id}, Value: {cell_value}")

    def _highlight_cell(self, row_id, column_id):
        # Remove existing tags
        for item in self.cell_area.tag_has("highlight"):
            self.cell_area.item(item, tags="")

        # Apply a tag to the selected cell
        self.cell_area.tag_configure("highlight", background="lightblue")

        # Highlight only the specific cell by modifying its value temporarily
        original_value = self.cell_area.set(row_id, column_id)
        self.cell_area.set(row_id, column_id, f"[{original_value}]")  # Add brackets to indicate selection
        self.cell_area.item(row_id, tags=("highlight",))

        # Restore the original value after a short delay
        self.root.after(100, lambda: self.cell_area.set(row_id, column_id, original_value))


class WsHeadindg(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, relief=tk.RAISED, borderwidth=1, **kwargs)
        self.widths = [(1, 30), (4, 10), (6, 60), (9, 15)]
        self.setGUI()

    def setGUI(self):
        canvas = tk.Canvas(self, height=20, background='green')
        canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        t_width = 0
        for i, (col, width) in enumerate(self.widths):
            label = tk.Label(canvas, text=f"Col {col}", width=width, anchor=tk.W)
            label.grid(row=0, column=i, sticky=tk.W + tk.E + tk.N + tk.S)
            canvas.grid_columnconfigure(i, weight=1)
            canvas.create_rectangle(t_width, 0, t_width + width, 20, fill="black")
            t_width += width


class Sketchpad(tk.Canvas):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.color = None
        self.bind("<Button-1>", self.save_posn)
        self.bind("<B1-Motion>", self.add_line)
        self.setGUI()
        self.setColor("black")  # Set default color

    def setGUI(self):
        id = self.create_rectangle(10, 10, 30, 30, fill="red", tags=("palette", "palettered"))
        self.tag_bind(id, "<Button-1>", lambda  x: self.setColor("red"))
        id = self.create_rectangle(10, 35, 30, 55, fill="blue", tags=("palette", "paletteblue"))
        self.tag_bind(id, "<Button-1>", lambda x: self.setColor("blue"))
        id = self.create_rectangle(10, 60, 30, 80, fill="black", tags=("palette", "paletteblack", "paletteSelected"))
        self.tag_bind(id, "<Button-1>", lambda x: self.setColor("black"))

    def setColor(self, color):
        self.color = color
        self.dtag("all", "paletteSelected")
        self.itemconfigure("palette", outline="white", width=5)
        self.addtag("paletteSelected", "withtag", "palette" + color)
        self.itemconfigure("paletteSelected", outline="#999999")
        
    def save_posn(self, event):
        self.lastx, self.lasty = event.x, event.y

    def add_line(self, event):
        self.create_line((self.lastx, self.lasty, event.x, event.y), fill=self.color)
        self.save_posn(event)


if __name__ == "__main__":
    # root = tk.Tk()
    # app = WorksheetUI(root)
    # root.mainloop()

    root = tk.Tk()
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    sketch = Sketchpad(root)
    sketch.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))

    root.mainloop()
