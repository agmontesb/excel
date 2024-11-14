import tkinter as tk
from tkinter import ttk
import time
import threading
from tkinter import filedialog

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Listbox and Scrollbar")
        self.geometry("500x500")
        self.setUI()
        self.lst_values = ('', '', '')
        self.attr_keys = ['sheet_name', 'tbl_name', 'tbl_range']
        self.clipboard_list = []
        self.clipboard_thread = None
        self.stop_event = threading.Event()

        # Bind focus events
        self.bind("<FocusIn>", self.stop_monitor)
        # self.bind("<FocusOut>", self.start_monitor)

    def setUI(self):
        root = self
        self.top_frame = tk.Frame(root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X)
        self.button = tk.Button(self.top_frame, text="Start", command=self.start_btn)
        self.button.pack(side=tk.LEFT)
        self.top_label = tk.Label(self.top_frame, text="Hello World")
        self.top_label.pack(side=tk.LEFT, fill=tk.X)

        self.bottom_frame = tk.Frame(root)
        self.bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.bottom_label = tk.Label(self.bottom_frame, text="")  # Initialize with empty text
        self.bottom_label.pack(side=tk.LEFT, fill=tk.BOTH)

        # Add buttons for edit and delete
        self.edit_button = tk.Button(self.top_frame, text="Edit", command=self.edit_btn, state=tk.DISABLED)
        self.edit_button.pack(side=tk.LEFT)
        self.delete_button = tk.Button(self.top_frame, text="Delete", command=self.delete_btn, state=tk.DISABLED)
        self.delete_button.pack(side=tk.LEFT)

        # Add export button
        self.export_button = tk.Button(self.top_frame, text="Export", command=self.export_data)
        self.export_button.pack(side=tk.LEFT)

        self.left_frame = tk.Frame(root)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y)
        self.scrollbar_left = tk.Scrollbar(self.left_frame)
        self.scrollbar_left.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox = tk.Listbox(self.left_frame, font=("Consolas", 10))
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        self.listbox.bind("<<ListboxSelect>>", self.listbox_click)  # Add event binding for clicks
        self.listbox.bind('<Return>', self.edit_btn)
        self.scrollbar_left.config(command=self.listbox.yview)
        self.listbox.config(yscrollcommand=self.scrollbar_left.set)

        self.right_frame = tk.Frame(root)
        self.right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        self.scrollbar_right = tk.Scrollbar(self.right_frame)
        self.scrollbar_right.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_box = tk.Text(self.right_frame, font=("Consolas", 12))
        self.text_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=tk.YES)
        self.scrollbar_right.config(command=self.text_box.yview)  # Configured scrollbar for textbox
        self.text_box.config(yscrollcommand=self.scrollbar_right.set)  # Configured textbox to use the scrollbar


    def start_btn(self):
        btn = self.button
        cur_state = btn["text"]
        nxt_state = ["Start", "Stop"][cur_state == "Start"]
        btn['text'] = nxt_state

        if nxt_state == "Stop":
            self.start_monitor()
        else:
            self.stop_monitor()

    def monitor_clipboard(self):
        while not self.stop_event.is_set():
            try:
                clipboard_content = self.clipboard_get()
            except:
                print(".", end="")
            else:
                if (clipboard_content,) not in self.clipboard_list:
                    self.clipboard_list.append((clipboard_content,))
                    self.listbox.insert(tk.END, clipboard_content[:21])
                    self.top_label['text'] = f"{len(self.clipboard_list)} items in clipboard"
            time.sleep(1) # 1 second delay

    def listbox_click(self, event):
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            try:
                content, *values = self.clipboard_list[index]
                self.text_box.delete("1.0", tk.END)  # Clear the text box
                self.text_box.insert(tk.END, content)
                # Enable Edit and Delete buttons when an item is selected
                self.edit_button.config(state=tk.NORMAL)
                self.delete_button.config(state=tk.NORMAL)

                # Update bottom frame with the last four elements
                if values:
                    bottom_text = ", ".join(f"{key} = {value}" for key, value in zip(self.attr_keys, values))
                    self.bottom_label.config(text=bottom_text)

            except IndexError:
                print("Index out of bounds")
        else:
            # Disable Edit and Delete buttons when no item is selected
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
            self.bottom_label.config(text="")  # Clear bottom frame text

    def edit_btn(self, event=None):
        selected_index = self.listbox.curselection()
        if selected_index:
            index = selected_index[0]  # Get the index of the selected item
            current_value = self.listbox.get(index)

            # Create a dialog box for editing
            edit_window = tk.Toplevel(self)
            edit_window.title("Edit Item")

            content, *values = self.clipboard_list[index]
            values = values or self.lst_values

            # Create widgets and organize them in a grid layout
            row = 0
            for wdg_name, value in zip(self.attr_keys, values):
                edit_label = tk.Label(edit_window, text=f"{wdg_name}:")
                edit_label.grid(row=row, column=0, padx=5, pady=5)
                edit_entry = tk.Entry(edit_window, name=wdg_name)
                edit_entry.insert(0, value)  # Pre-fill with current value
                edit_entry.grid(row=row, column=1, padx=5, pady=5)
                row += 1

            # **Focus on the first Entry widget after creating it**
            edit_window.nametowidget(self.attr_keys[0]).focus_set()

            # Create Save button
            save_button = tk.Button(edit_window, text="Save", command=self.save_changes(index, edit_window))
            save_button.grid(row=row, column=0, padx=5, pady=5)

            # Create Cancel button
            cancel_button = tk.Button(edit_window, text="Cancel", command=edit_window.destroy)
            cancel_button.grid(row=row, column=1, padx=5, pady=5)

            # Bind Escape key to Cancel button
            edit_window.bind("<Escape>", lambda event: cancel_button.invoke())

    def save_changes(self, index, edit_window):
        def _save_changes():
                values = tuple(
                    wdg.get()
                    for wdg_name in self.attr_keys
                    if (wdg := edit_window.nametowidget(wdg_name))
                )
                content = self.clipboard_list[index][0]
                self.clipboard_list[index] = (content, *values)
                self.lst_values = values
                self.listbox.delete(index)
                self.listbox.insert(index, values[1])
                edit_window.destroy()
                # Set focus on the edited listbox item
                self.listbox.selection_set(index)
                self.listbox_click(None)
        return _save_changes

    def delete_btn(self):
        selection = self.listbox.curselection()
        if selection:
            index = selection[0]
            if 0 <= index < len(self.clipboard_list):  # Check if index is valid
                del self.clipboard_list[index]
                self.listbox.delete(index)
                index = max(0, min(index, len(self.clipboard_list)))
                self.listbox.selection_set(index)
                self.listbox_click(None)

    def export_data(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            tpl_str = ''
            wb_struct = []
            with open(file_path, "w") as f:
                for item in self.clipboard_list:
                    if len(item) == 4:  # Check if the tuple has 4 elements
                        wb_struct.append(f'    {item[1:]}')
                        tpl_str += f'{item[2]} = """{item[0].replace(",", ".")}"""\n\n'
                wb_struct = ",\n".join(wb_struct)
                wb_struct = f'\nwb_structure = [\n{wb_struct}\n]\n\n'
                f.write(wb_struct)
                f.write(tpl_str)

    def stop_monitor(self, event=None):
        # Stop the thread when the window gets focus
        if self.clipboard_thread:
            self.stop_event.set()
            self.clipboard_thread.join()
            print("\nThread stopped")
            self.clipboard_thread = None
            self.button["text"] = "Start"

    def start_monitor(self, event=None):
        # Start the thread when the window loses focus
        if not self.clipboard_thread:
            self.clipboard_thread = threading.Thread(target=self.monitor_clipboard)
            self.clipboard_thread.start()
            self.iconify()

app = App()
app.mainloop()

