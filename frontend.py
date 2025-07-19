import sys
import traceback
import tkinter as tk


class Frontend(tk.Frame):
    def __init__(self, *args, context = None, **kwargs):
        super().__init__(*args, **kwargs)

        self.context = context or {}

        self.input = wdg = tk.Text(self, name='input', height=10, font=('Courier', 11))
        wdg.pack(side=tk.BOTTOM, fill=tk.X)
        wdg.bind("<Return>", self.on_input)
        wdg.bind("<Control-Return>", self.on_input)
        wdg.bind('<BackSpace>', self.backspace)
        wdg.bind('<Delete>', self.delete)
        wdg.bind('<Up>', self.history)
        wdg.bind('<Down>', self.history)
        wdg.bind('<Control-c>', self.clear_prompt)
        wdg.bind('<braceleft>', self.braceleft)
        wdg.event_add('<<Caret>>', '<End>', '<Home>', '<Right>', '<Left>')
        wdg.bind('<<Caret>>', self.caret)
        # wdg.bind('<Key>', self.Key)
        
        self.output = wdgo = tk.Text(self, name='output', font=('Courier', 11))
        wdgo.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        wdgo.tag_config('separator', background='black', font=('Courier', 5))
        wdgo.tag_config('cell', background='white')
        wdgo.tag_config('active_cell', background='white', relief=tk.GROOVE, borderwidth=3, lmargincolor='blue')
        wdgo.tag_config('inner_separator', background='white', font=('Courier', 5), lmargin1=0)
        wdgo.tag_config('in_id', foreground='red', lmargin1=20)
        wdgo.tag_config('input', background='lightblue', borderwidth=0)
        wdgo.tag_config('output', foreground='black')
        wdgo.tag_config('out_id', foreground='green', lmargin1=20)

        wdgo.bind('<Up>', self.on_active_cell)
        wdgo.bind('<Down>', self.on_active_cell)
        wdgo.bind('<Button-1>', self.on_active_cell)
        wdg.bind('<FocusIn>', self.on_focus)
        # wdgo.bind('<FocusOut>', self.on_focus)
        wdg.focus_set()


        # ************************************
        self.context['out_wdg'] = wdgo
        # ************************************

        self.history_list = []
        self.history_index = len(self.history_list)
        self.prompt = lambda: f"In [{len(self.history_list) + 1}]: "
        wdg.insert(tk.INSERT, self.prompt())

    def exec_code(self, text, toArchive=False):
        wdg = self.input
        wdg.delete("1.0", tk.END)
        to_insert, *suffix = text.split('\n')
        wdg.insert(tk.END, self.prompt() + to_insert)
        if suffix:
            n = len(self.prompt())
            prefix = f"\n{'...: ':>{n}}"
            for to_insert in suffix:
                wdg.insert(tk.END, prefix + to_insert)
        if toArchive:
            self.archive()

    def on_focus(self, event: tk.Event):
        wdg: tk.Text = event.widget
        if wdg == self.input:
            print(wdg.index(tk.INSERT), wdg.index(tk.END))
        pass

    def on_active_cell(self, event: tk.Event):
        wdg: tk.Text = event.widget
        keysym = event.keysym
        if keysym in ('Up', 'Down'):
            shift_pressed = event.state & 0x00001
            fnc = wdg.tag_nextrange if (bflag := keysym == 'Down') else wdg.tag_prevrange
            ndx = wdg.tag_ranges('active_cell')[bflag]
            cell_range = fnc('cell', ndx)
            if cell_range:
                crange = wdg.tag_ranges('active_cell')
                srange = wdg.tag_ranges(tk.SEL)
                if shift_pressed and not srange:
                    wdg.tag_add(tk.SEL, *crange[-2:])
                elif not shift_pressed and srange:
                    wdg.tag_remove(tk.SEL, srange[0], srange[-1])
                wdg.tag_remove('active_cell', crange[0], crange[-1])
                for tag in ('active_cell', tk.SEL)[:shift_pressed + 1]:
                    wdg.tag_add(tag, *cell_range)
            return 'break'
        elif event.type.name == 'ButtonPress' and event.num == 1:
            ndx = wdg.index(f"@{event.x},{event.y}")
            if 'cell' in wdg.tag_names(ndx):
                cell_range = wdg.tag_prevrange('cell', ndx)
                wdg.tag_remove('active_cell', '1.0', tk.END)
                wdg.tag_add('active_cell', *cell_range)
                wdg.focus_set()
            return 'break'
        pass

    def caret(self, event: tk.Event):
        wdg: tk.Text = event.widget
        line, col = map(int, wdg.index(tk.INSERT).split('.'))
        if event.keysym == 'Left':
            if col <= len(self.prompt()):
                col = len(self.prompt())
            else:
                col -= 1
        if event.keysym == 'Home':
            col = len(self.prompt())
        elif event.keysym == 'End':
            col = 'end'
        wdg.mark_set(tk.INSERT, f'{line}.{col}')
        return 'break'

    def braceleft(self, event: tk.Event):
        wdg = event.widget
        wdg.insert(tk.INSERT, '{}')
        wdg.mark_set(tk.INSERT, 'insert-1c')
        return 'break'

    def clear_prompt(self, event: tk.Event):
        wdg = event.widget
        wdg.delete('1.0', tk.END)
        wdg.insert(tk.INSERT, self.prompt())
        return 'break'

    def delete(self, event: tk.Event):
        wdg = event.widget
        line, col = map(int, wdg.index(tk.INSERT).split('.'))
        if wdg.index(f'{line}.end') == wdg.index(tk.INSERT):
            wdg.delete(f'{line}.end', f'{line + 1}.4')
            return 'break'
        wdg.delete('insert')
        return 'break'

    def backspace(self, event: tk.Event=None):
        wdg: tk.Text = self.input if event is None else event.widget
        left_limit = len(self.prompt())
        if wdg.index(tk.INSERT) == f"1.{left_limit}":
            return 'break'
        line, col = map(int, wdg.index(tk.INSERT).split('.'))
        if col <= left_limit:
            col = wdg.index(f'{line-1}.end-1c').split('.')[1]
            col = max(left_limit, int(col))
            wdg.delete(f'{line-1}.{col}', f'{line}.end')
            return 'break'
        wdg.delete('insert-1c')
        return 'break'

    def Key(self, event: tk.Event):
        wdg = event.widget
        print(event.keycode, event.keysym)

    def history(self, event: tk.Event):
        wdg = event.widget
        line, col = map(int, wdg.index(tk.INSERT).split('.'))
        last_line = int(wdg.index(tk.END).split('.')[0]) - 1
        if (event.keysym == 'Up' and line > 1) or (event.keysym == 'Down' and line < last_line):  # Up
            return
        n = len(self.prompt())
        if len(self.history_list) == 0 or (self.history_index == 1 + len(self.history_list) and col > n):
            return 'break'
        self.history_index += -1 if event.keysym == 'Up' else 1
        self.history_index = max(1, min(self.history_index, 1 + len(self.history_list)))
        wdg.delete('1.0', tk.END)
        content = self.history_list[self.history_index - 1] if self.history_index <= len(self.history_list) else self.prompt()
        content = self.prompt() + content.split(': ', 1)[1]    # Remove the prompt "In [?]: "
        wdg.insert(tk.INSERT, content)
        last_line = int(wdg.index(tk.END).split('.')[0]) - 1
        wdg.mark_set(tk.INSERT, f'{last_line}.end')
        return 'break'

    def on_input(self, event: tk.Event):
        wdg:tk.Text = event.widget
        currentline, endline = map(lambda x: int(wdg.index(x).split('.')[0]), (tk.INSERT, tk.END))
        with_control = event.state & 0x4
        if with_control or currentline != endline - 1: # Control key 0x4
            indent = wdg.get('1.0', tk.INSERT).count('{') - wdg.get('1.0', tk.INSERT).count('}')
            n = len(self.prompt())
            wdg.insert(tk.INSERT, f"\n{'...: ':>{n}}" + indent * '    ')
            pos = wdg.index(tk.INSERT)
            if wdg.get(pos) == '}':
                wdg.insert(tk.INSERT, '\n... ' + (indent - 1) * '    ')
            wdg.mark_set(tk.INSERT, pos)
            return 'break'
        self.archive()
        return self.clear_prompt(event)

    def pythonize(self, raw_text):
        n = len(self.prompt())
        lines = [x[n:] for x in raw_text.splitlines()]
        first_indent = len(lines[0]) - len(lines[0].lstrip())
        if first_indent:
            lines = [x[first_indent:] for x in lines]
        last_line, lines = lines[-1], lines[:-1]
        if lines:
            last_indent = len(last_line) - len(last_line.lstrip())
            if last_indent:
                lines.append(last_line)
                last_line = ''
        if last_line:
            prefix, suffix = (last_line.split(' = ', 1) + [''])[:2]
            if suffix and '(' not in prefix:
                lines.append(last_line)
                last_line = ''
        return '\n'.join(lines), last_line

    def archive(self):
        wdg = self.input
        raw_text = wdg.get("1.0", tk.END)
        wdg.delete("1.0", tk.END)
        self.history_list.append(raw_text.strip())
        self.history_index = 1 + len(self.history_list)
        wdg.insert(tk.INSERT, self.prompt())
        # text = self.pythonize(raw_text)
        self.write_input(raw_text)
        if not raw_text[len(self.prompt()):]:
            self.destroy()
        self.execute(raw_text)
        self.write('\n', tags=('inner_separator', 'cell'))
        self.write('\n', tags=('separator',))
        wdg.focus_set()
        return 'break'
    
    def execute(self, raw_text):
        sout = sys.stdout
        sys.stdout = self
        serr = sys.stderr
        sys.stderr = self
        to_exec, to_eval = self.pythonize(raw_text)
        if to_exec:
            try:
                exec(to_exec, self.context)
            except Exception:
                exc_type, exc_value, exc_traceback = sys.exc_info()
                traceback.print_exception(exc_type, exc_value, exc_traceback, file=sys.stderr)
                to_eval = ''
        try:
            to_eval[0]  # To detect if a to_eval line exists
        except IndexError:
            pass
        else:
            try:
                answ = eval(to_eval, self.context)
            except Exception:
                exc_type, exc_value, exc_traceback = sys.exc_info()
                traceback.print_exception(exc_type, exc_value, exc_traceback, file=sys.stderr)
            else:
                prefix = self.history_list[-1].split(': ', 1)[0].replace('In ', 'Out') + ': '
                txt = prefix + str(answ) + '\n'
                self.write('\n', tags=('inner_separator', 'cell'))
                self.write_output(txt)
        finally:
            sys.stdout = sout
            sys.stderr = serr

    def write(self, text, tags=None):
        if tags is None:
            tags = ('cell',)
        self.output.insert(tk.END, text, tags)
        self.output.see(tk.END)

    def write_input(self, text):
        self.write('\n', tags=('inner_separator', 'cell'))
        lines = text.strip('\n').split('\n')
        for line in lines:
            prefix, line = line.split(': ', 1)
            prefix = (prefix + ': ').replace('...:', '    ')
            self.write(prefix, tags=('in_id', 'cell'))
            self.write(line + '\n', tags=('input', 'cell'))

    def write_output(self, text):
        lines = text.strip('\n').split('\n')
        for line in lines:
            prefix, line = line.split(': ', 1)
            self.write(prefix + ': ', tags=('out_id', 'cell'))
            self.write(line + '\n', tags=('output', 'cell'))



def main():
    app = tk.Tk()
    app.title("My App")
    app.state('zoomed')
    fend = Frontend(app)
    fend.pack(side="top", fill="both", expand=True)

    app.mainloop()


if __name__ == '__main__':
    main()