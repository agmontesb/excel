import tkinter as tk
import tkinter.messagebox as msgbox
import os
import itertools
import logging

import mywidgets.userinterface as userinterface
from mywidgets.equations import equations_manager
from excelxml import OXLLoader

counter = itertools.count(start=1)

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


class ExcelXmlError(Exception):
    pass


class ExcelXmlViewer(tk.Toplevel):

    def __init__(self, master, zf: OXLLoader | str, geometry='640x480', *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.counter = itertools.count(start=1)
        self.title('Excel XML Viewer')
        self.messageVar = tk.StringVar()
        self.geometry(geometry)
        self.setGUI()
        self.filename = ''
        self.zf = None
        self.load_zf(zf)
        # self.state('zoomed')
        self.geometry("600x400")
        pass

    def load_zf(self, zf: OXLLoader | str, path: str | None = None):
        if isinstance(zf, str):
            self.btn_reload['state'] = 'normal'
            self.filename = zf
            zf = OXLLoader(zf)
        self.zf = zf
        path = path or f'{zf.path_obj.root}/xl/workbook.xml'
        self.load_xmlfile(path)

    def setGUI(self):
        file_path = '@layout/excelxml_viewer'
        xmlObj = userinterface.getLayout(file_path, withCss=True)
        userinterface.newPanelFactory(
            master=self,
            selpane=xmlObj,
            genPanelModule=None,
            setParentTo='master',
            registerWidget=self.register_widget,
        )
        equations_manager.set_initial_widget_states()

        filetree = self.filetree
        filetree.tag_configure('selected', background='light green')
        filetree.bind('<<TreeviewOpen>>', self.onTreeViewOpen)
        filetree.bind('<<ActiveSelection>>', self.onTreeSelection, '+')

        # self.urlFrame.setUrlContentProcessor(self.setContent)

        # self.txtEditor.setKeyHandler(self)
        # self.txtEditor.setHyperlinkManager(self.hyperLnkMngr)

        self.regexBar.setTextWidget(self.txtEditor.textw)
        self.regexBar.setTreeWidget(self.tree)
        self.regexBar.messageVar = self.messageVar
        # self.regexBar.setZoomManager(self.zoom)
        # self.txtEditor.textw.bind('<Button-3>', self.do_popup)

        self.statusBar.Message.config(textvariable=self.messageVar)
        self.btn_reload['command'] = self.reload

        self.bind("<FocusIn>", self.on_focus_in)
        self.bind("<FocusOut>", self.on_focus_out)

        pass

    def onVarChange(self, event=None, attr_data=None):
        # if event:
        #     attr_data = event.attr_data
        # var_name, value = attr_data
        # match var_name:
        #     case 'view_headings':
        #         self.sheetui.toggle_headings()
        #     case 'view_gridlines':
        #         self.sheetui.toggle_gridlines()
        #     case 'fml_r1c1':
        #         self.r1c1_flag = value
        #         self.sheetui.redraw_headings()
        #     case 'wb_loader':
        #         if value == 'openpyxl':
        #             self.wb_manager = openpyxl
        #         else:
        #             self.wb_manager = excelxml
        pass

    def onMenuClick(self, event):
        # menu_master, indx = event.widget, event.data
        # menu, menu_item = menu_master.cget("title"), menu_master.entrycget(indx, "label")
        # match menu:
        #     case 'file' | 'open recent':
        #         self.file_menu(menu_master, indx)
        #     case 'view':
        #         self.view_menu(menu_master, indx)
        pass

    def register_widget(self, master, xmlwidget, widget):
        attribs = xmlwidget.attrib
        name = attribs.get('name')
        if name in ('btn_reload', 'filetree', 'fpath', 'regexBar', 'txtEditor', 'tree', 'statusBar',):
            setattr(self, name, widget)
        pass

    def on_focus_in(self, event):
        if event.widget == self and self.filename:
            t1, t2 = map(os.path.getmtime, (self.filename, self.zf.zf.filename))
            if t1 != t2:
                answ = msgbox.askokcancel(
                    title="File Changed",
                    message=f"The file '{os.path.basename(self.filename)}' has changed on disk.\nDo you want to reload it?"
                )
                if answ:
                    self.reload()
                else:
                    self.btn_reload['background'] = 'red'

    def on_focus_out(self, event):
        if event.widget == self:
            pass

    def reload(self):
        if self.filename:
            path = '/' + self.fpath['text']
            self.load_zf(self.filename, path)
            self.btn_reload['background'] = 'SystemButtonFace'

    def setContent(self, data, newUrl=True):
        self.txtEditor.setContent(data)
        self.regexBar.getPatternMatch()

    def iid_to_path(self, iid):
        tree = self.filetree
        loader = self.zf
        root = loader.path_obj.root
        parts = iid.split('.')
        path = f'{root}/' + '/'.join(tree.item('.'.join(parts[:k + 1]), 'text') for k in range(len(parts)))
        return path.rstrip('/')
    
    def path_to_iid(self, path):
        tree = self.filetree
        fpath = self.zf.path_obj.root
        stack = path[len(self.zf.path_obj.root) + 1:].split('/')[::-1]
        iid = ''
        while stack:
            name = stack.pop()
            children = tree.get_children(iid)
            if not children or f'{iid}.dummy' in children:
                self.onTreeViewOpen(iid=iid)
                children = tree.get_children(iid)
            try:
                iid = [item for item in children if tree.item(item, 'text') == name][0]
            except IndexError:
                msg = f"Path not found in treeview: {fpath}"
                raise ExcelXmlError(msg)  
            fpath = f'{fpath}/{name}'
        return iid

    def onTreeViewOpen(self, event=None, *, iid=None):
        tree = self.filetree
        loader = self.zf
        if event is not None:
            iid = tree.focus()
        else:
            # iid = self.path_to_iid(path)
            tree.item(iid, open=True)
            if not iid:
                self.counter = itertools.count(start=1)
        counter = self.counter
        path = self.iid_to_path(iid)
        assert path is not None

        tree.delete(*tree.get_children(iid))
        root, d_names, f_names = next(loader.path_obj.walk(path))
        root = root.rstrip('/')
        path_id = iid
        for k, names in enumerate((d_names, f_names)):
            suffix = ('/', '')[k]
            for name in sorted(names):
                filename = f'{root}/{name}{suffix}'
                child_id = f'{path_id}.I{next(counter):0>4}'.strip('.')
                tree.insert(
                    path_id,
                    'end',
                    iid=child_id,
                    text=name,
                )
                if filename.endswith('/'):
                    tree.insert(child_id, 'end', iid=f'{child_id}.dummy', text='dummy')
        loader.path_obj.actual_dir = path
        pass

    def onTreeSelection(self, event):
        tree = self.filetree
        iid = event.widget.focus()
        fname = tree.item(iid, 'text')
        bflag = fname.endswith(('.xml', '.rels'))
        selected = tree.tag_has('selected')
        ndx = selected.index(iid)
        sel = selected[(ndx + bflag) % len(selected)]
        # sel,  = tree.tag_has('selected')
        tags = tuple(t for t in tuple(tree.item(sel, 'tags')) if t != 'selected')
        tree.item(sel, tags=tags)
        if bflag:
            logger.debug(f"{iid}, {tree.item(iid, 'text')}")
            path = self.iid_to_path(iid)
            self.load_xmlfile(path)
        pass

    def load_xmlfile(self, path):
        try:
            iid = self.path_to_iid(path)
        except ExcelXmlError:
            if path != f'{self.zf.path_obj.root}/xl/workbook.xml':
                logger.exception(f"Error loading XML file: {path}")
                self.after(0, lambda: msgbox.showerror(title='Error', message=f"Error loading XML file: {path}.\nLoading workbook.xml instead."))
                path = f'{self.zf.path_obj.root}/xl/workbook.xml'
                iid = self.path_to_iid(path)
            else:
                iid = ''
                path = ''

        if path.endswith(('.xml', '.rels')):
            self.filetree.item(iid, tags='selected')
            content = self.zf.load_xmlfile(path)
            self.setContent(content)
            self.fpath['text'] = path[1:]
            self.zf.path_obj.actual_dir = path

def main():
    root = tk.Tk()
    zf = r'C:\Users\agmontesb\Documents\GitHub\excel\tests\files\excel_module_test.xlsx'
    wbv = ExcelXmlViewer(root, zf, geometry='1200x1200')        # create the Toplevel viewer
    # wbv.transient(root)           # associate it with the hidden root
    wbv.protocol("WM_DELETE_WINDOW", root.destroy)  # close app when viewer closes
    wbv.state('zoomed')
    root.withdraw()  # hide the invisible root window
    wbv.mainloop()

if __name__ == '__main__':
    main()
