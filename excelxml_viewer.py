import shutil
import tempfile
import tkinter as tk
import os
import zipfile
import itertools
import logging

import mywidgets.userinterface as userinterface
from mywidgets.Widgets.Custom import navigationbar, WidgetsExplorer
from mywidgets.equations import equations_manager

counter = itertools.count(start=1)

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


class ExcelXmlViewer(tk.Toplevel):

    def __init__(self, master, zf: zipfile.ZipFile | str, geometry='640x480', *args, **kwargs):
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

    def load_zf(self, zf: zipfile.ZipFile | str):
        if isinstance(zf, str):
            self.btn_reload['state'] = 'normal'
            self.filename = zf
            tmpdir = tempfile.gettempdir()
            dstfile = shutil.copy(zf, tmpdir)
            shutil.copystat(zf, dstfile)
            zf = zipfile.ZipFile(dstfile)
        self.zf = zf
        filename = zf.filename
        root = f'/{os.path.basename(filename)}'
        namelist = [root + '/' + x for x in zf.namelist()]
        navigationbar.StrListObj.SEP = '/'
        self.path_obj = navigationbar.StrListObj(namelist, root)
        self.fpath['text'] = root[1:]
        self.onTreeViewOpen(path=root)
        self.onTreeViewOpen(path=f'{root}/xl')
        iid = self.path_to_iid(path=f'{root}/xl/workbook.xml')
        self.filetree.item(iid, tags='selected')
        self.load_xmlfile(f'{root}/xl/workbook.xml')

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

    def reload(self):
        if self.filename:
            self.load_zf(self.filename)

    def setContent(self, data, newUrl=True):
        self.txtEditor.setContent(data)
        self.regexBar.getPatternMatch()

    def iid_to_path(self, iid):
        tree = self.filetree
        parts = iid.split('.')
        path = f'{self.path_obj.root}/' + '/'.join(tree.item('.'.join(parts[:k + 1]), 'text') for k in range(len(parts)))
        return path
    
    def path_to_iid(self, path):
        tree = self.filetree
        rpath = self.path_obj.relpath(path, self.path_obj.root)
        if rpath == '.':
            return ''
        stack = rpath.split(self.path_obj.SEP)[::-1]
        child_id = ''
        while stack:
            name = stack.pop()
            children = tree.get_children(child_id)
            child_id = [item for item in children if tree.item(item, 'text') == name][0]
        return child_id

    def onTreeViewOpen(self, event=None, *, path=None):
        tree = self.filetree
        counter = self.counter
        if event is not None:
            iid = tree.focus()
            path = self.iid_to_path(iid)
        else:
            iid = self.path_to_iid(path)
            tree.item(iid, open=True)

        assert path is not None

        tree.delete(*tree.get_children(iid))
        root, d_names, f_names = next(self.path_obj.walk(path))
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
                    tree.insert(child_id, 'end', text='dummy')
        self.path_obj.actual_dir = path
        pass

    def onTreeSelection(self, event):
        tree = self.filetree
        iid = event.widget.focus()
        fname = tree.item(iid, 'text')
        if not fname.endswith('.xml'):
            tags = (t for t in tuple(tree.item(iid, 'tags')) if t != 'selected')
            tree.item(iid, tags=tags)
            return 'break'
        else:
            logger.debug(f"{iid}, {tree.item(iid, 'text')}")
            path = self.iid_to_path(iid)
            self.load_xmlfile(path)
        pass

    def load_xmlfile(self, path):
        rpath = self.path_obj.relpath(path, self.path_obj.root)
        content = self.zf.read(rpath).decode('utf-8')
        self.setContent(content)
        self.fpath['text'] = path[1:]


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
