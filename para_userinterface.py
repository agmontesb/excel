# -*- coding: utf-8 -*-

# tkinter reference: https://anzeljg.github.io/rin2/book2/2405/docs/tkinter/index.html
# Tcl8.5.19/Tk8.5.19 Documentation: http://tcl.tk/man/tcl8.5/contents.htm
import collections
import fnmatch
import importlib
import inspect
import os.path
import re
import sys
import tkinter as tk
from tkinter import ttk
import xml.etree.ElementTree as ET
import urllib.request
import urllib.parse
import functools
from contextlib import contextmanager
from typing import Callable, ClassVar, Any, MutableMapping, Iterable, Tuple, Literal, TypeAlias, Optional
from typing_extensions import Protocol, runtime_checkable
from types import ModuleType
from abc import abstractmethod
from mywidgets.Widgets.Custom import ImageProcessor as imageprocessor

from mywidgets.Widgets import specialwidgets
from mywidgets.Widgets.Custom import ImageProcessor as imgp
from mywidgets.Widgets.Custom.network import network
from mywidgets.equations import equations_manager
from mywidgets.Tools.uiStyle import uicss, cssgrid, cssflexbox
import mywidgets.cbwidgetstate as cbwidgetstate

MODULE_STACK = [(-1, ttk), (-1, tk), (-1, specialwidgets)]

getFontAwesomeIcon = imageprocessor.memoize(imageprocessor.getFontAwesomeIcon)

@runtime_checkable
class TreeLike(Iterable['TreeLike'], Protocol):
    tag: str
    attrib: dict[str, Optional[str]]

    def get(self, item: str, default: Optional[str] = None) -> Optional[str]:
        return self.attrib.get(item, default)

    def findall(self, name: str) -> Iterable['TreeLike']:
        ...


RegWidgetCb: TypeAlias = Optional[Callable[[tk.Tk | tk.Widget, TreeLike, tk.Widget], Any]]
GeoManager_Type = Literal['grid', 'place', 'pack']

pack_params = ["after", "anchor", "before", "expand", "fill", "in", "ipadx", "ipady", "padx", "pady", "side"]
grid_params = ["column", "columnspan", "in", "ipadx", "ipady", "padx", "pady", "row", "rowspan", "sticky"]
place_params = ["anchor", "bordermode", "height", "in", "relheight", "relwidth", "relx", "rely", "width", "x", "y"]


class BaseGeomngr(Protocol):
    manager: ClassVar[Literal['pack', 'grid', 'place']]
    container_attribs: ClassVar[set] = set()
    item_attribs: ClassVar[set] = set()
    slaves:list[tuple[str, dict]]

    @classmethod
    def filter_container_attribs(cls, attribs: dict):
        other_attribs = attribs.copy()
        item_attribs = {key: other_attribs.pop(key) for key in other_attribs.keys() & cls.container_attribs}
        return item_attribs, other_attribs

    @classmethod
    def filter_item_attribs(cls, attribs: dict):
        other_attribs = attribs.copy()
        item_attribs = {key: other_attribs.pop(key) for key in other_attribs.keys() & cls.item_attribs}
        return item_attribs, other_attribs

    def config_master(self, master):
        [
            getattr(master.nametowidget(name), self.manager)(**item_attrs)
            for name, item_attrs in self.slaves
        ]

    def register_item(self, widget, item_attrs: dict) -> None:
        if item_attrs.get('in'):
            item_attrs['in_'] = item_attrs.pop('in')
        self.slaves.append((str(widget), item_attrs))


class TkGridGeomngr(BaseGeomngr):
    manager = 'grid'
    item_attribs = {
        "column", "columnspan", "in", "ipadx", "ipady",
        "padx", "pady", "row", "rowspan", "sticky"
    }

    def __init__(self, **attribs):
        self.gridconfigure, _ = self.filter_container_attribs(attribs)
        self.slaves = []
        pass

    def config_master(self, master):
        super().config_master(master)
        for label in self.gridconfigure:
            method = getattr(master, label)
            for n, kwargs in self.gridconfigure.get(label):
                method(n, **kwargs)
        pass

    @classmethod
    def filter_container_attribs(cls, attribs: dict):
        other_attribs = attribs.copy()
        gridconfigure = {}
        for label in ('rowconfigure', 'columnconfigure'):
            n = len(label)
            gridconfigure[label] = [
                (key[n:].split('_'), eval(f'dict({other_attribs.pop(key)})')) for key in list(other_attribs.keys())
                if key.startswith(label)
            ]
        return gridconfigure, other_attribs


class TkPlaceGeomngr(BaseGeomngr):
    manager = 'place'

    item_attribs = {
        "anchor", "bordermode", "height", "in", "relheight", "relwidth",
        "relx", "rely", "width", "x", "y"
    }

    def __init__(self, **attribs):
        self.slaves = []
        pass


class TkPackGeomngr(BaseGeomngr):
    manager = 'pack'

    item_attribs = {
            "after", "anchor", "before", "expand", "fill", "in",
            "ipadx", "ipady", "padx", "pady", "side"
        }

    def __init__(self, **attribs):
        self.slaves = []
        pass


tkgeomanager_mapping = {'pack': TkPackGeomngr, 'place': TkPlaceGeomngr, 'grid': TkGridGeomngr}


# Este context manager se puede utilizar cuando se quiere adjuntar datos en la emision de un evento:
#   with event_data as event:
#       event.data = 'este es el mensaje'
#       widget.event_generate('<<el_evento_name>>')
#
@contextmanager
def event_data(**params):
    def decor(f):
        try:
            vars = inspect.getclosurevars(f)
            f = vars.nonlocals['f']
        except:
            pass

        def wrapper(*args, **kwargs):
            answ = f(*args, **kwargs)
            answ.__dict__.update(params)
            return answ

        return wrapper
        return decor

    wrapped = tk.Event
    tk.Event = decor(tk.Event)
    try:
        yield tk.Event
    finally:
        tk.Event = wrapped


def getFileUrl(resource_id, src=None):
    # @mywinzip:layout/winzip
    # '/Data/tkinter/tkWidgetStyle.xml'
    # @data:tkinter/tkWidgetStyle
    filename = resource_id
    if filename.startswith('@'):
        filename = filename[1:]
        resource_id, f_name = filename.split('/', 1)
        common_path = os.path.dirname(__file__).lower()
        try:
            pckg, d_name = resource_id.split(':', 1)
            match pckg.lower():
                case 'data' | 'tests' | 'doc' as case:
                    args = [pckg, d_name]
                    if case != 'data':
                        common_path = os.path.dirname(common_path)
                case _:
                    args = ['Tools', pckg, 'res', d_name]
            base_path = os.path.join(common_path, *args)
        except ValueError:
            d_name = resource_id
            src = os.path.dirname(src or sys.argv[0]).lower()
            # fnc = lambda x: os.path.commonpath((src, os.path.join(common_path, x))) == common_path
            # if any(map(fnc, ('tools', 'data'))):
            base_path = os.path.join(src, 'res', d_name)
            if not os.path.exists(base_path):
                base_path = src
        f_names = fnmatch.filter(os.listdir(base_path), f'{f_name}.*')
        try:
            filename = os.path.join(base_path, f_names[0])
        except IndexError:
            filename = None
    return filename


class SignedContent(str):
    MyTuple = collections.namedtuple('MyTuple', 'pckg res_type res_id src')
    MyTuple.resource_id = lambda x: f"@{x.pckg}:{x.res_type}/{x.res_id}" if x[:3] != ('', '', '') else ''
    MyTuple.__str__ = lambda x: x.src

    def __new__(cls, value, *, src=None):
        obj = str.__new__(cls, value)
        obj.src = cls.get_resourceid(src) if src else None
        return obj

    @classmethod
    def get_resourceid(cls, src_in):
        src = os.path.splitext(src_in.lower())[0].replace(os.sep, '/')
        if m := re.search(r'src/mywidgets/tools/(.+?)/res/(.+?)/(.+)', src):
            pckg, res_type, res_id = m.groups()
        elif m := re.search('/src/mywidgets/data/(.+?)/(.+)', src):
            pckg = 'data'
            res_type, res_id = m.groups()
        else:
            return cls.MyTuple('', '', '', src_in)
        obj = cls.MyTuple(pckg, res_type, res_id, src_in)
        return obj

def getContent(fileurl):
    fileurl = getFileUrl(fileurl)
    bflag = urllib.parse.urlparse(fileurl).scheme not in ('http', 'https', 'ftp')
    if bflag:
        if fileurl is None:
            raise FileNotFoundError(f"Resource not found: {fileurl}")
        basepath = os.path.dirname(__file__)
        layoutpath = os.path.join(basepath, fileurl)
        layoutfile = os.path.abspath(layoutpath)
        fileurl = f'file://{urllib.request.pathname2url(layoutfile)}'
    initConf = 'curl --user-agent "Mozilla/5.0 (Windows NT 6.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36" --cookie-jar "cookies.lwp" --location'
    net = network(initConf)
    content, resp_url = net.openUrl(fileurl)
    src = layoutfile.lower() if bflag else resp_url
    return SignedContent(content, src=src)


def getLayout(layoutfile, withCss=False, is_content=False):
    content = layoutfile if is_content else getContent(layoutfile)
    if withCss:
        tcase = os.path.splitext(str(content.src))[-1] or '.xml'
        wfactory = uicss.ElementFactory()
        return wfactory.getElements(content, tcase=tcase)
    return ET.XML(content)


def findGeometricManager(tag: GeoManager_Type) -> list[str]:
    if tag == 'grid':
        return grid_params
    if tag == 'place':
        return place_params
    return pack_params


def renderImage(filename):
    filename = getFileUrl(filename)
    try:
        _, ext = os.path.splitext(filename)
    except TypeError:      # El archivo no existe
        return
    if ext == '.xml':
        image_xml = getLayout(os.path.abspath(filename))
        attrib = image_xml.attrib.copy()
        match image_xml.tag:
            case 'fontawesome':
                charname = attrib.pop('charname')
                attrib['size'] = int(attrib['size'])
                attrib['isPhotoImage'] = bool(int(attrib.pop('isphotoimage', '1')))
                icon_img = getFontAwesomeIcon(charname, **attrib)
                return icon_img
            case 'texture':
                attrib['imagefile'] = os.path.abspath(attrib['imagefile'])
                if 'bordertexture' in attrib:
                    attrib['bordertexture'] = os.path.abspath(attrib['bordertexture'])

                for key in ('border', 'width', 'height', 'bordersize'):
                    try:
                        attrib[key] = eval(attrib[key])
                    except Exception as e:
                        raise

                positional = (attrib.pop('imagefile'), attrib.pop('width'), attrib.pop('height'))
                kwargs = {key: attrib.pop(key) for key in ('aspectratio',)}
                options = attrib
                texture = imageprocessor.getTexture(*positional, **kwargs, **options)
                return texture

    icon_img = imageprocessor.getCacheTexture(filename, None, None)
    return icon_img


def getWidgetInstance(master: tk.Tk | tk.Widget,
                      widgetname: str,
                      attributes: MutableMapping[str, str],
                      panelModule: list[Tuple[int, ModuleType]] | None = None) -> tk.Widget:
    '''

    :param master: widget. De ser derivado de basicWidget.
    :param widgetname: str. Nombre del widget que se quiere instanciar.
    :param attributes: dict. Atributos del widget a instanciar.
    :param panelModule: módulo. Contiene definiciones de widgets definidos por el usuario.
    :return: widget.
    '''

    # Si no se tiene panelModule se asume que es tkinter
    # panelModule = panelModule or tk
    panelModule = panelModule or MODULE_STACK
    for indx in range(len(panelModule) - 1, -1, -1):
        _, module = panelModule[indx]
        getWdgClass = getattr(module, 'getWidgetClass', lambda x: getattr(module, x, None))
        if module.__name__ == 'tkinter':
            widgetname = widgetname.title() if widgetname != 'labelframe' else 'LabelFrame'
        elif module.__name__ == 'tkinter.ttk':
            camel_case = {'labelframe': 'LabelFrame', 'panedwindow': 'PanedWindow'}
            widgetname = camel_case.get(widgetname, widgetname.title())
        widgetClass = getWdgClass(widgetname)
        if widgetClass:
            break

    if widgetClass is None:
        # ERROR: No se encontró definición del widget
        raise AttributeError(f'{widgetname} is not a tkinter widget or a user defined class. ')

    if 'image' in attributes:
        image_file = attributes['image']
        attributes['image'] = renderImage(image_file)
    if 'menuid' in attributes:
        pass
    widget = widgetClass(master, **attributes)

    return widget


def widgetFactory(master: tk.Tk | tk.Widget,
                  selPane: TreeLike,
                  panelModule: list[Tuple[int, ModuleType]] | None = None,
                  registerWidget: RegWidgetCb = None,
                  setParentTo: Literal['master', 'category', 'root'] = 'master',
                  k: int = -1) -> int:
    '''
    Crea la jerarquia de widgets que conforman una tkinter form.
    :param master: tkinter Form. Padre de la Form a generar.
    :param selPane:     element tree nade. Nodo raiz definido por el archivo de layout.
    :param panelModule: modulo. Define los widgets definidos por el usuario.
    :param setParentTo: str. Los widgets que se desprendan de selPane, serán creados como hijos de:
                             'master': Hace del master también el parent del widget.
                             'category': Frame 'category' parent de todos los widgets que contiene.
                             'root': Hace parent general al Frame que contiene toda la Form
    :param registerWidget: callable. Callable que debe aceptar dos parámetros: element tree node y
                        el widget recién procesado.
    :param k:           integer. Consecutivo utilizado por la función para crear el nombre
                        interno de los widgets a crear.
    :return:            2 members tuple. Posición 0: Valor consecutivo del últimmo widget agregado.
                        Posición 1: Lista de tuples que muestra el id del widget y
                        la ecuación que activa (enable) el widget.
    '''

    def set_geomanager(selpane) -> BaseGeomngr:
        if 'display' in selpane:
            match selpane['display']:
                case 'grid' | 'inline-grid':
                    master_geomngr = cssgrid.CssGrid
                case 'flex' | 'inline-flex':
                    master_geomngr = cssflexbox.CssFlexBox
                case _:
                    raise Exception('Not Geometric Manager supported')
        else:
            gmanager: Literal['pack', 'place', 'grid'] = selpane.get('geomanager', 'pack')
            master_geomngr: BaseGeomngr = tkgeomanager_mapping[gmanager]
        return master_geomngr(**selpane)

    kini = k
    panelModule = panelModule or [(kini, m) for _, m in MODULE_STACK]
    if selPane.get('lib'):
        # El contenedor declara que sus widgets están definidos en "lib" por lo cual se agrega
        # al camino de definiciones de los widgets.
        module_path = selPane.get('lib', '')
        assert module_path is not None
        module = importlib.import_module(module_path, __package__)
        panelModule.append((kini, module))
    master_geomngr = set_geomanager(selPane.attrib)
    for xmlwidget in selPane:
        k += 1
        is_widget = xmlwidget.tag != 'var'
        options: dict[str, Optional[str]] = dict.copy(xmlwidget.attrib)
        options.pop('class', None)
        # Se asigna como id del widget el último segmento del id definido.
        name_default = options.get('id', '').rsplit('/')[-1] or str(k)
        options.setdefault('name', name_default)  # Se asigna el consecutivo como nombre del widget.

        parent = master
        bparent = is_widget and ((setParentTo == 'category' and selPane.tag != 'category') or setParentTo == 'root')
        if bparent:
            # Cuando se elige la opción sameParent=True significa que todos los widgets que están
            # bajo una misma categoría serán creados como hijos del frame correspondiente a
            # la categoría pero puestos bajo el widget master
            in_default = master.winfo_name()
            if options.setdefault('in', in_default).startswith('.') and setParentTo != 'master':
                options['in'] = in_default + options['in']
            parent = master.nametowidget(master.winfo_parent())

        wdg_geomngr = set_geomanager(options)
        # wdg_geomngr = geomngr(**options)

        # Se filtran los atributos que manejan el geometric manager
        # Los relativos al widget como container
        options.pop('geomanager', None)
        options.pop('display', None)
        _, options = wdg_geomngr.filter_container_attribs(options)
        # Los relativos al widget como item del container en que está definido
        geomngr_opt, options = master_geomngr.filter_item_attribs(options)

        states_eq = dict([
            (key, options.pop(key)) for key in cbwidgetstate.STATES
            if key in options
        ])

        wType = xmlwidget.tag

        bflag = 'lib' in options
        if bflag:
            module_path = options.pop('lib')
            module = importlib.import_module(module_path, __package__)
            panelModule.append((kini, module))

        # El atributo 'read' se utiliza para especificar un widget que despliega un texto y
        # puede tomar los valores left-to-right ('LtoR'), top-to-bottom ('TtoB'),
        # bottom-to-top ('BtoT').
        # Con este atributo se pueden generar botones (normales, checkbutton, radiobutton) verticales
        # cuyo texto se reemplaza por una imagen de un texto vertical si se agrega un atibuto read con
        # valor 'TtoB' o 'BtoT'.
        read = options.pop('read', 'LtoR')

        widget = getWidgetInstance(parent, wType, options, panelModule=panelModule)

        if read in ('TtoB', 'BtoT'):
            text = widget.cget('text')
            angle = 90 if read == 'BtoT' else -90
            font = imgp._eqTkFont('C:\Windows\Fonts\ConsolaB.ttf', size=12)
            image = imgp.getLabel(text, font, 'black', angle=angle, isPhotoImage=True)
            widget.image = image
            widget.config(image=image)
            if wType in ('checkbutton', 'radiobutton'):
                widget.config(selectimage=image)

        if bflag:
            panelModule.pop()

        # if 'in' in geomngr_opt and isinstance(geomngr_opt['in'], (bytes, str)):
        ref_labels = geomngr_opt.keys() & ('in', 'before', 'after')
        if ref_labels:
            for ref_label in ref_labels:
                if not isinstance(geomngr_opt[ref_label], (bytes, str)):
                    continue
                parent_name = geomngr_opt.pop(ref_label).lstrip('.')
                parent_path = widget.winfo_parent()
                geomngr_opt[ref_label] = '.'.join((parent_path if parent_path != '.' else '', parent_name))
            # if 'in' in ref_labels:
            #     geomngr_opt['in_'] = geomngr_opt.pop('in')

        if is_widget and 'visible' not in states_eq:
            master_geomngr.register_item(widget, geomngr_opt)

        if registerWidget:
            registerWidget(master, xmlwidget, widget)

        for state, eq in states_eq.items():
            callback_closure = cbwidgetstate.STATES[state]
            if state == 'visible':
                kwargs = geomngr_opt
            elif state == 'enable':
                kwargs = {}
            cb = callback_closure(widget, **kwargs)
            equations_manager.add_equation(eq, cb)

        has_children = bool(len(list(xmlwidget)))
        if has_children:
            # wdg_geomngr.config_master(widget)
            k = widgetFactory(
                widget,
                xmlwidget,
                panelModule=panelModule,
                setParentTo=setParentTo,
                registerWidget=registerWidget,
                k=k
            )
    master_geomngr.config_master(master)
    if kini == panelModule[-1][0]:
        panelModule.pop()
    return k


def newPanelFactory(master: tk.Tk | tk.Widget,
                    selpane: TreeLike,
                    genPanelModule: list[Tuple[int, ModuleType]] | None = None,
                    setParentTo: Literal['master', 'category', 'root'] = 'master',
                    registerWidget: RegWidgetCb = None) -> int:
    '''
    Crea la jerarquia de widgets que conforman una tkinter form.
    :param master:      tkinter Form. Padre de la Form a generar.
    :param selpane:     element tree nade. Nodo raiz definido por el archivo de layout.
    :param genPanelModule: modulo. Define los widgets definidos por el usuario.
    :param setParentTo:       str. One of ['master', 'category', 'root'].
    :param registerWidget: callable. Callback function para registrar los widgets creados.
    :return:            2 members tuple. Posición 0: Valor consecutivo del últimmo widget agregado.
                        Posición 1: Lista de tuples que muestra el id del widget y
                        la ecuación que activa (enable) el widget.
    '''

    categories = selpane.findall('category')
    n = -1
    if len(categories) <= 1:
        try:
            selpane = categories[0]  # Case categories = [category]
        except IndexError:
            pass
        n = widgetFactory(
            master,
            selpane,
            panelModule=genPanelModule,
            registerWidget=registerWidget,
            setParentTo=setParentTo,
            k=n
        )
    else:
        def selectPane(id, panels):
            for pane in panels:
                pane.pack_forget()
            panels[id].pack(side=tk.TOP, fill=tk.BOTH, expand=tk.YES)
            pass

        panelArray = []
        panel_attrs = dict(relief=tk.RIDGE, bd=5, bg='white', padx=3, pady=3)
        leftPane = tk.Frame(master, **panel_attrs)
        leftPane.pack(side=tk.LEFT, fill=tk.Y)
        frstboton = None
        for k, elem in enumerate(categories):
            boton = tk.Radiobutton(
                leftPane,
                text=elem.get('label', ''),
                value=k,
                indicatoron=False
            )
            boton.pack(side=tk.TOP, fill=tk.X)
            frstboton = frstboton or boton
            panel = tk.Frame(master, name=f'categoria{k + 1}', **panel_attrs)
            panel.form = master  # type: ignore
            panelArray.append(panel)
        for child in leftPane.winfo_children():
            child.config(command=lambda x=child['value'], y=tuple(panelArray): selectPane(x, y))  # type: ignore
        assert isinstance(frstboton, tk.Radiobutton)
        frstboton.invoke()
        for root, selPane in zip(panelArray, categories):
            n = widgetFactory(
                root,
                selPane,
                panelModule=genPanelModule,
                registerWidget=registerWidget,
                setParentTo=setParentTo,
                k=n)
    return n


def menuclick_closure(widget, data=None):
    def menu_cb():
        with event_data(data=data) as event:
            widget.event_generate('<<MENUCLICK>>')
    return menu_cb


def menuFactory(
        parent: tk.Tk | tk.Widget,
        selPane: TreeLike,
        menu_map: dict[str, tk.Menu] = None,
        registerMenu: RegWidgetCb = None) -> tk.Menu:

    master = parent.winfo_toplevel()
    menu_map = menu_map or {}
    options = selPane.attrib.copy()
    label = options.pop('label')
    src = options.pop('src', None)
    if src:
        selPane = getLayout(src)
    menu_name = options.pop('name', label.lower())
    options['title'] = menu_name
    menu_map[label] = menu_master = tk.Menu(parent, name=menu_name, **options)
    labels = []
    cbs = []
    for k, item in enumerate(selPane):
        menu_type = item.tag
        options = item.attrib.copy()
        cb = None
        if menu_type not in ['separator', 'cascade']:
            try:
                command = options.pop('command')
                cb = getattr(master, command)
            except (AttributeError, KeyError):
                if 'variable' not in options:
                    # En el caso de los radio button o checbutton solo se asigna cb si no se ha definido
                    # una variable para ellos de tal manera que es posible capturar a través
                    # del cambio de valor que boton se pulsó. En el caso button no se puede definir variable.
                    cb = menuclick_closure(widget=menu_master, data=k)
            if options.get('accelerator', ''):
                accelerator = options['accelerator']
                accelerator = accelerator.replace('+', '-')
                try:
                    modifiers, key = accelerator.rsplit('-', 1)
                    modifiers = modifiers.lower()
                    modifiers_map = [('ctrl', 'Control'), ('alt', 'Alt'), ('shift', 'Shift')]
                    modifiers = functools.reduce(lambda t, x: t.replace(x[0], x[1]), modifiers_map, modifiers)
                    accelerator = f'{modifiers}-KeyPress-{key}'
                except ValueError:
                    accelerator = f'KeyPress-{accelerator}'
                master.bind(f'<{accelerator}>', lambda x, y=cb: y())
        elif menu_type == 'cascade':
            menu_wdg = menuFactory(master, item, menu_map, registerMenu)
            options['menu'] = menu_wdg
            options.pop('tearoff', None)
            options.pop('src', None)
        menu_master.add(menu_type, **options)
        labels.append(options.get('label', ''))
        if cb is not None:
            cbs.append((menu_master.index(tk.END), cb))
    try:
        registerMenu(parent, selPane, menu_master, labels)
    except:
        pass
    # Retrasando la asignación del command en este punto, se permite la inicialización del estado
    # de la 'entry asociada cuando esta es un checkbox.
    for index, cb in cbs:
        menu_master.entryconfigure(index, command=cb)
    return menu_master
