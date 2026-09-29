"""
A minimal tkinter stand-in, so the GUI's decision logic can be tested headlessly.

`main.py` is imported for its behaviour, not its pixels: which message the status
bar shows, whether a click toggles a checkbox, whether Apply stays usable after an
error. None of that needs a display, but it does need `import tkinter` to succeed
and widgets to accept the calls `main.py` makes.

Install with `install()` *before* importing `main`. Widgets record what they are
told rather than drawing anything; `FakeTreeview` keeps real parent/child
structure so tree logic is exercised for real, and `after()` runs its callback
immediately so a worker's completion is synchronous in tests.

This stub is deliberately dumb. If `main.py` starts calling a widget method that
isn't here, the test fails loudly with AttributeError - which is the point.
"""

import sys
import types


class _Var:
    """Stands in for StringVar / BooleanVar, including trace_add."""

    def __init__(self, value=None, master=None, **kwargs):
        self._value = value
        self._callbacks = []

    def get(self):
        return self._value

    def set(self, value):
        self._value = value
        for callback in list(self._callbacks):
            callback(None, None, "write")

    def trace_add(self, mode, callback):
        self._callbacks.append(callback)
        return f"trace{len(self._callbacks)}"


class StringVar(_Var):
    def __init__(self, value="", master=None, **kwargs):
        super().__init__(value, master, **kwargs)


class BooleanVar(_Var):
    def __init__(self, value=False, master=None, **kwargs):
        super().__init__(value, master, **kwargs)


class _Widget:
    """Records configuration; accepts every geometry call without complaint."""

    def __init__(self, master=None, **kwargs):
        self.master = master
        self.options = dict(kwargs)
        self.bindings = {}

    def configure(self, cnf=None, **kwargs):
        self.options.update(kwargs)
    config = configure

    def cget(self, key):
        return self.options.get(key)

    def __getitem__(self, key):
        return self.options.get(key)

    def pack(self, **kwargs):
        pass

    def grid(self, **kwargs):
        pass

    def columnconfigure(self, *args, **kwargs):
        pass

    def rowconfigure(self, *args, **kwargs):
        pass

    def bind(self, sequence, func, add=None):
        self.bindings[sequence] = func

    def set(self, *args):
        """Scrollbars are handed to Treeview as yscrollcommand/xscrollcommand."""

    def focus_set(self):
        pass

    def destroy(self):
        self.destroyed = True

    def winfo_exists(self):
        return not getattr(self, "destroyed", False)

    @property
    def state(self):
        return self.options.get("state", "normal")


class _Style(_Widget):
    def configure(self, *args, **kwargs):
        pass


class FakeTreeview(_Widget):
    """A Treeview with real parent/child structure and scripted hit-testing."""

    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self._items = {}
        self._roots = []
        self._counter = 0
        # Tests set these to say what the next click lands on.
        self.next_identify_row = ""
        self.next_identify_element = ""
        self._selection = ()
        self._headings = {}
        self._columns = {}
        self._tags = {}

    def insert(self, parent, index, text="", values=(), tags=(), iid=None,
               open=False, **kwargs):
        self._counter += 1
        iid = iid or f"I{self._counter:03d}"
        self._items[iid] = {"text": text, "parent": parent, "children": [],
                            "open": open, "values": tuple(values),
                            "tags": (tags,) if isinstance(tags, str) else tuple(tags)}
        if parent:
            self._items[parent]["children"].append(iid)
        else:
            self._roots.append(iid)
        return iid

    def delete(self, *iids):
        for iid in iids:
            self._drop(iid)

    def _drop(self, iid):
        entry = self._items.pop(iid, None)
        if entry is None:
            return
        for child in list(entry["children"]):
            self._drop(child)
        if iid in self._roots:
            self._roots.remove(iid)
        parent = entry["parent"]
        if parent in self._items and iid in self._items[parent]["children"]:
            self._items[parent]["children"].remove(iid)

    def get_children(self, item=""):
        if not item:
            return tuple(self._roots)
        return tuple(self._items[item]["children"])

    def item(self, iid, option=None, **kwargs):
        entry = self._items[iid]
        if kwargs:
            entry.update(kwargs)
            return None
        if option == "text":
            return entry["text"]
        return entry

    def parent(self, iid):
        return self._items[iid]["parent"]

    def identify_row(self, y):
        return self.next_identify_row

    def identify_element(self, x, y):
        return self.next_identify_element

    def column(self, column, **kwargs):
        self._columns.setdefault(column, {}).update(kwargs)

    def heading(self, column, **kwargs):
        self._headings.setdefault(column, {}).update(kwargs)

    def set(self, iid, column=None, value=None):
        entry = self._items[iid]
        columns = tuple(self.options.get("columns", ()))
        if column is None:
            return dict(zip(columns, entry["values"]))
        index = columns.index(column)
        if value is None:
            return entry["values"][index]
        values = list(entry["values"]) + [""] * (len(columns) - len(entry["values"]))
        values[index] = value
        entry["values"] = tuple(values)
        return None

    def tag_configure(self, tag, **kwargs):
        self._tags.setdefault(tag, {}).update(kwargs)

    def selection(self):
        return tuple(i for i in self._selection if i in self._items)

    def selection_set(self, *iids):
        if len(iids) == 1 and isinstance(iids[0], (list, tuple)):
            iids = tuple(iids[0])
        self._selection = tuple(iids)

    def selection_remove(self, *iids):
        self._selection = tuple(i for i in self._selection if i not in iids)

    def focus(self, iid=None):
        if iid is not None:
            self._focus = iid
        return getattr(self, "_focus", "")

    def see(self, iid):
        pass

    def exists(self, iid):
        return iid in self._items

    def yview(self, *args):
        pass

    def xview(self, *args):
        pass

    def text_of(self, iid):
        return self._items[iid]["text"]

    def values_of(self, iid):
        return self._items[iid]["values"]

    def tags_of(self, iid):
        return self._items[iid]["tags"]


class FakeRoot(_Widget):
    def __init__(self, *args, **kwargs):
        super().__init__(None, **kwargs)
        self.destroyed = False
        self.protocols = {}
        self.tk = _TkCall()

    def title(self, *a):
        pass

    def geometry(self, *a):
        pass

    def minsize(self, *a):
        pass

    def protocol(self, name, func):
        self.protocols[name] = func

    def after(self, delay, func=None, *args):
        # Run immediately so a worker's completion is synchronous in tests.
        if func is not None:
            func(*args)

    def mainloop(self):
        pass

    def destroy(self):
        self.destroyed = True


class _TkCall:
    def __init__(self):
        self.calls = []

    def call(self, *args):
        self.calls.append(args)


class FakeToplevel(_Widget):
    """A secondary window; records the calls a dialog makes and its fate."""

    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self.destroyed = False
        self.protocols = {}

    def title(self, *a):
        pass

    def geometry(self, *a):
        pass

    def minsize(self, *a):
        pass

    def resizable(self, *a):
        pass

    def transient(self, *a):
        pass

    def grab_set(self):
        pass

    def grab_release(self):
        pass

    def lift(self):
        pass

    def wait_window(self, *a):
        pass

    def protocol(self, name, func):
        self.protocols[name] = func

    def destroy(self):
        self.destroyed = True


class Dialogs:
    """Records what the app tried to show, and scripts what the user answers."""

    def __init__(self):
        self.reset()

    def reset(self):
        self.info = []
        self.errors = []
        self.warnings = []
        self.questions = []
        self.answer = True

    def showinfo(self, title, message, **kwargs):
        self.info.append((title, message))

    def showerror(self, title, message, **kwargs):
        self.errors.append((title, message))

    def showwarning(self, title, message, **kwargs):
        self.warnings.append((title, message))

    def askyesno(self, title, message, **kwargs):
        self.questions.append((title, message))
        return self.answer


dialogs = Dialogs()


def install():
    """Put the stub modules into sys.modules. Call before importing `main`."""
    tk = types.ModuleType("tkinter")
    for name in ("BOTH", "X", "Y", "LEFT", "RIGHT", "TOP", "BOTTOM", "END",
                 "VERTICAL", "HORIZONTAL", "DISABLED", "NORMAL", "W", "E", "NS", "EW"):
        setattr(tk, name, name.lower())
    tk.StringVar = StringVar
    tk.BooleanVar = BooleanVar
    tk.Tk = FakeRoot
    tk.Toplevel = FakeToplevel
    tk.Frame = _Widget
    # main.selftest() reports these; a frozen build that shipped without Tcl/Tk
    # would fail on them, which is the point of checking.
    tk.TkVersion = 8.6
    tk.TclVersion = 8.6

    ttk = types.ModuleType("tkinter.ttk")
    for name in ("Frame", "LabelFrame", "Label", "Entry", "Button",
                 "Checkbutton", "Scrollbar"):
        setattr(ttk, name, _Widget)
    ttk.Treeview = FakeTreeview
    ttk.Style = _Style

    messagebox = types.ModuleType("tkinter.messagebox")
    messagebox.showinfo = dialogs.showinfo
    messagebox.showerror = dialogs.showerror
    messagebox.showwarning = dialogs.showwarning
    messagebox.askyesno = dialogs.askyesno

    filedialog = types.ModuleType("tkinter.filedialog")
    filedialog.askdirectory = lambda **kwargs: ""
    filedialog.askopenfilename = lambda **kwargs: ""
    filedialog.asksaveasfilename = lambda **kwargs: ""

    tk.ttk = ttk
    tk.messagebox = messagebox
    tk.filedialog = filedialog

    sys.modules["tkinter"] = tk
    sys.modules["tkinter.ttk"] = ttk
    sys.modules["tkinter.messagebox"] = messagebox
    sys.modules["tkinter.filedialog"] = filedialog
    return tk
