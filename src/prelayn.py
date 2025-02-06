"""
Main module for PRELAYN application.

This module serves as the entry point for the PRELAYN application.
It handles GUI creation and execution of the core functionality.

Usage:
    $ python prelayn.py

Dependencies:
    - comtypes
    - ezdxf
    - pyautocad
    - pyautogui
    - pywin32

License:
    GPL-3.0 license - See LICENSE for more information.

Created on: February 5, 2025
Version: 1.0.0

Author: Antonio Fernández
Email: antfdez@uvigo.es
"""

import os
from pathlib import Path
import sys
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import font as tkfont
from tkinter import messagebox
from tkinter import ttk
import traceback
import webbrowser

from comtypes import COMError
import ezdxf
from pyautocad import Autocad
import pyautogui as pgui
from pywintypes import com_error
from win32com.client.dynamic import Dispatch


class PrefixAdder():
    """Class for prefixing layer names through different packages."""
    RESERVED = ("0", "Defpoints")
    LAYER_0 = "0"
    CLAYER = "$CLAYER"
    # `NAMES` must be hardcoded for `add_prefix_pyautogui` to work
    NAMES = ["Layer1", "Layer2", "Layer3", "Layer4"]

    def __init__(self, prefix, package, infile, outfile):
        """Initialize instance.

        Parameters
        ----------
        prefix : str
            Prefix to be added to the layer names.
        package : str
            Name of Python package to be used in the prefixing task.
        infile : pathlib.Path
            Path of the input file.
        outfile : pathlib.Path
            Path of the output file.
        """
        self.prefix = prefix
        self.package = package
        self.infile = infile
        self.outfile = outfile

    def __call__(self):
        """Perform the prefixing task through `self.package`."""
        try:
            func = getattr(self, f"add_prefix_{self.package}")
        except AttributeError:
            raise UnknownPythonPackage("Unknown Python package")
        func()
    
    def add_prefix_pyautocad(self):
        """`pyautocad`-based implementation."""
        acad = Autocad(create_if_not_exists=True)
        for layer in acad.doc.Layers:
            name = layer.Name
            if name not in self.RESERVED:
                new_name = self.prefix + name
                layer.Name = new_name

    def add_prefix_ezdxf(self):
        """`ezdxf`-based implementation."""
        doc = ezdxf.readfile(self.infile)
        clayer_name = doc.header[self.CLAYER]
        doc.header[self.CLAYER] = self.LAYER_0
        names = [layer.dxf.name for layer in doc.layers]
        for name in names:
            if name not in self.RESERVED:
                layer = doc.layers.get(name)
                new_name = self.prefix + name
                layer.rename(new_name)
        doc.header[self.CLAYER] = self.prefix + clayer_name
        doc.saveas(self.outfile)

    def add_prefix_win32com(self):
        """`win32com`-based implementation."""
        acad = Dispatch("AutoCAD.Application")
        acad.Visible = True
        doc = acad.Documents.Open(self.infile)
        for layer in doc.Layers:
            name = layer.Name
            if name not in self.RESERVED:
                new_name = self.prefix + name
                layer.Name = new_name
        doc.SaveAs(self.outfile)

    def add_prefix_pyautogui(self):
        """`pyautogui`-based implementation."""
        def typewrite(keyboard_input, delay=1):
            """Helper function for typing text in AutoCAD with pyautogui."""
            pgui.typewrite(keyboard_input)
            pgui.keyDown("enter")
            pgui.keyUp("enter")
            time.sleep(delay)

        # If the file is already open, the following line:
        # if os.system(f"start '{infile}'") == 0:
        # throws this error:
        # "The process cannot access the file
        # because it is being used by another process."
        # In contrast, `os.startfile()` works fine even if the file is open.
        try:
            os.startfile(f"{self.infile}")
            time.sleep(3)
            for name in self.NAMES:
                typewrite("-LAYER")
                typewrite("Rename")
                typewrite(name)
                typewrite(self.prefix + name)
                pgui.hotkey("escape")
            typewrite("SAVEAS")
            # Casting `self.outfile` to string is necessary
            typewrite(str(self.outfile))
            pgui.hotkey("alt", "s")
            pgui.hotkey("alt", "s")
        except:
            raise


# >---------- EXCEPTIONS ----------< #
class PrefixNotSpecified(ValueError):
    """User did not enter a prefix."""

class PrefixNotValid(ValueError):
    """Prefix contains illegal character(s)."""

class FileNotSpecified(ValueError):
    """File not specified."""

class ExtensionNotCompatible(ValueError):
    """File extension not compatible with selected Python package."""

class NoPythonPackageSelected(ValueError):
    """Python package not specified."""

class UnknownPythonPackage(ValueError):
    """Unknown Python package."""


def shorten_path(long_path, limit=50):
    """Utility function for limiting the lenght of a given path by
    replacing the middle parts by elipsis (...).

    Parameters
    ----------
    long_path : pathlib.Path
        Path to be shortened.
    limit : int (optional, default=50)
        Maximum length of the shortened path.

    Returns
    -------
    The shortened path as a string.
    """
    long_string = str(long_path)
    if len(long_string) <= limit:
        return long_string
    parts = long_path.parts
    head = parts[0] + "..."
    remaining = limit - len(head)
    last = parts[-1]
    if len(last) > remaining:
        return head + last[-remaining:]
    else:
        tail = []
        for part in parts[:0:-1]:
            if remaining > len(part):
                tail.append(part)
                tail.append(os.sep)
                remaining -= len(part + os.sep)
            else:
                break
        return head + "".join(tail[::-1])


class Application(tk.Frame):
    """
    Layout of the GUI
    (symbols taken from: https://marklodato.github.io/js-boxdrawing/)

        self
        ╔════════════════════════════════════════════════════╗
     0) ║       frm_heading                                  ║
        ║       ┌────────────────────────────────────┐       ║
        ║       │ lbl_acronym                        │       ║
        ║       ├────────────────────────────────────┤       ║
        ║       │ lbl_title                          │       ║
        ║       └────────────────────────────────────┘       ║
        ╠════════════════════════════════════════════════════╣
     1) ║       frm_settings                                 ║
        ║       ┌──────────────────┬─────────────────┐       ║
        ║       │ lbl_prefix       │ ent_prefix      │       ║
        ║       ├──────────────────┼─────────────────┤       ║
        ║       │ lbl_package      │ cbx_package     │       ║
        ║       └──────────────────┴─────────────────┘       ║
        ╠════════════════════════════════════════════════════╣
     2) ║  frm_source                                        ║
        ║  ┌───────────────┬───────────────┬──────────────┐  ║
        ║  │ lbl_infile    │ ent_infile    │ btn_infile   │  ║
        ║  ├───────────────┼───────────────┼──────────────┤  ║
        ║  │ lbl_infolder  │ ent_infolder  │ btn_infolder │  ║
        ║  └───────────────┴───────────────┴──────────────┘  ║
        ╠════════════════════════════════════════════════════╣
     3) ║  frm_destination                                   ║
        ║  ┌───────────────┬───────────────┬──────────────┐  ║
        ║  │ lbl_infile    │ ent_infile    │ btn_infile   │  ║
        ║  ├───────────────┼───────────────┼──────────────┤  ║
        ║  │ lbl_infolder  │ ent_infolder  │ btn_infolder │  ║
        ║  └───────────────┴───────────────┴──────────────┘  ║
        ╠════════════════════════════════════════════════════╣
     4) ║        frm_actions                                 ║
        ║        ┌───────────┬───────────┬───────────┐       ║
        ║        │ btn_run   │ btn_help  │ btn_exit  │       ║
        ║        └───────────┴───────────┴───────────┘       ║
        ╠════════════════════════════════════════════════════╣
     5) ║  ┌──────────────────────────────────────────────┐  ║
        ║  │ lbl_status                                   │  ║
        ║  └──────────────────────────────────────────────┘  ║
        ╚════════════════════════════════════════════════════╝

    """

    FONTSIZE = 10
    HUGE_FONTSIZE = 16
    LARGE_FONTSIZE = 13
    HEADING_FONT = "Courier"

    NUM_CHARS_SMALL = 16
    NUM_CHARS_MEDIUM = 30
    NUM_CHARS_LARGE = 50
    BUTTON_WIDTH = 9
    PADDING = 5

    ROW_HEADING = 0
    ROW_SETTINGS = 1
    ROW_SOURCE = 2
    ROW_DESTINATION = 3
    ROW_ACTIONS = 4
    ROW_STATUS = 5

    ILLEGAL = set("<>\\/\":;*?|,=`")
    
    PYAUTOCAD = "pyautocad"
    EZDXF = "ezdxf"
    WIN32COM = "win32com"
    PYAUTOGUI = "pyautogui"

    DWG = ".dwg"
    DXF = ".dxf"
    
    PACKAGES = (PYAUTOCAD, EZDXF, WIN32COM, PYAUTOGUI)
    REQUIRES_DWG = (PYAUTOGUI, WIN32COM)
    REQUIRES_DXF = (EZDXF,)

    def __init__(self, master=None):
        """Initialize GUI.

        Parameters
        ----------
        master : tkinter.Tk
            Root window.
        """
        super().__init__(master)
        self.master = master
        self.grid()

        self.NO_ERROR = ""

        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=self.FONTSIZE)
        self.master.option_add("*Font", default_font)

        self.cwd = Path.cwd()
        self.base_folder = self.get_base_folder()

        self.infolder = self.cwd
        self.outfolder = self.cwd

        self.sv_prefix = tk.StringVar(value="")
        self.sv_infile = tk.StringVar(value="")
        short_infolder = shorten_path(self.infolder, self.NUM_CHARS_LARGE)
        self.sv_infolder = tk.StringVar(value=short_infolder)
        self.sv_outfile = tk.StringVar(value="")
        short_outfolder = shorten_path(self.outfolder, self.NUM_CHARS_LARGE)
        self.sv_outfolder = tk.StringVar(value=short_outfolder)
        self.sv_status = tk.StringVar(value=self.NO_ERROR)

        self.master.title(Path(__file__).name)
        icon_path = self.base_folder.joinpath("python-icon-multisize.ico")
        self.master.wm_iconbitmap(icon_path)

        self.create_heading()
        self.create_settings()
        self.create_source()
        self.create_destination()
        self.create_actions()
        self.create_status()

        self.ent_prefix.focus_set()
        self.master.minsize(self.master.winfo_width(), self.master.winfo_height())
        self.master.resizable(0, 0)
        self.master.update()


    def get_base_folder(self):
        """Return the base folder, i.e.:
        - `sys._MEIPASS` if the program is run as a standalone
          executable bundled by Pyinstaller.
        - The folder that contains the script if the program
          is run in a normal Python process.
        """
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # Running in a PyInstaller bundle
            return Path(sys._MEIPASS)
        else:
            # Running in a normal Python process
            return Path(__file__).parent.resolve()        

    # >---------- HEADING ----------< #
    def create_heading(self):
        """Create widgets for the HEADING section of the GUI."""
        self.frm_heading = ttk.Frame(
            self,
            relief=tk.RIDGE,
            borderwidth=10,
        )
        self.frm_heading.grid(
            row=self.ROW_HEADING,
            padx=self.PADDING,
            pady=self.PADDING,
        )

        self.lbl_acronym = ttk.Label(self.frm_heading, text="PRELAYN")
        acronym_font = (self.HEADING_FONT, self.HUGE_FONTSIZE)
        self.lbl_acronym.config(font=acronym_font)
        self.lbl_acronym.grid(row=0, column=0)

        self.lbl_title = ttk.Label(
            self.frm_heading,
            text="PREfix LAyer Names of an AutoCAD drawing",
        )
        title_font = (self.HEADING_FONT, self.LARGE_FONTSIZE)
        self.lbl_title.config(font=title_font)
        self.lbl_title.grid(row=1, column=0)


    # >---------- SETTINGS ----------< #
    def create_settings(self):
        """Create widgets for the SETTINGS section of the GUI."""
        self.frm_settings = ttk.Frame(self)
        self.frm_settings.grid(
            row=self.ROW_SETTINGS,
            padx=self.PADDING,
            pady=self.PADDING,
        )

        self.lbl_prefix = self.factory_label(
            self.frm_settings,
            text="Prefix",
        )
        self.lbl_prefix.grid(row=0, column=0)

        self.ent_prefix = ttk.Entry(
            self.frm_settings,
            textvariable=self.sv_prefix,
            width=self.NUM_CHARS_MEDIUM,
        )
        self.ent_prefix.grid(row=0, column=1, sticky="EW")
        self.ent_prefix.bind(
            "<FocusOut>",
            self.callback_prefix_focusout,
        )

        self.lbl_package = self.factory_label(
            self.frm_settings, text="Python package")
        self.lbl_package.grid(row=1, column=0)

        self.cbox_package = ttk.Combobox(
            self.frm_settings,
            values=self.PACKAGES,
            state="readonly",
            width=self.NUM_CHARS_MEDIUM,
        )
        self.cbox_package.grid(row=1, column=1, sticky="EW")
        self.cbox_package.bind(
            "<<ComboboxSelected>>",
            self.callback_package_selected,
        )
        # self.cbox_package.bind(
        #     "<FocusOut>",
        #     self.callback_package_focusout,
        # )


    # >---------- SOURCE ----------< #
    def create_source(self):
        """Create widgets for the SOURCE section of the GUI."""
        self.frm_source = ttk.Frame(self)
        self.frm_source.grid(
            row=self.ROW_SOURCE,
            padx=self.PADDING,
            pady=self.PADDING,
        )   

        self.lbl_infile = self.factory_label(
            self.frm_source,
            text="Input file",
        )
        self.lbl_infile.grid(row=0, column=0)

        self.ent_infile = ttk.Entry(
            self.frm_source,
            textvariable=self.sv_infile,
            width=self.NUM_CHARS_LARGE,
        )
        self.ent_infile.grid(row=0, column=1, sticky="EW")
        self.ent_infile.bind(
            "<FocusOut>",
            self.callback_infile_focusout,
        )
        self.ent_infile.config(state="disabled")

        self.btn_infile = self.factory_button(
            self.frm_source,
            text="Browse...",
            command=self.callback_select_infile,
        )
        self.btn_infile.grid(row=0, column=2)
        self.btn_infile.config(state="disabled")

        self.lbl_infolder = self.factory_label(
            self.frm_source,
            text="Input folder",
        )
        self.lbl_infolder.grid(row=1, column=0)

        self.ent_infolder = ttk.Entry(
            self.frm_source,
            textvariable=self.sv_infolder,
            width=self.NUM_CHARS_LARGE,
        )
        self.ent_infolder.grid(row=1, column=1, sticky="EW")
        self.ent_infolder.config(state="disabled")

        self.btn_infolder = self.factory_button(
            self.frm_source,
            text="Change...",
            command=self.callback_select_infolder,
        )
        self.btn_infolder.grid(
            row=1,
            column=2,
            padx=self.PADDING,
        )
        self.btn_infolder.config(state="disabled")


    # >---------- DESTINATION ----------< #
    def create_destination(self):
        """Create widgets for the DESTINATION section of the GUI."""
        self.frm_destination = ttk.Frame(self)
        self.frm_destination.grid(
            row=self.ROW_DESTINATION,
            padx=self.PADDING,
            pady=self.PADDING,
        )

        self.lbl_outfile = self.factory_label(
            self.frm_destination,
            text="Output file",
        )
        self.lbl_outfile.grid(row=0, column=0)

        self.ent_outfile = ttk.Entry(
            self.frm_destination, textvariable=self.sv_outfile,
            width=self.NUM_CHARS_LARGE,
        )
        self.ent_outfile.grid(row=0, column=1, sticky="EW")
        self.ent_outfile.bind(
            "<FocusOut>",
            self.callback_outfile_focusout,
        )
        self.ent_outfile.config(state="disabled")

        self.btn_outfile = self.factory_button(
            self.frm_destination,
            text="Browse...",
            command=self.callback_select_outfile,
        )
        self.btn_outfile.grid(row=0, column=2)
        self.btn_outfile.config(state="disabled")

        self.lbl_outfolder = self.factory_label(
            self.frm_destination,
            text="Output folder",
        )
        self.lbl_outfolder.grid(row=1, column=0)

        self.ent_outfolder = ttk.Entry(
            self.frm_destination,
            textvariable=self.sv_outfolder,
            width=self.NUM_CHARS_LARGE,
        )
        self.ent_outfolder.grid(row=1, column=1, sticky="EW")
        self.ent_outfolder.config(state="disabled")

        self.btn_outfolder = self.factory_button(
            self.frm_destination,
            text="Change...",
            command=self.callback_select_outfolder,
        )
        self.btn_outfolder.grid(row=1, column=2, padx=self.PADDING)
        self.btn_outfolder.config(state="disabled")


    # >---------- ACTIONS ----------< #
    def create_actions(self):
        """Create widgets for the ACTIONS section of the GUI."""
        self.frm_actions = ttk.Frame(self)
        self.frm_actions.grid(
            row=self.ROW_ACTIONS,
            padx=self.PADDING,
            pady=self.PADDING,
        )

        self.btn_run = self.factory_button(
            self.frm_actions,
            text="Run",
            command=self.run,
        )
        self.btn_run.grid(row=0, column=0, padx=20)

        self.btn_help = self.factory_button(
            self.frm_actions,
            text="Help",
            command=self.help,
        )
        self.btn_help.grid(row=0, column=1, padx=20)

        self.btn_exit = self.factory_button(
            self.frm_actions,
            text="Exit",
            command=self.master.destroy,
        )
        self.btn_exit.grid(row=0, column=2, padx=20)


    # >---------- STATUS BAR ----------< #
    def create_status(self):
        """Create widgets for the STATUS section of the GUI."""
        self.lbl_status = ttk.Label(
            self,
            textvariable=self.sv_status,
            relief=tk.SUNKEN,
        )
        self.lbl_status.grid(
            row=self.ROW_STATUS,
            sticky="WE",
            padx=self.PADDING,
            pady=self.PADDING,
        )


    # >---------- CALLBACKS ----------< #
    def callback_prefix_focusout(self, event):
        """Function invoked when the input focus is moved out
        of a the Prefix entry."""
        self.do_checks(self.check_prefix)


    def callback_package_selected(self, event):
        package = self.cbox_package.get()
        new_state = "disabled" if package == self.PYAUTOCAD else "normal"
        self.ent_infile.config(state=new_state)
        self.ent_outfile.config(state=new_state)
        self.btn_infile.config(state=new_state)
        self.btn_infolder.config(state=new_state)
        self.btn_outfile.config(state=new_state)
        self.btn_outfolder.config(state=new_state)
        self.do_checks(self.check_package)


    def callback_package_focusout(self, event):
        self.do_checks(self.check_package)


    def callback_infile_focusout(self, event):
        self.do_checks(self.check_infile)


    def callback_outfile_focusout(self, event):
        self.do_checks(self.check_outfile)


    def get_file(self, initialdir, title):
        """Select a file through a dialog box.

        Parameters
        ----------
        initialdir : pathlib.Path
            The directory that the dialog starts in.
        title : str
            The title of the window.

        Returns
        -------
        file_path : pathlib.Path
            Path of the selected file.
        """
        if not initialdir or not initialdir.is_dir():
            initialdir = self.cwd
        package = self.cbox_package.get()
        # https://stackoverflow.com/questions/61456040/tkinter-ask-filedialog-avoid-internet-link
        # Files with extension .url are always shown
        if package in self.REQUIRES_DWG:
            filetypes = [("Drawing files", self.DWG)]
        elif package in self.REQUIRES_DXF:
            filetypes = [("Exchange format files", self.DXF)]
        else:
            filetypes = [("All files", ".*")]
        file_path = Path(
            filedialog.askopenfilename(
                master=self,
                initialdir=initialdir,
                title=title,
                filetypes = filetypes,
            )
        )
        return file_path


    def callback_select_infile(self):
        """Select input file and check that it is valid."""
        file_path = self.get_file(self.infolder,"Select input file")
        parent_folder = file_path.parent
        filename = file_path.name
        self.sv_infile.set(filename)
        self.infolder = parent_folder
        self.sv_infolder.set(shorten_path(parent_folder))
        self.do_checks(self.check_infile)


    def callback_select_outfile(self):
        """Select output file and check that it is valid."""
        file_path = self.get_file(self.infolder,"Select output file")
        parent_folder = file_path.parent
        filename = file_path.name
        self.sv_outfile.set(filename)
        self.outfolder = parent_folder
        self.sv_outfolder.set(shorten_path(parent_folder))
        self.do_checks(self.check_outfile)


    def get_folder(self, initialdir, title):
        """Select a folder through a dialog box.

        Parameters
        ----------
        initialdir : pathlib.Path
            The directory that the dialog starts in.
        title : str
            The title of the window.

        Returns
        -------
        dir_path : pathlib.Path
            Path of the selected directory.
        """
        if not initialdir or not initialdir.is_dir():
            initialdir = self.cwd
        dir_path = Path(
            filedialog.askdirectory(
                master=self,
                initialdir=initialdir,
                title=title,
                mustexist=True
            )
        )
        return dir_path


    def callback_select_infolder(self):
        """Select input folder and check if it exists."""
        dir_path = self.get_folder(self.infolder, "Select input folder")
        self.infolder = dir_path
        self.sv_infolder.set(shorten_path(dir_path))
        self.do_checks(self.check_infolder)


    def callback_select_outfolder(self):
        """Select output folder and check if it exists."""
        dir_path = self.get_folder(self.outfolder, "Select output folder")
        self.outfolder = dir_path
        self.sv_outfolder.set(shorten_path(dir_path))
        self.do_checks(self.check_outfolder)


    # >---------- CHECKS ----------< #
    def check_prefix(self):
        """Check that user has entered a valid prefix.

        Returns
        -------
        None.

        Raises
        ------
        PrefixNotSpecified or PrefixNotValid.
        """
        prefix = self.sv_prefix.get()
        if not prefix:
            raise PrefixNotSpecified("Prefix cannot be empty")
        if set(prefix).intersection(self.ILLEGAL):
            raise PrefixNotValid(f"Please enter a valid prefix")


    def check_package(self):
        """Check that user has selected a Python package
        from the dropdown list.

        Returns
        -------
        None.

        Raises
        ------
        NoPythonPackageSelected.
        """
        package = self.cbox_package.get()
        if not package:
            raise NoPythonPackageSelected("Please select a Python package")


    def check_infolder(self):
        """Check that input folder exists.

        Returns
        -------
        None.

        Raises
        ------
        FileNotFoundError.
        """
        if not self.infolder.is_dir():
            raise FileNotFoundError("Input folder not found")


    def check_outfolder(self):
        """Check that output folder exists.

        Returns
        -------
        None.

        Raises
        ------
        FileNotFoundError.
        """
        if not self.outfolder.is_dir():
            raise FileNotFoundEror("Output folder not found")


    def is_extension_compatible(self, extension):
        """Check if a file extension is compatible with the
        selected Python package.

        Parameters
        ----------
        extension : str
            Name of input or output file.

        Returns
        -------
        True if file extension is compatible with package,
        False otherwise.
        """
        package = self.cbox_package.get()
        dwg_fail = (extension != self.DWG) and (package in self.REQUIRES_DWG)
        dxf_fail = (extension != self.DXF) and (package in self.REQUIRES_DXF)
        return not(dxf_fail or dwg_fail)


    def check_infile(self):
        """Check that user has specified an input file, it exists,
        and its extension is compatible with the selected package.

        Returns
        -------
        None.

        Raises
        ------
        FileNotSpecified, FileNotFoundError or ExtensionNotCompatible.
        """
        folder = self.infolder
        filename = self.sv_infile.get()
        if not filename:
            raise FileNotSpecified("Please specify the input file")
        file_path = Path(folder).joinpath(filename)
        if not file_path.is_file():
            raise FileNotFoundError("Input file not found")
        suffix = file_path.suffix.casefold()
        if not self.is_extension_compatible(suffix):
            raise ExtensionNotCompatible(
                "Input file extension not compatible "
                "with selected Python package"
            )


    def check_outfile(self):
        """Check that user has specified an output file, and its
        extension is compatible with the selected Python package.

        Returns
        -------
        None.

        Raises
        ------
        FileNotSpecified, FileNotFoundError or ExtensionNotCompatible.
        """
        folder = self.outfolder
        filename = self.sv_outfile.get()
        if not filename:
            raise FileNotSpecified("Please specify the output file")
        suffix = Path(filename).suffix.casefold()
        if not self.is_extension_compatible(suffix):
            raise ExtensionNotCompatible(
                "Output file extension not compatible "
                "with selected Python package"
            )


    def do_checks(self, *checks):
        """Perform a series of checks by calling the passed methods.

        Parameters
        ----------
        checks : tuple
            Comma-separated methods to be called for checking that
            everything is fine before executing a prefixing task.
        """
        try:
            for check in checks:
                check()
        except Exception as exc:
            message = exc.args[0]
        else:
            message = self.NO_ERROR
        finally:
            self.sv_status.set(message)


    # >---------- FACTORIES ----------< #
    def factory_button(self, master, text, command):
        """Wrapper for creating buttons."""
        return ttk.Button(
            master,
            text=text,
            width=self.BUTTON_WIDTH,
            command=command,
        )

    def factory_label(self, master, text):
        """Wrapper for creating labels."""
        font = tkfont.nametofont("TkDefaultFont")
        bold_font = (font["family"], font["size"], "bold")
        return  ttk.Label(
            master,
            text=text+":",
            width=self.NUM_CHARS_SMALL,
            padding=5,
            anchor="w",
            font=bold_font,
        )


    def run(self):
        """Perform prefixing task."""
        self.do_checks(
            self.check_prefix,
            self.check_package,
        )
        package = self.cbox_package.get()
        if package in (self.REQUIRES_DWG + self.REQUIRES_DXF):
            self.do_checks(
                self.check_infolder,
                self.check_infile,
                self.check_outfolder,
                self.check_outfile,
            )
        status_after_checks = self.sv_status.get()
        if status_after_checks != self.NO_ERROR:
            return

        self.sv_status.set("Adding prefix to layer names...")
        # Necessary to update the status bar
        self.master.update()
        package = self.cbox_package.get()
        prefix = self.sv_prefix.get()
        infile = self.infolder.joinpath(self.sv_infile.get())
        outfile = self.outfolder.joinpath(self.sv_outfile.get())

        try:
            info = None
            PrefixAdder(prefix, package, infile, outfile)()

        except (COMError, com_error) as exc:
            info = handle_com_exception(exc)
            #display_exception_data(exc)

            mbox = messagebox.showerror(
                master=self,
                title=exc.__class__.__name__,
                message=traceback.format_exc(),
                detail='Please check open files in AutoCAD and try again',
            )

        except Exception as exc:
            info = f"{exc.__class__.__name__} >>> {exc.__doc__}"
            #display_exception_data(exc)

        else:
            info = "Done"

        finally:
            if info is not None:
                #print(f"{info}\n")
                self.sv_status.set(info)
                self.master.update()


    def help(self):
        """Display help file in default browser."""
        html_file = self.base_folder.joinpath("help.html")
        if html_file.is_file():
            successfully = webbrowser.open(html_file)
            if successfully:
                self.sv_status.set("Help is being displayed on the browser")
            else:
                self.sv_status.set("Unable to display help on the browser")
        else:
            self.sv_status.set(f'"{shorten_path(html_file)}" not found')



def handle_com_exception(exc):
    """Handle COMError and com_error."""
    # By analyzing the tracebacks of `COMError` and `com_error`
    # exceptions, I figured out this:
    #           +-----------+-------------+
    #           | COMError  | com_error   |
    #    +------+-----------+-------------+
    #    | arg2 | 'text'    | 'strerror'  |
    #    | arg3 | 'details' | 'excepinfo' |
    #    | idx  | 0         | 2           |
    #    +------+-----------+-------------+
    if isinstance(exc, COMError):
        arg2, arg3, idx = "text", "details", 0
    elif isinstance(exc, com_error):
        arg2, arg3, idx = "strerror", "excepinfo", 2

    name = exc.__class__.__name__
    if hasattr(exc, arg3) and exc.__dict__[arg3] is not None:
        try:
            seq = exc.__dict__[arg3]
            #for index, item in enumerate(seq):
            #    print(f"  $ {name}.{arg3}[{index}]: {item}")
            info = f"{name} >>> {seq[idx]}"
        except Exception as err:
            another = err.__class__.__name__
            info = (f"{another} >>> {name}.{arg3}[{idx}]")
    else:
        txt = exc.__dict__[arg2]
        if txt is not None:
            info = f"{name} >>> {txt}"
        else:
            info = f"{name} >>> No information found for this error"
    return info


def display_exception_data(exc):
    """Utility function for debugging."""
    tb = exc.__traceback__
    print("========== dir(exc.__traceback__) ==========")
    for attr in dir(tb):
        print(f"{attr}: {getattr(tb, attr, "???")}")
    excname = exc.__class__.__name__
    print(f"vvvvvvvvvv handle_exception({excname}) vvvvvvvvvv")
    for key, value in (exc.__dict__.items()):
        print(f"{excname}.{key}: {value}")
    print(f"^^^^^^^^^^ handle_exception({excname}) ^^^^^^^^^^\n")


if __name__ == "__main__":
    root = tk.Tk()    
    app = Application(master=root)
    app.mainloop()
