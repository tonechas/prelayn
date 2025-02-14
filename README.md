# prelayn
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)
[![Python 3.12](https://img.shields.io/badge/python-3.12-blue.svg)](https://www.python.org/downloads/release/python-3120/)

Application for automatically **PRE**fixing **LAY**er **N**ames of AutoCAD drawings. This task can be performed using four different libraries:

- `win32com` ([`pywin32`](https://github.com/mhammond/pywin32))
- [`pyautocad`](https://github.com/reclosedev/pyautocad)
- [`pyautogui`](https://github.com/asweigart/pyautogui)
- [`ezdxf`](https://github.com/mozman/ezdxf)

## Installation

1. Clone this repository:
```console
C:\Users\Me>git clone https://github.com/tonechas/prelayn.git
```

2. Create a virtual environment:
```console
C:\Users\Me>python -m venv path\to\venvs\directory\myvenv python=3.12
```

3. Activate the virtual environment:
```console
C:\Users\Me>path\to\venvs\directory\myvenv\Scripts\activate
```

4. Install the dependencies:
```console
(myvenv) C:\Users\Me>cd prelayn
(myvenv) C:\Users\Me\prelayn>pip install -r requirements.txt
```

## Usage
To run the application, execute the following command:
```console
(myvenv) C:\Users\Me\prelayn>python src\prelayn.py
```

After that, a graphical user interface will pop up, boasting a help button that makes the program easy to use.

<img src="./imgs/gui.JPG" alt="GUI" width="auto">

## Platform support

The project supports the following operating systems:

| Operating System | Supported |
|------------------|-----------|
| Windows          | Yes       |
| Mac OS           | No        |
| Linux            | No        |
| Other Unix-like  | No        |

## Compatibility
The project has been developed and tested on a Windows 10 Pro OS using Python 3.12.0 and AutoCAD 2023.
