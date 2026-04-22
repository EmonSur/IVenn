# IVenn

IVenn is a Python API for exploring overlaps between up to 6 sets using [InteractiVenn](https://www.interactivenn.net/)-style diagrams. It can load sets from Python or from an Excel file, display them in a viewer, step through union views, inspect regions and sets, and export results as SVG, PNG, or Excel files.

---

## Features

- Draw InteractiVenn-style diagrams for up to 6 sets  
- Define sets in Python or load them from Excel  
- Explore intersections through union views  
- Support both list unions and tree unions  
- Open an interactive GUI viewer  
- Inspect the elements in a selected region  
- Inspect the elements in a selected set  
- Export diagrams as SVG or PNG  
- Export original sets and intersections to Excel  
- Change diagram theme, text size, opacity, and percentage display  

---

## Requirements

Before installing IVenn, make sure you have:

- **Python 3.10 or newer**
- **pip**

You can check your Python version with:

```
python --version
```

or:

```
python3 --version
```

---

## Installation

At the moment, IVenn is used by cloning the repository and installing it locally.

### 1. Clone the repository

```
git clone https://github.com/EmonSur/IVenn.git
cd IVenn
```

### 2. Create a virtual environment

#### Windows (PowerShell)

```
python -m venv .venv
.venv\Scripts\Activate.ps1
```

#### Linux / macOS

```
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install dependencies

```
pip install .
```

### 4. Test that the installation works

```
python examples/load_sets.py
```

---

## Quick start

### Option 1: Load sets from Excel

```
from ivenn import IVenn

v = IVenn.from_excel("my_sets.xlsx")
v.draw()
```

### Option 2: Create sets directly in Python

```
from ivenn import IVenn, Set

a = Set("Set A", ["1", "2", "3"])
b = Set("Set B", ["2", "3", "4"])
c = Set("Set C", ["3", "4", "5"])

v = IVenn(a, b, c)
v.draw()
```

---

## Excel input format

- each **column** is treated as one set  
- the **column name** becomes the set label  
- each non-empty cell becomes an element  

---

## Viewer

```
v.draw()
```

The viewer supports:

- zooming and panning  
- stepping through union views  
- theme and display settings  
- clicking regions to inspect elements  
- exporting results  

---

## Basic methods

```
v.set_sizes()
v.regions()
v.region_elements()
v.get_region("AB")
v.top_intersections()
```

---

## Working with unions

### List unions

```
v.set_unions("AB,CD;ABC")
```

### Tree unions

```
v.set_unions("((A,B),C)")
```

---

## Export

```
v.export_svg("diagram.svg")
v.export_png("diagram.png")
v.export_sets("sets.xlsx")
v.export_intersections("intersections.xlsx")
```

---

## Troubleshooting

If imports fail, reinstall:

```
pip install .
```
