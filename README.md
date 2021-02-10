# python-excel
This is a project that aims at enhancing excel spreadsheets with Python.

# About the project

## Transitioning from Excel to Python

Many industries are increasingly moving to Python as it can offer substantial benefits over Excel. A Python program is easily reproducible, doesn't suffer scalability issues and is more efficient than Excel. 

## Working with Excel and Python

As Excel is still more used than Python, and many people do not know how to code in Python, one way to operate a transition to Python could be to use both Excel and Python. This will be the case in this project.

## What is this project about

The goal of this project is to demonstrate what can be done by combining Python and Excel. Because of my studies in finance, this demonstration will have a financial dimension to it: We will import price data from exchanges and organize it

## How to use the project

### Create a virtual environment

Open Windows powershell and navigate to the cloned repository

```cd python-excel```

Create a virtual environment with ```python -m venv venv``` and activate it with ```venv\scripts\activate```

Install all dependencies using ```pip install requirements.txt```

```python -m ipykernel install --user --name=venv```

## xlwings Excel addin

In the ```/venv``` directory you will find two xlwings .dll files. Move them to ```/venv/Scripts```

Check your version of xlwings with ```pip list``` and go to https://github.com/xlwings/xlwings/releases to download the appropriate .xlam file
