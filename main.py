import re
import sys
from collections import OrderedDict
from tkinter import *
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import xlrd

from GUI import App
from process import Gitt
from hslrgb import *

matplotlib.rcParams['font.sans-serif'] = ['SimHei']

def main():
    root = Tk()
    root.title("GittView")
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
