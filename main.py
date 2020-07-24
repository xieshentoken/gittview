import re
import sys
from collections import OrderedDict
from tkinter import *
from tkinter import colorchooser, filedialog, messagebox, simpledialog, ttk

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import scipy.signal as signal
import xlrd

from GUI import App
from process import Gitt


def main():
    root = Tk()
    root.title("gitt")
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()