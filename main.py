import sys
import os

if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))

os.chdir(base_dir)
sys.path.insert(0, base_dir)

import tkinter as tk
from app import ConciliacaoCartaoApp

if __name__ == '__main__':
    root = tk.Tk()
    app = ConciliacaoCartaoApp(root)
    root.mainloop()
