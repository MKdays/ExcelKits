from pathlib import Path
import glob
import shutil
import os

destination = r"C:\Users\MK\Desktop\dest"
path = r"C:\Users\MK\Desktop\TEST"
option = "/**"
output = glob.glob(path + option, recursive=True)

print(output)