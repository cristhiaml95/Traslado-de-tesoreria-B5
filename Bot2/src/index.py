import win32com.client
import subprocess
import time
from openpyxl import Workbook
from openpyxl import load_workbook
import re
from usefulFunctions import *
from usefulObjets import *

if __name__=='__main__':
    x = sapInterfaceJob()
    x.fullProcess()
