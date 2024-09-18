# import sys
# sys.path.append("C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\Working_attributes")
# from Imports import * 



import sys
try:
    sys.path.append("C:\\Users\\KlimovaAnnaA\\Documents\\MyFiles\\Projects\\Working_attributes")
except:
    sys.path.append("z:\\Anna_Klimova\\Working_attributes")
    
from Defs import merge_SalesUnits
from Defs import merge_Mapping
from Defs import Period
from Defs import new_list
from Defs import export_from_RISKCUSTOM
from Defs import add_in_currency_column
from Defs import concat_columns
from Defs import export_from_WHWEEK
from Defs import CCY_tech_dict
from Defs import is_approximately_equal_for_cols
from Defs import holding_list


import pandas as pd
import numpy as np
from datetime import date, timedelta
import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.axes as ax
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl.drawing.image import Image
import win32com.client as win32
import os
from PIL import ImageGrab
import re

olApp = win32.Dispatch('Outlook.Application')
olNS = olApp.GetNameSpace('MAPI')

Output_path = 'z:\\Anna_Klimova\\OCP\\Archive\\'

