# Get-SSIDs-info-for-multiple-Meraki-Networks-in-Excel-file
Get-SSIDs info for multiple Meraki Networks in Excel file. Using Meraki APIv1 with Python3.



Following modules/libraries necessary:

import requests
import json
import pandas as pd
import numpy as np
import os
import time
import logging
import openpyxl


Script was tested on multiple Meraki Networks using Windows Power Shell.

Script will automatically create report file in excel format and place it into the same folder as the script.
Script will also automatically create LOG file with useful info and place it into the same folder as the script.

User must fill two cells in Config_File_for_GET.xlsx:
   - API Key
   - Org ID

File "Config_File_for_GET.xlsx" must be in the same folder as the sript
