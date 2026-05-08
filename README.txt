# Excel Python Setup Guide

Follow these steps to set up Python, install the required packages, and enable the Excel add-ins.

## 1. Download And Install Python

1. Go to the official Python download page:

   https://www.python.org/downloads/

2. Download and install Python for Windows.

3. During installation, make sure to check:

   "Add python.exe to PATH"

4. Complete the installation.

## 2. Open Terminal

Open **Command Prompt** or **PowerShell**.

You can open it by pressing `Windows + R`, type:

    "cmd"

and press `Enter`.

## 3. Install Required Python Packages

In the terminal, run these commands one by one:

   powershell
     i)  pip install xlwings

     ii) pip install numpy

    iii) pip install matplotlib


Wait for each command to finish before running the next one.

## 4. Open Excel Options

1. Open Microsoft Excel.
2. Click "File"
3. Click "Options"

## 5. Open Add-ins

1. In the Excel Options window, click "Add-ins"
2. At the bottom, find the "Manage" section.
3. Select "Excel Add-ins"
4. Click "Go"

## 6. Enable Required Add-ins

In the Add-ins window, check these options:

   i) Analysis ToolPak - VBA

  ii) Solver Add-in

Also check the "xlwings" option if it is available.

## 7. Add xlwings If It Is Not Available

If "xlwings" is not available in the Add-ins list:

1. Click "Browse"
2. Go to this location:

   C:\Users\your_user\AppData\Local\Python\pythoncore-3.14-64\Lib\site-packages\xlwings\addin

Note: Important: replace `your_user` with your Windows username.
Example:

C:\Users\Ahmar\AppData\Local\Python\pythoncore-3.14-64\Lib\site-packages\xlwings\addin

3. Select this file:

   "xlwings.xlam"
  

4. Click "OK"
5. Now the "xlwings" option should appear in the Add-ins list.
6. Check the "xlwings" option.
7. Click "OK"
8. Close the Excel Options window.



## 8. Check xlwings Ribbon

After enabling the add-in, you should see an "xlwings" option in the top ribbon bar of Excel.

## 9. Open The Excel File

Now everything is set up.

Open the Excel file and enjoy.

Bonus: If you want to modify or change coding according to your requirement 
Download and install Visual Studio code from https://code.visualstudio.com/download
download required extensions like python, python debugger etc and enjoy

Author: Engr Ahmar Shahzad
UET Lahore main branch graduate 
Email: ahmarshah01@gmail.com
Email: ahmarshah136@gmail.com
Ph# +92 3324402178
