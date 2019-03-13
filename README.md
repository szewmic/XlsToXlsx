# XLS to XLSX Converter!

Hi! I've created this app for my current work place, because we had trouble with a lot of files saved in old .xls format instead of much lighter .xlsx. 

Application is pretty simple and use Microsoft Excels Interop so **you need to have Microsoft Excel** installed on your workstation. 

Technologies used here are: **C# .NET, WinForms** GUI and simple **SmartUI** design pattern.

## Problem
We had hundreds of .xls files on our network disc with weight up to **80 MB** each! Opening manually everyone of them and saving with newer format would take us enormous amount of time.

## Solution
Simple Application with needed for us functionallity. That searches selected folder and all of his subfolders for files with ".xls" extension then converting them to ".xlsx" extensions.

Since I couldn't find any good free library that can help me with changing extensions, I decided to use Microsoft Excels Interop.

**Running this application on our network locations reduced the used disc space up to 50%!**

# Application functionallity
![alt text](https://i.ibb.co/RPcnmvS/xlstoxlsxappimg.png)

Functionallity includes:
- selecting path to folder which folders and subfolders will be searched for ".xls" extension files,
- decide minimum weight from which files will be converted (no point to convert already leighweight files),
- checkbox to decide if you want delete previous ".xls" files or just let them be
- progress bar obviously shows you improvment in converting task.

After process is done you will get **Log file** with .txt extension. That looks like this:
![alt text](https://i.ibb.co/6JVqywv/logimg.png)
And contains informations about every new succesfully created ".xlsx" file and every optionally deleted ".xls" file. It shows also information about their weights and the if you selected option to remove old ".xls" files **you will get info about how much space on disc you gained!**


## Authors
- Michał Szewczak