# Excel-Module-Logger
## How to use
A module for Microsoft Excel that allows logging to a file.

All you have to do is download the Logger.bas and then in Excel under Visual Basic, Right Click on Modules and select "Import File" Select Logger.bas from where you downloaded it.

See Customization part of the readme on how to change how the log file is set up to be more of how you want it.

You can now call the sub in the following way.
```visual-basic
Call WriteToLog("Action")
```
Where Action is replaced by what you want written.


For example if in one of your macros you want to display that the user selected a sheet you would have the following in the sheet code.

```visual-basic
Private Sub Worksheet_Activate()
Call WriteToLog(Currentsheet.Name & " was selected")
End Sub
```

This will write something like "Sheet1 was selected" in the action part of the log.

For a better example of what you can do with this kind of code, in realses I've included an excel file with several examples.


## The Log File

Depending some of the constants you can change will effect the name of the file, however for now the set up of the file itself is all the same.

Below is an example

```
Date           Time            User       | Action
-------------------------------------------------------------------------------------------------
2015-07-20     09:33:58     GamerMan7799 | Workbook Saved
2015-07-20     09:34:01     GamerMan7799 | Report Menu opened
2015-07-20     09:34:57     GamerMan7799 | Workbook Saved
2015-07-20     09:36:11     GamerMan7799 | Workbook Saved
2015-07-20     09:36:13     GamerMan7799 | Workbook Closed
```



## Customization

There are currently 7 constant that you can change to cause the log to be saved differently, they are:

**gBlLogAction** - If logging is currently enabled or not, simple True or False. Default : True

**strBaseLogFileName** - The name of the log file, is string. Default : log.txt

**gBlAppendDateToLog** - If the date is added to the start of the log file, usuful for sorting on a computer. Default : True

**strLogFolderPath** - The path to the folder where the log files will be saved. If it is left blank it will be saved to the same folder as the workbook. Default : ""

**mbytDateFormat** - The Format that the date will take in the the file name (is that is on) and in the log. Has the following options:
* 0 = YYYY-MM-DD (Ex 2015-07-22) DEFAULT
* 1 = MM-DD-YYYY (Ex 07-22-2015)
* 2 = DD-MM-YYYY (Ex 22-07-2015)
* 3 = MM-DD (EX 07-22)
* 4 = YYYYMMDD (Ex 20150722)

**mblnUseMilitaryTime** - If military time should be used. If true time AM/PM will not be included. 3:00 PM would become 15:00 instead. Default : True

**mbytTimeFormat** - The format that the time will take in the log file. Has the following options:
* 0 = HH:MM:SS (EX 15:24:23 or 3:24:23 PM) DEFAULT
* 1 = HH:MM (EX 15:24 or 3:24 PM)
* 2 = HHMMSS (EX 152423 or 32423 PM)
* 3 = HHMM (EX 1524 or 324 PM)



