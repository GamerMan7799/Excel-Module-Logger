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

Some time later I will likely make this easier to customize.


## Customization

There are currently 4 constant that you can change to cause the log to be saved differently, they are:

**gBlLogAction** - If logging is currently enabled or not, simple True or False. Default : True

**strBaseLogFileName** - The name of the log file, is string. Default : log.txt

**gBlAppendDateToLog** - If the date is added to the start of the log file, usuful for sorting on a computer. Default : True

**strLogFolderPath** - The path to the folder where the log files will be saved. If it is left blank it will be saved to the same folder as the workbook. Default : ""



