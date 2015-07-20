Attribute VB_Name = "Logger"

'---------------------------------------------------------------------------------------
' Module    : Logger
' Author    : Patrick Rye
' Date      : 7/20/2015
' Purpose   : Holds a function which will hold whatever is entered for it to a file.
'---------------------------------------------------------------------------------------
Option Explicit
'   ----------------------------------------------------------------------------
Private Const gBlLogAction As Boolean = True 'Whether all actions should be logged or not
Private Const strBaseLogFileName As String = "log.txt" 'The base name of the file
Private Const gBlAppendDateToLog As Boolean = True 'Whether the date should be added to the front of the log file or not.
Private Const strLogFolderPath As String = "" 'The file path for the folder the log is placed into, leave blank to save to same folder as excel file.
'   ----------------------------------------------------------------------------
'Use the line to call this sub, remove ' before and replace *Log* with what you want logged
'Call WriteToLog(*Log*)
'   ----------------------------------------------------------------------------
Public Sub WriteToLog(strLine As String)
Dim strLogFilePath As String
Dim strUserName As String
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim oFile As Object
If gBlLogAction <> True Then GoTo EndLogger 'If logging is not set to true then go to the end of the function
Dim strDate As String
Dim strTime As String
strDate = DateMaker() 'Make the Date
strTime = TimeMaker() 'Make the time
strUserName = (Environ$("Username")) 'Get the user name from the environmental variables
'Make the File Path of the log file.
If strLogFolderPath <> "" Then
    strLogFilePath = strLogFolderPath
Else
    strLogFilePath = ThisWorkbook.Path
End If
If Right(strLogFilePath, 1) <> "\" Then strLogFilePath = strLogFilePath & "\"
If gBlAppendDateToLog Then strLogFilePath = strLogFilePath & strDate & " "
strLogFilePath = strLogFilePath & strBaseLogFileName

'Checks if the log file already exits or not.
If Dir(strLogFilePath) <> "" Then
    'File Does exist, open it for appending
    Set oFile = fso.OpenTextFile(strLogFilePath, 8, True)
Else
    'File does not exist, create it with headers
    Set oFile = fso.CreateTextFile(strLogFilePath)
    oFile.WriteLine "Date           Time            User       | Action"
    oFile.WriteLine "-------------------------------------------------------------------------------------------------"
End If
oFile.WriteLine strDate & "     " & strTime & "     " & strUserName & "   | " & strLine
oFile.Close
EndLogger:
Set fso = Nothing
Set oFile = Nothing
End Sub
'   ----------------------------------------------------------------------------
Private Function DateMaker() As String
'Writes the date in the following Format YYYY-MM-DD
Dim strDate As String
strDate = Year(Now) & "-"
If Month(Now) < 10 Then
    strDate = strDate & "0" & Month(Now) & "-"
Else
    strDate = strDate & Month(Now) & "-"
End If
If Day(Now) < 10 Then
    strDate = strDate & "0" & Day(Now)
Else
    strDate = strDate & Day(Now)
End If
DateMaker = strDate
End Function
'   ----------------------------------------------------------------------------
Private Function TimeMaker() As String
'Writes the time in the following format (military time) HH:MM:SS 
Dim strTime As String
If Hour(Now) < 10 Then
    strTime = "0" & Hour(Now) & ":"
Else
    strTime = Hour(Now) & ":"
End If
If Minute(Now) < 10 Then
    strTime = strTime & "0" & Minute(Now) & ":"
Else
    strTime = strTime & Minute(Now) & ":"
End If
If Second(Now) < 10 Then
    strTime = strTime & "0" & Second(Now)
Else
    strTime = strTime & Second(Now)
End If
TimeMaker = strTime
End Function

