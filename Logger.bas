Attribute VB_Name = "Logger"

'---------------------------------------------------------------------------------------
' Module    : Logger
' Author    : Patrick Rye
' Date      : 7/22/2015
' Version	: 1.1.0
' Purpose   : Holds a procedure which will write stuff to a log.
'---------------------------------------------------------------------------------------
Option Explicit
'   ----------------------------------------------------------------------------
Private Const gBlLogAction As Boolean = True 'Whether all actions should be logged or not
Private Const strBaseLogFileName As String = "log.txt" 'The base name of the file
Private Const gBlAppendDateToLog As Boolean = True 'Whether the date should be added to the front of the log file or not.
Private Const mbytDateFormat as Byte = 0 'What format the date is is
'0 = YYYY-MM-DD (default)
'1 = MM-DD-YYYY
'2 = DD-MM-YYYY
'3 = MM-DD
'4 = YYYYMMDD
Private Const mbytTimeFormat as Byte = 0 'The format of the time
'0 = HH:MM:SS (default)
'1 = HH:MM
'2 = HHMMSS
'3 = HHMM
Private const mblnUseMilitaryTime as Boolean = True 'If military time should be used (15:00 instead of 3:00 PM)
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
'Writes the date in the selected format
Dim strDate As String
Dim strMonth as String
Dim strDay as String
'Add a 0 to the front of the month and day if its less than 10.
'Makes it look nicer.
If Month(Now) < 10 Then
	strMonth = "0" & Month(Now)
Else
	strMonth = Month(Now)
End If
If Day(Now) < 10 Then
    strDay = "0" & Day(Now)
Else
    strDay = Day(Now)
End If
Select Case mbytDateFormat
	Case 0
		strDate = Year(Now) & "-" & strMonth & "-" & strDay
		'Ex 2015-07-22
	Case 1
		strDate = strMonth & "-" & strDay & "-" & Year(Now)
		'Ex 07-22-2015
	Case 2
		strDate = strDay & "-" & strMonth & "-" & Year(Now)
		'Ex 22-07-2015
	case 3
		strDate = strMonth & "-" & strDay
		'Ex 07-22
	case 4
		strDate = Year(Now) & strMonth & strDay
		'Ex 20150722
	case else
		strDate = Year(Now) & "-" & strMonth & "-" & strDay
end select
DateMaker = strDate
End Function
'   ----------------------------------------------------------------------------
Private Function TimeMaker() As String
'Writes the time in the selected format
Dim strTime As String
Dim strHour as String
Dim strMinute as String
Dim strSecond as String
Dim strAMPM as String

if mblnUseMilitaryTime then
	if Hour(Now) < 10 then
		strHour = "0" & Hour(Now)
	else
		strHour = Hour(now)
	end if
else
	if Hour(now) > 12 then
		strAMPM = "PM"
		strHour = (Hour(Now) - 12)
	elseIf Hour(Now) = 12 then
		strAMPM = "PM"
		strHour = "12"
	else
		strAMPM = "AM"
		strHour = Hour(now)
	end if
end if	

If Minute(Now) < 10 Then
    strMinute = "0" & Minute(Now)
Else
    strMinute = Minute(Now)
End If
If Second(Now) < 10 Then
    strSecond = "0" & Second(now)
Else
    strSecond = Second(now)
End If
Select case mbytTimeFormat
	Case 0
		strTime = strHour & ":" & strMinute & ":" & strSecond 
		'EX 15:24:23 or 3:24:23 PM
	case 1
		strTime = strHour & ":" & strMinute
		'EX 15:24 or 3:24 PM
	case 2
		strTime = strHour & strMinute & strSecond
		'EX 152423 or 32423 PM
	case 3
		strTime = strHour & strMinute
		'EX 1524 or 324 PM
	Case Else
		strTime = strHour & ":" & strMinute & ":" & strSecond
end Select
If Not(mblnUseMilitaryTime) then strTime = strTime & " " & strAMPM
TimeMaker = strTime
End Function

