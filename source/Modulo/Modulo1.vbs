'***********************************************************************
'Macros by Eric Bentzen, 22th August 2015.
'How to make a calendar for picking dates with VBA only - no ActiveX
'and lack of compatibility between versions.
'You can format the selected dates in the two userform
'procedures: "FillFirstDay" and "FillSecondDay".

'Bug fix December 2015: Dates were not formatted correctly,
'if the date format in the system settings was MM/DD/Year
'This is now fixed by the userform's ReturnDate function.

'You can find more VBA and macro stuff at:
'http://sitestory.dk/excel_vba/vba-start-page.htm
'***********************************************************************
Option Explicit
Public colLabelEvent As Collection 'Collection of labels for event handling
Public colLabels As Collection     'Collection of the date labels
Public bSecondDate As Boolean      'True if finding second date
Public sActiveDay As String        'Last day selected
Public lDays As Long               'Number of days in month
Public lFirstDay As Long           'Day selected, e.g. 19th
Public lStartPos As Long
Public lSelMonth As Long           'The selected month
Public lSelYear As Long            'The selected year
Public lSelMonth1 As Long          'Used to check if same date is selected twice
Public lSelYear1 As Long           'Used to check if same date is selected twice

Public cnn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public strSQL As String
Public Const pword = "KyouruKenji"

Public Sub OpenDB()
    If cnn.State = adStateOpen Then cnn.Close
    cnn.ConnectionString = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & _
    ActiveWorkbook.Path & Application.PathSeparator & Application.ActiveWorkbook.Name & ";ReadOnly=0;"
    cnn.Open
End Sub

Public Sub closeRS()
    If rs.State = adStateOpen Then rs.Close
    rs.CursorLocation = adUseClient
End Sub



