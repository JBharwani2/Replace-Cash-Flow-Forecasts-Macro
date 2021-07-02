Sub ReplaceForecast()
    '
    ' ReplaceForecast Macro
    ' Created by Jeremy Bharwani on 7/1/21
    ' (questions- email jcb926@gmail.com)
    '
    ' This macro takes a selection of worksheets and a user input date to update the latest month from a forecast to an actual value. The date is compared
    ' with each row to find the one which was specified as the input. Then certain columns within that row are updated to a new formula that is linked to the
    ' corresponding month's CHS file. The new data is updated to the color red and left selected to be reviewed.
    '
    '

    'VARIABLES -------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim fileDate As String
    Dim year As String
    Dim month As String
    Dim checkDate As String
    Dim checkMonth As String
    Dim checkYear As String
    Dim row As Integer
    Dim count As Integer
    row = 10
    count = 0

    'USER INPUT -------------------------------------------------------------------------------------------------------------------------------------------------------
    'Gets user input for the month and year of this batch of files
    fileDate = InputBox("CONFIRM ALL SHEETS ARE SELECTED, then input the month and year in this format: 0521")
    month = Left(fileDate, 2)
    year = Right(fileDate, 2)

    'MAIN PROCESS -----------------------------------------------------------------------------------------------------------------------------------------------------
    Application.ScreenUpdating = False 'aviods the visible opening and closing of the new workbooks

    'Iterates through each worksheet that was initially selected
    For Each ws In ActiveWindow.SelectedSheets

        'Find row of next forecast
        Do
            row = row + 1
            checkMonth = Left(ws.Cells(row, 1), 2)
            checkYear = Right(ws.Cells(row, 1), 2)
            If Right(checkMonth, 1) = "/" Then
                checkMonth = "0" & Left(checkMonth, 1)
            End If

            checkDate = checkMonth & checkYear
        Loop Until checkDate = fileDate

        'Update rows
        ws.Range("C" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$43"
        ws.Range("D" & row).Clear()
        ws.Range("E" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$30 + 'S:\acct\JLS\20" &
            year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$35"
        ws.Range("F" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$33"
        ws.Range("G" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$31"
        ws.Range("H" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$32"
        ws.Range("K" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$167"
        ws.Range("L" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$169 + 'S:\acct\JLS\20" &
            year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$170"
        ws.Range("M" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$177"
        ws.Range("P" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$222"
        ws.Range("V" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$175 + 'S:\acct\JLS\20" &
            year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!$K$177 + 'S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" &
                year & "chs.xls]" & ws.Name & "'!$K$179"
        count = count + 1
    Next ws

    'Changes updated cells to red
    Range("C75,D75,E75,F75,G75,H75,K75,L75,M75,P75,V75").Select
    Selection.Font.Color = RGB(255, 0, 0)

    'Completion message
    Application.ScreenUpdating = True
    MsgBox Str(count) & " sheets have been successfuly updated."

End Sub

