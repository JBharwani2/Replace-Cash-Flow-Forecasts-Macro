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
    ' Time: 1.2 seconds per worksheet
    ' References: "Microsoft VBScript Regular Expressions 5.5"
    '

    'VARIABLES -------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim checkSheet As Worksheet
    Dim fileDate As String
    Dim year As String
    Dim month As String
    Dim checkDate As String
    Dim checkMonth As String
    Dim checkYear As String
    Dim row As Integer
    Dim count As Integer
    Dim bonusRow As Integer
    Dim regexOne As Object
    Set regexOne = New RegExp
   regexOne.Pattern = "\d{4}"
    count = 0

    'USER INPUT -------------------------------------------------------------------------------------------------------------------------------------------------------
    'Gets user input for the month and year of this batch of files
    fileDate = InputBox("CONFIRM ALL SHEETS ARE SELECTED, then input the month and year in this format: 0521")
    month = Left(fileDate, 2)
    year = Right(fileDate, 2)

    If regexOne.Test(fileDate) And Len(fileDate) = 4 Then
        Application.ScreenUpdating = False 'aviods the visible opening and closing of the new workbooks
    Else
        MsgBox "Invalid date entered"
        End
    End If

    'MAIN PROCESS -----------------------------------------------------------------------------------------------------------------------------------------------------


    'Iterates through each worksheet that was initially selected
    For Each ws In ActiveWindow.SelectedSheets

        'Find row of next forecast
        row = 10
        Do
            row = row + 1
            checkMonth = Left(ws.Cells(row, 1), 2)
            checkYear = Right(ws.Cells(row, 1), 2)
            If Right(checkMonth, 1) = "/" Then
                checkMonth = "0" & Left(checkMonth, 1)
            End If

            checkDate = checkMonth & checkYear
        Loop Until checkDate = fileDate

        'Update rows with new formulas
        ws.Range("C" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("C" & (row - 1)).Formula, 5)

        ws.Range("D" & row).Clear()

        ws.Range("E" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" _
            & Mid(ws.Range("E" & (row - 1)).Formula, InStr(ws.Range("E" & (row - 1)).Formula, "+") - 5, 5) _
                & " + 'S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("E" & (row - 1)).Formula, 5)

        ws.Range("F" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("F" & (row - 1)).Formula, 5)
        ws.Range("G" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("G" & (row - 1)).Formula, 5)
        ws.Range("H" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("H" & (row - 1)).Formula, 5)
        ws.Range("K" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("K" & (row - 1)).Formula, 6)

        ws.Range("L" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" _
            & Mid(ws.Range("L" & (row - 1)).Formula, InStr(ws.Range("L" & (row - 1)).Formula, "+") - 6, 6) _
                & " + 'S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("L" & (row - 1)).Formula, 6)

        ws.Range("M" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("M" & (row - 1)).Formula, 6)
        ws.Range("P" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("P" & (row - 1)).Formula, 6)

        ws.Range("V" & row).Formula = "='S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" _
            & Mid(ws.Range("V" & (row - 1)).Formula, InStr(ws.Range("V" & (row - 1)).Formula, "+") - 6, 6) & " + 'S:\acct\JLS\20" _
                & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" _
                    & Mid(Mid(ws.Range("V" & (row - 1)).Formula, InStr(ws.Range("V" & (row - 1)).Formula, "+") + 1), InStr(Mid(ws.Range("V" & (row - 1)).Formula, InStr(ws.Range("V" & (row - 1)).Formula, "+") + 1), "+") - 6, 6) _
                        & " + 'S:\acct\JLS\20" & year & "\" & year & "-CHS\[" & month & "-" & year & "chs.xls]" & ws.Name & "'!" & Right(ws.Range("V" & (row - 1)).Formula, 6)


        'Changes updated cells to red
        ws.Range("C" & row & "," & "E" & row & "," & "F" & row & "," & "G" & row & "," & "H" & row & "," & "K" & row & "," & "L" & row & "," & "M" & row & "," & "P" & row & "," & "V" & row) _
            .Font.Color = RGB(255, 0, 0)

        'Count number of updated sheets
        count = count + 1
    Next ws

    'Completion message
    Application.ScreenUpdating = True
    MsgBox Str(count) & " sheets have been successfuly updated."

End Sub
