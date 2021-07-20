Sub ReplaceCashFlowForecasts()
    '
    ' ReplaceCashFlowForecasts Macro
    ' Created by Jeremy Bharwani on 7/7/21
    ' Updated by Jeremy Bharwani on 7/19/21
    ' (questions- email jcb926@gmail.com)
    '
    ' This macro takes a selection of worksheets and a user input date to update the latest month from a forecast to an actual value. The date is compared
    ' with each row to find the one which was specified as the input. Then certain columns within that row are updated to a new formula that is linked to the
    ' corresponding month's CHS file. The new data is updated to the color red and left selected to be reviewed.
    '
    ' Time: < 2 seconds per worksheet
    ' References: none
    '

    'VARIABLES -------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim ws As Worksheet
    Dim fileDate As String
    Dim year As String
    Dim month As String
    Dim checkDate As String
    Dim checkMonth As String
    Dim checkYear As String
    Dim previousMonth As String
    Dim row As Integer
    Dim count As Integer
    count = 0

    'USER INPUT -------------------------------------------------------------------------------------------------------------------------------------------------------    
    'Gets user input for the month and year of this batch of files
    fileDate = Left(Right(ThisWorkbook.Name, 9), 5)
    month = Left(fileDate, 2)
    year = Right(fileDate, 2)

    'Asks user if they are sure that they want to run the macro with the selected sheets
    CarryOn = MsgBox("You have selected " & ActiveWindow.SelectedSheets.count & " sheets for the month of " & month & "-" & year &
        ". Do you want to proceed in updating this month's forecasts to actual data?", vbYesNo, "Macro Run Confirmation")

    If CarryOn = vbNo Then
        End
    End If

    'MAIN PROCESS -----------------------------------------------------------------------------------------------------------------------------------------------------
    'Iterates through each worksheet that was initially selected
    Application.ScreenUpdating = False 'aviods the visible opening and closing of the new workbooks
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
        Loop Until checkDate = month & year

        'Copies down the formulas from the row above
        ws.Range("C" & row).Formula = ws.Range("C" & (row - 1)).Formula
        ws.Range("D" & row).Clear()
        ws.Range("E" & row).Formula = ws.Range("E" & (row - 1)).Formula
        ws.Range("F" & row).Formula = ws.Range("F" & (row - 1)).Formula
        ws.Range("G" & row).Formula = ws.Range("G" & (row - 1)).Formula
        ws.Range("H" & row).Formula = ws.Range("H" & (row - 1)).Formula
        ws.Range("K" & row).Formula = ws.Range("K" & (row - 1)).Formula
        ws.Range("L" & row).Formula = ws.Range("L" & (row - 1)).Formula
        ws.Range("M" & row).Formula = ws.Range("M" & (row - 1)).Formula
        ws.Range("P" & row).Formula = ws.Range("P" & (row - 1)).Formula
        ws.Range("V" & row).Formula = ws.Range("V" & (row - 1)).Formula

        'Updates formula to have a new date
        previousMonth = Mid(ws.Range("C" & (row - 1)).Formula, InStr(ws.Range("C" & (row - 1)).Formula, "[") + 1, 2)
        ws.Range("C" & row & ":" & "V" & row).Replace What:=previousMonth & "-" & year, Replacement:=month & "-" & year, LookAt:=xlPart,
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False,
            ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

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


