Sub GoalSeekEndBal()
'
' GoalSeekEndBal Macro
' Created by Jeremy Bharwani on 7/19/21
' (questions- email jcb926@gmail.com)
'
' This macro uses What-If Analysis to goal seek an ending balance of 0. Your cursor must be on the bottom-most
' cell of the end balance column to set that cell to 0.
'

Set EndingBal = ActiveCell
    Set ChangeValue = Cells(EndingBal.row + 3, EndingBal.Column - 5)
    
    Range(EndingBal.Address).GoalSeek Goal:=0, ChangingCell:=Range(ChangeValue.Address)
End Sub
