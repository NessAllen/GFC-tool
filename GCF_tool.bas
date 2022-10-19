Attribute VB_Name = "GCF_tool"
Option Compare Text

Sub GCF()                                        'tool for general conditioning formatting across multiple cells

    Dim MergeRange As Range, StartCells As Range, EndCells As Range, InputDirection As String, OverrideFormat As Integer, _
    InputAnswer As Integer, StartRow As Integer, StartCol As Integer, EndRow As Integer, EndCol As Integer, WS As Worksheet, _
    oldRng As Range, cond As Variant
    
    Set StartCells = Application.InputBox("Enter or click where you would like formatting to start", "Obtain Range Object", Type:=8) 'prompts user to select range in their active workbook
    StartRow = StartCells.Row
    StartCol = StartCells.Column
    

    Set EndCells = Application.InputBox("Enter or click where you would like formatting to stop", "Obtain Range Object", Type:=8) 'where to end pasting
    EndRow = EndCells.Row
    EndCol = EndCells.Column

    InputAnswer = InputBox("How many cells would you like between each format paste?") 'increments

    InputDirection = InputBox("What direction would you like your formatting to paste? Your options are Left, Right, Down, Up.")
    If InputDirection = "Up" Or InputDirection = "Down" Or InputDirection = "Left" Or InputDirection = "Right" Then
    Else
        MsgBox ("Invalid option entered. Please retry.")
        Exit Sub

    End If

    OverrideFormat = MsgBox("Would you like to override original conditional formatting?", vbQuestion + _
                                                                                          vbYesNo + vbDefaultButton2, "General Conditional Format Tool") 'option to override any preexisting formatting. if (6) yes, formatting will be deleted prior to paste.

    Set WS = ActiveSheet
    Set MergeRange = StartCells

    Application.ScreenUpdating = False

    For Each SelectCell In StartCells

        StartRow = SelectCell.Row
        StartCol = SelectCell.Column


        If InputDirection = "Right" Then
      
            For cond = 1 To SelectCell.FormatConditions.Count
                For i = StartCol + InputAnswer To EndCol Step InputAnswer
                    If OverrideFormat = 6 And cond = 1 Then WS.Cells(StartRow, i).FormatConditions.Delete 'for "Right" answer
                    Set oldRng = SelectCell.FormatConditions(cond).AppliesTo
                    Set MergeRange = Union(MergeRange, oldRng, WS.Cells(StartRow, i))
                Next
                SelectCell.FormatConditions(cond).ModifyAppliesToRange MergeRange
            Next
                                               

        ElseIf InputDirection = "Left" Then
                                
            For cond = 1 To SelectCell.FormatConditions.Count
                For i = StartCol - InputAnswer To EndCol Step -InputAnswer
                    If OverrideFormat = 6 And cond = 1 Then WS.Cells(StartRow, i).FormatConditions.Delete 'for "Left" answer
                    Set oldRng = SelectCell.FormatConditions(cond).AppliesTo
                    Set MergeRange = Union(MergeRange, oldRng, WS.Cells(StartRow, i))
                Next
                SelectCell.FormatConditions(cond).ModifyAppliesToRange MergeRange
            Next
        
        ElseIf InputDirection = "Up" Then
                                    
            For cond = 1 To SelectCell.FormatConditions.Count
                For i = StartRow - InputAnswer To EndRow Step -InputAnswer
                    If OverrideFormat = 6 And cond = 1 Then WS.Cells(i, StartCol).FormatConditions.Delete 'for "Up" answer
                    Set oldRng = SelectCell.FormatConditions(cond).AppliesTo
                    Set MergeRange = Union(MergeRange, oldRng, WS.Cells(i, StartCol))
                Next
                SelectCell.FormatConditions(cond).ModifyAppliesToRange MergeRange
            Next
        
        ElseIf InputDirection = "Down" Then
    
            For cond = 1 To SelectCell.FormatConditions.Count
                For i = StartRow + InputAnswer To EndRow Step InputAnswer
                    If OverrideFormat = 6 And cond = 1 Then WS.Cells(i, StartCol).FormatConditions.Delete 'for "Down" answer
                    Set oldRng = SelectCell.FormatConditions(cond).AppliesTo
                    Set MergeRange = Union(MergeRange, oldRng, WS.Cells(i, StartCol))
                Next
                SelectCell.FormatConditions(cond).ModifyAppliesToRange MergeRange
            Next
            
        End If
    
    Next

    Application.ScreenUpdating = True
    MsgBox ("Conditional formatting update complete.")

End Sub
