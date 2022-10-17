Attribute VB_Name = "GCF_tool"
Option Compare Text

Sub GCF()                                        'tool for general conditioning formatting across multiple cells

    Dim MergeRange As Range, StartCells As Range, EndCells As Range, InputDirection As String, OverrideFormat As Integer, _
    InputAnswer As Integer, StartRow As Integer, StartCol As Integer, EndRow As Integer, EndCol As Integer, WS As Worksheet

    Set StartCells = Application.InputBox("Enter or click where you would like formatting to start", "Obtain Range Object", Type:=8) 'prompts user to select range in their active workbook
    StartRow = StartCells.Row
    StartCol = StartCells.Column
    'MsgBox ("The range is" & StartCells.Address) 'test to ensure that range was picked up
    

    Set EndCells = Application.InputBox("Enter or click where you would like formatting to stop", "Obtain Range Object", Type:=8) 'where to end pasting
    EndRow = EndCells.Row
    EndCol = EndCells.Column
    'MsgBox ("The range is" & EndCells.Address)

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
    'MsgBox ("The range is" & MergeRange.Address)

    If InputDirection = "Right" Then
     
        For i = StartCol + InputAnswer To EndCol Step InputAnswer
            If OverrideFormat = 6 Then WS.Cells(StartRow, i).FormatConditions.Delete 'for "Right" answer
            Set MergeRange = Union(MergeRange, WS.Cells(StartRow, i))
            StartCells.FormatConditions(1).ModifyAppliesToRange MergeRange
        
        Next
    
        For Each cell In MergeRange.FormatConditions(cell)
        MergeRange.FormatConditions(cell).ModifyAppliesToRange MergeRange + 1
        Next
        

    ElseIf InputDirection = "Left" Then
    
        For i = StartCol - InputAnswer To EndCol Step -InputAnswer
        
            If OverrideFormat = 6 Then WS.Cells(StartRow, i).FormatConditions.Delete 'for "Left" answer
            Set MergeRange = Union(MergeRange, WS.Cells(StartRow, i))
            StartCells.FormatConditions(1).ModifyAppliesToRange MergeRange
        Next
        
    ElseIf InputDirection = "Up" Then
    
        For i = StartRow - InputAnswer To EndRow Step -InputAnswer
        
            If OverrideFormat = 6 Then WS.Cells(i, StartCol).FormatConditions.Delete 'for "Up" answer
            Set MergeRange = Union(MergeRange, WS.Cells(i, StartCol))
            StartCells.FormatConditions(1).ModifyAppliesToRange MergeRange
        Next
        
    ElseIf InputDirection = "Down" Then
    
        For i = StartRow + InputAnswer To EndRow Step InputAnswer
        
            If OverrideFormat = 6 Then WS.Cells(i, StartCol).FormatConditions.Delete 'for "Down" answer
            Set MergeRange = Union(MergeRange, WS.Cells(i, StartCol))
            StartCells.FormatConditions(1).ModifyAppliesToRange MergeRange
        Next
            
            
    End If
    


    'Set UnionRange = Range("E1")
    'For i = 1 To 20 Step InputAnswer
    'Set UnionRange = Union(UnionRange, WS.Cells(i, 5))
    'Next
    'ActiveCell.FormatConditions(1).ModifyAppliesToRange UnionRange

End Sub


