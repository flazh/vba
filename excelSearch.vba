'Immediate Window in excel vba
'To enable Debug.Print output press ctrl+g

'Sub to remove part of formula from activesheet
Sub chceckFormulas()
  For Each cl In ActiveSheet.UsedRange
    If cl.HasFormula() = True Then
      cl.Formula = Replace(cl.Formula, "stringToBeReplaced", "replaceWithThis")
    End if
  Next cl
End Sub


'Sub to go thru all worksheets and cells within
Sub forEachWs()
  For Each ws In ActiveWorkbook.Worksheets
    For Each cl In ws.UsedRange
      If cl.Formula Then
        Debug.Print cl.Address & "---" & cl.Formula & "---" & cl.Value
      End If
    Next cl
  Next ws
End Sub

'Sub to unhide hidden names in names manager in all opened files
Sub unhideAllNamesAllOpenedWorkbooks()
  For Each wb In Application.Workbooks
    For Each n In wb.Names
      n.Visible = True
    Next n
  Next wb
End Sub

'Sub with common operations
Sub commonOperations()

  'get last cell of filled cells
  Range("A1").End(xlUp).Offset(1, 0).Value = myvalue

  
  
End Sub.