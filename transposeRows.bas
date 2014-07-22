Sub transposeByDirections()
  Dim cCell, perRow, perCol, cId
  perRow = ActiveCell.Value
  perCol = ActiveCell.Offset(0, 1).Value
  stRow = ActiveCell.Offset(0, 2).Value
  Set ori = ActiveSheet
  Set trans = ActiveWorkbook.Sheets.Add
  r = 1
  cId = ""
  For i = stRow To perRow
      If i = stRow Then
          k = 1
      Else
          k = 2
      End If
      For k = k To perCol
          trans.Cells(1, r).Value = ori.Cells(i, k).Value
          r = r + 1
      Next k
  Next i

End Sub
