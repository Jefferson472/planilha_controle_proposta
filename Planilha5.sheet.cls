Private Sub Worksheet_Change(ByVal Target As Range)
Dim LR As Long
Dim procv As Variant
Dim tabela As ListObject
Set tabela = Planilha3.ListObjects(1)
  If Target.Count > 1 Then Exit Sub
  If Target.Column <> 3 Then Exit Sub
    Target.Offset(, -1) = ""
    LR = Cells(Rows.Count, 2).End(xlUp).Row
    procv = Application.VLookup(Target.Value, Planilha3.Range(tabela), 4, False)
    Target.Offset(, -1).FormulaArray = _
        "=Max(if(Left(B2:B" & LR & ",7)=Upper(Left(""" & procv & """, 7)),Right(B2:B" & LR & ",3)+0,0))"

    Target.Offset(, -1).Value = _
        UCase(procv) & "-" & Format(Target.Offset(, -1).Value + 1, "000")
    'Target.Offset(, -2).Value = Target.Offset(-1, -2).Value + 1
End Sub


