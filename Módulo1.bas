Attribute VB_Name = "Módulo1"
Option Explicit
Public Arr As Variant
Public L As Object
Public TotalColuna As Double

Sub Inserir()

Dim tabela As ListObject
Dim n As Integer
Set tabela = Planilha1.ListObjects(1)

tabela.ListRows.Add
n = tabela.Range.Rows.Count

'COD_PROPOSTA
Dim LR As Long
Dim cod_proposta As String
LR = Planilha5.Cells(Rows.Count, 2).End(xlUp).Row + 1
Planilha5.Range("C" & LR).Value = UserForm1.cboPropriedades
cod_proposta = Planilha5.Range("B" & LR)

Dim Data As Date
Data = Date

'DATABASE
tabela.Range(n, 3).Value = UserForm1.cboPropriedades
tabela.Range(n, 1).Value = cod_proposta
tabela.Range(n, 4).Value = UserForm1.cboCategoria
tabela.Range(n, 5).Value = cod_proposta & " - " & UserForm1.txtescopo
tabela.Range(n, 7).Value = UserForm1.cboFornecedor
tabela.Range(n, 2).Value = Data

'LAYOUT
Planilha10.Range("I4").Value = cod_proposta
Planilha10.Range("I5").Value = UserForm1.cboPropriedades
Planilha10.Range("H8").Value = UserForm1.cboCategoria
Planilha10.Range("B10").Value = UserForm1.txtescopo
Planilha10.Range("O6").Value = Data

End Sub

Sub InserirItem()

Dim tabela As ListObject
Dim n As Integer
Dim valor_bdi As Double
Set tabela = Planilha2.ListObjects(1)

UserForm1.ListBox1.RowSource = ""

tabela.ListRows.Add
n = tabela.Range.Rows.Count

If UserForm1.cboCategoria = "Hora Extra" Then
    valor_bdi = UserForm1.txtValor
Else
    valor_bdi = UserForm1.txtValor * ((UserForm1.txtBdi.Value / 100) + 1)
End If

'BASEITENS
tabela.Range(n, 2).Value = UserForm1.txtitem
tabela.Range(n, 3).Value = UserForm1.txtqnt
tabela.Range(n, 4).Value = UserForm1.cboUnid
tabela.Range(n, 5).Value = UserForm1.txtValor
tabela.Range(n, 7).Value = valor_bdi

'BASEITENS_GERAL
Set tabela = Planilha9.ListObjects(1)
tabela.ListRows.Add
n = tabela.Range.Rows.Count
tabela.Range(n, 1).Value = Planilha10.Range("I4").Value
tabela.Range(n, 2).Value = UserForm1.txtitem
tabela.Range(n, 3).Value = UserForm1.txtqnt
tabela.Range(n, 4).Value = UserForm1.cboUnid
tabela.Range(n, 5).Value = UserForm1.txtValor
tabela.Range(n, 7).Value = valor_bdi

Call Atulizar_ListBox
Call Valor_Total

End Sub
Sub Inserir_Layout()

Dim n As Integer, linha_atual As Integer

Planilha10.Activate
Range("B11").Select
    
Do Until IsEmpty(ActiveCell)
    ActiveCell.Offset(1, 0).Select
    n = n + 1
Loop

linha_atual = n + 11

If n > 1 Then
    Planilha10.Range("A12").EntireRow.Copy
    Planilha10.Rows(linha_atual).EntireRow.Insert
End If

ActiveCell.Value = n

'LAYOUT
Planilha10.Cells(linha_atual, 4).Value = UserForm1.txtitem
Planilha10.Cells(linha_atual, 22).Value = UserForm1.txtqnt
Planilha10.Cells(linha_atual, 23).Value = UserForm1.cboUnid
Planilha10.Cells(linha_atual, 24).Value = UserForm1.txtValor * ((UserForm1.txtBdi.Value / 100) + 1)


End Sub

Sub Atulizar_ListBox()

    On Error Resume Next
    Dim tabela As ListObject
    'Att ListBox1
    Set tabela = Planilha2.ListObjects(1)
    UserForm1.ListBox1.RowSource = tabela.DataBodyRange.Address(, , , True)
    
    'Att ListBox2
    'DESTAIVADA
    'Set tabela = Planilha1.ListObjects(1)
    'UserForm1.ListBox2.RowSource = tabela.DataBodyRange.Address(, , , True)
    
End Sub

Sub Atualizar_ComboBox()

'-------------------------------------------------------------------------------
'FORM1

'Att ComboBox Proprieades
Dim ult_linha As Integer, Linha As Integer
ult_linha = Planilha3.Range("A" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cboPropriedades.AddItem Planilha3.Range("A1").Cells(Linha, 1)
Next Linha

'Att ComboBox Categorias
ult_linha = Planilha3.Range("J" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cboCategoria.AddItem Planilha3.Range("J1").Cells(Linha, 1)
Next Linha

'Att ComboBox Unidades
ult_linha = Planilha3.Range("L" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cboUnid.AddItem Planilha3.Range("L1").Cells(Linha, 1)
Next Linha

'Att ComboBox Fornecedores
ult_linha = Planilha3.Range("N" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cboFornecedor.AddItem Planilha3.Range("N1").Cells(Linha, 1)
Next Linha

'-------------------------------------------------------------------------------
'FORM2

'Att ComboBox código da proposta
ult_linha = Planilha1.Range("A" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cbo_cod.AddItem Planilha1.Range("A1").Cells(Linha, 1)
Next Linha

'Att ComboBox Proprieades
ult_linha = Planilha3.Range("A" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cbo_propriedade.AddItem Planilha3.Range("A1").Cells(Linha, 1)
Next Linha

'Att ComboBox Categorias
ult_linha = Planilha3.Range("J" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cbo_categoria.AddItem Planilha3.Range("J1").Cells(Linha, 1)
Next Linha

'Att ComboBox Fornecedores
ult_linha = Planilha3.Range("N" & Rows.Count).End(xlUp).Row
For Linha = ult_linha To 2 Step -1
    UserForm1.cbo_fornecedor.AddItem Planilha3.Range("N1").Cells(Linha, 1)
Next Linha

End Sub

Sub Valor_Total()

Dim tabela As ListObject
Dim n As Integer
Set tabela = Planilha1.ListObjects(1)
n = tabela.Range.Rows.Count

tabela.Range(n, 6).Value = tabela.Range(n, 6).Value + ((UserForm1.txtValor * ((UserForm1.txtBdi.Value / 100) + 1)) * UserForm1.txtqnt)

tabela.Range(n, 8).Value = tabela.Range(n, 8).Value + (UserForm1.txtValor.Value * UserForm1.txtqnt)
UserForm1.txtValor_Total = tabela.Range(n, 8).Value

End Sub

Sub Salvar_Como()

Dim escreva_nome As String
If UserForm1.name_proposta.Value = "" Then
    escreva_nome = "Proposta - " & Planilha10.Range("I4").Value & " - " & UserForm1.cboPropriedades & " - " & UserForm1.cboCategoria
Else:
    escreva_nome = "Proposta - " & Planilha10.Range("I4").Value & " - " & UserForm1.cboPropriedades & " - " & UserForm1.cboCategoria & " - " & UserForm1.name_proposta
End If

Planilha10.Activate

'Impede que o Excel atualize a tela
Application.ScreenUpdating = False
'Impede que o Excel exiba alertas
Application.DisplayAlerts = False

'Seta uma variável para se referir a nova pasta de trabalho
Dim NovoWB As Workbook
'Cria esta nova aba
Set NovoWB = Workbooks.Add(xlWBATWorksheet)
With NovoWB
'Copia a aba atual para o novo arquivo, como a segunda aba
ThisWorkbook.ActiveSheet.Copy After:=.Worksheets(.Worksheets.Count)
'Deleta a primeira aba do arquivo criado (Aba em branco)
.Worksheets(1).Delete
'Salva o novo arquivo para a mesma pasta do arquivo atual
'Troque "Novo Arquivo" para um outro nome qualquer que preferir
.SaveAs ThisWorkbook.Path & "\" & escreva_nome & ".xlsx"
.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    ActiveWorkbook.Path & "\" & escreva_nome, IgnorePrintAreas:=False
'Fecha o novo arquivo
.Close False
End With

'Permite que o Excel volte a atualizar a tela
Application.ScreenUpdating = False
'Permite que o Excel volte a exibir alertas
Application.DisplayAlerts = False

Planilha10.Delete

End Sub

Sub Limpa_Baseitens()
    Dim tabela As ListObject
    Set tabela = Planilha2.ListObjects(1)
    UserForm1.ListBox1.RowSource = ""
    tabela.DataBodyRange.Rows.Delete
    Call Atulizar_ListBox
'    WorkPlanilha10.Delete
End Sub

Sub Limpa_Form()
    UserForm1.cboPropriedades = ""
    UserForm1.cboCategoria = ""
    UserForm1.cboFornecedor = ""
    UserForm1.txtescopo = ""
End Sub

Sub Form_Clique()
    UserForm1.Show vbModeless
End Sub

Sub Limpa_Form2()
    UserForm1.cbo_cod.Value = ""
    UserForm1.cbo_propriedade.Value = ""
    UserForm1.cbo_categoria.Value = ""
    UserForm1.cbo_fornecedor.Value = ""
    UserForm1.txt_valor_forn.Value = ""
    UserForm1.txt_valor_cliente.Value = ""
    UserForm1.txt_aprov.Value = ""
    UserForm1.txt_data.Value = ""
    UserForm1.txt_po_cliente.Value = ""
    UserForm1.txt_PO_Empresa.Value = ""
End Sub

Sub Aprovar()
    Dim tabela As ListObject
    Dim n As String, L As Integer, po_cliente As String, valor_forn As Double, valor_cliente As Double, PO_Empresa As String
    Set tabela = Planilha1.ListObjects(1)
    
    po_cliente = UserForm1.txt_po_cliente.Value
    PO_Empresa = UserForm1.txt_PO_Empresa.Value
    valor_forn = UserForm1.txt_valor_forn.Value
    valor_cliente = UserForm1.txt_valor_cliente.Value
    
    n = UserForm1.ListBox2.Value
    L = tabela.Range.Columns().Find(n, , , xlWhole).Row
    tabela.Range(L, 10).Value = "APROVADO"
    tabela.Range(L, 11).Value = Date
    tabela.Range(L, 13).Value = po_cliente
    tabela.Range(L, 12).Value = PO_Empresa
    tabela.Range(L, 8).Value = valor_forn
    tabela.Range(L, 6).Value = valor_cliente
        
End Sub

Sub Lança_E1_PLAN()
    Dim L As Integer
    Dim tabela As ListObject

    'Acha a última linha preenchida
    Dim LR As Long
    LR = Planilha7.Cells(Rows.Count, 2).End(xlUp).Row + 1

    Planilha7.Range("B" & LR).Value = 1
    Planilha7.Range("F" & LR).Value = 1
    Planilha7.Range("J" & LR).Value = UserForm1.cbo_cod.Value
    Planilha7.Range("L" & LR).Value = 57800
    Planilha7.Range("M" & LR).Value = "'005"
    
    'Lançamentos relacionados ao centro de custo
    Set tabela = Planilha3.ListObjects(1)
    L = tabela.Range.Columns(1).Find(UserForm1.cbo_propriedade.Value, , , xlWhole).Row
    Planilha7.Range("C" & LR).Value = Planilha3.Range("B" & L).Value
    Planilha7.Range("K" & LR).Value = Planilha3.Range("B" & L).Value
    Planilha7.Range("I" & LR).Value = UserForm1.cbo_propriedade.Value

    'Lançamento relacionado ao fornecedor
    Set tabela = Planilha3.ListObjects(4)
    L = tabela.Range.Columns(1).Find(UserForm1.cbo_fornecedor.Value, , , xlWhole).Row
    Planilha7.Range("D" & LR).Value = Planilha3.Range("O" & L).Value
    Planilha7.Range("G" & LR).Value = UserForm1.txt_valor_forn.Value
    If Planilha3.Range("Q" & L).Value = "Material" Then
        Planilha7.Range("E" & LR).Value = "BRPREEMB002"
    Else
        Planilha7.Range("E" & LR).Value = "BRPREEMB001"
    End If
End Sub

'DESATIVADA
Sub E1()
    Dim LR As Long, lin_E1 As Variant
    LR = Planilha7.Cells(Rows.Count, 2).End(xlUp).Row

    lin_E1 = Planilha7.Range("B" & LR, "M" & LR)
End Sub

'código pronto para filtro
Sub Filtro()

On Error GoTo Erro

Dim QtdLinhas As Double
Dim Campo1, Campo2, Campo3, Campo4, Campo5, Campo6 As String
Dim Coluna1, Coluna2, Coluna3, Coluna4, Colunadata As Double

Dim Plan As String
Dim Range As String
Dim RangeColuna1 As String

Set L = UserForm1.ListBox2
Plan = Worksheets("DATABASE").name
Range = "A2:O"
RangeColuna1 = "A:A"
TotalColuna = 15

Campo1 = UserForm1.cbo_cod
Coluna1 = 1

Campo2 = UserForm1.cbo_propriedade
Coluna2 = 3

Campo3 = UserForm1.cbo_categoria
Coluna3 = 4

Campo4 = UserForm1.cbo_fornecedor
Coluna4 = 7

'Campo5 = TDInicio
'Campo6 = TDFim
'Colunadata = 2

On Error Resume Next
L.Clear

L.ColumnCount = TotalColuna
L.ColumnWidths = "75;60;120;75;150;50;50;100;100;150"




QtdLinhas = WorksheetFunction.CountA(Sheets(Plan).Range(RangeColuna1))
Arr = Sheets(Plan).Range(Range & QtdLinhas).CurrentRegion

If Campo5 <> "" And Campo6 <> "" Then

        Dim Inicio As Date
        Dim Fim As Date
        Inicio = Campo5
        Fim = Campo6

End If


If Campo1 = "" And Campo2 = "" And Campo3 = "" And Campo4 = "" And Campo5 = "" And Campo6 = "" Then
L.List = Arr
Arr = Nothing
Exit Sub
End If

    Dim QtdLinhaFiltro As Double
    QtdLinhaFiltro = Empty
          
            
   Dim arrayItems()
          
   Dim Linha As Double
      
   Dim Coluna As Double
   Dim valor_celula As String
   Dim Data As Date
   Dim i As Integer
   
   If Campo5 <> "" And Campo6 <> "" Then
   
   QtdLinhaFiltro = 1
   
   For i = LBound(Arr, 1) To UBound(Arr, 1)
           
        Data = Empty
        
        valor_celula = Arr(i, Coluna1)
        If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo1))) = VBA.UCase(Campo1) Then
          
            valor_celula = Arr(i, Coluna2)
            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo2))) = VBA.UCase(Campo2) Then
            
                     valor_celula = Arr(i, Coluna3)
                    If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo3))) = VBA.UCase(Campo3) Then
                    
                             valor_celula = Arr(i, Coluna4)
                            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo4))) = VBA.UCase(Campo4) Then
                              
                              On Error Resume Next
                                Data = Arr(i, Colunadata)
                                 If Data >= Inicio And Data <= Fim Then
                                 
                                  QtdLinhaFiltro = QtdLinhaFiltro + 1
                                     
                                 End If
                                 
                          End If
                
                  End If
             
           End If
          
        End If
        
          
     Next i
   
   
   
    ReDim arrayItems(1 To QtdLinhaFiltro, 1 To TotalColuna)
   
   Linha = 2
   
        For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        Data = Empty
        
        valor_celula = Arr(i, Coluna1)
        If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo1))) = VBA.UCase(Campo1) Then
          
            valor_celula = Arr(i, Coluna2)
            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo2))) = VBA.UCase(Campo2) Then
            
                     valor_celula = Arr(i, Coluna3)
                    If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo3))) = VBA.UCase(Campo3) Then
                    
                             valor_celula = Arr(i, Coluna4)
                            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo4))) = VBA.UCase(Campo4) Then
                              
                              On Error Resume Next
                                Data = Arr(i, Colunadata)
                                 If Data >= Inicio And Data <= Fim Then
                                   
                                    For Coluna = 1 To TotalColuna
                                         
                                         arrayItems(Linha, Coluna) = Arr(i, Coluna)
                                          
                                    Next Coluna
                                                                        
                                    Linha = Linha + 1
                                                                       
                                    
                                 End If
                                 
                          End If
                
                  End If
             
           End If
          
        End If
        
          
     Next i
   
   
  L.List = arrayItems()
  Call cabecalho
     
   
   Exit Sub
   
   End If
   
   
   QtdLinhaFiltro = 1
   
   For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        valor_celula = Arr(i, Coluna1)
        If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo1))) = VBA.UCase(Campo1) Then
          
            valor_celula = Arr(i, Coluna2)
            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo2))) = VBA.UCase(Campo2) Then
            
                     valor_celula = Arr(i, Coluna3)
                    If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo3))) = VBA.UCase(Campo3) Then
                    
                             valor_celula = Arr(i, Coluna4)
                            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo4))) = VBA.UCase(Campo4) Then
                                 
                                  QtdLinhaFiltro = QtdLinhaFiltro + 1
                                     
                          End If
                
                  End If
             
           End If
          
        End If
        
          
     Next i
   
      
    ReDim arrayItems(1 To QtdLinhaFiltro, 1 To TotalColuna)
   
   Linha = 2
    
   
     For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        valor_celula = Arr(i, Coluna1)
        If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo1))) = VBA.UCase(Campo1) Then
          
            valor_celula = Arr(i, Coluna2)
            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo2))) = VBA.UCase(Campo2) Then
            
                     valor_celula = Arr(i, Coluna3)
                    If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo3))) = VBA.UCase(Campo3) Then
                    
                             valor_celula = Arr(i, Coluna4)
                            If VBA.UCase(VBA.Left(valor_celula, VBA.Len(Campo4))) = VBA.UCase(Campo4) Then
                    
                    
                                    For Coluna = 1 To TotalColuna
                                         
                                         arrayItems(Linha, Coluna) = Arr(i, Coluna)
                                    
                                    Next Coluna
                                    
                                    Linha = Linha + 1
                        
                          End If
                
                  End If
             
           End If
          
        End If
        
          
     Next i
   
   
  L.List = arrayItems()
  Call cabecalho

 Arr = Nothing
 
 

Exit Sub
Erro:
MsgBox "Erro!", vbCritical, "ERRO"

End Sub
Sub cabecalho()

Dim Cr As Double
Cr = 1
Dim Cl As Double
Cl = 0

With L
        
        .AddItem
        
        Do
        .List(0, Cl) = Arr(1, Cr)
        Cl = Cl + 1
        Cr = Cr + 1
        Loop Until Cr = TotalColuna + 1


End With

L.ColumnWidths = "75;60;120;75;150;50;50;100;100;150"

Arr = Nothing

End Sub

Sub att_Form2()
    Dim nlin As Integer
    nlin = UserForm1.ListBox2.ListIndex
    If nlin = -1 Then Exit Sub
    
    UserForm1.cbo_cod.Value = UserForm1.ListBox2.List(nlin, 0)
    UserForm1.cbo_propriedade.Value = UserForm1.ListBox2.List(nlin, 2)
    UserForm1.cbo_categoria.Value = UserForm1.ListBox2.List(nlin, 3)
    UserForm1.cbo_fornecedor.Value = UserForm1.ListBox2.List(nlin, 6)
    UserForm1.txt_valor_forn.Value = UserForm1.ListBox2.List(nlin, 7)
    UserForm1.txt_valor_cliente.Value = UserForm1.ListBox2.List(nlin, 5)
    UserForm1.txt_aprov.Value = UserForm1.ListBox2.List(nlin, 9)
    UserForm1.txt_po_cliente.Value = UserForm1.ListBox2.List(nlin, 12)
    UserForm1.txt_PO_Empresa.Value = UserForm1.ListBox2.List(nlin, 11)
    If UserForm1.ListBox2.List(nlin, 10) = 0 Then
        UserForm1.txt_data.Value = ""
    Else
        UserForm1.txt_data.Value = FormatDateTime(UserForm1.ListBox2.List(nlin, 10), vbShortDate)
    End If
End Sub
