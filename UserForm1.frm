VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8475.001
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   15585
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_E1_Click()
    Call Lança_E1_PLAN
End Sub

Private Sub btn_filtro_Click()
    Call Filtro
End Sub

Private Sub btn_gerar_Click()
    On Error GoTo ErrorHandler
        Worksheets("LAYOUT_BASE").Copy Before:=Worksheets("LAYOUT_BASE")
        Worksheets("LAYOUT_BASE (2)").name = "LAYOUT_TEMP"
        Worksheets("LAYOUT_TEMP").Visible = True
        Call Inserir
    Exit Sub
ErrorHandler:
        Worksheets("LAYOUT_BASE (2)").Delete
End Sub

Private Sub btn_limpar_Click()
    Call Limpa_Form2
End Sub

Private Sub btnLimpar_Click()
    Call Limpa_Baseitens
    Call Limpa_Form
End Sub

Private Sub btn_aprovar_Click()
    Call Aprovar
    Call Filtro
    Call att_Form2
End Sub

Private Sub ListBox2_Change()
    Call att_Form2
End Sub

Private Sub UserForm_Initialize()
    Call Atulizar_ListBox
    Call Atualizar_ComboBox
    Call Filtro
    txtBdi.Value = 35
End Sub

Private Sub SpinBdi_SpinDown()
    txtBdi.Value = SpinBdi.Value
End Sub

Private Sub SpinBdi_SpinUp()
    txtBdi.Value = SpinBdi.Value
End Sub

Private Sub SpinQnt_SpinDown()
    txtqnt.Value = SpinQnt.Value
End Sub

Private Sub SpinQnt_SpinUp()
    txtqnt.Value = SpinQnt.Value
End Sub

Private Sub txtValor_AfterUpdate()
    'txtValor.Text = FormatCurrency(txtValor.Text)
End Sub
Private Sub AddItem_Click()
    Call InserirItem
    Call Inserir_Layout
End Sub

Private Sub btnsave_Click()
    Call Salvar_Como
    MsgBox ("Proposta salva com sucesso")
End Sub

Private Sub name_proposta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case 65 To 127
    Case Asc(" ")
    Case Asc("-")
    Case Asc(".")
    Case Else
        KeyAscii = 0
End Select
End Sub
