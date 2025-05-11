VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6405
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm: UserForm1
' Controles esperados:
'   TextBox: txtNF, txtData, txtMes, txtAno, txtCNPJ, txtValor
'   ComboBox: cboEmissor, cboNatureza, cboCentro, cboPagamento
'   CommandButton: btnAdd, btnCanc
Option Explicit

Private Sub Label1_Click()
    ' (se não for usar, pode excluir este handler)
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lblEmissor_Click()

End Sub

Private Sub UserForm_Initialize()
    With Me.cboEmissor
        .Clear
        .List = ThisWorkbook.Names("ListaEmissor").RefersToRange.Value
    End With
    With Me.cboNatureza
        .Clear
        .List = ThisWorkbook.Names("ListaNaturezas").RefersToRange.Value
    End With
    With Me.cboPagamento
        .Clear
        .List = ThisWorkbook.Names("ListaPagamentos").RefersToRange.Value
    End With
End Sub

Private Sub txtData_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsDate(Me.txtData.Value) Then
        Me.txtMes.Value = Month(CDate(Me.txtData.Value))
        Me.txtAno.Value = Year(CDate(Me.txtData.Value))
    Else
        Me.txtMes.Value = ""
        Me.txtAno.Value = ""
    End If
End Sub

Private Sub btnAdd_Click()
    Dim ws   As Worksheet, ws2   As Worksheet
    Dim tbl  As ListObject, tbl2 As ListObject
    Dim lr   As ListRow, lr2     As ListRow
    Dim dt   As Date
    Dim val  As Double

    ' Validações obrigatórias
    If Trim(Me.txtNF.Value) = "" Then MsgBox "Informe o Nº da Nota Fiscal.", vbExclamation: Exit Sub
    If Not IsDate(Me.txtData.Value) Then MsgBox "Data inválida.", vbExclamation: Exit Sub Else dt = CDate(Me.txtData.Value)
    If Me.cboEmissor.ListIndex = -1 Then MsgBox "Selecione um Emissor.", vbExclamation: Exit Sub
    If Me.cboNatureza.ListIndex = -1 Then MsgBox "Selecione a Natureza.", vbExclamation: Exit Sub
    If Trim(Me.cboCentro.Value) = "" Then MsgBox "Informe o Centro de Custo.", vbExclamation: Exit Sub
    If Not IsNumeric(Me.txtValor.Value) Then MsgBox "Valor inválido.", vbExclamation: Exit Sub Else val = CDbl(Me.txtValor.Value)
    If Me.cboPagamento.ListIndex = -1 Then MsgBox "Selecione o Tipo de Pagamento.", vbExclamation: Exit Sub

    ' === Inserir na BASE ===
    Set ws = ThisWorkbook.Worksheets("BASE")
    Set tbl = ws.ListObjects("Tabela1")
    Set lr = tbl.ListRows.Add(1)

    ' Copia formatação da linha 2 para a nova
    tbl.ListRows(2).Range.Copy
    lr.Range.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Preenche valores na BASE
    Call PreencheLinha(lr, tbl, dt, val)

    ' === Inserir na BASE (2) ===
    Set ws2 = ThisWorkbook.Worksheets("BASE (2)")
    Set tbl2 = ws2.ListObjects("Tabela19")
    Set lr2 = tbl2.ListRows.Add(1)

    ' Copia formatação da linha 2 para a nova
    tbl2.ListRows(2).Range.Copy
    lr2.Range.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' Preenche valores na BASE (2)
    Call PreencheLinha(lr2, tbl2, dt, val)

    ' Atualiza Tabelas Dinâmicas e Dashboard
    ThisWorkbook.RefreshAll

    ' Limpa form e notifica
    With Me
        .txtNF.Value = ""
        .txtData.Value = ""
        .txtMes.Value = ""
        .txtAno.Value = ""
        .txtCNPJ.Value = ""
        .txtValor.Value = ""
        .cboEmissor.ListIndex = -1
        .cboNatureza.ListIndex = -1
        .cboCentro.Value = ""
        .cboPagamento.ListIndex = -1
    End With

    MsgBox "Lançamento Adicionado e Dashboards atualizados.", vbInformation
    Unload Me
End Sub


Private Sub PreencheLinha(lr As ListRow, tbl As ListObject, dt As Date, val As Double)
    With lr.Range
        .Cells(1, tbl.ListColumns("N° NOTA FISCAL").Index).Value = Me.txtNF.Value
        .Cells(1, tbl.ListColumns("EMISSOR").Index).Value = Me.cboEmissor.Value
        .Cells(1, tbl.ListColumns("DATA DE EMISSÃO").Index).Value = dt
        .Cells(1, tbl.ListColumns("MÊS").Index).Value = Me.txtMes.Value
        .Cells(1, tbl.ListColumns("ANO").Index).Value = Me.txtAno.Value
        .Cells(1, tbl.ListColumns("CNPJ").Index).Value = Me.txtCNPJ.Value
        .Cells(1, tbl.ListColumns("NATUREZA DA OPERAÇÃO").Index).Value = Me.cboNatureza.Value
        .Cells(1, tbl.ListColumns("CENTRO DE CUSTO").Index).Value = Me.cboCentro.Value
        .Cells(1, tbl.ListColumns("R$ TOTAL").Index).Value = val
        .Cells(1, tbl.ListColumns("TIPO DE PAGAMENTO").Index).Value = Me.cboPagamento.Value
    End With
End Sub

Private Sub btnCanc_Click()
    Unload Me
End Sub

