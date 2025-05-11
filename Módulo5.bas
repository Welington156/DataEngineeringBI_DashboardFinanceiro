Attribute VB_Name = "Módulo5"
Sub ImportarCamposDoXMLParaBASE()

    Dim xmlDoc As Object
    Dim xmlNode As Object
    Dim wsBase As Worksheet
    Dim wsXML As Worksheet
    Dim camposBase As Collection
    Dim celula As Range
    Dim campo As String
    Dim valorExtraido As String
    Dim i As Long

    ' Definir as planilhas
    Set wsBase = ThisWorkbook.Sheets("BASE")
    Set wsXML = ThisWorkbook.Sheets("XML TESTE")
    Set camposBase = New Collection
    
    ' Coletar os nomes dos campos na linha 1 da aba BASE
    For Each celula In wsBase.Range("A1", wsBase.Cells(1, wsBase.Cells(1, Columns.Count).End(xlToLeft).Column))
        camposBase.Add Trim(celula.Value)
    Next celula

    ' Carregar XML da célula A1
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.LoadXML wsXML.Range("A1").Value

    ' Verificar se o XML foi carregado com sucesso
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "Erro ao carregar o XML: " & xmlDoc.parseError.reason, vbCritical
        Exit Sub
    End If

    ' Preencher os campos na linha 2 da aba BASE
    For i = 1 To camposBase.Count
        campo = camposBase(i)
        valorExtraido = ""

        ' Tenta buscar o valor no XML usando o nome do campo
        On Error Resume Next
        Set xmlNode = xmlDoc.SelectSingleNode("//*[local-name()='" & campo & "']")
        If Not xmlNode Is Nothing Then
            valorExtraido = xmlNode.Text
        End If
        On Error GoTo 0

        ' Preenche o valor se existir
        If Len(valorExtraido) > 0 Then
            wsBase.Cells(2, i).Value = valorExtraido
        End If
    Next i

    MsgBox "Dados importados com sucesso!", vbInformation

End Sub

