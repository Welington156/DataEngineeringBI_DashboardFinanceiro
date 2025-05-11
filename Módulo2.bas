Attribute VB_Name = "M�dulo2"
Sub ConverterSaidaParaNegativo()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim tipo As String
    Dim valor As Variant
    
    ' Ajuste para sua planilha
    Set ws = ThisWorkbook.Worksheets("BASE (2)")
    
    ' Pega a �ltima linha com dados na coluna G (onde est� o Tipo)
    ultimaLinha = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    For i = 2 To ultimaLinha
        tipo = LCase(Trim(ws.Cells(i, "G").Value))
        If tipo = "sa�da" Or tipo = "saida" Then
            ' Pega o valor atual (pode vir como n�mero ou texto formatado)
            valor = ws.Cells(i, "I").Value
            
            ' Se for texto (ex.: "(1.234,56)"), tenta converter para n�mero
            If Not IsNumeric(valor) Then
                valor = ws.Cells(i, "I").Text
                ' Remove par�nteses e separadores para transformar em n�mero
                valor = Replace(valor, "(", "-")
                valor = Replace(valor, ")", "")
                valor = Replace(valor, ".", "")   ' milhares
                valor = Replace(valor, ",", ".")  ' decimal
                valor = CDbl(valor)
            End If
            
            ' Escreve o n�mero negativo � o formato Cont�bil permanece
            ws.Cells(i, "I").Value = -Abs(valor)
        End If
    Next i
    
    MsgBox "Convers�o conclu�da!", vbInformation
End Sub



