Attribute VB_Name = "Módulo2"
Sub ConverterSaidaParaNegativo()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim tipo As String
    Dim valor As Variant
    
    ' Ajuste para sua planilha
    Set ws = ThisWorkbook.Worksheets("BASE (2)")
    
    ' Pega a última linha com dados na coluna G (onde está o Tipo)
    ultimaLinha = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    
    For i = 2 To ultimaLinha
        tipo = LCase(Trim(ws.Cells(i, "G").Value))
        If tipo = "saída" Or tipo = "saida" Then
            ' Pega o valor atual (pode vir como número ou texto formatado)
            valor = ws.Cells(i, "I").Value
            
            ' Se for texto (ex.: "(1.234,56)"), tenta converter para número
            If Not IsNumeric(valor) Then
                valor = ws.Cells(i, "I").Text
                ' Remove parênteses e separadores para transformar em número
                valor = Replace(valor, "(", "-")
                valor = Replace(valor, ")", "")
                valor = Replace(valor, ".", "")   ' milhares
                valor = Replace(valor, ",", ".")  ' decimal
                valor = CDbl(valor)
            End If
            
            ' Escreve o número negativo — o formato Contábil permanece
            ws.Cells(i, "I").Value = -Abs(valor)
        End If
    Next i
    
    MsgBox "Conversão concluída!", vbInformation
End Sub



