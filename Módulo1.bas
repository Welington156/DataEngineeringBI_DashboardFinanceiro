Attribute VB_Name = "Módulo1"
Option Explicit

Sub PreencherParcelas()
    Dim sheetNames As Variant
    Dim nomePlanilha As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dtEmi As Date
    Dim tipo As String
    Dim totalVal As Double
    Dim nParc As Integer
    Dim valParc As Double
    
    ' Lista de folhas a processar
    sheetNames = Array("BASE", "BASE (2)")
    
    For Each nomePlanilha In sheetNames
        Set ws = ThisWorkbook.Sheets(nomePlanilha)
        
        ' Encontra a última linha com dado na coluna J (Tipo de Pagamento)
        lastRow = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
        
        For i = 2 To lastRow
            dtEmi = ws.Cells(i, "C").Value              ' C = Data de Emissão
            tipo = UCase(Trim(ws.Cells(i, "J").Value))  ' J = Tipo de Pagamento
            totalVal = ws.Cells(i, "I").Value          ' I = R$ Total
            
            ' Limpa colunas K:R (datas e valores das parcelas)
            ws.Range(ws.Cells(i, "K"), ws.Cells(i, "R")).ClearContents
            
            ' Define número de parcelas e preenche datas
            Select Case tipo
                Case "A VISTA"
                    nParc = 1
                    ws.Cells(i, "K").Value = DateAdd("d", 1, dtEmi)
                Case "1+1"
                    nParc = 2
                    ws.Cells(i, "K").Value = DateAdd("d", 1, dtEmi)
                    ws.Cells(i, "M").Value = DateAdd("d", 30, dtEmi)
                Case "1+2"
                    nParc = 3
                    ws.Cells(i, "K").Value = DateAdd("d", 1, dtEmi)
                    ws.Cells(i, "M").Value = DateAdd("d", 30, dtEmi)
                    ws.Cells(i, "O").Value = DateAdd("d", 45, dtEmi)
                Case "1+3"
                    nParc = 4
                    ws.Cells(i, "K").Value = DateAdd("d", 1, dtEmi)
                    ws.Cells(i, "M").Value = DateAdd("d", 30, dtEmi)
                    ws.Cells(i, "O").Value = DateAdd("d", 45, dtEmi)
                    ws.Cells(i, "Q").Value = DateAdd("d", 60, dtEmi)
                Case Else
                    nParc = 0
            End Select
            
            ' Se tiver parcelas, divide o valor e preenche as colunas de valor
            If nParc > 0 Then
                valParc = totalVal / nParc
                ws.Cells(i, "L").Value = valParc   ' R$ 1ª Parcela
                If nParc >= 2 Then ws.Cells(i, "N").Value = valParc   ' R$ 2ª Parcela
                If nParc >= 3 Then ws.Cells(i, "P").Value = valParc   ' R$ 3ª Parcela
                If nParc >= 4 Then ws.Cells(i, "R").Value = valParc   ' R$ 4ª Parcela
            End If
        Next i
    Next nomePlanilha
    
    MsgBox "Preenchimento concluído!", vbInformation
End Sub


