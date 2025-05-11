Attribute VB_Name = "Módulo3"
Sub EnviarResumoDashboard()
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim destinatario As String
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim tipo As String
    Dim valor As Double
    Dim totalGeral As Double
    Dim totalEntrada As Double
    Dim totalSaida As Double
    Dim totalEstorno As Double
    Dim numNotas As Long

    ' Solicita o e-mail
    destinatario = InputBox("Digite o endereço de e-mail do destinatário:", "Enviar Relatório")
    If destinatario = "" Then
        MsgBox "Envio cancelado.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets("BASE") ' ajuste se sua aba tiver outro nome
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' coluna A = Nº NOTA FISCAL

    ' Loop pelas linhas para calcular os totais
    For i = 2 To ultimaLinha
        tipo = Trim(ws.Cells(i, "G").Value) ' G = Tipo (Entrada/Saída)
        valor = ws.Cells(i, "I").Value             ' I = R$ TOTAL

        If IsNumeric(valor) Then
            totalGeral = totalGeral + valor
            If tipo = "Entrada" Then
                totalEntrada = totalEntrada + valor
            ElseIf tipo = "Saída" Or tipo = "Saida" Then
                totalSaida = totalSaida + valor
            ElseIf tipo = "Estornos" Then
                totalEstorno = totalEstorno + valor
            End If
            numNotas = numNotas + 1
        End If
    Next i

    ' Corpo do e-mail
    Dim corpoEmail As String
    corpoEmail = "Olá," & vbCrLf & vbCrLf & _
                 "Segue abaixo o resumo do relatório de notas fiscais:" & vbCrLf & vbCrLf & _
                 "Total de Notas Processadas: " & numNotas & vbCrLf & _
                 "Valor Total Geral: R$ " & Format(totalGeral, "#,##0.00") & vbCrLf & _
                 "Total de Entradas: R$ " & Format(totalEntrada, "#,##0.00") & vbCrLf & _
                 "Total de Saídas: R$ " & Format(totalSaida, "#,##0.00") & vbCrLf & _
                 "Total de Estornos: R$ " & Format(totalEstorno, "#,##0.00") & vbCrLf & vbCrLf & _
                 "Atenciosamente," & vbCrLf & "Equipe Financeira"

    ' Inicia o Outlook e prepara o e-mail
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If outlookApp Is Nothing Then
        MsgBox "Erro ao abrir o Outlook.", vbCritical
        Exit Sub
    End If

    Set outlookMail = outlookApp.CreateItem(0)

    With outlookMail
        .To = destinatario
        .Subject = "Resumo das Notas Fiscais"
        .Body = corpoEmail
        .Display ' Ou .Send para enviar direto
    End With

    MsgBox "Resumo pronto para envio!", vbInformation
End Sub


