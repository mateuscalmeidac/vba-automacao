Sub EnviarEmailsCrachas()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ws As Worksheet
    Dim i As Integer
    Dim nome As String
    Dim email As String
    Dim serie As String
    Dim modelo As String
    Dim centroCusto As String
    Dim corpoEmail As String
    Dim assunto As String
    Dim mesAtual As String

    Set ws = ThisWorkbook.Sheets(1)
    Set OutlookApp = CreateObject("Outlook.Application")
    mesAtual = Format(Date, "mmmm/yyyy")

    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        nome = ws.Cells(i, 1).Value
        email = ws.Cells(i, 2).Value
        serie = ws.Cells(i, 3).Value
        modelo = ws.Cells(i, 4).Value
        centroCusto = ws.Cells(i, 5).Value

        Set OutlookMail = OutlookApp.CreateItem(0)

        assunto = "Solicitação dos contadores de crachás - " & mesAtual

        corpoEmail = "Olá, " & nome & "!" & vbNewLine & vbNewLine & _
                     "Espero que esteja bem. Favor enviar os contadores da impressora de crachás referentes ao mês de " & mesAtual & "." & vbNewLine & vbNewLine & _
                     "Série: " & serie & vbNewLine & _
                     "Modelo: " & modelo & vbNewLine & _
                     "Centro de Custo: " & centroCusto & vbNewLine & vbNewLine & _
                     "Obrigado!" & vbNewLine & _
                     "Matheus Cruz - TI" & vbNewLine & vbNewLine & _
                     "CASO VOCÊ JÁ TENHA FEITO O ENVIO, FAVOR DESCONSIDERAR"

        With OutlookMail
            .To = email
            .Subject = assunto
            .Body = corpoEmail
            .Send
        End With
    Next i

    MsgBox "Todos os e-mails foram enviados com sucesso!", vbInformation
End Sub
