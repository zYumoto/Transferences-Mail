Option Explicit

Sub SendEmails()
    Dim wsData As Worksheet, wsText As Worksheet, wsSubject As Worksheet
    Dim wsTemp As Worksheet
    Dim OutlookApp As Object, OutlookMail As Object
    Dim emailBody As String, preText As String, postText As String
    Dim lastRow As Long, client As String
    Dim cell As Range, tempTo As String, tempCc As String, tempBcc As String
    Dim htmlTable As String, tempSubject As String
    Dim dataRange As Range, uniqueClients As Collection, key As Variant
    Dim rngVisible As Range, row As Range

    Set wsData = ThisWorkbook.Sheets("BKGS TRANSF BY LINERS LIST")
    Set wsText = ThisWorkbook.Sheets("MAILTEXT EDIT")
    Set wsSubject = ThisWorkbook.Sheets("SUBJECT MESSAGE EDIT")

    ' Cria aba temporária
    Set wsTemp = ThisWorkbook.Sheets.Add(After:=wsData)
    wsTemp.Name = "Working Sheet"

    ' Copia os dados para aba temporária
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).row
    wsTemp.Range("A1:K1").Value = wsData.Range("A5:K5").Value
    wsTemp.Range("A2:K" & lastRow - 4).Value = wsData.Range("A6:K" & lastRow).Value

    ' Monta texto pré e pós com quebra de linha
    For Each cell In wsText.Range("A1:A2")
        preText = preText & cell.Text & "<br>"
    Next
    For Each cell In wsText.Range("A3:A71")
        If cell.Text = "" Then
            postText = postText & "<br>"
        Else
            postText = postText & cell.Text & "<br>"
        End If
    Next

    ' Coleta clientes únicos
    Set uniqueClients = New Collection
    On Error Resume Next
    For Each cell In wsTemp.Range("A2:A" & wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).row)
        uniqueClients.Add cell.Value, CStr(cell.Value)
    Next
    On Error GoTo 0

    ' Inicializa Outlook
    Set OutlookApp = CreateObject("Outlook.Application")

    ' Loop por cliente
    For Each key In uniqueClients
        wsTemp.Rows(1).AutoFilter Field:=1, Criteria1:=key
        Set rngVisible = wsTemp.Range("A1:K" & wsTemp.Cells(wsTemp.Rows.Count, 1).End(xlUp).row).SpecialCells(xlCellTypeVisible)

        ' Coleta To e Cc
        tempTo = ""
        tempCc = ""
        For Each row In rngVisible.Rows
            If row.row = 1 Then GoTo SkipRow ' pula cabeçalho
            If InStr(1, tempTo, row.Cells(10).Value) = 0 And row.Cells(10).Value <> "" Then
                tempTo = tempTo & row.Cells(10).Value & ";"
            End If
            If InStr(1, tempCc, row.Cells(11).Value) = 0 And row.Cells(11).Value <> "" Then
                tempCc = tempCc & row.Cells(11).Value & ";"
            End If
SkipRow:
        Next

        ' Verifica se o e-mail do remetente é válido
        If tempTo = "" Then
            MsgBox "Não foi possível encontrar o remetente para o cliente: " & key, vbExclamation, "Erro no E-mail"
            GoTo SkipClient
        End If

        tempBcc = wsSubject.Range("B1").Value
        If tempCc <> "" Then
            tempCc = tempCc & wsSubject.Range("B2").Value
        Else
            tempCc = wsSubject.Range("B2").Value
        End If
        tempSubject = wsSubject.Range("B4").Value

        ' Monta tabela HTML (apenas colunas A:E)
        htmlTable = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse; font-family:Calibri; font-size:11pt;'>"
        htmlTable = htmlTable & "<tr style='background-color:#f2f2f2; font-weight:bold;'>"
        htmlTable = htmlTable & "<td>EmailClient</td><td>BookingNumber</td><td>BookingClient</td><td>Transhipment</td><td>PortOfDischarge</td></tr>"

        For Each row In rngVisible.Rows
            If row.row = 1 Then GoTo SkipHeader
            htmlTable = htmlTable & "<tr>"
            htmlTable = htmlTable & "<td>" & CStr(row.Cells(1).Value) & "</td>"
            htmlTable = htmlTable & "<td>" & CStr(row.Cells(2).Value) & "</td>"
            htmlTable = htmlTable & "<td>" & CStr(row.Cells(3).Value) & "</td>"
            htmlTable = htmlTable & "<td>" & CStr(row.Cells(4).Value) & "</td>"
            htmlTable = htmlTable & "<td>" & CStr(row.Cells(5).Value) & "</td>"
            htmlTable = htmlTable & "</tr>"
SkipHeader:
        Next
        htmlTable = htmlTable & "</table><br>"

        ' Cria e-mail
        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = tempTo
            .Cc = tempCc
            .Bcc = tempBcc
            .Subject = tempSubject
            .HTMLBody = preText & htmlTable & postText
            
            On Error GoTo RemetenteErro
            .SentOnBehalfOfName = "BR241-mscbrazil.transferencecustomerservice@msc.com"
            On Error GoTo 0
            
            .Send
        End With

        wsTemp.AutoFilterMode = False
        GoTo ContinuaEnvio

RemetenteErro:
        MsgBox "Erro ao definir o remetente 'From'. Verifique se você tem permissão para enviar como 'BR241-mscbrazil.transferencecustomerservice@msc.com'.", vbCritical, "Erro de Remetente"
        On Error GoTo 0
        wsTemp.AutoFilterMode = False
        GoTo SkipClient

ContinuaEnvio:
SkipClient:
    Next key

    ' Deleta aba temporária
    Application.DisplayAlerts = False
    wsTemp.Delete
    Application.DisplayAlerts = True

    MsgBox "E-mails enviados com sucesso!", vbInformation
End Sub

-------------------------------------------------------------------

Function Traduzir(texto As String, Optional idiomaDestino As String = "pt") As String
    Dim http As Object
    Dim URL As String
    Dim resultado As String
    Dim inicio As Long, fim As Long

    On Error GoTo erro

    Set http = CreateObject("MSXML2.XMLHTTP")
    
    
    URL = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=" & idiomaDestino & "&dt=t&q=" & WorksheetFunction.EncodeURL(texto)
    
    http.Open "GET", URL, False
    http.Send
    
    resultado = http.responseText
    
    inicio = InStr(1, resultado, """") + 1
    fim = InStr(inicio, resultado, """") - 1
    Traduzir = Mid(resultado, inicio, fim - inicio + 1)
    Exit Function

erro:
    Traduzir = "Erro na tradução"
End Function

