<%@ Language="VBScript" CodePage="65001" %>
<!--#include virtual="/admin/int/prima.asp"-->
<%
Response.CodePage = 65001
Response.Charset  = "UTF-8"

' ============================================================
' Proxy download Excel — P&L Prodotti
' Scarica il file .xlsx da FastAPI e lo passa al browser.
' Mantiene l'autenticazione admin (prima.asp) e passa il token.
' ============================================================
Server.ScriptTimeout = 600

Const INTERNAL_TOKEN = "change_me_long_random_string"
Const REPORTS_API    = "http://127.0.0.1:8000/api/reports/pl_prodotti/export.xlsx"

' Ricostruisce la querystring ricevuta dal browser e la passa a FastAPI
Dim qs, url
qs  = Request.ServerVariables("QUERY_STRING")
url = REPORTS_API
If qs <> "" Then url = url & "?" & qs

' Chiamata HTTP server-side a FastAPI
Dim http
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.setTimeouts 5000, 10000, 30000, 600000  ' resolve, connect, send, receive (ms)
http.open "GET", url, False
http.setRequestHeader "X-Internal-Token", INTERNAL_TOKEN
http.send

If http.status <> 200 Then
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/plain; charset=utf-8"
    Response.Write "Errore generazione file (" & http.status & "): " & http.responseText
    Response.End
End If

' Usa il Content-Disposition che arriva da FastAPI (include il nome file con le date)
Dim disp
disp = http.getResponseHeader("Content-Disposition")
If disp <> "" Then
    Response.AddHeader "Content-Disposition", disp
Else
    Response.AddHeader "Content-Disposition", "attachment; filename=""pl-prodotti-export.xlsx"""
End If

Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
Response.BinaryWrite http.responseBody
Response.End
%>
