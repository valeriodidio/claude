<!--#include virtual="/admin/int/prima.asp"-->
<!--#include virtual="/admin/int/menu.asp"-->
<%
' ============================================================
' Proxy ASP per scaricare l'Excel generato da FastAPI.
' Mantiene l'autenticazione admin (prima.asp verifica la sessione)
' e passa il file al browser come allegato.
' ============================================================

Server.ScriptTimeout = 600

INTERNAL_TOKEN = "change_me_long_random_string"   ' <-- stesso di actual_turnover_general.asp
REPORTS_API = "http://127.0.0.1:8000/api/reports/turnover/export.xlsx"

' Ricostruisce la query string in arrivo e chiama FastAPI
Dim qs
qs = Request.ServerVariables("QUERY_STRING")

Dim url
url = REPORTS_API
if qs <> "" then url = url & "?" & qs

Dim http
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.setTimeouts 5000, 10000, 30000, 600000   ' resolve, connect, send, receive
http.open "GET", url, False
http.setRequestHeader "X-Internal-Token", INTERNAL_TOKEN
http.send

if http.status <> 200 then
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/plain; charset=utf-8"
    Response.Write "Errore generazione Excel (" & http.status & "): " & http.responseText
    Response.End
end if

' Nome file dal backend (se presente) altrimenti default
Dim fname
fname = "actual_turnover.xlsx"
Dim disp
disp = http.getResponseHeader("Content-Disposition")
if disp <> "" then
    Response.AddHeader "Content-Disposition", disp
else
    Response.AddHeader "Content-Disposition", "attachment; filename=""" & fname & """"
end if

Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
Response.BinaryWrite http.responseBody
Response.End
%>
