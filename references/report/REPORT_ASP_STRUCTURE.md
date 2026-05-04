# Struttura ASP — Pagina Report

Questo documento descrive come costruire correttamente un file `.asp` per i report,
inclusa la gestione dei token, le chiamate AJAX al backend Python, la codifica UTF-8
dei caratteri speciali (€, à, è, ecc.) e il proxy per download Excel.

---

## Header obbligatorio — encoding UTF-8

**Questo è il punto più critico.** Classic ASP su IIS usa di default CodePage 1252
(Windows-1252). Se il backend Python restituisce JSON UTF-8, i caratteri € e gli
accenti vengono corrotti.

La soluzione è dichiarare UTF-8 come CodePage **in cima al file, prima di qualsiasi
output**:

```asp
<%@ Language="VBScript" CodePage="65001" %>
<!--#include virtual="/admin/int/prima.asp"-->
<!--#include virtual="/admin/int/menu.asp"-->
<%
Response.CodePage  = 65001
Response.Charset   = "UTF-8"
%>
```

> **Ordine fondamentale**: `<%@ CodePage="65001" %>` deve essere la **prima riga
> assoluta** del file, prima degli `#include`. Metterlo dopo causa encoding errato.

---

## Struttura completa file ASP

```asp
<%@ Language="VBScript" CodePage="65001" %>
<!--#include virtual="/admin/int/prima.asp"-->
<!--#include virtual="/admin/int/menu.asp"-->
<%
Response.CodePage = 65001
Response.Charset  = "UTF-8"

' Token iniettato dal deploy.ps1 — NON modificare manualmente
Const INTERNAL_TOKEN = "change_me_long_random_string"
Const API_BASE       = "http://127.0.0.1:8000/api/reports/nome_report"
%>
<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="utf-8">
<title>Nome Report</title>
<!-- CSS: vedi REPORT_UI_DESIGN.md -->
<style>
:root { /* palette */ }
/* ... tutti i CSS del design system ... */
</style>
</head>
<body>
<div id="Content">
  <!-- Contenuto HTML del report: filter bar, KPI, tabs, tabelle -->
</div>

<script>
// Token passato da VBScript → JS per le chiamate AJAX
var INTERNAL_TOKEN = "<%= INTERNAL_TOKEN %>";
var API_BASE       = "<%= API_BASE %>";

// Funzione generica per chiamare il backend
function apiGet(endpoint, params, onSuccess, onError) {
    var qs = Object.keys(params).map(function(k) {
        var v = params[k];
        if (Array.isArray(v)) {
            return v.map(function(x) {
                return encodeURIComponent(k) + "=" + encodeURIComponent(x);
            }).join("&");
        }
        return encodeURIComponent(k) + "=" + encodeURIComponent(v);
    }).join("&");

    $.ajax({
        url: API_BASE + "/" + endpoint + (qs ? "?" + qs : ""),
        method: "GET",
        headers: { "X-Internal-Token": INTERNAL_TOKEN },
        success: onSuccess,
        error: function(xhr) {
            if (typeof onError === "function") {
                onError(xhr.status, xhr.responseText);
            }
        }
    });
}
</script>
</body>
</html>
```

---

## Token INTERNAL_TOKEN — pattern placeholder

Il token viene sostituito automaticamente da `deploy.ps1` al momento del deploy.
Nel file sorgente (quello che carichi in staging) **lascia sempre il placeholder**:

```asp
Const INTERNAL_TOKEN = "change_me_long_random_string"
```

`deploy.ps1` cerca questa stringa in tutti i `.asp` e la sostituisce con il valore
reale letto dal `.env` del server:

```powershell
# da deploy.ps1 (estratto)
$content = Get-Content $aspFile -Raw -Encoding UTF8
$content = $content -replace [regex]::Escape($CFG.TokenPlaceholder), $realToken
Set-Content $aspFile -Value $content -Encoding UTF8
```

**Non inserire mai il token reale nei file sorgente.**

---

## Lettura JSON UTF-8 da Python — ADODB.Stream

Quando si chiama il backend Python da VBScript lato server (non AJAX),
`MSXML2.ServerXMLHTTP` restituisce il body come binario.
Per decodificarlo correttamente come UTF-8 (e preservare €, à, ecc.):

```asp
Dim http
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.open "GET", url, False
http.setRequestHeader "X-Internal-Token", INTERNAL_TOKEN
http.send

If http.status = 200 Then
    ' Decodifica il body binario come UTF-8
    Dim stream
    Set stream = Server.CreateObject("ADODB.Stream")
    stream.Open
    stream.Type     = 1        ' adTypeBinary
    stream.Write http.responseBody
    stream.Position = 0
    stream.Type     = 2        ' adTypeText
    stream.Charset  = "UTF-8"
    Dim jsonText
    jsonText = stream.ReadText
    stream.Close
    Set stream = Nothing

    ' jsonText ora contiene il JSON correttamente decodificato
    ' Puoi passarlo alla pagina con Response.Write o usarlo in VBScript
    Response.Write jsonText
Else
    Response.Status = "500 Internal Server Error"
    Response.Write "Errore API: " & http.status
End If
Set http = Nothing
```

> **Senza ADODB.Stream**: se usi `http.responseText` direttamente, IIS lo decodifica
> con la CodePage di sistema (1252), corrompendo € e caratteri multi-byte.

---

## Conversione data DD/MM/YYYY → YYYY-MM-DD

I daterangepicker restituiscono la data in formato italiano (DD/MM/YYYY).
Il backend FastAPI vuole il formato ISO (YYYY-MM-DD). Converti in JavaScript:

```javascript
function itToIso(s) {
    // "25/04/2026" → "2026-04-25"
    if (!s) return "";
    var p = s.split("/");
    if (p.length !== 3) return s;
    return p[2] + "-" + p[1] + "-" + p[0];
}

// Uso nel load:
var range = $("#daterange").val().split(" - ");
var dal   = itToIso(range[0]);
var al    = itToIso(range[1]);
```

---

## Download Excel — file ASP proxy

Per scaricare file binari (XLSX) generati da FastAPI, crea un ASP separato
che fa da proxy: legge il binario e lo inoltra al browser.

```asp
<%@ Language="VBScript" CodePage="65001" %>
<!--#include virtual="/admin/int/prima.asp"-->
<!--#include virtual="/admin/int/menu.asp"-->
<%
Response.CodePage = 65001
Response.Charset  = "UTF-8"

Server.ScriptTimeout = 600

Const INTERNAL_TOKEN = "change_me_long_random_string"
Const REPORTS_API    = "http://127.0.0.1:8000/api/reports/nome_report/export.xlsx"

' Ricostruisce la query string e chiama FastAPI
Dim qs : qs = Request.ServerVariables("QUERY_STRING")
Dim url : url = REPORTS_API
If qs <> "" Then url = url & "?" & qs

Dim http
Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.setTimeouts 5000, 10000, 30000, 600000   ' resolve, connect, send, receive
http.open "GET", url, False
http.setRequestHeader "X-Internal-Token", INTERNAL_TOKEN
http.send

If http.status <> 200 Then
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/plain; charset=utf-8"
    Response.Write "Errore generazione Excel (" & http.status & "): " & http.responseText
    Response.End
End If

' Usa il Content-Disposition del backend se presente, altrimenti default
Dim fname : fname = "report.xlsx"
Dim disp  : disp  = http.getResponseHeader("Content-Disposition")
If disp <> "" Then
    Response.AddHeader "Content-Disposition", disp
Else
    Response.AddHeader "Content-Disposition", "attachment; filename=""" & fname & """"
End If

Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
Response.BinaryWrite http.responseBody
Response.End
%>
```

Il link di download dalla pagina principale:

```javascript
function exportXlsx() {
    var params = buildQueryParams();  // stessa query string del report
    // Aggiungi il token come query param (BinaryWrite non permette response.write del token)
    var url = "/admin/asp_nuovo/nome_report/download-nome-xlsx.asp?" + params;
    window.location.href = url;
}
```

> **Nota**: per il download Excel il token viene trasmesso via header server-side
> dall'ASP proxy, non dal browser. Il browser chiama solo l'ASP (autenticato da
> `prima.asp`), e l'ASP chiama FastAPI con il token nell'header.

---

## Struttura cartelle ASP

```
admin/
└── asp_nuovo/
    └── nome_report/
        ├── nome_report.asp              ← pagina principale
        └── download-nome-xlsx.asp       ← proxy download Excel (se serve)
```

In produzione i file finiscono in:
```
D:\www\yeppon.it\admin\asp_nuovo\nome_report\
```

---

## Checklist file ASP

- [ ] Prima riga: `<%@ Language="VBScript" CodePage="65001" %>`
- [ ] Subito dopo include: `Response.CodePage = 65001` + `Response.Charset = "UTF-8"`
- [ ] `INTERNAL_TOKEN = "change_me_long_random_string"` (placeholder, NON il valore reale)
- [ ] Chiamate AJAX con header `"X-Internal-Token": INTERNAL_TOKEN`
- [ ] Conversione date IT→ISO prima di chiamare il backend
- [ ] Se leggi JSON lato VBScript: usa `ADODB.Stream` con `Charset = "UTF-8"`
- [ ] Se scarichi Excel: crea ASP proxy separato con `Response.BinaryWrite`
