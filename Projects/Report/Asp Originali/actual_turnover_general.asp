<!--#include virtual="/admin/int/prima.asp"-->
<!--#include virtual="/admin/int/menu.asp"-->
<!--#include virtual="/int/connection.asp"-->
<script>window.PRODUCT_NAME_UPDATE_URL = '/admin/int/update_nome_prodotto.asp';</script>
<title>Ecommerce Turnover</title>

<link rel="stylesheet" href="//code.jquery.com/ui/1.9.2/themes/base/jquery-ui.css">
<script src="//code.jquery.com/jquery-1.8.3.js"></script>
<script src="//code.jquery.com/ui/1.9.2/jquery-ui.js"></script>


<script type="text/javascript" src="https://cdn.jsdelivr.net/jquery/latest/jquery.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/momentjs/latest/moment.min.js"></script>
<script type="text/javascript" src="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.min.js"></script>
<link rel="stylesheet" type="text/css" href="https://cdn.jsdelivr.net/npm/daterangepicker/daterangepicker.css" />
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/jquery.tablesorter/2.9.1/jquery.tablesorter.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.4.0/Chart.min.js"></script>


<script type="text/javascript">	
	function collapseaction(id){
        $("#"+id).toggle();
    };
</script>

<style>
table {
	margin:0px 0px 15px 0px;
	padding:0px;
	font-size:12px;
	line-height:13px;
	color:#000;
	}
td {
	margin:0px 0px 15px 0px;
	padding:0px;
	font-size:10px;
	line-height:11px;
	color:#000;
	border:1px;
	border-style:dotted;
	border-color:#ccc;
	}
th {
	margin:0px 0px 15px 0px;
	padding:0px;
	font-size:10px;
	line-height:11px;
	color:#000;
	border:1px;
	border-style:dotted;
	border-color:#ccc;
	}
</style>


<script type="text/javascript">
$(function() {

    var start = moment().subtract(29, 'days');
    var end = moment();

    function cb(start, end) {
        $('#date-choose').val(start.format('DD/MM/YYYY') + ' - ' + end.format('DD/MM/YYYY'));
		document.getElementById("myform").submit();
    }

    $('#reportrange').daterangepicker({
        startDate: start,
        endDate: end,
        ranges: {
           'Today': [moment(), moment()],
           'Yesterday': [moment().subtract(1, 'days'), moment().subtract(1, 'days')],
           'Last 7 Days': [moment().subtract(6, 'days'), moment()],
           'Last 30 Days': [moment().subtract(29, 'days'), moment()],
           'This Month': [moment().startOf('month'), moment().endOf('month')],
           'Last Month': [moment().subtract(1, 'month').startOf('month'), moment().subtract(1, 'month').endOf('month')],
		   'This Year': [moment().startOf('year'), moment()]
        }
    }, cb);

    // $('#date-choose').val(start.format('DD/MM/YYYY') + ' - ' + end.format('DD/MM/YYYY'));

});
</script>
<%
token=request("token")
if token<>"" and token<>"958153f1b8b96ec4c4eb2147429105d9" then
	response.write "non fare il furbo"
	response.end
end if
Server.ScriptTimeout = 50000000
datechoose=request("date-choose")
'fixed_category="'GED','PED','Salute e Sport','Outlet'"
if datechoose<>"" then
	'response.write datechoose
	'response.end
	dataDal = left(datechoose,10)
	dataAl = right(datechoose,10)
else
	datechoose=date() & " - " & date()
end if


dated=cstr(request("dated"))
if dated<>"" then
	dataDal = DateAdd("d",dated,date())
end if

if dataDal = "" or dated="0" then 
	dataDal = date()
elseif dated="-60" then
	dataDal = "01/01/"&year(date()) 
end if


if dated="-1" then
	DataAl = DateAdd("d",dated,date())
else
	if DataAl = "" or dated="0" then 
		DataAl = date()
	end if
end if

forni = request("forni")
if forni="" then
	forni="0"
end if

categoria = request("categoria")
if categoria="" then
	categoria="0"
end if

marca = request("marca")
if marca="" then
	marca="0"
end if
cat2=request("cat2")
if cat2="" then
	cat2="0"
end if

codice=request("codice")
if codice="" then
	codice="0"
end if
flag=request("flag")
if flag="" then	
	flag=false
end if
visualizza=request("visualizza")
dettaglio=request("dettaglio")
'response.write codice
if codice="0" then

		Sel1 = "SELECT distinct p.id_p, " &_
			"IF(c1.nome='Auto' OR c1.nome='Bicicletta' OR c1.nome='Nautica' OR c1.nome='Moto','Auto/Bici/Moto/Nautica', " &_
			"IF(c1.nome='Giocattoli e Prima Infanzia','Giocattoli/Baby Care', " &_
			"IF(c1.nome='Casa Ferramenta e Sicurezza','Casa/Fai Da Te', " &_
			"IF(c1.nome='Moda','Moda/Cosmetica', " &_
			"IF(c1.nome='Fotografia','Foto', " &_
			"IF(c1.nome='Videogiochi','Videogiochi', " &_
			"IF(c1.nome='Informatica' OR (c1.nome='TV Audio Video' AND NOT(c2.seo='televisori')) ,'IT', " &_
			"IF(c1.nome='Telefoni e Smartphone','Telefonia', " &_
			"IF(c2.seo='Televisori','TV', " &_
			"IF(c1.nome='Elettrodomestici e Clima' AND NOT(c2.seo='grandi-elettrodomestici') AND NOT(c3.nome='Forni elettrici'),'PED', " &_
			"IF(c2.seo='grandi-elettrodomestici' OR c3.nome='Forni elettrici','GED','Altro'))))))))))) as cate3 " &_
			"from prodotti p " &_
			"inner join ordini_carrello c on p.id_p = c.id_p " &_
			"inner join ordini_cliente o on c.id_ordine = o.id "&_
			"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 AND e.site = CASE WHEN o.pixmania = 3 THEN 'it' WHEN o.pixmania = 23 THEN 'de' WHEN o.pixmania = 24 THEN 'fr'  WHEN o.pixmania = 25 THEN 'es'  WHEN o.pixmania = 26 THEN 'uk' ELSE 'it' END " &_
			"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
			"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
			"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
			"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
			"where o.data >= '"&datainnum(Datadal)&"' AND o.data <= '"&datainnum(DataAl)&"' AND c.quantita>0  AND ((o.step4_status = 1 OR (o.id_spe=111 AND o.step4_status <> 1)) AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
			  
		   
			if request("Marketplace") <> "" then
				Sel1 = Sel1 &"and o.pixmania = "&request("marketplace")&" "
			end if
			if request("forni") <> "" then
				  Sel1 = Sel1 &"and c.id_forni_nav = '"&request("forni")&"' "
			end if
			if request("cat2") <> "" then
				Sel1 = Sel1 &"and p.id_c3 = '"&request("cat2")&"' "
			end if
			
			if request("cat1") <> "" then
				Sel1 = Sel1 &"and p.id_c2 in ("&request("cat1")&") "
			end if
			if request("marca") <> "" then
					  Sel1 = Sel1 &"and p.id_m = '"&request("marca")&"' "
			end if

	
	
			
			Sel1 = Sel1 & "group by cate3,p.id_p "
		
			if request("categoria") <> "" and request("categoria") <> "0" then
				Sel1 = Sel1 &"having cate3= '"&request("categoria")&"' "
			end if
			'response.write sel1
			'response.end
			set rs1 = Conn.execute(Sel1)
			listidp=""
			do while not rs1.eof
			'response.write "aarrivato"
			'response.end
				listidp=listidp & rs1("id_p") & ","
			rs1.movenext
			loop
			rs1.close
			set rs1 = nothing
			'response.write listidp
			if listidp<>"" then
				listidp=left(listidp,len(listidp)-1)
			else	
				listidp=0
			end if
'response.write forni & " - " & categoria & " - " & listdp
'response.end
selForni = "select id, nome, id_forni_nav from fornitori where attivo = 1 order by nome"
set rsForni = Conn.execute(selForni)

selcat = 	"SELECT DISTINCT vcat.category AS nome " &_
			"FROM " &_
			"category_name vcat "&_
			"ORDER BY " &_ 
				"1"
'response.write selcat
'response.end
set rsCat = Conn.execute(selcat)

selMarca = "select distinct m.id_m, m.nome from marca as m join prodotti as p on p.id_m=m.id_m where p.id_p in ("&listidp&") order by nome"
set rsMarca = Conn.execute(selMarca)

selcat2 = "select distinct c2.id_c2, c2.nome from cat_2 as c2 join prodotti as p on p.id_c2=c2.id_c2 where p.id_p in ("&listidp&") order by nome"
'response.write selcat2
set rscat2 = Conn.execute(selcat2)

selweek="select DATE_FORMAT(current_date(),'%v') as week"
set rsweek = Conn.execute(selweek)
%>

<div id="Content">
	<h1>E-commerce Actual Turnover</h1>
<div>
	<form action="actual_turnover_general.asp" method="get" id="myform">
				
				<div id="reportrange" style=" width: 15%">
					<i class="fa fa-calendar"></i>
					Data<input type="text" id="date-choose" name="date-choose" value="<%=datechoose%>" style="width: 100%; vertical-align:middle;"/><i class="fa fa-caret-down"></i>
				</div>
				<br>MarketPlace 
                    <select name="Marketplace">
    <option value="">Tutti</option>
    <%SQL = "SELECT IntValue, name FROM marketplace"
    Set rs = Conn.Execute(sql)
    Do While NOT rs.EOF%>
    <option value="<%=rs(0)%>"><%=rs(1)%></option>
    <%rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    %>
    </select>
				Fornitore
                    <select name="Forni">
                        <option value="" selected>Tutti</option>
                        <%do while not rsForni.eof%>
                            <option value="<%=rsForni("id_forni_nav")%>" ><%=rsForni("nome")%></option>
						<%	on error resume next
                        rsForni.movenext
                        loop
                        rsForni.close
                        set rsForni = nothing%>
                    </select>
                Categoria
                    <select name="cat2">
                        <option value="" selected>Tutti</option>
                        <%do while not rscat2.eof%>
                            <option value="<%=rscat2("id_c2")%>" <%if cint(rscat2("id_c2"))=cint(cat2) then %>selected<%end if%>><%=rscat2("nome")%></option>
                        <%rscat2.movenext
                        loop 
                        rscat2.close
                        set rscat2 = nothing%>
                    </select>
                Marca
                    <select name="marca">
                        <option value="" selected>Tutti</option>
                        <%do while not rsMarca.eof%>
                            <option value="<%=rsMarca("id_m")%>" <%if cint(rsMarca("id_m"))=cint(marca) then %>selected<%end if%>><%=rsMarca("nome")%></option>
                        <%rsMarca.movenext
                        loop 
                        rsMarca.close
                        set rsMarca = nothing%>
                    </select>
					<input type="hidden" name="categoria" value="<%=categoria%>">
					<input type="hidden" name="dettaglio" value="<%=dettaglio%>">
			   <input type="submit" value="Cerca" name="B1" />

			
            </form>
			
</div>
<!--
<a class="show-trend-month" >Mostra Trend Mensile</a>
<div id="trend-mensile" name="trend-mensile"></div>
-->
<a class="show-trend" week-data="<%=rsweek("week")%>" visualizza-data="week" div-data="trend-week" onclick="collapseaction('trend-week');">Mostra Trend Settimana</a><br>
<a class="show-trend" week-data="<%=rsweek("week")%>" visualizza-data="hour" div-data="trend-hour" onclick="collapseaction('trend-hour');">Mostra Trend Orario</a><br>
<a class="show-trend" week-data="<%=rsweek("week")%>" visualizza-data="region" div-data="trend-region" onclick="collapseaction('trend-region');">Mostra Trend Per Regione</a><br>
<a class="show-trend" week-data="<%=rsweek("week")%>" visualizza-data="province" div-data="trend-province" onclick="collapseaction('trend-province');">Mostra Trend Per Provincia</a><br>
<a class="detmktp" week-data="<%=rsweek("week")%>" Datadal-data="<%=dataDal%>" DataAl-data="<%=DataAl%>" Marketplace-data="<%=request("Marketplace")%>">Mostra Dettaglio MarketPlace</a><br>
<div id="trend-week" name="trend-week" style="display:none"></div>
<div id="trend-hour" name="trend-hour" style="display:none"></div>
<div id="trend-region" name="trend-region" style="display:none"></div>
<div id="trend-province" name="trend-province" style="display:none"></div>


<div style=" max-height: 400px; overflow: auto;">
<table id="category_table" class="tablesorter" data-filter="#filter" data-filter-text-only="true" data-filter-minimum="2" id="projectSpreadsheet">
	 	<thead>
        <tr bgcolor="#ccc">
			<th width="20%" scope="col">Categoria Prodotto</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #5cd65c;">Pezzi Venduti Yeppon</th>
				<th width="5%" class="header" scope="col" style="background-color: #5cd65c;">Prezzo Medio Yeppon</th>
				<th width="5%" class="header" scope="col" style="background-color: #5cd65c;">Fatturato Yeppon</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #5cd65c;">GGP Yeppon</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #1aa3ff;">Pezzi Venduti CH</th>
				<th width="5%" class="header" scope="col" style="background-color: #1aa3ff;">Prezzo Medio CH</th>
				<th width="5%" class="header" scope="col" style="background-color: #1aa3ff;">Fatturato CH</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #1aa3ff;">GGP CH</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #ffcc66;">Pezzi Venduti Mktp</th>
				<th width="5%" class="header" scope="col" style="background-color: #ffcc66;">Prezzo Medio Mktp</th>
				<th width="5%" class="header" scope="col" style="background-color: #ffcc66;">Fatturato Mktp</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #ffcc66;">GGP Mktp</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #9966ff;">Pezzi Venduti Totali</th>
				<th width="5%" class="header" scope="col" style="background-color: #9966ff;">Prezzo Medio Totale</th>
				<th width="5%" class="header" scope="col" style="background-color: #9966ff;">Fatturato Totale</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #9966ff;">GGP Totale</th>
        </tr>
		</thead>
<%
        Sel3 = "SELECT   " &_
		"vcat.category as cate3, " &_
		"m.nome as marca, " &_
		"  SUM(IF(pixmania=0,c.quantita,0)) AS qta_yeppon, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon ,SUM(IF(pixmania=41,c.quantita,0)) AS qta_coinvest, (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41,IFNULL(e.contributo,0),0))*IF(pixmania=41 ,c.quantita,0)),2)-SUM(IF(pixmania=41 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest,SUM(IF(pixmania<>0 AND pixmania<>10 AND pixmania<>41,c.quantita,0)) AS qta_mktp, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND  ( pixmania<>41),c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND  ( pixmania<>41 ) AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)/1.22,0))) AS tot_mktp, SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=41,c.quantita,0)) + SUM(IF(pixmania<>0 AND pixmania<>10 AND  ( pixmania<>41 ),c.quantita,0)) as qta_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41 ,IFNULL(e.contributo,0),0))*IF(pixmania=41,c.quantita,0)),2)-SUM(IF(pixmania=41 ,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND  ( pixmania<>41),c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND  ( pixmania<>41 ) AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22)) as tot_tot, " &_ 
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2) as acquisto_yeppon, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2) as acquisto_coinvest, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND pixmania<>41,c.quantita,0)),2) as acquisto_mktp, (ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND pixmania<>41,c.quantita,0)),2)) as acquisto_tot " &_
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 AND e.site = CASE WHEN o.pixmania = 3 THEN 'it' WHEN o.pixmania = 23 THEN 'de' WHEN o.pixmania = 24 THEN 'fr'  WHEN o.pixmania = 25 THEN 'es'  WHEN o.pixmania = 26 THEN 'uk' ELSE 'it' END " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"LEFT OUTER JOIN fornitori_ordini f on c.id = f.id_car " &_
		"where o.data >= '"&datainnum(Datadal)&"' AND o.data <= '"&datainnum(DataAl)&"' AND c.quantita>0  AND ((o.step4_status = 1 OR (o.id_spe=111 AND o.step4_status <> 1)) AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526  "
          
       
		if request("Marketplace") <> "" then
			Sel3 = Sel3 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      sel3 = sel3 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		if request("cat1") <> "" then
			sel3 = sel3 &"and p.id_c2 in ("&request("cat1")&") "
		end if

		if marca <> "0" then
			  sel3 = sel3 &"and p.id_m = '"&marca&"' "
		end if
		
		Sel3 = Sel3 & "group by cate3 "
		
		Sel3 = Sel3 & "having cate3<>'Bibes' ORDER BY (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)) DESC"
		
		'response.write sel3
		'response.end
		
		set rs3 = Conn.execute(Sel3)
		
		tot_qta_yeppon=0
		tot_fatturato_yeppon=0
		tot_qta_coinvest=0
		tot_fatturato_coinvest=0
		tot_qta_mktp=0
		tot_fatturato_mktp=0
		tot_qta_tot=0
		tot_fatturato_tot=0
		iCount=0
		contaq=0
		Do While NOT rs3.EOF
            qta_yeppon  = cdbl(rs3("qta_yeppon"))
            fatturato_yeppon        = rs3("tot_yeppon")
            qta_coinvest  = cdbl(rs3("qta_coinvest"))
            fatturato_coinvest        = rs3("tot_coinvest")
            qta_mktp  = cdbl(rs3("qta_mktp"))
            fatturato_mktp       = rs3("tot_mktp")
			qta_tot = cdbl(rs3("qta_tot"))
            fatturato_tot       = rs3("tot_tot")
		
				fatturato_tot       = rs3("tot_tot")
				acquisto_yeppon      = rs3("acquisto_yeppon")
				acquisto_coinvest       = rs3("acquisto_coinvest")				
				acquisto_mktp       = rs3("acquisto_mktp")
				acquisto_tot       = rs3("acquisto_tot")
			
				tot_qta_yeppon   = tot_qta_yeppon + qta_yeppon
				tot_fatturato_yeppon   = tot_fatturato_yeppon + fatturato_yeppon
				tot_qta_coinvest   = tot_qta_coinvest + qta_coinvest
				tot_fatturato_coinvest   = tot_fatturato_coinvest + fatturato_coinvest
				tot_qta_mktp   = tot_qta_mktp + qta_mktp
				tot_fatturato_mktp   = tot_fatturato_mktp + fatturato_mktp
				tot_qta_tot = tot_qta_tot + qta_tot
				tot_fatturato_tot= tot_fatturato_tot + fatturato_tot
				tot_acquistato_yeppon=tot_acquistato_yeppon+acquisto_yeppon
				tot_acquistato_coinvest=tot_acquistato_coinvest+acquisto_coinvest
				tot_acquistato_mktp=tot_acquistato_mktp+acquisto_mktp
				tot_acquistato_tot=tot_acquistato_tot+acquisto_tot
				
				if tot_fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=tot_fatturato_yeppon/tot_qta_yeppon
					margine_yeppon=cdbl((tot_fatturato_yeppon-tot_acquistato_yeppon)/tot_fatturato_yeppon)*100
				end if
				if tot_fatturato_coinvest= 0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=tot_fatturato_coinvest/tot_qta_coinvest
					margine_coinvest=cdbl((tot_fatturato_coinvest-tot_acquistato_coinvest)/tot_fatturato_coinvest)*100
				end if
				if tot_fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=tot_fatturato_mktp/tot_qta_mktp
					margine_mktp=cdbl((tot_fatturato_mktp-tot_acquistato_mktp)/tot_fatturato_mktp)*100
				end if
				if tot_fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=tot_fatturato_tot/tot_qta_tot
					margine_tot=cdbl((tot_fatturato_tot-tot_acquistato_tot)/tot_fatturato_tot)*100
				end if
			rs3.movenext
			contaq=1
			
		iCount = iCount + 1
		if iCount = 10 then
			Response.Flush
			iCount = 0
		end if
			
        loop
%>
<thead>
    <tr>
        <td><a name="totali" href="?token=<%=token%>&date-choose=<%=datechoose%> ">Totale</a></td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_yeppon%></td>
			<td align="left" ><%=formatnumber(avg_price_yeppon,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_yeppon,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_coinvest%></td>
			<td align="left" ><%=formatnumber(avg_price_coinvest,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_coinvest,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_mktp%></td>
			<td align="left" ><%=formatnumber(avg_price_mktp,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_mktp,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_tot%></td>
			<td align="left" ><%=formatnumber(avg_price_tot,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_tot,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_tot,2)%>%</td>


    </tr>
</thead>
<%
		
		if contaq=1 then
			rs3.movefirst
		end if
		iCount=0
        Do While NOT rs3.EOF

            qta_yeppon  = cdbl(rs3("qta_yeppon"))
            fatturato_yeppon        = rs3("tot_yeppon")
            qta_coinvest  = cdbl(rs3("qta_coinvest"))
            fatturato_coinvest        = rs3("tot_coinvest")
            qta_mktp  = cdbl(rs3("qta_mktp"))
            fatturato_mktp       = rs3("tot_mktp")
			qta_tot = cdbl(rs3("qta_tot"))
            fatturato_tot       = rs3("tot_tot")			
			acquisto_yeppon      = rs3("acquisto_yeppon")
				acquisto_coinvest       = rs3("acquisto_coinvest")				
				acquisto_mktp       = rs3("acquisto_mktp")
				acquisto_tot       = rs3("acquisto_tot")
				if fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=fatturato_yeppon/qta_yeppon
					margine_yeppon=cdbl((fatturato_yeppon-acquisto_yeppon)/fatturato_yeppon)*100
				end if
				if fatturato_coinvest= 0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=fatturato_coinvest/qta_coinvest
					margine_coinvest=cdbl((fatturato_coinvest-acquisto_coinvest)/fatturato_coinvest)*100
				end if
				if fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=fatturato_mktp/qta_mktp
					margine_mktp=cdbl((fatturato_mktp-acquisto_mktp)/fatturato_mktp)*100
				end if
				if fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=fatturato_tot/qta_tot
					margine_tot=cdbl((fatturato_tot-acquisto_tot)/fatturato_tot)*100
				end if

            if bgcolor = "#EEE" then
                bgcolor = "#CCC;"
			else
                bgcolor = "#EEE"
            end if
            if rs3("cate3")=categoria then
				bgcolor = "#FA8072"
			end if
    %>
        <tr style="background-color:<%=bgcolor%>;">
 			<td><a name="<%=rs3("cate3")%>" href="?token=<%=token%>&date-choose=<%=datechoose%>&categoria=<%=rs3("cate3")%> "><%=rs3("cate3")%></a></td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_yeppon%></td>
				<td align="left"><%=formatnumber(avg_price_yeppon,2)%></td>
				<td align="left"><%=formatnumber(fatturato_yeppon,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_coinvest%></td>
				<td align="left"><%=formatnumber(avg_price_coinvest,2)%></td>
				<td align="left"><%=formatnumber(fatturato_coinvest,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_mktp%></td>
				<td align="left"><%=formatnumber(avg_price_mktp,2)%></td>
				<td align="left"><%=formatnumber(fatturato_mktp,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_tot%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_tot,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_tot,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_tot,2)%>%</td>

        </tr>
 
		
		<%

	if rs3("cate3")=categoria then

        Sel5 = "SELECT   " &_
		"vcat.category as cate3, " &_
		"m.nome as marca, " &_
		"p.id_c3 as id_c2, " &_
		"c3.nome as cate2, " &_
		"  SUM(IF(pixmania=0,c.quantita,0)) AS qta_yeppon, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon ,SUM(IF(pixmania=41,c.quantita,0)) AS qta_coinvest, (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41,IFNULL(e.contributo,0),0))*IF(pixmania=41,c.quantita,0)),2)) AS tot_coinvest,SUM(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)) AS qta_mktp, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41 AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)/1.22,0))) AS tot_mktp, SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=41,c.quantita,0)) + SUM(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)) as qta_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41,IFNULL(e.contributo,0),0))*IF(pixmania=41,c.quantita,0)),2)) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41 AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22)) as tot_tot, " &_ 
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2) as acquisto_yeppon, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2) as acquisto_coinvest, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2) as acquisto_mktp, (ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)) as acquisto_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0))/SUM(IF(pixmania=0,c.quantita,0)),2) as avg_price_yeppon,ROUND((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(o.importo_spe<>0,0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))/1.22)/SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) as avg_price_coinvest, ROUND((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22))/SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2) as avg_price_mktp, ROUND(((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22)) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(o.importo_spe<>0,0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))/1.22))/ (SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0)))),2) as avg_price_tot  " &_
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 AND e.site = CASE WHEN o.pixmania = 3 THEN 'it' WHEN o.pixmania = 23 THEN 'de' WHEN o.pixmania = 24 THEN 'fr'  WHEN o.pixmania = 25 THEN 'es'  WHEN o.pixmania = 26 THEN 'uk' ELSE 'it' END " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"LEFT OUTER JOIN fornitori_ordini f on c.id = f.id_car " &_
		"where o.data >= '"&datainnum(Datadal)&"' AND o.data <= '"&datainnum(DataAl)&"' AND c.quantita>0  AND ((o.step4_status = 1 OR (o.id_spe=111 AND o.step4_status <> 1)) AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526   "
          
		'AND vcat.category IN ("&fixed_category&")
		
		if request("Marketplace") <> "" then
			Sel5 = Sel5 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      Sel5 = Sel5 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		if categoria <> "" then
		      Sel5 = Sel5 &"and vcat.category = '"&categoria&"' "
		end if
		
		if request("cat1") <> "" then
			Sel5 = Sel5 &"and p.id_c2 in ("&request("cat1")&") "
		end if

		if marca <> "0" then
			  Sel5 = Sel5 &"and p.id_m = '"&marca&"' "
		end if
		
		Sel5 = Sel5 & "group by c3.nome "
		
		Sel5 = Sel5 & "ORDER BY (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)) DESC"
		
		'response.write Sel5
		'response.end
		
		set rs5 = Conn.execute(Sel5)
		
		
		iCount=0
        Do While NOT rs5.EOF

            qta_yeppon  = cdbl(rs5("qta_yeppon"))
            fatturato_yeppon        = rs5("tot_yeppon")
            qta_coinvest  = cdbl(rs5("qta_coinvest"))
            fatturato_coinvest        = rs5("tot_coinvest")
            qta_mktp  = cdbl(rs5("qta_mktp"))
            fatturato_mktp       = rs5("tot_mktp")
			qta_tot = cdbl(rs5("qta_tot"))
            fatturato_tot       = rs5("tot_tot")			
			acquisto_yeppon      = rs5("acquisto_yeppon")
			acquisto_coinvest       = rs5("acquisto_coinvest")				
			acquisto_mktp       = rs5("acquisto_mktp")
			acquisto_tot       = rs5("acquisto_tot")
			if fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=fatturato_yeppon/qta_yeppon
					margine_yeppon=cdbl((fatturato_yeppon-acquisto_yeppon)/fatturato_yeppon)*100
				end if
				if fatturato_coinvest= 0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=fatturato_coinvest/qta_coinvest
					margine_coinvest=cdbl((fatturato_coinvest-acquisto_coinvest)/fatturato_coinvest)*100
				end if
				if fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=fatturato_mktp/qta_mktp
					margine_mktp=cdbl((fatturato_mktp-acquisto_mktp)/fatturato_mktp)*100
				end if
				if fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=fatturato_tot/qta_tot
					margine_tot=cdbl((fatturato_tot-acquisto_tot)/fatturato_tot)*100
				end if

            if bgcolor = "#EEE" then
                bgcolor = "#CCC;"
			else
                bgcolor = "#EEE"
            end if
            if rs5("id_c2")=cint(cat2) then
				bgcolor = "#00FF00"
			end if
    %>
        <tr style="background-color:<%=bgcolor%>;">
 			<td align="right"><a name="<%=rs5("cate2")%>" href="?token=<%=token%>&date-choose=<%=datechoose%>&categoria=<%=categoria%>&cat2=<%=rs5("id_c2")%> "><%=rs5("cate2")%></a></td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_yeppon%></td>
				<td align="left"><%=formatnumber(avg_price_yeppon,2)%></td>
				<td align="left"><%=formatnumber(fatturato_yeppon,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_coinvest%></td>
				<td align="left"><%=formatnumber(avg_price_coinvest,2)%></td>
				<td align="left"><%=formatnumber(fatturato_coinvest,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_mktp%></td>
				<td align="left"><%=formatnumber(avg_price_mktp,2)%></td>
				<td align="left"><%=formatnumber(fatturato_mktp,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_tot%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_tot,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_tot,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_tot,2)%>%</td>

        </tr>
 
		
		<%

		rs5.movenext
		
		iCount = iCount + 1
		if iCount = 10 then
			Response.Flush
			iCount = 0
		end if
		
		loop
		rs5.close		
		
		end if
		
		
		
		rs3.movenext
		
		iCount = iCount + 1
		if iCount = 10 then
			Response.Flush
			iCount = 0
		end if
		
		loop
		rs3.close
		%>
   </table>
   </div>
   </td>
<div id="det_mktp" name="det_mktp" ></div>
   <%
	'tabella marche
	'response.write categoria
	'response.end
	if categoria<>"0" then
	
	%>
<div style=" max-height: 400px; overflow: auto;">
	<table id="marca_table" class="tablesorter" data-filter="#filter" data-filter-text-only="true" data-filter-minimum="2" id="projectSpreadsheet">
	 	<thead>
        <tr bgcolor="#ccc">
			<th width="20%" align="center" scope="col">Marca</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #5cd65c;">Pezzi Venduti Yeppon</th>
				<th width="5%" class="header" scope="col" style="background-color: #5cd65c;">Prezzo Medio Yeppon</th>
				<th width="5%" class="header" scope="col" style="background-color: #5cd65c;">Fatturato Yeppon</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #5cd65c;">GGP Yeppon</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #1aa3ff;">Pezzi Venduti CH</th>
				<th width="5%" class="header" scope="col" style="background-color: #1aa3ff;">Prezzo Medio CH</th>
				<th width="5%" class="header" scope="col" style="background-color: #1aa3ff;">Fatturato CH</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #1aa3ff;">GGP CH</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #ffcc66;">Pezzi Venduti Mktp</th>
				<th width="5%" class="header" scope="col" style="background-color: #ffcc66;">Prezzo Medio Mktp</th>
				<th width="5%" class="header" scope="col" style="background-color: #ffcc66;">Fatturato Mktp</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #ffcc66;">GGP Mktp</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #9966ff;">Pezzi Venduti Totali</th>
				<th width="5%" class="header" scope="col" style="background-color: #9966ff;">Prezzo Medio Totale</th>
				<th width="5%" class="header" scope="col" style="background-color: #9966ff;">Fatturato Totale</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #9966ff;">GGP Totale</th>
        </tr>
		</thead>
<%
        Sel2 = " SELECT m.nome as marca,SUM(IF(pixmania=0,c.quantita,0)) AS qta_yeppon, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon ,SUM(IF(pixmania=41,c.quantita,0)) AS qta_coinvest, (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41,IFNULL(e.contributo,0),0))*IF(pixmania=41,c.quantita,0)),2)) AS tot_coinvest,SUM(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)) AS qta_mktp, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41 AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)/1.22,0))) AS tot_mktp, SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=41,c.quantita,0)) + SUM(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)) as qta_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41,IFNULL(e.contributo,0),0))*IF(pixmania=41,c.quantita,0)),2)) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41 AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22)) as tot_tot, " &_ 
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2) as acquisto_yeppon, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2) as acquisto_coinvest, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2) as acquisto_mktp, (ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)) as acquisto_tot, " &_
		"vcat.category as cate3, " &_
		"m.id_m as id_m " &_
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 AND e.site = CASE WHEN o.pixmania = 3 THEN 'it' WHEN o.pixmania = 23 THEN 'de' WHEN o.pixmania = 24 THEN 'fr'  WHEN o.pixmania = 25 THEN 'es'  WHEN o.pixmania = 26 THEN 'uk' ELSE 'it' END " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"LEFT OUTER JOIN fornitori_ordini f on c.id = f.id_car " &_
		"where o.data >= '"&datainnum(Datadal)&"' AND o.data <= '"&datainnum(DataAl)&"' AND c.quantita>0  AND ((o.step4_status = 1 OR (o.id_spe=111 AND o.step4_status <> 1)) AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526  "
          
		'AND vcat.category IN ("&fixed_category&")
       
		if request("Marketplace") <> "" then
			Sel2 = Sel2 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      Sel2 = Sel2 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		if cat2 <> "0" then
		      sel2 = sel2 &"and p.id_c3 = '"&cat2&"' "
		end if

		if request("cat1") <> "" then
			sel2 = sel2 &"and p.id_c2 in ("&request("cat1")&") "
		end if		

		Sel2 = Sel2 & "group by cate3,marca "
		
		Sel2 = Sel2 &"having cate3= '"&request("categoria")&"' "
		
		Sel2 = Sel2 & "ORDER BY (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)) DESC"
		
		set rs2 = Conn.execute(Sel2)
		
		tot_qta_yeppon=0
		tot_fatturato_yeppon=0
		tot_qta_coinvest=0
		tot_fatturato_coinvest=0
		tot_qta_mktp=0
		tot_fatturato_mktp=0
		tot_qta_tot=0
		tot_fatturato_tot=0
		tot_acquistato_yeppon=0
		tot_acquistato_coinvest=0
		tot_acquistato_mktp=0
		tot_acquistato_tot=0
		iCount=0
		contaq=0
		Do While NOT rs2.EOF
            qta_yeppon  = cdbl(rs2("qta_yeppon"))
            fatturato_yeppon        = rs2("tot_yeppon")
            qta_coinvest  = cdbl(rs2("qta_coinvest"))
            fatturato_coinvest        = rs2("tot_coinvest")
            qta_mktp  = cdbl(rs2("qta_mktp"))
            fatturato_mktp       = rs2("tot_mktp")
			qta_tot = cdbl(rs2("qta_tot"))
            fatturato_tot       = rs2("tot_tot")
			acquisto_yeppon      = rs2("acquisto_yeppon")
			acquisto_coinvest       = rs2("acquisto_coinvest")				
			acquisto_mktp       = rs2("acquisto_mktp")
			acquisto_tot       = rs2("acquisto_tot")
		
			tot_qta_yeppon   = tot_qta_yeppon + qta_yeppon
			tot_fatturato_yeppon   = tot_fatturato_yeppon + fatturato_yeppon
			tot_qta_coinvest   = tot_qta_coinvest + qta_coinvest
			tot_fatturato_coinvest   = tot_fatturato_coinvest + fatturato_coinvest
			tot_qta_mktp   = tot_qta_mktp + qta_mktp
			tot_fatturato_mktp   = tot_fatturato_mktp + fatturato_mktp
			tot_qta_tot = tot_qta_tot + qta_tot
			tot_fatturato_tot= tot_fatturato_tot + fatturato_tot
			tot_acquistato_yeppon=tot_acquistato_yeppon+acquisto_yeppon
			tot_acquistato_coinvest=tot_acquistato_coinvest+acquisto_coinvest
			tot_acquistato_mktp=tot_acquistato_mktp+acquisto_mktp
			tot_acquistato_tot=tot_acquistato_tot+acquisto_tot
			
				if tot_fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=tot_fatturato_yeppon/tot_qta_yeppon
					margine_yeppon=cdbl((tot_fatturato_yeppon-tot_acquistato_yeppon)/tot_fatturato_yeppon)*100
				end if
				if tot_fatturato_coinvest= 0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=tot_fatturato_coinvest/tot_qta_coinvest
					margine_coinvest=cdbl((tot_fatturato_coinvest-tot_acquistato_coinvest)/tot_fatturato_coinvest)*100
				end if
				if tot_fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=tot_fatturato_mktp/tot_qta_mktp
					margine_mktp=cdbl((tot_fatturato_mktp-tot_acquistato_mktp)/tot_fatturato_mktp)*100
				end if
				if tot_fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=tot_fatturato_tot/tot_qta_tot
					margine_tot=cdbl((tot_fatturato_tot-tot_acquistato_tot)/tot_fatturato_tot)*100
				end if
			rs2.movenext
			contaq=1
		iCount = iCount + 1
		if iCount = 10 then
			Response.Flush
			iCount = 0
		end if
			
        loop
%>
<thead>
    <tr>
        <td><a name="totali" href="?token=<%=token%>&date-choose=<%=datechoose%>&categoria=<%=categoria%>&cat2=<%=cat2%>&dated=<%=dated%> ">Totale</a></td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_yeppon%></td>
			<td align="left" ><%=formatnumber(avg_price_yeppon,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_yeppon,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_coinvest%></td>
			<td align="left" ><%=formatnumber(avg_price_coinvest,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_coinvest,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_mktp%></td>
			<td align="left" ><%=formatnumber(avg_price_mktp,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_mktp,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_tot%></td>
			<td align="left" ><%=formatnumber(avg_price_tot,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_tot,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_tot,2)%>%</td>


    </tr>
</thead>

<%
		
		if contaq=1 then
			rs2.movefirst
		end if
		
		iCount=0
        Do While NOT rs2.EOF

            qta_yeppon  = cdbl(rs2("qta_yeppon"))
            fatturato_yeppon        = rs2("tot_yeppon")
            qta_coinvest  = cdbl(rs2("qta_coinvest"))
            fatturato_coinvest        = rs2("tot_coinvest")
            qta_mktp  = cdbl(rs2("qta_mktp"))
            fatturato_mktp       = rs2("tot_mktp")
			qta_tot = cdbl(rs2("qta_tot"))
            fatturato_tot       = rs2("tot_tot")			
			acquisto_yeppon      = rs2("acquisto_yeppon")
				acquisto_coinvest       = rs2("acquisto_coinvest")				
				acquisto_mktp       = rs2("acquisto_mktp")
				acquisto_tot       = rs2("acquisto_tot")
				
				if fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=fatturato_yeppon/qta_yeppon
					margine_yeppon=cdbl((fatturato_yeppon-acquisto_yeppon)/fatturato_yeppon)*100
				end if
				if fatturato_coinvest= 0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=fatturato_coinvest/qta_coinvest
					margine_coinvest=cdbl((fatturato_coinvest-acquisto_coinvest)/fatturato_coinvest)*100
				end if
				if fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=fatturato_mktp/qta_mktp
					margine_mktp=cdbl((fatturato_mktp-acquisto_mktp)/fatturato_mktp)*100
				end if
				if fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=fatturato_tot/qta_tot
					margine_tot=cdbl((fatturato_tot-acquisto_tot)/fatturato_tot)*100
				end if
            if bgcolor = "#EEE" then
                bgcolor = "#CCC;"
            else
                bgcolor = "#EEE"
            end if
			if cint(rs2("id_m"))=cint(marca) then
				bgcolor = "#00FF00"
			end if

    %>
        <tr style="background-color:<%=bgcolor%>;">
			<td><a name="<%=rs2("id_m")%>" href="?token=<%=token%>&date-choose=<%=datechoose%>&categoria=<%=rs2("cate3")%>&cat2=<%=cat2%>&marca=<%=rs2("id_m")%>"><%=rs2("marca")%></a></td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_yeppon%></td>
				<td align="left"><%=formatnumber(avg_price_yeppon,2)%></td>
				<td align="left"><%=formatnumber(fatturato_yeppon,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_coinvest%></td>
				<td align="left"><%=formatnumber(avg_price_coinvest,2)%></td>
				<td align="left"><%=formatnumber(fatturato_coinvest,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_mktp%></td>
				<td align="left"><%=formatnumber(avg_price_mktp,2)%></td>
				<td align="left"><%=formatnumber(fatturato_mktp,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_tot%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_tot,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_tot,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_tot,2)%>%</td>

        </tr>

		
		<%
		
		rs2.movenext
		
		iCount = iCount + 1
		if iCount = 10 then
			Response.Flush
			iCount = 0
		end if
		loop
		rs2.close
		end if
		%>

	</table>	
</div>
<div style=" overflow: auto;">
 <%
	'tabella prodotti
	'response.write rs2("id_m") & "-" & marca
end if
if dettaglio="" then	
	%>

	<table id="prodotti_table" class="tablesorter" data-filter="#filter" data-filter-text-only="true" data-filter-minimum="2" id="projectSpreadsheet">
			<thead>
			<tr>
<%
if codice="0" then
%>
                <th bgcolor="#ccc" width="5%">Codice</th>
				<th bgcolor="#ccc" width="5%">Marca</th>
				<th bgcolor="#ccc" width="10%" >Nome</th>
<%
end if
%>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #5cd65c;">Pezzi Venduti Yeppon</th>
				<th width="5%" class="header" scope="col" style="background-color: #5cd65c;">Prezzo Medio Yeppon</th>
				<th width="5%" class="header" scope="col" style="background-color: #5cd65c;">Fatturato Yeppon</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #5cd65c;">GGP Yeppon</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #1aa3ff;">Pezzi Venduti CH</th>
				<th width="5%" class="header" scope="col" style="background-color: #1aa3ff;">Prezzo Medio CH</th>
				<th width="5%" class="header" scope="col" style="background-color: #1aa3ff;">Fatturato CH</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #1aa3ff;">GGP CH</th>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #ffcc66;">Pezzi Venduti Mktp</th>
				<th width="5%" class="header" scope="col" style="background-color: #ffcc66;">Prezzo Medio Mktp</th>
				<th width="5%" class="header" scope="col" style="background-color: #ffcc66;">Fatturato Mktp</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #ffcc66;">GGP Mktp</th>
<%
if codice="0" then
%>
				<th width="5%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #9966ff;">Pezzi Venduti Totali</th>
				<th width="5%" class="header" scope="col" style="background-color: #9966ff;">Prezzo Medio Totale</th>
				<th width="5%" class="header" scope="col" style="background-color: #9966ff;">Fatturato Totale</th>
				<th width="5%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #9966ff;">GGP Totale</th>
<%
end if
%>
			</tr>
			</thead>
	 

	<%     if bgcolor = "#DEDEDE" then
				bgcolor = "#FFF;"
			else
				bgcolor = "#DEDEDE"
			end if

	%>

		<%
		

			Sel4 = "SELECT p.id_p, p.codice, p.nome AS nomeProd,p.imgt AS imgp, SUM(IF(pixmania=0,c.quantita,0)) AS qta_yeppon, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon ,SUM(IF(pixmania=41,c.quantita,0)) AS qta_coinvest, (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41,IFNULL(e.contributo,0),0))*IF(pixmania=41,c.quantita,0)),2)) AS tot_coinvest,SUM(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)) AS qta_mktp, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41 AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22)) AS tot_mktp, SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=41,c.quantita,0)) + SUM(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)) as qta_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=41,IFNULL(e.contributo,0),0))*IF(pixmania=41,c.quantita,0)),2)) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)-sum(IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41 AND o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22)) as tot_tot, " &_ 
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2) as acquisto_yeppon, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2) as acquisto_coinvest, " &_
			"ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2) as acquisto_mktp, (ROUND(SUM(c.prezzo_acquisto*IF(pixmania=0,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania=41,c.quantita,0)),2)+ROUND(SUM(c.prezzo_acquisto*IF(pixmania<>0 AND pixmania<>10 AND   pixmania<>41,c.quantita,0)),2)) as acquisto_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0))/SUM(IF(pixmania=0,c.quantita,0)),2) as avg_price_yeppon,ROUND((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(o.importo_spe<>0,0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))/1.22)/SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) as avg_price_coinvest, ROUND((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22))/SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2) as avg_price_mktp, ROUND(((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22)) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(o.importo_spe<>0,0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id))))/1.22)/ (SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0)))),2) as avg_price_tot, " &_
			"vcat.category as cate3," &_
			"m.nome as marca, " &_
			"p.id_m as id_m " &_
			"from prodotti p " &_
			"inner join ordini_carrello c on p.id_p = c.id_p " &_
			"inner join fornitori forn using(id_forni_nav) " &_
			"inner join ordini_cliente o on c.id_ordine = o.id "&_
			"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
			"LEFT JOIN (select distinct codice,dataInizio,dataFine,contributo,site FROM ebay_offerte) AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 AND e.site = CASE WHEN o.pixmania = 3 THEN 'it' WHEN o.pixmania = 23 THEN 'de' WHEN o.pixmania = 24 THEN 'fr'  WHEN o.pixmania = 25 THEN 'es'  WHEN o.pixmania = 26 THEN 'uk' ELSE 'it' END " &_
			"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
			"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
			"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
			"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
			"LEFT OUTER JOIN fornitori_ordini f on c.id = f.id_car " &_
			"where o.data >= '"&datainnum(Datadal)&"' AND o.data <= '"&datainnum(DataAl)&"' AND c.quantita>0 AND pixmania<>10 AND ((o.step4_status = 1 OR (o.id_spe=111 AND o.step4_status <> 1)) AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526   "
			  
			'AND vcat.category IN ("&fixed_category&")
		   
			if request("Marketplace") <> "" then
				Sel4 = Sel4 &"and o.pixmania = "&request("marketplace")&" "
			end if
			if request("forni") <> "" then
				  sel4 = sel4 &"and c.id_forni_nav = '"&request("forni")&"' "
			end if
			if cat2 <> "0" then
				sel4 = sel4 &"and p.id_c3 = '"&cat2&"' "
			end if
			if marca <> "0" then
				  sel4 = sel4 &"and p.id_m = '"&marca&"' "
			end if

			if codice <> "0" then
				  sel4 = sel4 &"and p.codice = '"&codice&"' "
			end if

			if request("cat1") <> "" then
				sel4 = sel4 &"and p.id_c2 in ("&request("cat1")&") "
			end if			
			
			Sel4 = Sel4 & "group by cate3,p.id_p "
		
			if request("categoria") <> "" and request("categoria") <> "0" then
				sel4 = sel4 &"having cate3= '"&request("categoria")&"' "
			end if
			
			Sel4 = Sel4 & "having cate3<>'Bibes' ORDER BY (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) + ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)) DESC"
			
			'response.write sel4
			'response.end
			set rs4 = Conn.execute(Sel4)
			

	''
			tot_qta_yeppon=0
			tot_fatturato_yeppon=0
			tot_qta_coinvest=0
			tot_fatturato_coinvest=0
			tot_qta_mktp=0
			tot_fatturato_mktp=0
			tot_qta_tot=0
			tot_fatturato_tot=0
			tot_acquistato_yeppon=0
			tot_acquistato_coinvest=0
			tot_acquistato_mktp=0
			tot_acquistato_tot=0
			
			iCount=0
			contaq=0
			Do While NOT rs4.EOF
				qta_yeppon  = cdbl(rs4("qta_yeppon"))
				fatturato_yeppon        = rs4("tot_yeppon")
				qta_coinvest  = cdbl(rs4("qta_coinvest"))
				fatturato_coinvest        = rs4("tot_coinvest")
				qta_mktp  = cdbl(rs4("qta_mktp"))
				fatturato_mktp       = rs4("tot_mktp")
				qta_tot = cdbl(rs4("qta_tot"))
				fatturato_tot       = rs4("tot_tot")
				acquisto_yeppon      = rs4("acquisto_yeppon")
				acquisto_coinvest       = rs4("acquisto_coinvest")				
				acquisto_mktp       = rs4("acquisto_mktp")
				acquisto_tot       = rs4("acquisto_tot")
			
				tot_qta_yeppon   = tot_qta_yeppon + qta_yeppon
				tot_fatturato_yeppon   = tot_fatturato_yeppon + fatturato_yeppon
				tot_qta_coinvest   = tot_qta_coinvest + qta_coinvest
				tot_fatturato_coinvest   = tot_fatturato_coinvest + fatturato_coinvest
				tot_qta_mktp   = tot_qta_mktp + qta_mktp
				tot_fatturato_mktp   = tot_fatturato_mktp + fatturato_mktp
				tot_qta_tot = tot_qta_tot + qta_tot
				tot_fatturato_tot= tot_fatturato_tot + fatturato_tot
				tot_acquistato_yeppon=tot_acquistato_yeppon+acquisto_yeppon
				tot_acquistato_coinvest=tot_acquistato_coinvest+acquisto_coinvest
				tot_acquistato_mktp=tot_acquistato_mktp+acquisto_mktp
				tot_acquistato_tot=tot_acquistato_tot+acquisto_tot
				
				if tot_fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=tot_fatturato_yeppon/tot_qta_yeppon
					margine_yeppon=cdbl((tot_fatturato_yeppon-tot_acquistato_yeppon)/tot_fatturato_yeppon)*100
				end if
				if tot_fatturato_coinvest= 0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=tot_fatturato_coinvest/tot_qta_coinvest
					margine_coinvest=cdbl((tot_fatturato_coinvest-tot_acquistato_coinvest)/tot_fatturato_coinvest)*100
				end if
				if tot_fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=tot_fatturato_mktp/tot_qta_mktp
					margine_mktp=cdbl((tot_fatturato_mktp-tot_acquistato_mktp)/tot_fatturato_mktp)*100
				end if
				if tot_fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=tot_fatturato_tot/tot_qta_tot
					margine_tot=cdbl((tot_fatturato_tot-tot_acquistato_tot)/tot_fatturato_tot)*100
				end if
				
				
				
				rs4.movenext
				contaq=1
			iCount = iCount + 1
			if iCount = 10 then
				Response.Flush
				iCount = 0
			end if
				
			loop
	%>
<%
if codice="0" then
%>
	<thead>
		<tr style="height: 20px">
			<td colspan="3">Totali</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_yeppon%></td>
			<td align="left" ><%=formatnumber(avg_price_yeppon,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_yeppon,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_coinvest%></td>
			<td align="left" ><%=formatnumber(avg_price_coinvest,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_coinvest,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_mktp%></td>
			<td align="left" ><%=formatnumber(avg_price_mktp,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_mktp,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_tot%></td>
			<td align="left" ><%=formatnumber(avg_price_tot,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_tot,2)%></td>
			<td align="left"  style=" border-right: 5px solid black; "><%=formatnumber(margine_tot,2)%>%</td>


		</tr>
	</thead>
<%
end if
%>
	<%
			
			if contaq=1 then
				rs4.movefirst
			end if
			iCount=0
			Do While NOT rs4.EOF

				id_p        = rs4("id_p")
				'codice      = rs4("codice")
				nomeProd    = rs4("nomeProd")
				imgprod    = rs4("imgp")
				imgprod = "https://photo.yeppon.it"&imgprod
				qta_yeppon  = cdbl(rs4("qta_yeppon"))
				fatturato_yeppon        = rs4("tot_yeppon")
				qta_coinvest  = cdbl(rs4("qta_coinvest"))
				fatturato_coinvest        = rs4("tot_coinvest")
				qta_mktp  = cdbl(rs4("qta_mktp"))
				fatturato_mktp       = rs4("tot_mktp")
				qta_tot = cdbl(rs4("qta_tot"))
				fatturato_tot      = rs4("tot_tot")
				acquisto_yeppon      = rs4("acquisto_yeppon")
				acquisto_coinvest       = rs4("acquisto_coinvest")				
				acquisto_mktp       = rs4("acquisto_mktp")
				acquisto_tot       = rs4("acquisto_tot")
				if fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=fatturato_yeppon/qta_yeppon
					margine_yeppon=cdbl((fatturato_yeppon-acquisto_yeppon)/fatturato_yeppon)*100
				end if
				if fatturato_coinvest= 0 or qta_coinvest=0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=fatturato_coinvest/qta_coinvest
					margine_coinvest=cdbl((fatturato_coinvest-acquisto_coinvest)/fatturato_coinvest)*100
				end if
				if fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=fatturato_mktp/qta_mktp
					margine_mktp=cdbl((fatturato_mktp-acquisto_mktp)/fatturato_mktp)*100
				end if
				if fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=fatturato_tot/qta_tot
					margine_tot=cdbl((fatturato_tot-acquisto_tot)/fatturato_tot)*100
				end if
				
				'selspedebay="CALL fnEbaySpeseSpedizione("","","","peso","prezzo")
				
				
				if isnull(tot4) then tot4 = 0

				if bgcolor = "#EEE" then
					bgcolor = "#CCC;"
				else
					bgcolor = "#EEE"
				end if
				if error = 1 then bgcolor = "red"
				url_prod = getProductUrlComplete(id_p)
				url_prod = replace(url_prod,"https://www.yeppon.it/p-","https://yeppon-uat.spotview.dev/products/")

		%>

			<tr style="background-color:<%=bgcolor%>; height: 20px;">
<%
if codice="0" then
%>
                <td style="vertical-align:middle;" ><a style="text-decoration: none; color: black; " href="<%= url_prod %>" target="_blank"><%= rs4("codice") %></a></td>
				<td style="vertical-align:middle" ><%= rs4("marca") %></td>
				<td style="vertical-align:middle" ><span class="product-name-editable-wrap"><span class="product-name-editable" data-id-p="<%=id_p%>" data-original-name="<%=Server.HTMLEncode(nomeProd & "")%>"><span class="product-name-editable-text"><%=Server.HTMLEncode(nomeProd & "")%></span></span></span></td>
<%
end if
%>
				<td align="left" style=" border-left: 5px solid black; "><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=yeppon"> <%end if%><%=qta_yeppon%></td>
				<td align="left" ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=yeppon"> <%end if%><%=formatnumber(avg_price_yeppon,2)%></td>
				<td align="left" ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=yeppon"> <%end if%><%=formatnumber(fatturato_yeppon,2)%></td>
				<td align="left" style=" border-right: 5px solid black;" ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=yeppon"> <%end if%><%=formatnumber(margine_yeppon,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; " ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=ebaypromo"> <%end if%><%=qta_coinvest%></td>
				<td align="left" ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=ebaypromo"> <%end if%><%=formatnumber(avg_price_coinvest,2)%></td>
				<td align="left" ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=ebaypromo"> <%end if%><%=formatnumber(fatturato_coinvest,2)%></td>
				<td align="left" style=" border-right: 5px solid black; " ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=ebaypromo"> <%end if%><%=formatnumber(margine_coinvest,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; " ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=mktp"> <%end if%><%=qta_mktp%></td>
				<td align="left" ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=mktp"> <%end if%><%=formatnumber(avg_price_mktp,2)%></td>
				<td align="left" ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=mktp"> <%end if%><%=formatnumber(fatturato_mktp,2)%></td>
				<td align="left" style=" border-right: 5px solid black; " ><% if codice<>"0" then %> <a style="text-decoration: none;color:black" href="?token=<%=token%>&dated=<%=dated%>&flag=true&codice=<%=codice%>&visualizza=mktp"> <%end if%><%=formatnumber(margine_mktp,2)%>%</td>

<%
if codice="0" then
%>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_tot%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_tot,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_tot,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold; <%if margine_tot<0 then%>color:red;<%end if%>"><%=formatnumber(margine_tot,2)%>%</td>
<%
end if
%>

			</tr>
		<%
			

			rs4.movenext
			
			iCount = iCount + 1
			if iCount = 10 then
				Response.Flush
				iCount = 0
			end if
			
			loop
			rs4.close

		
		%>

</table>
</div>
<%
end if
	if codice<>"0" then 
%>
<div id="trend-storico" name="trend-storico"></div>
<%
	if flag then
%>
<canvas id="myLineChart" width="100%" height="30"></canvas>


<%
sql="SELECT MONTH(o.data) AS mese,o.data,date_format(o.data,'%M') as mese_string,  "
if visualizza="yeppon" then
	sql=sql & "ROUND(AVG(c.prezzo/(c.iva/100+1)),2) "
elseif visualizza="ebaypromo" then
	sql=sql & "(ROUND(AVG((c.prezzo-ifnull(e.contributo,0))/(c.iva/100+1)),2)-ROUND(AVG(IF(o.importo_spe=0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)/1.22),2))+IF(pixmania=0,0,IF(ISNULL(e.codice),0,ifnull(e.contributo,0))) "
elseif visualizza="mktp" then
	sql=sql & "(ROUND(AVG(c.prezzo/(c.iva/100+1)),2)-ROUND(AVG(IF(o.importo_spe<>0,0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)/1.22)),2)) " 
end if
sql=sql & " AS tot_tot, "
if visualizza="yeppon" then
	sql=sql & "SUM(IF(pixmania=0,c.quantita,0)) "
elseif visualizza="ebaypromo" then
	sql=sql & "SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))) "
elseif visualizza="mktp" then
	sql=sql & "SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))) " 
end if
sql=sql & " AS qta_tot " &_
	"FROM " &_
		"prodotti p " &_
		"INNER JOIN ordini_carrello c ON p.id_p = c.id_p  " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"INNER JOIN ordini_cliente o ON c.id_ordine = o.id  " &_
		"INNER JOIN v_category_name vcat ON vcat.id_p=p.id_p  " &_
		"LEFT JOIN (select distinct codice,dataInizio,dataFine,contributo FROM ebay_offerte) AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 AND e.site = CASE WHEN o.pixmania = 3 THEN 'it' WHEN o.pixmania = 23 THEN 'de' WHEN o.pixmania = 24 THEN 'fr'  WHEN o.pixmania = 25 THEN 'es'  WHEN o.pixmania = 26 THEN 'uk' ELSE 'it' END  " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3  " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2  " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1  " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"LEFT JOIN ordini_fattura AS `of` ON o.id=of.id_o " &_
		"LEFT JOIN geo_province AS gp ON gp.id_geo_province=of.provincia " &_
		"LEFT JOIN geo_region AS gr ON gr.id_geo_region=gp.geo_region_id " &_
	"WHERE  " &_
		"o.data >='"&datainnum(Datadal)&"' AND o.data <='"&datainnum(DataAl)&"'  AND ((o.step4_status = 1 OR (o.id_spe=111 AND o.step4_status <> 1)) AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) AND pixmania<>10 AND c_annullato = 0 AND c.id_p <> 88526 " &_
		"AND p.codice='"&codice&"' " &_
	"GROUP BY  " &_
		"mese,o.data " &_
	"HAVING  " &_
		"qta_tot>0 " &_
	"ORDER BY  " &_
		"mese,o.data "
	'response.write sql
	'response.end
	'AND vcat.category IN ("&fixed_category&")
set rs = Conn.execute(sql)
vavgprice=""
vqta=""
vdata=""""
max=0
i=1
if not rs.EOF then
%>
<table width="100%">
<tr>
	<td style="vertical-align: top">
	<table width="100%">
		<tr>
			<th colspan="3">Q1</th>
		</tr>
<%
	meseiniz=""
	c=0
	while not rs.EOF and c=0
	if rs("mese")<=3 then
		if meseiniz<>rs("mese") then 
%>
		<tr>
			<th colspan="3"><%=rs("mese_string")%></th>
		</tr>
		<tr>
			<th>Day</th>
			<th>Avg Price</th>
			<th>Qta</th>
		</tr>
		
<%		
		meseiniz=rs("mese")
		end if
		vavgprice=vavgprice & replace(rs("tot_tot"),",",".") & ","
		vqta=vqta & replace(cint(rs("qta_tot")),",",".") & ","
		if cint(rs("qta_tot"))>max then
			max=cint(rs("qta_tot"))
		end if
		vdata=vdata & numindata(rs("data")) & """" & "," & """"
%>
		<tr>
			<td><%=numindata(rs("data"))%></td>
			<th><%=rs("tot_tot")%></th>
			<th><%=rs("qta_tot")%></th>
		</tr>
<%
	rs.movenext
	else
	c=1
	end if
		
	wend

%>
	</table>
	</td>
	<td style="vertical-align: top">
	<table width="100%">
		<tr>
			<th colspan="3">Q2</th>
		</tr>
<%
	meseiniz=""
	c=0
	while not rs.EOF and c=0
	if rs("mese")>=4 and rs("mese")<=6 then
		if meseiniz<>rs("mese") then 
%>
		<tr>
			<th colspan="3"><%=rs("mese_string")%></th>
		</tr>
		<tr>
			<th>Day</th>
			<th>Avg Price</th>
			<th>Qta</th>
		</tr>
		
<%		
		meseiniz=rs("mese")
		end if
		vavgprice=vavgprice & replace(rs("tot_tot"),",",".") & ","
		vqta=vqta & replace(cint(rs("qta_tot")),",",".") & ","
		if cint(rs("qta_tot"))>max then
			max=cint(rs("qta_tot"))
		end if
		vdata=vdata & numindata(rs("data")) & """" & "," & """"
%>
		<tr>
			<td><%=numindata(rs("data"))%></td>
			<th><%=rs("tot_tot")%></th>
			<th><%=rs("qta_tot")%></th>
		</tr>
<%
		rs.movenext
	else
		c=1
	end if
	wend

%>
	</table>
	</td>
	<td style="vertical-align: top">
	<table width="100%">
		<tr>
			<th colspan="3">Q3</th>
		</tr>
<%
	meseiniz=""
	c=0
	while not rs.EOF and c=0
	if rs("mese")>=7 and rs("mese")<=9 then
		if meseiniz<>rs("mese") then 
%>
		<tr>
			<th colspan="3"><%=rs("mese_string")%></th>
		</tr>
		<tr>
			<th>Day</th>
			<th>Avg Price</th>
			<th>Qta</th>
		</tr>
		
<%		
		meseiniz=rs("mese")
		end if
		vavgprice=vavgprice & replace(rs("tot_tot"),",",".") & ","
		vqta=vqta & replace(cint(rs("qta_tot")),",",".") & ","
		if cint(rs("qta_tot"))>max then
			max=cint(rs("qta_tot"))
		end if
		vdata=vdata & numindata(rs("data")) & """" & "," & """"
%>
		<tr>
			<td><%=numindata(rs("data"))%></td>
			<th><%=rs("tot_tot")%></th>
			<th><%=rs("qta_tot")%></th>
		</tr>
<%
		rs.movenext
	else
		c=1
	end if
	wend

%>
	</table>
	</td>
	<td style="vertical-align: top">
	<table width="100%">
		<tr>
			<th colspan="3">Q4</th>
		</tr>
<%
	meseiniz=""
	c=0
	while not rs.EOF and c=0
	if rs("mese")>=9 and rs("mese")<=12 then
		if meseiniz<>rs("mese") then 
%>
		<tr>
			<th colspan="3"><%=rs("mese_string")%></th>
		</tr>
		<tr>
			<th>Day</th>
			<th>Avg Price</th>
			<th>Qta</th>
		</tr>
		
<%		
		meseiniz=rs("mese")
		end if
		vavgprice=vavgprice & replace(rs("tot_tot"),",",".") & ","
		vqta=vqta & replace(cint(rs("qta_tot")),",",".") & ","
		if cint(rs("qta_tot"))>max then
			max=cint(rs("qta_tot"))
		end if
		vdata=vdata & numindata(rs("data")) & """" & "," & """"
%>
		<tr>
			<td><%=numindata(rs("data"))%></td>
			<th><%=rs("tot_tot")%></th>
			<th><%=rs("qta_tot")%></th>
		</tr>
<%
		rs.movenext
	else
		c=1
	end if
	wend
	vavgprice=left(vavgprice,len(vavgprice)-1)
	vqta=left(vqta,len(vqta)-1)
	vdata=left(vdata,len(vdata)-2)
%>
	</table>
	</td>
</tr>
</table>
<%
	end if
%>
<script>
// Definisco i dati da mostrare nel grafico
var data = {
	labels: [<%=vdata%>],
	datasets: [
		{
			label: "Avg Price",
			fillColor: "rgba(0,0,0)",
			strokeColor: "rgba(0,0,0)",
			pointColor: "rgba(0,0,0)",
			pointStrokeColor: "rgba(0,0,0)",
			pointHighlightFill: "rgba(0,0,0)",
			pointHighlightStroke: "rgba(0,0,0)",
			data: [<%=vavgprice%>]
		}
	]
};
// Ottengo il contesto 2D del Canvas in cui mostrare il grafico
var ctx = document.getElementById("myLineChart").getContext("2d");

// Crea il grafico e visualizza i dati
var myNewChart = new Chart(ctx , {
    type: "line",
    data: {
			labels: [<%=vdata%>],
			datasets: [{ 
					data: [<%=vavgprice%>],
					yAxisID: 'A',
					label: "Avg Price",
					borderColor: "#3e95cd",
					fill: false
				  },
				  {
					data: [<%=vqta%>],
					yAxisID: 'B',
					label: "Qta",
					borderColor: "#008000",
					fill: false,
					showLine: false
				  }
			]			  
	},
	options: {
    "hover": {
      "animationDuration": 0
    },
    "animation": {
      "duration": 1,
      "onComplete": function() {
        var chartInstance = this.chart,
          ctx = chartInstance.ctx;

        ctx.font = Chart.helpers.fontString(Chart.defaults.global.defaultFontSize, Chart.defaults.global.defaultFontStyle, Chart.defaults.global.defaultFontFamily);
        ctx.textAlign = 'center';
        ctx.textBaseline = 'bottom';
		var sopra=5
        this.data.datasets.forEach(function(dataset, i) {
          var meta = chartInstance.controller.getDatasetMeta(i);
          meta.data.forEach(function(bar, index) {
            var data = dataset.data[index];
			if (sopra>0)
			{
				sopra=(-15);
			}
			else
			{
				sopra=5;
			}
			ctx.fillStyle = "#000000";
            ctx.fillText(data, bar._model.x, bar._model.y - sopra);
          });
        });
      }
    },
    legend: {
      "display": true
    },
    tooltips: {
      "enabled": false
    },
    scales: {
      yAxes: [{
        id: 'A',
        type: 'linear',
        position: 'left',
        display: true,
        gridLines: {
          display: false
			},
		ticks: {
          max: Math.max(...data.datasets[0].data) + 10,
          display: true,
          beginAtZero: false
			}
		},
		{
        id: 'B',
        type: 'linear',
        position: 'right',
        display: true,
        gridLines: {
          display: false
			},
		ticks: {
          max: <%=replace(max*1.2,",",".")%>,
		  min: 0,
          display: true,
          beginAtZero: false
        }
		}
      ],
      xAxes: [{
        gridLines: {
          display: false
        },
        ticks: {
          beginAtZero: true
        }
      }]
    }
  }
});
</script>
<%	
	
	end if			
end if



%>
<div><a href="xls_actual_turnover.asp?date-choose=<%=datechoose%>&categoria=<%=categoria%>&cat2=<%=cat2%>&marca=<%=marca%>&codice=<%=codice%> ">esporta excel</div>

<script src="/admin/js/jquery.stickytableheaders.min.js" type="text/javascript"></script>
<script type="text/javascript">
 $(function () {
 	$("#prodotti_table").stickyTableHeaders();
	
	$.tablesorter.addParser({
		id: "colcpl",
		is: function(s) {
		  return /^[0-9]?[0-9, \.]*$/.test(s);
		},
		format: function(s) {
		  return jQuery.tablesorter.formatFloat(s.replace(/\./g, '').replace(/,/g, '.'));
		},
		type: "numeric"
	  });
	
	$("#category_table").tablesorter({
		widgets: 'zebra',
		headers: {
		2: { sorter: 'colcpl' },
		3: { sorter: 'colcpl' },
		4: { sorter: 'colcpl' },
		5: { sorter: 'colcpl' },
		6: { sorter: 'colcpl' },
		7: { sorter: 'colcpl' },
		8: { sorter: 'colcpl' },
		9: { sorter: 'colcpl' }
	}
	});
	
	$("#marca_table").tablesorter({
		widgets: 'zebra'
	});
	$("#prodotti_table").tablesorter({
		widgets: 'zebra'
	});
	
	
  });
</script>
<script type="text/javascript">  
     function submitform(){
		document.getElementById("myform").submit();
    }
	
	function showtrendhour(date) {
		
		
		$.ajax( {
			type: "GET",
			url: "/admin/ajax/get-trend-mensile.asp?data="+encodeURIComponent(date),
			success: function(response) {
				$("#trend-hour").html(response);
			},
			error: function(xhr, textStatus, errorThrown) {
				alert("Errore caricamento: " + (xhr.status ? xhr.status + " " + xhr.statusText : textStatus));
			}
		});
		
	};

	function showtrend(week,visualizza,div) {
		
		
		$.ajax( {
			type: "GET",
			url: "/admin/ajax/get-trend-mensile.asp?week="+ encodeURIComponent(week) + "&visualizza="+ encodeURIComponent(visualizza),
			success: function(response) {
				$("#"+div).html(response);
			},
			error: function(xhr, textStatus, errorThrown) {
				alert("Errore caricamento: " + (xhr.status ? xhr.status + " " + xhr.statusText : textStatus));
			}
		});
		
	};
</script>
<script type="text/javascript">
 $( document ).ready(function() {
 

	$(".show-trend-hour").click(function() {
		
		
		$.ajax( {
			type: "GET",
			url: "/admin/ajax/get-trend-mensile.asp",
			success: function(response) {
				$("#trend-mensile").html(response);
			},
			error: function(xhr, textStatus, errorThrown) {
				alert("Errore caricamento: " + (xhr.status ? xhr.status + " " + xhr.statusText : textStatus));
			}
		});
		
	});


	
	$(".show-trend").click(function() {
		
		var week=$(this).attr("week-data");
		var visualizza=$(this).attr("visualizza-data");
		var div=$(this).attr("div-data");
		
		$.ajax( {
			type: "GET",
			url: "/admin/ajax/get-trend-mensile.asp?week="+ encodeURIComponent(week) + "&visualizza="+ encodeURIComponent(visualizza),
			success: function(response) {
				$("#"+div).html(response);
			},
			error: function(xhr, textStatus, errorThrown) {
				alert("Errore caricamento: " + (xhr.status ? xhr.status + " " + xhr.statusText : textStatus));
			}
		});
		
	}); 

	$(".detmktp").click(function() {
		
		var Datadal=$(this).attr("Datadal-data");
		var DataAl=$(this).attr("DataAl-data");
		var Marketplace=$(this).attr("Marketplace-data") || "";
		var url = "/admin/ajax/get-trend-mensile.asp?detmktp=si&Datadal="+ encodeURIComponent(Datadal)+ "&DataAl="+ encodeURIComponent(DataAl);
		if (Marketplace) url += "&Marketplace="+ encodeURIComponent(Marketplace);
		
		$.ajax( {
			type: "GET",
			url: url,
			success: function(response) {
				$("#det_mktp").html(response);
			},
			error: function(xhr, textStatus, errorThrown) {
				alert("Errore caricamento: " + (xhr.status ? xhr.status + " " + xhr.statusText : textStatus));
			}
		});
		
	});
	
	
	$(".show-storico").click(function() {
		
		var Datadal=$(this).attr("Datadal-data");
		var DataAl=$(this).attr("DataAl-data");
		var visualizza=$(this).attr("visualizza-data");
		var codice=$(this).attr("codice-data");
		
		$.ajax( {
			type: "GET",
			url: "/admin/ajax/get-trend-mensile.asp?codice=" + encodeURIComponent(codice) + "&visualizza=" + encodeURIComponent(visualizza) + "&DataAl=" + encodeURIComponent(DataAl) + "&Datadal=" + encodeURIComponent(Datadal),
			success: function(response) {
				$("#trend-storico").html(response);
			},
			error: function(xhr, textStatus, errorThrown) {
				alert("Errore caricamento: " + (xhr.status ? xhr.status + " " + xhr.statusText : textStatus));
			}
		});
		
	});
	
  });
</script>

