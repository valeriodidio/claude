<!--#include virtual="/int/config.asp"-->
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/2.4.0/Chart.min.js"></script>

<%
		Server.ScriptTimeout = 50000000
		
		visualizza=request("visualizza")
		week=request("week")
		codice=request("codice")
		Datadal=request("Datadal")
		DataAl=request("DataAl")
		detmktp=request("detmktp")
		data=request("data")
		daydiff=request("daydiff")
		if week<>"" and visualizza="week" then

		Sel3 = "SELECT   " &_
		"DATE_FORMAT(o.data,'%a') as mese, o.data as data, DATE_FORMAT(o.data,'%v') as week, IF(DATE_FORMAT(o.data,'%a')='mon',concat(right(o.data,2),'/',mid(o.data,5,2)),0) as data_lun, IF(DATE_FORMAT(o.data,'%a')='mon',concat(DAY(SUBDATE(o.data,interval -1 DAY)),'/',MONTH(SUBDATE(o.data,interval -1 DAY))),0) as data_mar, IF(DATE_FORMAT(o.data,'%a')='mon',concat(DAY(SUBDATE(o.data,interval -2 DAY)),'/',MONTH(SUBDATE(o.data,interval -2 DAY))),0) as data_mer, IF(DATE_FORMAT(o.data,'%a')='mon',concat(DAY(SUBDATE(o.data,interval -3 DAY)),'/',MONTH(SUBDATE(o.data,interval -3 DAY))),0) as data_gio, IF(DATE_FORMAT(o.data,'%a')='mon',concat(DAY(SUBDATE(o.data,interval -4 DAY)),'/',MONTH(SUBDATE(o.data,interval -4 DAY))),0) as data_ven, IF(DATE_FORMAT(o.data,'%a')='mon',concat(DAY(SUBDATE(o.data,interval -5 DAY)),'/',MONTH(SUBDATE(o.data,interval -5 DAY))),0) as data_sab, IF(DATE_FORMAT(o.data,'%a')='mon',concat(DAY(SUBDATE(o.data,interval -6 DAY)),'/',MONTH(SUBDATE(o.data,interval -6 DAY))),0) as data_dom,ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon , (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest,ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=30,c.quantita,0)),2) AS tot_sd, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania in (0,30),0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania in (0,30),0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp,  ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) as tot_tot " &_
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"where DATE_FORMAT(o.data,'%v') ="&week&" AND year(o.data)=year(current_date()) AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 and o.pixmania <> 10 "
          
       
		if request("Marketplace") <> "" then
			Sel3 = Sel3 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      sel3 = sel3 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Sel3 = Sel3 & "group by DATE_FORMAT(o.data,'%a') "
		
		Sel3 = Sel3 & "ORDER BY IF(DATE_FORMAT(o.data,'%w')=0,8,DATE_FORMAT(o.data,'%w')) "
		
		'response.write sel3
		'response.end
		
		set rsmonth = Conn.execute(Sel3)
		
		

%>
<table>
	<thead>
		<tr>
			<th width="9%"><a onclick="showtrend('<%=week-1%>','week','trend-week','')" ><-</a><span>Week <%=week%></span><a onclick="showtrend('<%=week+1%>','week','trend-week','')">-></a></th>
			<th width="13%">Lunedi&nbsp<%=rsmonth("data_lun")%></th>
			<th width="13%">Martedi&nbsp<%=rsmonth("data_mar")%></th>
			<th width="13%">Mercoledi&nbsp<%=rsmonth("data_mer")%></th>
			<th width="13%">Giovedi&nbsp<%=rsmonth("data_gio")%></th>
			<th width="13%">Venerdi&nbsp<%=rsmonth("data_ven")%></th>
			<th width="13%">Sabato&nbsp<%=rsmonth("data_sab")%></th>
			<th width="13%">Domenica&nbsp<%=rsmonth("data_dom")%></th>
			<th width="13%">Totale&nbsp</th>
		</tr>
	</thead>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #5cd65c;">Yeppon</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

            fatturato_yeppon        = rsmonth("tot_yeppon")
            


%>
				<td align="left" ><%=formatnumber(fatturato_yeppon,2)%></td>

<%
	rsmonth.movenext
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #1aa3ff;">Promo Ebay</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

            fatturato_coinvest        = rsmonth("tot_coinvest")
            


%>
				<td align="left" ><%=formatnumber(fatturato_coinvest,2)%></td>
<%
	rsmonth.movenext
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #ffcc66;">MKTP</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

            fatturato_mktp       = rsmonth("tot_mktp")



%>
				<td align="left" ><%=formatnumber(fatturato_mktp,2)%></td>

<%
	rsmonth.movenext
wend
%>
</tr>

	<tr>
	<td style=" border-left: 5px solid black;background-color: #1aa3ff;">SuperDisty Amzn</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

            fatturato_sd        = rsmonth("tot_sd")
            


%>
				<td align="left" ><%=formatnumber(fatturato_sd,2)%></td>
<%
	rsmonth.movenext
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #9966ff;">Totale</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

            fatturato_tot       = rsmonth("tot_tot")			
			


%>
				<td align="left" style=" font-weight:bold;" > <%=formatnumber(fatturato_tot,2)%></td>
<%
	rsmonth.movenext
wend
%>
</tr>
</table>
        	
<%
elseif week<>"" and visualizza="hour" then
Sel3 = "SELECT   " &_
		"IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) as mese, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon , (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp,  ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) as tot_tot " &_ 
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"where DATE_FORMAT(o.data,'%v') ="&week&" AND year(o.data)=year(current_date()) AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
          
       
		if request("Marketplace") <> "" then
			Sel3 = Sel3 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      sel3 = sel3 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Sel3 = Sel3 & "group by IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) "
		
		Sel3 = Sel3 & "ORDER BY CAST(IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) AS DECIMAL(2,0)) "
		
		'response.write sel3
		'response.end
		
		set rsmonth = Conn.execute(Sel3)

%>
<div>Actual Turnover Orario Week: <%=week%></div>
<table>
	<thead>

<%
app=0
c=0
for i=12 to 24 step 12
%>
		<tr>
			
<% 
if c=0 then 
	c=1
%>
	<th width="16%"><a onclick="showtrend('<%=week-1%>','hour','trend-hour','')" ><-</a><span>Week <%=week%></span><a onclick="showtrend('<%=week+1%>','hour','trend-hour','')">-></a></th>
<%
else
%>
	<th width="16%"></th>
<%
end if
j=1
rsmonth.movefirst
while not rsmonth.EOF
	if j>app and j<=i then
%>
	
			<th width="7%" ><%=rsmonth("mese")%>-<%=rsmonth("mese")+1%></th>
<%
	end if
	rsmonth.movenext
	j=j+1
wend
%>
		</tr>

	</thead>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #5cd65c;">Yeppon</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF 
			if j>app and j<=i then

            fatturato_yeppon        = rsmonth("tot_yeppon")


%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_yeppon,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #1aa3ff;">Promo Ebay</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF
			if j>app and j<=i then

            fatturato_coinvest        = rsmonth("tot_coinvest")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_coinvest,2)%></td>
<%
	end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #ffcc66;">MKTP</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF 
			if j>app and j<=i then

            fatturato_mktp       = rsmonth("tot_mktp")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_mktp,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #9966ff;">Totale</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF
			if j>app and j<=i then

            fatturato_tot       = rsmonth("tot_tot")			



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_tot,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
app=i
next
%>
</tr>
</table>	


<%
elseif daydiff<>"" and visualizza="hour-vs-day" then
Selfirstday = "SELECT   " &_
		"IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) as ora, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon , (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp,  ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) as tot_tot,date_format(o.data,'%d/%m/%Y') as dayformat " &_ 
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"where o.data = SUBDATE(CURRENT_DATE, INTERVAL "&daydiff&" DAY) AND year(o.data)=year(current_date()) AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
          
       
		if request("Marketplace") <> "" then
			Selfirstday = Selfirstday &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      Selfirstday = Selfirstday &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Selfirstday = Selfirstday & "group by IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) "
		
		Selfirstday = Selfirstday & "ORDER BY CAST(IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) AS DECIMAL(2,0)) "
		
		'response.write sel3
		'response.end
		
		set rsfirstday = Conn.execute(Selfirstday)
		
Selyesterday = "SELECT   " &_
		"IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) as ora, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon , (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp,  ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) as tot_tot " &_ 
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"where o.data = SUBDATE(CURRENT_DATE, INTERVAL "&daydiff+1&" DAY) AND year(o.data)=year(current_date()) AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
          
       
		if request("Marketplace") <> "" then
			Selyesterday = Selyesterday &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      Selyesterday = Selyesterday &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Selyesterday = Selyesterday & "group by IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) "
		
		Selyesterday = Selyesterday & "ORDER BY CAST(IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) AS DECIMAL(2,0)) "
		
		'response.write Selyesterday & "<br>"
		'response.end
		
		set rsyesterday = Conn.execute(Selyesterday)

Sellastweekday = "SELECT   " &_
		"IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) as ora, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon , (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp,  ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) as tot_tot " &_ 
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"where o.data = SUBDATE(CURRENT_DATE, INTERVAL "&daydiff+7&" DAY) AND year(o.data)=year(current_date()) AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
          
       
		if request("Marketplace") <> "" then
			Sellastweekday = Sellastweekday &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      Sellastweekday = Sellastweekday &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Sellastweekday = Sellastweekday & "group by IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) "
		
		Sellastweekday = Sellastweekday & "ORDER BY CAST(IF(LENGTH(o.ora)=6,LEFT(o.ora,2)-1,IF(LENGTH(o.ora)=5,LEFT(o.ora,1)-1,0)) AS DECIMAL(2,0)) "
		
		'response.write Sellastweekday
		'response.end
		
		set rslastweekday = Conn.execute(Sellastweekday)

%>
<div>Actual Turnover Orario Week: <%=rsfirstday("dayformat")%></div>
<table id = "hoursvs" >
	<thead>

<%
app=0
c=0
for i=6 to 24 step 6
%>
		<tr>
			
<% 
if c=0 then 
	c=1
%>
	<th width="12%"><a onclick="showtrend('','hour-vs-day','trend-hour-vs-day','<%=daydiff+1%>')" ><-</a><span><%=rsfirstday("dayformat")%></span><a onclick="showtrend('','hour-vs-day','trend-hour-vs-day','<%=daydiff+1%>')">-></a></th>
<%
else
%>
	<th width="12%"></th>
<%
end if
j=1
rsfirstday.movefirst
while not rsfirstday.EOF
	if j>app and j<=i then
%>
	
			<th width="12%" colspan="3"><%=rsfirstday("ora")%>-<%=rsfirstday("ora")+1%></th>
<%
	end if
	rsfirstday.movenext
	j=j+1
wend
%>
		</tr>	
		<tr>
			
	<th width="12%"></th>

<%
j=1
rsfirstday.movefirst
while not rsfirstday.EOF
	if j>app and j<=i then
%>
	
			<th width="4%" >today</th>
			<th width="4%" >yesterday</th>
			<th width="4%" >7 day ago</th>
<%
	end if
	rsfirstday.movenext
	j=j+1
wend
%>
		</tr>
	</thead>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #5cd65c;">Yeppon</td>
<%
rsfirstday.movefirst
rsyesterday.movefirst
rslastweekday.movefirst
j=1

while not rsfirstday.EOF 
			if j>app and j<=i then

            fatturato_yeppon        = rsfirstday("tot_yeppon")
			fatturato_yeppony       = rsyesterday("tot_yeppon")
			fatturato_yeppon7       = rslastweekday("tot_yeppon")


%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_yeppon,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_yeppony,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_yeppon7,2)%></td>
<%
			end if
	rsfirstday.movenext
	rsyesterday.movenext
	rslastweekday.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #1aa3ff;">Promo Ebay</td>
<%
rsfirstday.movefirst
rsyesterday.movefirst
rslastweekday.movefirst
j=1
while not rsfirstday.EOF
			if j>app and j<=i then

            fatturato_coinvest        = rsfirstday("tot_coinvest")
			fatturato_coinvesty       = rsyesterday("tot_coinvest")
			fatturato_coinvest7       = rslastweekday("tot_coinvest")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_coinvest,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_coinvesty,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_coinvest7,2)%></td>
<%
	end if
	rsfirstday.movenext
	rsyesterday.movenext
	rslastweekday.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #ffcc66;">MKTP</td>
<%
rsfirstday.movefirst
rsyesterday.movefirst
rslastweekday.movefirst
j=1
while not rsfirstday.EOF 
			if j>app and j<=i then

            fatturato_mktp       = rsfirstday("tot_mktp")
			fatturato_mktpy      = rsyesterday("tot_mktp")
			fatturato_mktp7      = rslastweekday("tot_mktp")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_mktp,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_mktpy,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_mktp7,2)%></td>
<%
			end if
	rsfirstday.movenext
	rsyesterday.movenext
	rslastweekday.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #9966ff;">Totale</td>
<%
rsfirstday.movefirst
rsyesterday.movefirst
rslastweekday.movefirst
j=1
while not rsfirstday.EOF
			if j>app and j<=i then

            fatturato_tot       = rsfirstday("tot_tot")	
			fatturato_toty      = rsyesterday("tot_tot")
			fatturato_tot7      = rslastweekday("tot_tot")			



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_tot,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_toty,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_tot7,2)%></td>
<%
			end if
	rsfirstday.movenext
	rsyesterday.movenext
	rslastweekday.movenext
	j=j+1
wend
app=i
next
%>
</tr>
</table>	
<p><a style="text-decoration:none;color:#000;background-color:#ddd;border:1px solid #ccc;padding:8px;" onclick="XlsExport('hoursvs')">Exporta Tabella in Excel</a></p>


<%
elseif week<>"" and visualizza="region" then
Sel3 = "SELECT " &_
			"IFNULL(gr.internal_name,'ESTERO') AS mese,  " &_
			"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon ,  " &_
			"(ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest,  " &_
			"(ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp,  " &_
			"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) AS tot_tot  " &_
		"FROM  " &_
			"prodotti p  " &_
			"INNER JOIN ordini_carrello c ON p.id_p = c.id_p  " &_
			"inner join fornitori forn using(id_forni_nav) " &_
			"INNER JOIN ordini_cliente o ON c.id_ordine = o.id  " &_
			"INNER JOIN v_category_name vcat ON vcat.id_p=p.id_p  " &_
			"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0  " &_
			"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3  " &_
			"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2  " &_
			"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1  " &_
			"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
			"LEFT JOIN ordini_fattura AS `of` ON o.id=of.id_o " &_
			"LEFT JOIN geo_province AS gp ON gp.id_geo_province=of.provincia " &_
			"LEFT JOIN geo_region AS gr ON gr.id_geo_region=gp.geo_region_id " &_
		"where DATE_FORMAT(o.data,'%v') ="&week&" AND year(o.data)=year(current_date()) AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
      
		if request("Marketplace") <> "" then
			Sel3 = Sel3 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      sel3 = sel3 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Sel3 = Sel3 & "group by IFNULL(gr.id_geo_region,'ESTERO') "
		
		Sel3 = Sel3 & "ORDER BY tot_tot DESC "
		
		'response.write sel3
		'response.end
		
		set rsmonth = Conn.execute(Sel3)

%>
<div>Actual Turnover Per Regione Week: <%=week%></div>
<table>
	<thead>

<%
app=0
c=0
for i=10 to 20 step 10
%>
		<tr>
<% 
if c=0 then 
	c=1
%>
	<th width="16%"><a onclick="showtrend('<%=week-1%>','region','trend-region','')" ><-</a><span>Week <%=week%></span><a onclick="showtrend('<%=week+1%>','region','trend-region','')">-></a></th>
<%
else
%>
	<th width="16%"></th>
<%
end if
j=1
rsmonth.movefirst
while not rsmonth.EOF
	if j>app and j<=i then
%>
	
			<th width="7%" ><%=rsmonth("mese")%></th>
<%
	end if
	rsmonth.movenext
	j=j+1
wend
%>
		</tr>

	</thead>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #5cd65c;">Yeppon</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF 
			if j>app and j<=i then

            fatturato_yeppon        = rsmonth("tot_yeppon")


%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_yeppon,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #1aa3ff;">Promo Ebay</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF
			if j>app and j<=i then

            fatturato_coinvest        = rsmonth("tot_coinvest")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_coinvest,2)%></td>
<%
	end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #ffcc66;">MKTP</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF 
			if j>app and j<=i then

            fatturato_mktp       = rsmonth("tot_mktp")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_mktp,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #9966ff;">Totale</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF
			if j>app and j<=i then

            fatturato_tot       = rsmonth("tot_tot")			



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_tot,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
app=i
next
%>
</tr>
</table>	


<%

elseif week<>"" and visualizza="province" then
Sel3 = "SELECT " &_
			"IFNULL(gp.internal_name,'ESTERO') AS mese,  " &_
			"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon ,  " &_
			"(ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest,  " &_
			"(ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp,  " &_
			"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) AS tot_tot  " &_
		"FROM  " &_
			"prodotti p  " &_
			"INNER JOIN ordini_carrello c ON p.id_p = c.id_p  " &_
			"inner join fornitori forn using(id_forni_nav) " &_
			"INNER JOIN ordini_cliente o ON c.id_ordine = o.id  " &_
			"INNER JOIN v_category_name vcat ON vcat.id_p=p.id_p  " &_
			"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0  " &_
			"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3  " &_
			"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2  " &_
			"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1  " &_
			"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
			"LEFT JOIN ordini_fattura AS `of` ON o.id=of.id_o " &_
			"LEFT JOIN geo_province AS gp ON gp.id_geo_province=of.provincia " &_
			"LEFT JOIN geo_region AS gr ON gr.id_geo_region=gp.geo_region_id " &_
		"where DATE_FORMAT(o.data,'%v') ="&week&" AND year(o.data)=year(current_date()) AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
      
		if request("Marketplace") <> "" then
			Sel3 = Sel3 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      sel3 = sel3 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Sel3 = Sel3 & "group by IFNULL(gp.id_geo_province,'ESTERO') "
		
		Sel3 = Sel3 & "ORDER BY tot_tot DESC LIMIT 30"
		
		'response.write sel3
		'response.end
		
		set rsmonth = Conn.execute(Sel3)

%>
<div>Actual Turnover Per Provincia Week: <%=week%></div>
<table>
	<thead>

<%
app=0
c=0
for i=10 to 30 step 10
%>
		<tr>
<% 
if c=0 then 
	c=1
%>
	<th width="16%"><a onclick="showtrend('<%=week-1%>','province','trend-province','')" ><-</a><span>Week <%=week%></span><a onclick="showtrend('<%=week+1%>','province','trend-province','')">-></a></th>
<%
else
%>
	<th width="16%"></th>
<%
end if
j=1
rsmonth.movefirst
while not rsmonth.EOF
	if j>app and j<=i then
%>
	
			<th width="7%" ><%=rsmonth("mese")%></th>
<%
	end if
	rsmonth.movenext
	j=j+1
wend
%>
		</tr>
	</thead>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #5cd65c;">Yeppon</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF 
			if j>app and j<=i then

            fatturato_yeppon        = rsmonth("tot_yeppon")


%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_yeppon,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #1aa3ff;">Promo Ebay</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF
			if j>app and j<=i then

            fatturato_coinvest        = rsmonth("tot_coinvest")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_coinvest,2)%></td>
<%
	end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #ffcc66;">MKTP</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF 
			if j>app and j<=i then

            fatturato_mktp       = rsmonth("tot_mktp")



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_mktp,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #9966ff;">Totale</td>
<%
rsmonth.movefirst
j=1
while not rsmonth.EOF
			if j>app and j<=i then

            fatturato_tot       = rsmonth("tot_tot")			



%>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(fatturato_tot,2)%></td>
<%
			end if
	rsmonth.movenext
	j=j+1
wend
app=i
next
%>
</tr>
</table>	

<%

elseif data="mensile" then 

Sel3 = "SELECT   " &_
		"DATE_FORMAT(o.data,'%M') as mese, " &_
		" SUM(IF(pixmania=0,c.quantita,0)) AS qta_yeppon, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) AS tot_yeppon ,SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))) AS qta_coinvest, (ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22) AS tot_coinvest,SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))) AS qta_mktp, (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) AS tot_mktp, (SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0)))) as qta_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0)),2) + ((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id)))))) as tot_tot, ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,c.quantita,0))/SUM(IF(pixmania=0,c.quantita,0)),2) as avg_price_yeppon,ROUND((ROUND(SUM(((c.prezzo)/(c.iva/100+1)+IF(pixmania=3 AND NOT ISNULL(e.codice),IFNULL(e.contributo,0),0))*IF(pixmania=3 AND NOT ISNULL(e.codice),c.quantita,0)),2)-SUM(IF(pixmania=3 AND NOT ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))/1.22)/SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) as avg_price_coinvest, ROUND((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0))))/SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2) as avg_price_mktp, ROUND(((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))) + (ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)-sum(IF(pixmania=0,0,IF(ISNULL(e.codice),0,fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id))))))/ (SUM(IF(pixmania=0,c.quantita,0)) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))) + SUM(IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0)))),2) as avg_price_tot, " &_ 
			"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=0,c.quantita,0)),2) as acquisto_yeppon, " &_
			"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2) as acquisto_coinvest, " &_
			"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2) as acquisto_mktp, (ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=0,c.quantita,0)),2)+ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=0,0,IF(ISNULL(e.codice),0,c.quantita))),2)+ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=0,0,IF(ISNULL(e.codice),c.quantita,0))),2)) as acquisto_tot " &_
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"LEFT OUTER JOIN fornitori_ordini f on c.id = f.id_car " &_
		"where o.data >='"&left(datainnum(Datadal),4)&"0101' AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
          
       
		if request("Marketplace") <> "" then
			Sel3 = Sel3 &"and o.pixmania = "&request("marketplace")&" "
		end if
		if request("forni") <> "" then
		      sel3 = sel3 &"and c.id_forni_nav = '"&request("forni")&"' "
		end if
		
		Sel3 = Sel3 & "group by MONTH(o.data) "
		
		Sel3 = Sel3 & "ORDER BY MONTH(o.data) "
		
		'response.write sel3
		'response.end
		
		set rsmonth = Conn.execute(Sel3)

%>
<table>
	<thead>
		<tr>
			<th></th>
<%
while not rsmonth.EOF
%>
	
			<th colspan="4"><%=rsmonth("mese")%></th>
<%
	rsmonth.movenext
wend
%>
		</tr>
		<tr>
			<th></th>
<%
rsmonth.movefirst
while not rsmonth.EOF
%>
			<th width="7%" class="header" scope="col" style=" border-left: 5px solid black;">Pezzi Venduti Totali</th>
			<th width="7%" class="header" scope="col" >Prezzo Medio Totale</th>
			<th width="7%" class="header" scope="col" >Fatturato Totale</th>
			<th width="7%" class="header" scope="col"  style=" border-right: 5px solid black;">GGP Totale</th>
<%
	rsmonth.movenext
wend
%>
		</tr>
	</thead>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #5cd65c;">Yeppon</td>
<%
rsmonth.movefirst
while not rsmonth.EOF
			qta_yeppon  = cdbl(rsmonth("qta_yeppon"))
            fatturato_yeppon        = rsmonth("tot_yeppon")
		
			acquisto_yeppon      = rsmonth("acquisto_yeppon")

				if fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=fatturato_yeppon/qta_yeppon
					margine_yeppon=cdbl((fatturato_yeppon-acquisto_yeppon)/fatturato_yeppon)*100
				end if



%>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_yeppon%></td>
				<td align="left" ><%=formatnumber(avg_price_yeppon,2)%></td>
				<td align="left" ><%=formatnumber(fatturato_yeppon,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(margine_yeppon,2)%>%</td>
<%
	rsmonth.movenext
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #1aa3ff;">Promo Ebay</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

            qta_coinvest  = cdbl(rsmonth("qta_coinvest"))
            fatturato_coinvest        = rsmonth("tot_coinvest")

				acquisto_coinvest       = rsmonth("acquisto_coinvest")				


				if fatturato_coinvest= 0 then
					margine_coinvest=0
					avg_price_coinvest=0
				else
					avg_price_coinvest=fatturato_coinvest/qta_coinvest
					margine_coinvest=cdbl((fatturato_coinvest-acquisto_coinvest)/fatturato_coinvest)*100
				end if



%>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_coinvest%></td>
				<td align="left" ><%=formatnumber(avg_price_coinvest,2)%></td>
				<td align="left" ><%=formatnumber(fatturato_coinvest,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(margine_coinvest,2)%>%</td>
<%
	rsmonth.movenext
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #ffcc66;">MKTP</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

            qta_mktp  = cdbl(rsmonth("qta_mktp"))
            fatturato_mktp       = rsmonth("tot_mktp")
			
				acquisto_mktp       = rsmonth("acquisto_mktp")

				if fatturato_mktp= 0 then
					margine_mktp=0
					avg_price_mktp=0
				else
					avg_price_mktp=fatturato_mktp/qta_mktp
					margine_mktp=cdbl((fatturato_mktp-acquisto_mktp)/fatturato_mktp)*100
				end if



%>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_mktp%></td>
				<td align="left" ><%=formatnumber(avg_price_mktp,2)%></td>
				<td align="left" ><%=formatnumber(fatturato_mktp,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(margine_mktp,2)%>%</td>
<%
	rsmonth.movenext
wend
%>
</tr>
	<tr>
	<td style=" border-left: 5px solid black;background-color: #9966ff;">Totale</td>
<%
rsmonth.movefirst
while not rsmonth.EOF

			qta_tot = cdbl(rsmonth("qta_tot"))
            fatturato_tot       = rsmonth("tot_tot")			

				acquisto_tot       = rsmonth("acquisto_tot")

				if fatturato_tot= 0 then
					margine_tot=0
					avg_price_tot=0
				else
					avg_price_tot=fatturato_tot/qta_tot
					margine_tot=cdbl((fatturato_tot-acquisto_tot)/fatturato_tot)*100
				end if


%>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_tot%></td>
				<td align="left" ><%=formatnumber(avg_price_tot,2)%></td>
				<td align="left" ><%=formatnumber(fatturato_tot,2)%></td>
				<td align="left" style=" border-right: 5px solid black; "><%=formatnumber(margine_tot,2)%>%</td>
<%
	rsmonth.movenext
wend
%>
</tr>
</table>


<%
end if

if detmktp<>"" then
%>
<table id="detmktp_table" class="tablesorter" data-filter="#filter" data-filter-text-only="true" data-filter-minimum="2" id="projectSpreadsheet">
	 	<thead>
        <tr bgcolor="#ccc">
			<th width="20%" scope="col">Categoria Prodotto</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #5cd65c;">Pezzi Venduti Amazon</th>
				<th width="4%" class="header" scope="col" style="background-color: #5cd65c;">Prezzo Medio Amazon</th>
				<th width="4%" class="header" scope="col" style="background-color: #5cd65c;">Fatturato Amazon</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #5cd65c;">GGP Amazon</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #ffcc66;">Pezzi Venduti Amazon DE</th>
				<th width="4%" class="header" scope="col" style="background-color: #ffcc66;">Prezzo Medio Amazon DE</th>
				<th width="4%" class="header" scope="col" style="background-color: #ffcc66;">Fatturato Amazon DE</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #ffcc66;">GGP Amazon DE</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #9966ff;">Pezzi Venduti Amazon FR</th>
				<th width="4%" class="header" scope="col" style="background-color: #9966ff;">Prezzo Medio Amazon FR</th>
				<th width="4%" class="header" scope="col" style="background-color: #9966ff;">Fatturato Amazon FR</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #9966ff;">GGP Amazon FR</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: yellow;">Pezzi Venduti Amazon ES</th>
				<th width="4%" class="header" scope="col" style="background-color: yellow;">Prezzo Medio Amazon ES</th>
				<th width="4%" class="header" scope="col" style="background-color: yellow;">Fatturato Amazon ES</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: yellow;">GGP Amazon ES</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #1aa3ff;">Pezzi Venduti Ebay</th>
				<th width="4%" class="header" scope="col" style="background-color: #1aa3ff;">Prezzo Medio Ebay</th>
				<th width="4%" class="header" scope="col" style="background-color: #1aa3ff;">Fatturato Ebay</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #1aa3ff;">GGP Ebay</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #af6901;">Pezzi Venduti Manomano</th>
				<th width="4%" class="header" scope="col" style="background-color: #af6901;">Prezzo Medio Manomano</th>
				<th width="4%" class="header" scope="col" style="background-color: #af6901;">Fatturato Manomano</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #af6901;">GGP Manomano</th>
				<th width="4%" class="header" scope="col" style=" color:white;border-left: 5px solid black;background-color: #293847;">Pezzi Venduti Cdiscount</th>
				<th width="4%" class="header" scope="col" style="color:white;background-color: #293847;">Prezzo Medio Cdiscount</th>
				<th width="4%" class="header" scope="col" style="color:white;background-color: #293847;">Fatturato Cdiscount</th>
				<th width="4%" class="header" scope="col"  style=" color:white;border-right: 5px solid black;background-color: #293847;">GGP Cdiscount</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #6b8e23;">Pezzi Venduti Digitec</th>
				<th width="4%" class="header" scope="col" style="background-color: #6b8e23;">Prezzo Medio Digitec</th>
				<th width="4%" class="header" scope="col" style="background-color: #6b8e23;">Fatturato Digitec</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #6b8e23;">GGP Digitec</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #8b4513;">Pezzi Venduti Digitec DE</th>
				<th width="4%" class="header" scope="col" style="background-color: #8b4513;">Prezzo Medio Digitec DE</th>
				<th width="4%" class="header" scope="col" style="background-color: #8b4513;">Fatturato Digitec DE</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #8b4513;">GGP Digitec DE</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #4682b4;">Pezzi Venduti Stockly</th>
				<th width="4%" class="header" scope="col" style="background-color: #4682b4;">Prezzo Medio Stockly</th>
				<th width="4%" class="header" scope="col" style="background-color: #4682b4;">Fatturato Stockly</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #4682b4;">GGP Stockly</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #228b22;">Pezzi Venduti Leroy Merlin</th>
				<th width="4%" class="header" scope="col" style="background-color: #228b22;">Prezzo Medio Leroy Merlin</th>
				<th width="4%" class="header" scope="col" style="background-color: #228b22;">Fatturato Leroy Merlin</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #228b22;">GGP Leroy Merlin</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #c71585;">Pezzi Venduti Mediamarket</th>
				<th width="4%" class="header" scope="col" style="background-color: #c71585;">Prezzo Medio Mediamarket</th>
				<th width="4%" class="header" scope="col" style="background-color: #c71585;">Fatturato Mediamarket</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #c71585;">GGP Mediamarket</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #e62e05;">Pezzi Venduti Kaufland</th>
				<th width="4%" class="header" scope="col" style="background-color: #e62e05;">Prezzo Medio Kaufland</th>
				<th width="4%" class="header" scope="col" style="background-color: #e62e05;">Fatturato Kaufland</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #e62e05;">GGP Kaufland</th>
				<th width="4%" class="header" scope="col" style=" border-left: 5px solid black;background-color: #ff69b4;">Pezzi Venduti H&A</th>
				<th width="4%" class="header" scope="col" style="background-color: #ff69b4;">Prezzo Medio H&A</th>
				<th width="4%" class="header" scope="col" style="background-color: #ff69b4;">Fatturato H&A</th>
				<th width="4%" class="header" scope="col"  style=" border-right: 5px solid black;background-color: #ff69b4;">GGP H&A</th>
        </tr>
		</thead>	
<%	
        Sel3 = "SELECT   " &_
		"vcat.category as cate3, " &_
		"m.nome as marca, " &_
		" SUM(IF(pixmania=2,c.quantita,0)) AS qta_amazon, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=2,c.quantita,0)),2) AS tot_amazon, " &_
		"SUM(IF(pixmania=7,c.quantita,0)) AS qta_amazon_de, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=7,c.quantita,0)),2) AS tot_amazon_de, " &_
		"SUM(IF(pixmania=8,c.quantita,0)) AS qta_amazon_fr, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=8,c.quantita,0)),2) AS tot_amazon_fr," &_
		"SUM(IF(pixmania=9,c.quantita,0)) AS qta_amazon_es, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=9,c.quantita,0)),2) AS tot_amazon_es, " &_
		"(SUM(IF(pixmania=3 AND ISNULL(e.codice),c.quantita,0))+SUM(IF(pixmania=23,c.quantita,0))+SUM(IF(pixmania=24,c.quantita,0))) AS qta_ebay, " &_
		"((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=3 AND ISNULL(e.codice),c.quantita,0)),2)-sum(IF(pixmania=3 AND ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))+ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=23,c.quantita,0)),2)+ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=24,c.quantita,0)),2)) AS tot_ebay, " &_
		"SUM(IF(pixmania=22,c.quantita,0)) AS qta_manomano, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=22,c.quantita,0)),2) AS tot_manomano, " &_
		"SUM(IF(pixmania=28,c.quantita,0)) AS qta_cdiscount, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=28,c.quantita,0)),2) AS tot_cdiscount, " &_
		"SUM(IF(pixmania=41,c.quantita,0)) AS qta_digitec, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=41,c.quantita,0)),2) AS tot_digitec, " &_
		"SUM(IF(pixmania=57,c.quantita,0)) AS qta_digitec_de, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=57,c.quantita,0)),2) AS tot_digitec_de, " &_
		"SUM(IF(pixmania=52,c.quantita,0)) AS qta_stockly, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=52,c.quantita,0)),2) AS tot_stockly, " &_
		"SUM(IF(pixmania IN (44,50,49,51),c.quantita,0)) AS qta_leroy, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania IN (44,50,49,51),c.quantita,0)),2) AS tot_leroy, " &_
		"SUM(IF(pixmania=53,c.quantita,0)) AS qta_mediamarket, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=53,c.quantita,0)),2) AS tot_mediamarket, " &_
		"SUM(IF(pixmania IN (47,54,55),c.quantita,0)) AS qta_kaufland, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania IN (47,54,55),c.quantita,0)),2) AS tot_kaufland, " &_
		"SUM(IF(o.order_type=3 AND mktp_import.reference_marketplace_id = 1027,c.quantita,0)) AS qta_ha, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(o.order_type=3 AND mktp_import.reference_marketplace_id = 1027,c.quantita,0)),2) AS tot_ha, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=2,c.quantita,0))/SUM(IF(pixmania=2,c.quantita,0)),2) as avg_price_amazon, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=7,c.quantita,0))/SUM(IF(pixmania=7,c.quantita,0)),2) as avg_price_amazon_de, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=8,c.quantita,0))/SUM(IF(pixmania=8,c.quantita,0)),2) as avg_price_amazon_fr," &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=9,c.quantita,0))/SUM(IF(pixmania=9,c.quantita,0)),2) as avg_price_amazon_ES, " &_
		"ROUND(((ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=3 AND ISNULL(e.codice),c.quantita,0)),2)-sum(IF(pixmania=3 AND ISNULL(e.codice),fnEbaySpeseSpedizione(p.id_p,p.codice,p.id_c1,p.id_c2,p.id_c3,p.peso,c.prezzo, forn.id),0)))+ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=23,c.quantita,0)),2)+ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=24,c.quantita,0)),2))/NULLIF(SUM(IF(pixmania=3 AND ISNULL(e.codice),c.quantita,0))+SUM(IF(pixmania=23,c.quantita,0))+SUM(IF(pixmania=24,c.quantita,0)),0),2) as avg_price_ebay, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=22,c.quantita,0))/SUM(IF(pixmania=22,c.quantita,0)),2) as avg_price_manomano, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=28,c.quantita,0))/SUM(IF(pixmania=28,c.quantita,0)),2) as avg_price_cdiscount, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=41,c.quantita,0))/SUM(IF(pixmania=41,c.quantita,0)),2) as avg_price_digitec, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=57,c.quantita,0))/SUM(IF(pixmania=57,c.quantita,0)),2) as avg_price_digitec_de, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=52,c.quantita,0))/SUM(IF(pixmania=52,c.quantita,0)),2) as avg_price_stockly, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania IN (44,50,49,51),c.quantita,0))/SUM(IF(pixmania IN (44,50,49,51),c.quantita,0)),2) as avg_price_leroy, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania=53,c.quantita,0))/SUM(IF(pixmania=53,c.quantita,0)),2) as avg_price_mediamarket, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(pixmania IN (47,54,55),c.quantita,0))/SUM(IF(pixmania IN (47,54,55),c.quantita,0)),2) as avg_price_kaufland, " &_
		"ROUND(SUM(c.prezzo/(c.iva/100+1)*IF(o.order_type=3 AND mktp_import.reference_marketplace_id = 1027,c.quantita,0))/SUM(IF(o.order_type=3 AND mktp_import.reference_marketplace_id = 1027,c.quantita,0)),2) as avg_price_ha, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=2,c.quantita,0)),2) as acquisto_amazon, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=7,c.quantita,0)),2) as acquisto_amazon_de, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=8,c.quantita,0)),2) as acquisto_amazon_fr, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=9,c.quantita,0)),2) as acquisto_amazon_es, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*(IF(pixmania=3 AND ISNULL(e.codice),c.quantita,0)+IF(pixmania=23,c.quantita,0)+IF(pixmania=24,c.quantita,0))),2) as acquisto_ebay, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=22,c.quantita,0)),2) as acquisto_manomano, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=28,c.quantita,0)),2) as acquisto_cdiscount, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=41,c.quantita,0)),2) as acquisto_digitec, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=57,c.quantita,0)),2) as acquisto_digitec_de, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=52,c.quantita,0)),2) as acquisto_stockly, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania IN (44,50,49,51),c.quantita,0)),2) as acquisto_leroy, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania=53,c.quantita,0)),2) as acquisto_mediamarket, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(pixmania IN (47,54,55),c.quantita,0)),2) as acquisto_kaufland, " &_
		"ROUND(SUM(IF(f.prezzo > 0,f.prezzo,c.prezzo_acquisto)*IF(o.order_type=3 AND mktp_import.reference_marketplace_id = 1027,c.quantita,0)),2) as acquisto_ha " &_
		"from prodotti p " &_
		"inner join ordini_carrello c on p.id_p = c.id_p " &_
		"inner join fornitori forn using(id_forni_nav) " &_
		"inner join ordini_cliente o on c.id_ordine = o.id "&_
		"inner join v_category_name vcat on vcat.id_p=p.id_p "&_
		"LEFT JOIN marketplace_import_order_info AS mktp_import ON o.id = mktp_import.ordini_cliente_id " &_
		"LEFT JOIN ebay_offerte AS e ON e.codice=p.codice AND DATE_FORMAT(dataInizio,'%Y%m%d')<=o.data AND DATE_FORMAT(dataFine,'%Y%m%d')>=o.data AND e.contributo>0 " &_
		"LEFT JOIN cat_3 AS c3 ON c3.id_c3=p.id_c3 " &_
		"LEFT JOIN cat_2 AS c2 ON c2.id_c2=p.id_c2 " &_
		"LEFT JOIN cat_1 AS c1 ON c1.id_c1=p.id_c1 " &_
		"LEFT JOIN marca AS m ON m.id_m=p.id_m " &_
		"LEFT OUTER JOIN fornitori_ordini f on c.id = f.id_car " &_
		"where o.data >= '"&datainnum(Datadal)&"' AND o.data <= '"&datainnum(DataAl)&"' AND (o.step4_status = 1 AND (o.order_type=1 OR (o.pixmania IN (21,42,35,52) AND o.order_type=3) OR (o.order_type=3 AND mktp_import.reference_marketplace_id = 1027) OR o.order_type=5)) and c_annullato = 0 AND c.id_p <> 88526 "
		
		if request("Marketplace") <> "" then
			Sel3 = Sel3 & " and o.pixmania = "&request("Marketplace")&" "
		end if
		
		Sel3 = Sel3 & "group by cate3 "
		
		Sel3 = Sel3 & "ORDER BY (tot_amazon + tot_amazon_de + tot_amazon_fr + tot_amazon_es + tot_ebay + tot_manomano + tot_cdiscount + tot_digitec + tot_digitec_de + tot_stockly + tot_leroy + tot_mediamarket + tot_kaufland + tot_ha) DESC"
		
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
		tot_qta_tot_es=0
		tot_fatturato_tot_es=0
		tot_qta_manomano = 0
		tot_fatturato_manomano = 0
		tot_acquistato_cdiscount = 0
		tot_fatturato_cdiscount = 0
		tot_qta_cdiscount = 0
		tot_qta_digitec = 0
		tot_fatturato_digitec = 0
		tot_acquistato_digitec = 0
		tot_qta_digitec_de = 0
		tot_fatturato_digitec_de = 0
		tot_acquistato_digitec_de = 0
		tot_qta_stockly = 0
		tot_fatturato_stockly = 0
		tot_acquistato_stockly = 0
		tot_qta_leroy = 0
		tot_fatturato_leroy = 0
		tot_acquistato_leroy = 0
		tot_qta_mediamarket = 0
		tot_fatturato_mediamarket = 0
		tot_acquistato_mediamarket = 0
		tot_qta_kaufland = 0
		tot_fatturato_kaufland = 0
		tot_acquistato_kaufland = 0
		tot_qta_ha = 0
		tot_fatturato_ha = 0
		tot_acquistato_ha = 0
		iCount=0
		contaq=0
		Do While NOT rs3.EOF
            qta_yeppon  = cdbl(rs3("qta_amazon"))
            fatturato_yeppon        = rs3("tot_amazon")
            qta_coinvest  = cdbl(rs3("qta_ebay"))
            fatturato_coinvest        = rs3("tot_ebay")
            qta_mktp  = cdbl(rs3("qta_amazon_de"))
            fatturato_mktp       = rs3("tot_amazon_de")
			qta_tot = cdbl(rs3("qta_amazon_fr"))
            fatturato_tot       = rs3("tot_amazon_fr")
			qta_tot_es = cdbl(rs3("qta_amazon_es"))
            fatturato_tot_es       = rs3("tot_amazon_es")
			qta_manomano = cdbl(rs3("qta_manomano"))
            fatturato_manomano       = rs3("tot_manomano")
			fatturato_cdiscount = rs3("tot_cdiscount")
			qta_cdiscount = cdbl(rs3("qta_cdiscount"))
			qta_digitec = cdbl(rs3("qta_digitec"))
			fatturato_digitec = rs3("tot_digitec")
			qta_digitec_de = cdbl(rs3("qta_digitec_de"))
			fatturato_digitec_de = rs3("tot_digitec_de")
			qta_stockly = cdbl(rs3("qta_stockly"))
			fatturato_stockly = rs3("tot_stockly")
			qta_leroy = cdbl(rs3("qta_leroy"))
			fatturato_leroy = rs3("tot_leroy")
			qta_mediamarket = cdbl(rs3("qta_mediamarket"))
			fatturato_mediamarket = rs3("tot_mediamarket")
			qta_kaufland = cdbl(rs3("qta_kaufland"))
			fatturato_kaufland = rs3("tot_kaufland")
			qta_ha = cdbl(rs3("qta_ha"))
			fatturato_ha = rs3("tot_ha")
			acquisto_yeppon      = rs3("acquisto_amazon")
			acquisto_coinvest       = rs3("acquisto_ebay")				
			acquisto_mktp       = rs3("acquisto_amazon_de")
			acquisto_tot       = rs3("acquisto_amazon_fr")
			acquisto_tot_es       = rs3("acquisto_amazon_es")
			acquisto_manomano       = rs3("acquisto_manomano")
			acquisto_cdiscount       = rs3("acquisto_cdiscount")
			acquisto_digitec       = rs3("acquisto_digitec")
			acquisto_digitec_de       = rs3("acquisto_digitec_de")
			acquisto_stockly       = rs3("acquisto_stockly")
			acquisto_leroy       = rs3("acquisto_leroy")
			acquisto_mediamarket       = rs3("acquisto_mediamarket")
			acquisto_kaufland       = rs3("acquisto_kaufland")
			acquisto_ha       = rs3("acquisto_ha")
			tot_qta_yeppon   = tot_qta_yeppon + qta_yeppon
			tot_fatturato_yeppon   = tot_fatturato_yeppon + fatturato_yeppon
			tot_qta_coinvest   = tot_qta_coinvest + qta_coinvest
			tot_fatturato_coinvest   = tot_fatturato_coinvest + fatturato_coinvest
			tot_qta_mktp   = tot_qta_mktp + qta_mktp
			tot_fatturato_mktp   = tot_fatturato_mktp + fatturato_mktp
			tot_qta_tot = tot_qta_tot + qta_tot
			tot_fatturato_tot= tot_fatturato_tot + fatturato_tot
			tot_qta_tot_es = tot_qta_tot_es + qta_tot_es
			tot_fatturato_tot_es= tot_fatturato_tot_es + fatturato_tot_es
			tot_qta_manomano = tot_qta_manomano + qta_manomano
			tot_fatturato_manomano= tot_fatturato_manomano + fatturato_manomano
			tot_acquistato_yeppon=tot_acquistato_yeppon+acquisto_yeppon
			tot_acquistato_coinvest=tot_acquistato_coinvest+acquisto_coinvest
			tot_acquistato_mktp=tot_acquistato_mktp+acquisto_mktp
			tot_acquistato_tot=tot_acquistato_tot+acquisto_tot
			tot_acquistato_tot_es=tot_acquistato_tot_es+acquisto_tot_es
			tot_acquistato_manomano=tot_acquistato_manomano+acquisto_manomano
			tot_acquistato_cdiscount=tot_acquistato_cdiscount+acquisto_cdiscount
			tot_fatturato_cdiscount = tot_fatturato_cdiscount + fatturato_cdiscount 
			tot_qta_cdiscount   =  tot_qta_cdiscount + qta_cdiscount
			tot_qta_digitec = tot_qta_digitec + qta_digitec
			tot_fatturato_digitec = tot_fatturato_digitec + fatturato_digitec
			tot_acquistato_digitec = tot_acquistato_digitec + acquisto_digitec
			tot_qta_digitec_de = tot_qta_digitec_de + qta_digitec_de
			tot_fatturato_digitec_de = tot_fatturato_digitec_de + fatturato_digitec_de
			tot_acquistato_digitec_de = tot_acquistato_digitec_de + acquisto_digitec_de
			tot_qta_stockly = tot_qta_stockly + qta_stockly
			tot_fatturato_stockly = tot_fatturato_stockly + fatturato_stockly
			tot_acquistato_stockly = tot_acquistato_stockly + acquisto_stockly
			tot_qta_leroy = tot_qta_leroy + qta_leroy
			tot_fatturato_leroy = tot_fatturato_leroy + fatturato_leroy
			tot_acquistato_leroy = tot_acquistato_leroy + acquisto_leroy
			tot_qta_mediamarket = tot_qta_mediamarket + qta_mediamarket
			tot_fatturato_mediamarket = tot_fatturato_mediamarket + fatturato_mediamarket
			tot_acquistato_mediamarket = tot_acquistato_mediamarket + acquisto_mediamarket
			tot_qta_kaufland = tot_qta_kaufland + qta_kaufland
			tot_fatturato_kaufland = tot_fatturato_kaufland + fatturato_kaufland
			tot_acquistato_kaufland = tot_acquistato_kaufland + acquisto_kaufland
			tot_qta_ha = tot_qta_ha + qta_ha
			tot_fatturato_ha = tot_fatturato_ha + fatturato_ha
			tot_acquistato_ha = tot_acquistato_ha + acquisto_ha
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
				if tot_fatturato_tot_es= 0 then
					margine_tot_es=0
					avg_price_tot_es=0
				else
					avg_price_tot_es=tot_fatturato_tot_es/tot_qta_tot_es
					margine_tot_es=cdbl((tot_fatturato_tot_es-tot_acquistato_tot_es)/tot_fatturato_tot_es)*100
				end if
				if tot_fatturato_manomano= 0 then
					margine_manomano=0
					avg_price_manomano=0
				else
					avg_price_manomano=tot_fatturato_manomano/tot_qta_manomano
					margine_manomano=cdbl((tot_fatturato_manomano-tot_acquistato_manomano)/tot_fatturato_manomano)*100
				end if
				if tot_fatturato_cdiscount= 0 then
					margine_cdiscount=0
					avg_price_cdiscount=0
				else
					avg_price_cdiscount=tot_fatturato_cdiscount/tot_qta_cdiscount
					margine_cdiscount=cdbl((tot_fatturato_cdiscount-tot_acquistato_cdiscount)/tot_fatturato_cdiscount)*100
				end if
				if tot_fatturato_digitec= 0 then
					margine_digitec=0
					avg_price_digitec=0
				else
					avg_price_digitec=tot_fatturato_digitec/tot_qta_digitec
					margine_digitec=cdbl((tot_fatturato_digitec-tot_acquistato_digitec)/tot_fatturato_digitec)*100
				end if
				if tot_fatturato_digitec_de= 0 then
					margine_digitec_de=0
					avg_price_digitec_de=0
				else
					avg_price_digitec_de=tot_fatturato_digitec_de/tot_qta_digitec_de
					margine_digitec_de=cdbl((tot_fatturato_digitec_de-tot_acquistato_digitec_de)/tot_fatturato_digitec_de)*100
				end if
				if tot_fatturato_stockly= 0 then
					margine_stockly=0
					avg_price_stockly=0
				else
					avg_price_stockly=tot_fatturato_stockly/tot_qta_stockly
					margine_stockly=cdbl((tot_fatturato_stockly-tot_acquistato_stockly)/tot_fatturato_stockly)*100
				end if
				if tot_fatturato_leroy= 0 then
					margine_leroy=0
					avg_price_leroy=0
				else
					avg_price_leroy=tot_fatturato_leroy/tot_qta_leroy
					margine_leroy=cdbl((tot_fatturato_leroy-tot_acquistato_leroy)/tot_fatturato_leroy)*100
				end if
				if tot_fatturato_mediamarket= 0 then
					margine_mediamarket=0
					avg_price_mediamarket=0
				else
					avg_price_mediamarket=tot_fatturato_mediamarket/tot_qta_mediamarket
					margine_mediamarket=cdbl((tot_fatturato_mediamarket-tot_acquistato_mediamarket)/tot_fatturato_mediamarket)*100
				end if
				if tot_fatturato_kaufland= 0 then
					margine_kaufland=0
					avg_price_kaufland=0
				else
					avg_price_kaufland=tot_fatturato_kaufland/tot_qta_kaufland
					margine_kaufland=cdbl((tot_fatturato_kaufland-tot_acquistato_kaufland)/tot_fatturato_kaufland)*100
				end if
				if tot_fatturato_ha= 0 then
					margine_ha=0
					avg_price_ha=0
				else
					avg_price_ha=tot_fatturato_ha/tot_qta_ha
					margine_ha=cdbl((tot_fatturato_ha-tot_acquistato_ha)/tot_fatturato_ha)*100
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
        <td>Totale</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_yeppon%></td>
			<td align="left" ><%=formatnumber(avg_price_yeppon,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_yeppon,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_mktp%></td>
			<td align="left" ><%=formatnumber(avg_price_mktp,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_mktp,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_tot%></td>
			<td align="left" ><%=formatnumber(avg_price_tot,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_tot,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_tot,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_tot_es%></td>
			<td align="left" ><%=formatnumber(avg_price_tot_es,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_tot_es,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_tot_es,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_coinvest%></td>
			<td align="left" ><%=formatnumber(avg_price_coinvest,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_coinvest,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_manomano%></td>
			<td align="left" ><%=formatnumber(avg_price_manomano,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_manomano,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_manomano,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_cdiscount%></td>
			<td align="left" ><%=formatnumber(avg_price_cdiscount,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_cdiscount,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_cdiscount,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_digitec%></td>
			<td align="left" ><%=formatnumber(avg_price_digitec,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_digitec,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_digitec,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_digitec_de%></td>
			<td align="left" ><%=formatnumber(avg_price_digitec_de,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_digitec_de,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_digitec_de,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_stockly%></td>
			<td align="left" ><%=formatnumber(avg_price_stockly,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_stockly,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_stockly,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_leroy%></td>
			<td align="left" ><%=formatnumber(avg_price_leroy,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_leroy,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_leroy,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_mediamarket%></td>
			<td align="left" ><%=formatnumber(avg_price_mediamarket,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_mediamarket,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_mediamarket,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_kaufland%></td>
			<td align="left" ><%=formatnumber(avg_price_kaufland,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_kaufland,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_kaufland,2)%>%</td>
			<td align="left" style=" border-left: 5px solid black;"><%=tot_qta_ha%></td>
			<td align="left" ><%=formatnumber(avg_price_ha,2)%></td>
			<td align="left" ><%=formatnumber(tot_fatturato_ha,2)%></td>
			<td align="left"  style=" border-right: 5px solid black;"><%=formatnumber(margine_ha,2)%>%</td>

    </tr>
</thead>
<%
		
		if contaq=1 then
			rs3.movefirst
		end if
		iCount=0
        Do While NOT rs3.EOF

            qta_yeppon  = cdbl(rs3("qta_amazon"))
            fatturato_yeppon        = rs3("tot_amazon")
            qta_coinvest  = cdbl(rs3("qta_ebay"))
            fatturato_coinvest        = rs3("tot_ebay")
            qta_mktp  = cdbl(rs3("qta_amazon_de"))
            fatturato_mktp       = rs3("tot_amazon_de")
			qta_tot = cdbl(rs3("qta_amazon_fr"))
            fatturato_tot       = rs3("tot_amazon_fr")
			qta_tot_es = cdbl(rs3("qta_amazon_es"))
			fatturato_tot_es       = rs3("tot_amazon_es")
			qta_manomano = cdbl(rs3("qta_manomano"))
            fatturato_manomano       = rs3("tot_manomano")			
			fatturato_cdiscount      = rs3("tot_cdiscount")
			qta_cdiscount = cdbl(rs3("qta_cdiscount"))
			qta_digitec = cdbl(rs3("qta_digitec"))
			fatturato_digitec = rs3("tot_digitec")
			qta_digitec_de = cdbl(rs3("qta_digitec_de"))
			fatturato_digitec_de = rs3("tot_digitec_de")
			qta_stockly = cdbl(rs3("qta_stockly"))
			fatturato_stockly = rs3("tot_stockly")
			qta_leroy = cdbl(rs3("qta_leroy"))
			fatturato_leroy = rs3("tot_leroy")
			qta_mediamarket = cdbl(rs3("qta_mediamarket"))
			fatturato_mediamarket = rs3("tot_mediamarket")
			qta_kaufland = cdbl(rs3("qta_kaufland"))
			fatturato_kaufland = rs3("tot_kaufland")
			qta_ha = cdbl(rs3("qta_ha"))
			fatturato_ha = rs3("tot_ha")
			acquisto_yeppon      = rs3("acquisto_amazon")
			acquisto_coinvest    = rs3("acquisto_ebay")				
			acquisto_mktp       = rs3("acquisto_amazon_de")
			acquisto_tot       = rs3("acquisto_amazon_fr")
			acquisto_tot_es       = rs3("acquisto_amazon_es")
			acquisto_manomano       = rs3("acquisto_manomano")
			acquisto_cdiscount  = rs3("acquisto_cdiscount")			
			acquisto_digitec  = rs3("acquisto_digitec")
			acquisto_digitec_de  = rs3("acquisto_digitec_de")
			acquisto_stockly  = rs3("acquisto_stockly")
			acquisto_leroy  = rs3("acquisto_leroy")
			acquisto_mediamarket  = rs3("acquisto_mediamarket")
			acquisto_kaufland  = rs3("acquisto_kaufland")
			acquisto_ha  = rs3("acquisto_ha")
				if fatturato_yeppon= 0 then
					margine_yeppon=0
					avg_price_yeppon=0
				else
					avg_price_yeppon=fatturato_yeppon/qta_yeppon
					margine_yeppon=cdbl((fatturato_yeppon-acquisto_yeppon)/fatturato_yeppon)*100
				end if
				'response.write fatturato_coinvest & " " & qta_coinvest
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
				if fatturato_tot_es= 0 then
					margine_tot_es=0
					avg_price_tot_es=0
				else
					avg_price_tot_es=fatturato_tot_es/qta_tot_es
					margine_tot_es=cdbl((fatturato_tot_es-acquisto_tot_es)/fatturato_tot_es)*100
				end if
				if fatturato_manomano= 0 then
					margine_manomano=0
					avg_price_manomano=0
				else
					avg_price_manomano=fatturato_manomano/qta_manomano
					margine_manomano=cdbl((fatturato_manomano-acquisto_manomano)/fatturato_manomano)*100
				end if
				if fatturato_cdiscount= 0 then
					margine_cdiscount=0
					avg_price_cdiscount=0
				else
					avg_price_cdiscount=fatturato_cdiscount/qta_cdiscount
					margine_cdiscount=cdbl((fatturato_cdiscount-acquisto_cdiscount)/fatturato_cdiscount)*100
				end if
				if fatturato_digitec= 0 then
					margine_digitec=0
					avg_price_digitec=0
				else
					avg_price_digitec=fatturato_digitec/qta_digitec
					margine_digitec=cdbl((fatturato_digitec-acquisto_digitec)/fatturato_digitec)*100
				end if
				if fatturato_digitec_de= 0 then
					margine_digitec_de=0
					avg_price_digitec_de=0
				else
					avg_price_digitec_de=fatturato_digitec_de/qta_digitec_de
					margine_digitec_de=cdbl((fatturato_digitec_de-acquisto_digitec_de)/fatturato_digitec_de)*100
				end if
				if fatturato_stockly= 0 then
					margine_stockly=0
					avg_price_stockly=0
				else
					avg_price_stockly=fatturato_stockly/qta_stockly
					margine_stockly=cdbl((fatturato_stockly-acquisto_stockly)/fatturato_stockly)*100
				end if
				if fatturato_leroy= 0 then
					margine_leroy=0
					avg_price_leroy=0
				else
					avg_price_leroy=fatturato_leroy/qta_leroy
					margine_leroy=cdbl((fatturato_leroy-acquisto_leroy)/fatturato_leroy)*100
				end if
				if fatturato_mediamarket= 0 then
					margine_mediamarket=0
					avg_price_mediamarket=0
				else
					avg_price_mediamarket=fatturato_mediamarket/qta_mediamarket
					margine_mediamarket=cdbl((fatturato_mediamarket-acquisto_mediamarket)/fatturato_mediamarket)*100
				end if
				if fatturato_kaufland= 0 then
					margine_kaufland=0
					avg_price_kaufland=0
				else
					avg_price_kaufland=fatturato_kaufland/qta_kaufland
					margine_kaufland=cdbl((fatturato_kaufland-acquisto_kaufland)/fatturato_kaufland)*100
				end if
				if fatturato_ha= 0 then
					margine_ha=0
					avg_price_ha=0
				else
					avg_price_ha=fatturato_ha/qta_ha
					margine_ha=cdbl((fatturato_ha-acquisto_ha)/fatturato_ha)*100
				end if
            if bgcolor = "#EEE" then
                bgcolor = "#CCC;"
			else
                bgcolor = "#EEE"
            end if
    %>
        <tr style="background-color:<%=bgcolor%>;">
 			<td><%=rs3("cate3")%></a></td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_yeppon%></td>
				<td align="left"><%=formatnumber(avg_price_yeppon,2)%></td>
				<td align="left"><%=formatnumber(fatturato_yeppon,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_yeppon,2)%>%</td>
								<td align="left" style=" border-left: 5px solid black;"><%=qta_mktp%></td>
				<td align="left"><%=formatnumber(avg_price_mktp,2)%></td>
				<td align="left"><%=formatnumber(fatturato_mktp,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_mktp,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_tot%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_tot,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_tot,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_tot,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_tot_es%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_tot_es,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_tot_es,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_tot_es,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_coinvest%></td>
				<td align="left"><%=formatnumber(avg_price_coinvest,2)%></td>
				<td align="left"><%=formatnumber(fatturato_coinvest,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_coinvest,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_manomano%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_manomano,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_manomano,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_manomano,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_cdiscount%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_cdiscount,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_cdiscount,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_cdiscount,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_digitec%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_digitec,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_digitec,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_digitec,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_digitec_de%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_digitec_de,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_digitec_de,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_digitec_de,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_stockly%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_stockly,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_stockly,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_stockly,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_leroy%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_leroy,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_leroy,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_leroy,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_mediamarket%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_mediamarket,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_mediamarket,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_mediamarket,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black; font-weight:bold;"><%=qta_kaufland%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(avg_price_kaufland,2)%></td>
				<td align="left" style="font-weight:bold;"><%=formatnumber(fatturato_kaufland,2)%></td>
				<td align="left" style=" border-right: 5px solid black; font-weight:bold;"><%=formatnumber(margine_kaufland,2)%>%</td>
				<td align="left" style=" border-left: 5px solid black;"><%=qta_ha%></td>
				<td align="left"><%=formatnumber(avg_price_ha,2)%></td>
				<td align="left"><%=formatnumber(fatturato_ha,2)%></td>
				<td align="left" style=" border-right: 5px solid black;"><%=formatnumber(margine_ha,2)%>%</td>
        </tr>
<%
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
<%
end if
%>


<!--#include virtual="/int/connection-close.asp"-->