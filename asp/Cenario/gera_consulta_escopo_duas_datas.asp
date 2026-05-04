<%
dim vet_Cenario1()
dim vet_Cenario2()
dim vet_Cenario11()
dim vet_Cenario22()

mega=request("selMegaProcesso")
onda=request("selOnda")
situacao=request("selStatus")
assunto=request("selAssunto")

if mega<>0 then
	compl=compl+" MEPR_CD_MEGA_PROCESSO=" & mega & " AND"
end if

if onda<>0 then
	compl=compl+" ONDA_CD_ONDA=" & onda & " AND"
end if

if situacao<>"0" then
	compl=compl+" CENA_TX_SITUACAO='" & situacao & "' AND"
end if

data_cons1=request("data01")
data_cons2=request("data02")

dia=day(data_cons1)
mes=month(data_cons1)
ano=year(data_cons1)

dia1=Right("00" & day(data_cons1),2)
mes1=Right("00" & month(data_cons1),2)
ano1=Right("0000" & year(data_cons1),4)

dia2=Right("00" & day(data_cons2),2)
mes2=Right("00" & month(data_cons2),2)
ano2=Right("0000" & year(data_cons2),4)

dia1=left(data_cons1, 2)
mes1=left((right(data_cons1,7)),2)
ano1=right(data_cons1, 4)

dia2=left(data_cons2, 2)
mes2=left((right(data_cons2,7)),2)
ano2=right(data_cons2, 4)

data_inicio = ano1 & "-" & mes1 & "-" & dia1
data_termino = ano2 & "-" & mes2 & "-" & dia2

'data_inicio = dateadd ("dd",-1,data_inicio)
'data_termino = dateadd ("dd",1,data_termino)

data_inicio = (Cdate(data_inicio)-1)
data_termino = (Cdate(data_termino)+1)

'response.Write(data_inicio) & "<p>"
'response.Write(data_termino) & "<p>"

set db = Server.CreateObject("ADODB.Connection")
'conecta="Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"
'db.Open conecta
db.Open Session("Conn_String_Cogest_Gravacao")
Const adUseClient = 3
db.CursorLocation = adUseClient

if len(compl)>0 then
	compl=left(compl,(len(compl))-4)
	compl=" WHERE" + compl
end if

'response.Write("   comple    ")
'response.Write(compl)
'response.Write("   fim comple    ")

set rs=db.execute("SELECT DISTINCT * FROM " & Session("PREFIXO") & "CENARIO" & compl & " ORDER BY CENA_CD_CENARIO")

intTotalRegistro = rs.RecordCount

'response.Write("  aaa   ")
'response.Write(intTotalRegistro)

redim vet_Cenario1(intTotalRegistro)
redim vet_Cenario2(intTotalRegistro)
redim vet_Cenario11(intTotalRegistro)
redim vet_Cenario22(intTotalRegistro)

'For next i to intTotalRegistro
'    vet_Cenario1(i) = ""
'next

int_indice1 = 0
int_indice2 = 0
situacao=0

if not rs.EOF then

   ssql=""
   ssql="SELECT CENA_CD_CENARIO, MAX(ATUA_DT_ATUALIZACAO) AS DATA "
   ssql=ssql + " , CEVA_TX_SITUACAO FROM " & session("prefixo") & "CENARIO_VALIDACAO"
   ssql=ssql + " GROUP BY CENA_CD_CENARIO, CEVA_TX_SITUACAO"
   ssql=ssql + " HAVING (MAX(ATUA_DT_ATUALIZACAO) < CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102))"
   'RESPONSE.Write(ssql) & "<P>"	
   set temp2=db.execute(ssql)	
   if temp2.eof=false then    

      do until rs.eof=true   
         VALOR5=0
         ssql="" 
         ssql="SELECT CENA_CD_CENARIO, (ATUA_DT_ATUALIZACAO) AS DATA, CEVA_TX_SITUACAO FROM " & session("prefixo") & "CENARIO_VALIDACAO"
         ssql=ssql + " GROUP BY CENA_CD_CENARIO, CEVA_TX_SITUACAO, ATUA_DT_ATUALIZACAO"
         ssql=ssql + " HAVING (CENA_CD_CENARIO = '" & RS("CENA_CD_CENARIO") & "') "
         ssql=ssql + " AND (ATUA_DT_ATUALIZACAO < CONVERT(DATETIME, '" & data_inicio & " 00:00:00', 102)) "
		 ssql=ssql + " ORDER BY ATUA_DT_ATUALIZACAO"
		 'RESPONSE.Write(ssql) & "<P>"	         
	     set temp=db.execute(ssql)
         if temp.eof=true then
			valor5=4			
         else
            do until temp.eof=true	
			   situacao=temp("CEVA_TX_SITUACAO")
			   temp.movenext
		    loop
	     end if	
	 	 if situacao<>0 and VALOR5<>4 then
		    int_indice1 = int_indice1 + 1
            vet_Cenario1(int_indice1) =  rs("CENA_CD_CENARIO") 
		 end if		
         rs.movenext
		 situacao=0
      loop
   else
      'todos os cenários estão fora do escopo
   end if
else
   ' não possui cenarios
end if
'*********************************************************************
' encontra escopo para segunda data
'*********************************************************************
on error resume next
rs.movefirst
err.clear()

if not rs.EOF then

   ssql=""
   ssql="SELECT CENA_CD_CENARIO, MAX(ATUA_DT_ATUALIZACAO) AS DATA "
   ssql=ssql + " , CEVA_TX_SITUACAO FROM " & session("prefixo") & "CENARIO_VALIDACAO"
   ssql=ssql + " GROUP BY CENA_CD_CENARIO, CEVA_TX_SITUACAO"
   ssql=ssql + " HAVING (MAX(ATUA_DT_ATUALIZACAO) < CONVERT(DATETIME, '" & data_termino & " 00:00:00', 102))"
   RESPONSE.Write(ssql) & "<P>"	
   set temp2=db.execute(ssql)	
   if temp2.eof=false then    

      do until rs.eof=true   
         VALOR5=0
         ssql="" 
         ssql="SELECT CENA_CD_CENARIO, (ATUA_DT_ATUALIZACAO) AS DATA, CEVA_TX_SITUACAO FROM " & session("prefixo") & "CENARIO_VALIDACAO"
         ssql=ssql + " GROUP BY CENA_CD_CENARIO, CEVA_TX_SITUACAO, ATUA_DT_ATUALIZACAO"
         ssql=ssql + " HAVING (CENA_CD_CENARIO = '" & RS("CENA_CD_CENARIO") & "') "
         ssql=ssql + " AND (ATUA_DT_ATUALIZACAO < CONVERT(DATETIME, '" & data_termino & " 00:00:00', 102)) "
		 ssql=ssql + " ORDER BY ATUA_DT_ATUALIZACAO"
		 'RESPONSE.Write(ssql) & "<P>"	         
	     set temp=db.execute(ssql)
         if temp.eof=true then
			valor5=4			
         else
            do until temp.eof=true	
			   situacao=temp("CEVA_TX_SITUACAO")
			   temp.movenext
		    loop
	     end if	
	 	 if situacao<>0 and VALOR5<>4 then
		    int_indice2 = int_indice2 + 1
            vet_Cenario2(int_indice2) =  rs("CENA_CD_CENARIO") 
		 end if		
         rs.movenext
		 situacao=0
      loop
   else
      'todos os cenários estão fora do escopo
   end if
else
   ' não possui cenarios
end if

'	ls_Min_Data = CDate("15/10/2001")	
'ls_dt_referencia = DoDateTime(Date(), 2, 1033)
'DoDateTime(ls_Sexta, 2, 1033)
'ls_Segunda = DateAdd("d", 0, ls_Segunda)

function DoDateTime(str, nNamedFormat, nLCID)				
	dim strRet								
	dim nOldLCID																		
	strRet = str								
	If (nLCID > -1) Then							
		oldLCID = Session.LCID						
	End If																			
	On Error Resume Next																	
	If (nLCID > -1) Then							
		Session.LCID = nLCID						
	End If																			
	If ((nLCID < 0) Or (Session.LCID = nLCID)) Then				
		strRet = FormatDateTime(str, nNamedFormat)			
	End If																			
	If (nLCID > -1) Then							
		Session.LCID = oldLCID						
	End If																			
	DoDateTime = strRet							
End Function

j = 0
For i = 1 to intTotalRegistro
    if vet_Cenario1(i) <> "" then
       if SequentialSearchStringArray(vet_Cenario2 , vet_Cenario1(i)) = -1 then
	      j = j + 1
	      vet_Cenario11(j) = vet_Cenario1(i)
	   end if
	end if   
next

j = 0
For i = 1 to intTotalRegistro
    if vet_Cenario2(i) <> "" then
       if SequentialSearchStringArray(vet_Cenario1 , vet_Cenario2(i)) = -1 then
	      j = j + 1
	      vet_Cenario22(j) = vet_Cenario2(i)
	   end if
	end if   
next


Public Function SequentialSearchStringArray(sArray() , sFind ) 
   Dim j     
   Dim iLBound 
   Dim iUBound 

   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   'iLBound = 0
   'iUBound = 10

   For j = iLBound To iUBound
      If sArray(j) = sFind Then SequentialSearchStringArray = j: Exit Function
   Next 

   SequentialSearchStringArray = -1
End Function


%>
<html>
<head>
<STYLE type=text/css>
BODY{
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="gera_consulta_escopo.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">2 </td>
  </tr>
</table>
  <p align="center" style="word-spacing: 0; margin-top: 0"><font color="#000080" face="Verdana" size="3">Consulta 
    de Escopo de Cenário</font></p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#000080" face="Verdana" size="3">Data 
    1: <%=request("data01")%> - Data 2: <%=request("data02")%></font></b></p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <% a=1
  if a <> 1 then %>
  <table border="0" width="95%">
    <tr>
      <td width="62%" bgcolor="#000080"><font size="2" face="Verdana" color="#FFFFFF"><b>Cen&aacute;rios 
        Exclu&iacute;dos do Escopo<%=data_inicio%> </b></font></td>
      <td width="13%" bgcolor="#000080" align="center"><font size="2" face="Verdana" color="#FFFFFF"><b>Status
        do Cenário</b></font></td>
    </tr>
	<% FOR j = 1 TO intTotalRegistro 
	       if vet_Cenario1(j) <> "" then
	%>
    <tr>
      <td width="62%" bgcolor="#D8D8D8"><font face="Verdana" size="1"><B><%=vet_Cenario1(j)%><%'=rs("CENA_CD_CENARIO")%></B>- <%'=rs("CENA_TX_TITULO_CENARIO")%></font></td>
      <td width="13%" align="center" bgcolor="#CCD2C6"><font face="Verdana" size="1"><%=str_Situacao%></font></td>
    </tr>
	   <%  end if
	NEXT %>
  </table>
  <table border="0" width="95%">
    <tr>
      <td width="62%" bgcolor="#000080"><font size="2" face="Verdana" color="#FFFFFF"><b>Cen&aacute;rios 
        Inclu&iacute;dos no Escopo<%=data_termino%> </b></font></td>
      <td width="13%" bgcolor="#000080" align="center"><font size="2" face="Verdana" color="#FFFFFF"><b>Status
        do Cenário</b></font></td>
    </tr>
	<% FOR j = 1 TO intTotalRegistro 
           if vet_Cenario1(j) <> "" then
	%>
    <tr>
      <td width="62%" bgcolor="#D8D8D8"><font face="Verdana" size="1"><B><%=vet_Cenario2(j)%><%'=rs("CENA_CD_CENARIO")%></B>- <%'=rs("CENA_TX_TITULO_CENARIO")%></font></td>
      <td width="13%" align="center" bgcolor="#CCD2C6"><font face="Verdana" size="1"><%=valor3%></font></td>
    </tr>
	     <% end if
	NEXT %>
  </table>
  <p>
 <% end if %>
  </p>
  <p><font size="2" face="Verdana" color="#0000FF"><b>Data 1 - <%=request("data01")%> 
    </b></font> </p>
  <table border="0" width="95%">
    <tr> 
      <td width="62%" bgcolor="#000080"><font size="2" face="Verdana" color="#FFFFFF"><b>Cen&aacute;rios 
        Exclu&iacute;dos do Escopo</b></font></td>
      <td width="13%" bgcolor="#000080" align="center"><font size="2" face="Verdana" color="#FFFFFF"><b>Status 
        do Cenário</b></font></td>
    </tr>
    <% int_Tot = 0
	FOR j = 1 TO intTotalRegistro 
        	if vet_Cenario11(j) <> "" then
			   int_Tot = int_Tot + 1
			   str_SQL = ""
			   str_SQL = str_SQL & " SELECT CENA_TX_TITULO_CENARIO, CENA_TX_SITUACAO, "
               str_SQL = str_SQL & " CENA_CD_CENARIO"
               str_SQL = str_SQL & " FROM CENARIO"
               str_SQL = str_SQL & " WHERE (CENA_CD_CENARIO = '" & vet_Cenario11(j) & "')"
			   set rs_Cenario = db.Execute(str_SQL)
			   if not rs_Cenario.EOF then
			      str_Ds_Cenario = rs_Cenario("CENA_TX_TITULO_CENARIO")
				  str_Situacao = rs_Cenario("CENA_TX_SITUACAO")
				  rs_Cenario.close
	              select case str_Situacao
                        case "EE"
                   			str_Situacao="EM ELABORAÇÃO"
                  		case "DS"
                 			str_Situacao="DESENHADO"
                 		case "DF"
                			str_Situacao="DEFINIDO"
                		case "PT"
                 			str_Situacao="PRONTO PARA TESTE"
                		case "TD"
                 			str_Situacao="TESTADO NO PED"
                		case "TQ"
                			str_Situacao="TESTADO NO PEQ"
                		'case else
                			'str_Situacao="SEM STATUS DEFINIDO"
             	  end select
				else   		 
 				    str_Situacao="CENÁRIO NÃO ENCONTRADO" 
				end if  
	%>
    <tr> 
      <td width="62%" bgcolor="#D8D8D8"><font face="Verdana" size="1"><B><a href="gera_rel_geral.asp?id=<%=vet_Cenario11(j)%>"><%=vet_Cenario11(j)%> </a></B>- <%=str_Ds_Cenario%>
        </font></td>
      <td width="13%" align="center" bgcolor="#CCD2C6"><font face="Verdana" size="1"><%=str_Situacao%></font></td>
    </tr>
    <% end if
	NEXT %>
    <tr>
      <td bgcolor="#D8D8D8"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Total 
          : <%=int_Tot%></font></div></td>
      <td align="center" bgcolor="#CCD2C6">&nbsp;</td>
    </tr>
  </table>
  <p><font size="2" face="Verdana" color="#0000FF"><b>Data 2 - <%=request("data02")%> 
    </b></font></p>
  <table border="0" width="95%">
    <tr> 
      <td width="62%" bgcolor="#000080"><font size="2" face="Verdana" color="#FFFFFF"><b>Cen&aacute;rios 
        Inclu&iacute;dos no Escopo</b></font></td>
      <td width="13%" bgcolor="#000080" align="center"><font size="2" face="Verdana" color="#FFFFFF"><b>Status 
        do Cenário</b></font></td>
    </tr>
    <% int_Tot = 0
	FOR j = 1 TO intTotalRegistro 
        	if vet_Cenario22(j) <> "" then
			   int_Tot = int_Tot + 1
			   str_SQL = ""
			   str_SQL = str_SQL & " SELECT CENA_TX_TITULO_CENARIO, CENA_TX_SITUACAO, "
               str_SQL = str_SQL & " CENA_CD_CENARIO"
               str_SQL = str_SQL & " FROM CENARIO"
               str_SQL = str_SQL & " WHERE (CENA_CD_CENARIO = '" & vet_Cenario22(j) & "')"
			   set rs_Cenario = db.Execute(str_SQL)
			   if not rs_Cenario.EOF then
			      str_Ds_Cenario = rs_Cenario("CENA_TX_TITULO_CENARIO")
				  str_Situacao = rs_Cenario("CENA_TX_SITUACAO")
				  rs_Cenario.close
	              select case str_Situacao
                        case "EE"
                   			str_Situacao="EM ELABORAÇÃO"
                  		case "DS"
                 			str_Situacao="DESENHADO"
                 		case "DF"
                			str_Situacao="DEFINIDO"
                		case "PT"
                 			str_Situacao="PRONTO PARA TESTE"
                		case "TD"
                 			str_Situacao="TESTADO NO PED"
                		case "TQ"
                			str_Situacao="TESTADO NO PEQ"
                		case else
                			str_Situacao="SEM STATUS DEFINIDO"
             	  end select		
				else   		 
 				    str_Situacao="CENÁRIO NÃO ENCONTRADO" 
				end if  				   			   
	%>
    <tr> 
      <td width="62%" bgcolor="#D8D8D8"><font face="Verdana" size="1"><B><a href="gera_rel_geral.asp?id=<%=vet_Cenario22(j)%>"><%=vet_Cenario22(j)%></a> </B>- <%=str_Ds_Cenario%>
        </font></td>
      <td width="13%" align="center" bgcolor="#CCD2C6"><font face="Verdana" size="1"><%=str_Situacao%></font></td>
    </tr>
    <% end if
	NEXT %>
    <tr>
      <td bgcolor="#D8D8D8"><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Total 
          : <%=int_Tot%></font></div></td>
      <td align="center" bgcolor="#CCD2C6">&nbsp;</td>
    </tr>
  </table>
</form>
</body>

</html>
