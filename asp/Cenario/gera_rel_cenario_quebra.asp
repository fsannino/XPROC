<%@LANGUAGE="VBSCRIPT"%> 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

onda=0
mega=0
situacao=0

onda=request("selOnda")
mega=request("selMegaProcesso")
situacao=request("selStatus")

if request("soma")=1 then
	response.redirect "gera_rel_cenario_quebra_perc.asp?selOnda=" & onda & "&SelMegaProcesso=" & mega & "&selStatus=" & situacao
end if

compl=""

if onda<>0 then
	ssql="SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA<>4 AND ONDA_CD_ONDA=" & ONDA
else
	ssql="SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_CD_ONDA"
end if

set onda=db.execute(ssql)

if mega<>0 then
	compl=compl + " MEPR_CD_MEGA_PROCESSO=" & mega & " AND"
end if

on error resume next
if situacao<>0 then
	if err.number<>0 then
		compl=compl + " CENA_TX_SITUACAO='" & situacao & "' AND"
	end if
end if

If len(compl)>0 then
	compl=left(compl,((len(compl))-4))
	compl=" AND " + COMPL
END IF

ORDEM= " ORDER BY ONDA_CD_ONDA, MEPR_CD_MEGA_PROCESSO, SUMO_NR_CD_SEQUENCIA, CENA_TX_SITUACAO"

%>
<html>
<head>
<STYLE type=text/css>
BODY {
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
<form name="frm1" method="post" action="gera_rel_cenario_quebra.asp">
  <input type="hidden" name="INC" size="20" value="1"> 
<%if request("excel")=0 then%>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"><a href="gera_rel_cenario_quebra.asp?excel=1&selOnda=<%=request("selOnda")%>&SelMegaProcesso=<%=request("selMegaProcesso")%>&selStatus=<%=request("selStatus")%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if%>
  <table width="90%" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso" height="34">
    <tr>
      <td width="10%" height="1"></td>
      <td width="71%" height="1"> 
      </td>
    </tr>
    <tr>
      <td width="10%" height="21"></td>
      <td width="71%" height="21"> 
      <font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="3"><b>Relatório
      de Cenário</b></font> 
      </td>
    </tr>
  </table>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
  <%
  	TOTAL=0
  	
  	do until onda.eof=true
  	
  	VALOR_ONDA=ONDA("ONDA_TX_DESC_ONDA")
    
      		SUP=0
      		MES=0
      		COM=0
      		EMP=0
      		MAN=0
      		POS=0
      		PRD=0
      		QUA=0
      		REC=0
      		PLC=0
      		FIN=0
      		RHU=0
      		GER=0
      		TIN=0
      		
      		EE=0
      		DF=0
      		DS=0
      		PT=0
      		TQ=0
			TD=0


    ssql=""
    ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE ONDA_CD_ONDA=" & ONDA("ONDA_CD_ONDA") & COMPL & ORDEM 
    
    set rs=db.execute(ssql)
    
    if rs.eof=false then
    %>
     <table border="0" width="823" cellspacing="1">
    <tr>
      <td width="71" bgcolor="#330099">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#FFFFFF"><b>Onda</b></font></p>
      </td>
      <td width="165" bgcolor="#330099">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#FFFFFF"><b>Mega-Processo</b></font></p>
      </td>
      <td width="185" bgcolor="#330099">
        <font face="Verdana" size="1" color="#FFFFFF"><b>Assunto</b></font>
      </td>
      <td width="105" bgcolor="#330099">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#FFFFFF"><b>Cenário</b></font></p>
      </td>
      <td width="130" bgcolor="#330099">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#FFFFFF"><b>Status</b></font></td>
      <td width="133" bgcolor="#330099">
        <font face="Verdana" size="1" color="#FFFFFF"><b>Escopo</b></font></td>
    </tr>
    <%
    end if
    
    tem=0
    
    MEGA_ANT=0
    MEGA_ATUAL=0
    
    val_assunto="" 
    
    do until rs.eof=true
    
    MEGA_ATUAL=RS("MEPR_CD_MEGA_PROCESSO")
   
    %>
    <tr>
      <td width="71">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=VALOR_ONDA%></font></p>
      </td>
      <%
      SELECT CASE RS("MEPR_CD_MEGA_PROCESSO")
      CASE 1
      		SUP=SUP+1
      CASE 2
      		MES=MES+1
      CASE 3
      		COM=COM+1
      CASE 4
      		EMP=EMP+1
      CASE 5
      		MAN=MAN+1
      CASE 6
      		POS=POS+1
      CASE 7	
      		PRD=PRD+1
      CASE 8
      		QUA=QUA+1
      CASE 9
      		REC=REC+1
      CASE 10
      		PLC=PLC+1
      CASE 11
      		FIN=FIN+1
      CASE 12
      		RHU=RHU+1
      CASE 13
      		GER=GER+1
      CASE 14
      		TIN=TIN+1
      END SELECT
      %>
      <td width="165">
      <%
      SET RS2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS("MEPR_CD_MEGA_PROCESSO"))
      IF MEGA_ATUAL<>MEGA_ANT THEN
      	VALOR_MEGA=RS2("MEPR_TX_DESC_MEGA_PROCESSO")
      ELSE
      	VALOR_MEGA=" "
      END IF
      %>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=VALOR_MEGA%></font></p>
      </td>
      <td width="185">
      <font size="1">
      <%
        set temp=db.execute("SELECT * FROM SUB_MODULO WHERE SUMO_NR_CD_SEQUENCIA=" & RS("SUMO_NR_CD_SEQUENCIA"))
        if temp.eof=false then
	        val_assunto=temp("SUMO_TX_DESC_SUB_MODULO")     
	     end if
	     
	   %>
      </font>
        <font size="1" face="Verdana"><%=val_assunto%></font>
      </td>
      <td width="105">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=RS("CENA_CD_CENARIO")%></font></p>
      </td>
      <td width="130">
      <%
      VALOR=""
      
      SELECT CASE RS("CENA_TX_SITUACAO")
		
      CASE "EE"
      		EE=EE+1
      		VALOR="EM ELABORAÇÃO"
      CASE "DF"
      		DF=DF+1
      		VALOR="DEFINIDO"
      CASE "DS"
      		DS=DS+1
      		VALOR="DESENHADO"
      CASE "PT"
      		PT=PT+1
      		VALOR="PRONTO PARA TESTE"
      CASE "TQ"
      		TQ=TQ+1
      		VALOR="TESTADO NO PEQ"
      CASE "TD"
			TD=TD+1      
      		VALOR="TESTADO NO PED"
      END SELECT
      %>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=VALOR%></font></p>
      </td>
      <td width="133">
      <font size="1">
      <%
		escopo=rs("CENA_TX_SITUACAO_VALIDACAO")      
		
		select case escopo
		case 0
			val_escopo="FORA DO ESCOPO"      
		case 1
			val_escopo="DENTRO DO ESCOPO"      
		end select      

      %>
      </font>
      <font face="Verdana" size="1"><%=val_escopo%></font>
      </td>
    </tr>
    <%
    tem=tem+1
    VALOR_ONDA=" "
    MEGA_ANT=RS("MEPR_CD_MEGA_PROCESSO")
    rs.movenext
    loop
    
    if tem>0 then
    
    TOTAL=TOTAL+TEM
    %>
    </table>
    <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;
  
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1">Total
  de Cenários por Onda : <b> <%=TEM%></b></font>
  
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1">Total
  de Cenários Por Mega-Processo : </font>
  <table border="0" width="494">
    <tr>
      <td width="31" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>SUP</b></font></td>
      <td width="32" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>MES</b></font></td>
      <td width="32" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>COM</b></font></td>
      <td width="32" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>EMP</b></font></td>
      <td width="33" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>MAN</b></font></td>
      <td width="26" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>POS</b></font></td>
      <td width="26" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>PRD</b></font></td>
      <td width="28" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>QUA</b></font></td>
      <td width="25" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>REC</b></font></td>
      <td width="25" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>PLC</b></font></td>
      <td width="25" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>FIN</b></font></td>
      <td width="34" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>RHU</b></font></td>
      <td width="36" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>GER</b></font></td>
      <td width="27" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>TIN</b></font></td>
    </tr>
    <tr>
      <td width="31" align="center"><b><font face="Verdana" size="1"><%=SUP%></font></b></td>
      <td width="32" align="center"><b><font face="Verdana" size="1"><%=MES%></font></b></td>
      <td width="32" align="center"><b><font face="Verdana" size="1"><%=COM%></font></b></td>
      <td width="32" align="center"><b><font face="Verdana" size="1"><%=EMP%></font></b></td>
      <td width="33" align="center"><b><font face="Verdana" size="1"><%=MAN%></font></b></td>
      <td width="26" align="center"><b><font face="Verdana" size="1"><%=POS%></font></b></td>
      <td width="26" align="center"><b><font face="Verdana" size="1"><%=PRD%></font></b></td>
      <td width="28" align="center"><b><font face="Verdana" size="1"><%=QUA%></font></b></td>
      <td width="25" align="center"><b><font face="Verdana" size="1"><%=REC%></font></b></td>
      <td width="25" align="center"><b><font face="Verdana" size="1"><%=PLC%></font></b></td>
      <td width="25" align="center"><b><font face="Verdana" size="1"><%=FIN%></font></b></td>
      <td width="34" align="center"><b><font face="Verdana" size="1"><%=RHU%></font></b></td>
      <td width="36" align="center"><b><font face="Verdana" size="1"><%=GER%></font></b></td>
      <td width="27" align="center"><b><font face="Verdana" size="1"><%=TIN%></font></b></td>
    </tr>
  </table>
  
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1">Total
  de Cenários por Status :</font>
  <table border="0" width="691">
    <tr>
      <td width="103" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Em
        Elaboração</b></font></td>
      <td width="96" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Definido</b></font></td>
      <td width="102" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Desenhado</b></font></td>
      <td width="117" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Pronto
        para Teste</b></font></td>
      <td width="108" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Testado no
        PED</b></font></td>
      <td width="127" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Testado no
        PEQ</b></font></td>
    </tr>
    <tr>
      <td width="103" align="center"><b><font size="1" face="Verdana"><%=EE%></font></b></td>
      <td width="96" align="center"><b><font size="1" face="Verdana"><%=DF%></font></b></td>
      <td width="102" align="center"><b><font size="1" face="Verdana"><%=DS%></font></b></td>
      <td width="117" align="center"><b><font size="1" face="Verdana"><%=PT%></font></b></td>
      <td width="108" align="center"><b><font size="1" face="Verdana"><%=TD%></font></b></td>
      <td width="127" align="center"><b><font size="1" face="Verdana"><%=TQ%></font></b></td>
    </tr>
    <tr>
	  <%
		V1=FORMATPERCENT(EE/TEM)
		V2=FORMATPERCENT(DF/TEM)
		V3=FORMATPERCENT(DS/TEM)
		V4=FORMATPERCENT(PT/TEM)
		V5=FORMATPERCENT(TD/TEM)
		V6=FORMATPERCENT(TQ/TEM)
	  %>
	  <td width="103" align="center"><b><font size="1" face="Verdana"><%=LEFT(V1,(LEN(V1))-4)%>%</font></b></td>
      <td width="96" align="center"><b><font size="1" face="Verdana"><%=LEFT(V2,(LEN(V2))-4)%>%</font></b></td>
      <td width="102" align="center"><b><font size="1" face="Verdana"><%=LEFT(V3,(LEN(V3))-4)%>%</font></b></td>
      <td width="117" align="center"><b><font size="1" face="Verdana"><%=LEFT(V4,(LEN(V4))-4)%>%</font></b></td>
      <td width="108" align="center"><b><font size="1" face="Verdana"><%=LEFT(V5,(LEN(V5))-4)%>%</font></b></td>
      <td width="127" align="center"><b><font size="1" face="Verdana"><%=LEFT(V6,(LEN(V6))-4)%>%</font></b></td> </tr>
    <tr>
      <td width="199" align="center" colspan="2"><b><font size="1" face="Verdana" color="#FFFFFF">-</font></b></td>
      <td width="102" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">-</font></b></td>
      <td width="117" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">-</font></b></td>
      <td width="108" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">-</font></b></td>
      <td width="127" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">-</font></b></td>
    </tr>
    <tr>
      <td width="199" align="center" colspan="2">
        <p align="left"><font size="1" face="Verdana">Percentual
        Acumulativo</font></p>
      </td>
      <td width="102" align="center"></td>
      <td width="117" align="center"></td>
      <td width="108" align="center"></td>
      <td width="127" align="center"></td>
    </tr>
    <tr>
      <td width="103" align="center" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="96" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Definido</b></font></td>
      <td width="102" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Desenhado</b></font></td>
      <td width="117" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Pronto
        para Teste</b></font></td>
      <td width="108" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Testado no
        PED</b></font></td>
      <td width="127" align="center" bgcolor="#C0C0C0"><font face="Verdana" size="1"><b>Testado no
        PEQ</b></font></td>
    </tr>
    <%
	VV1=FORMATPERCENT((DF/TEM)+(DS/TEM)+(PT/100)+(TD/TEM)+(TQ/TEM))
	VV2=FORMATPERCENT((DS/TEM)+(PT/100)+(TD/TEM)+(TQ/TEM))
	VV3=FORMATPERCENT((PT/100)+(TD/TEM)+(TQ/TEM))
	VV4=FORMATPERCENT((TD/TEM)+(TQ/TEM))
	VV5=FORMATPERCENT((TQ/TEM))
	%>
	<tr>
      <td width="103" align="center" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="96" align="center"><b><font size="1" face="Verdana"><%=LEFT(VV1,(LEN(VV1))-4)%>%</font></b></td>
      <td width="102" align="center"><b><font size="1" face="Verdana"><%=LEFT(VV2,(LEN(VV2))-4)%>%</font></b></td>
      <td width="117" align="center"><b><font size="1" face="Verdana"><%=LEFT(VV3,(LEN(VV3))-4)%>%</font></b></td>
      <td width="108" align="center"><b><font size="1" face="Verdana"><%=LEFT(VV4,(LEN(VV4))-4)%>%</font></b></td>
      <td width="127" align="center"><b><font size="1" face="Verdana"><%=LEFT(VV5,(LEN(VV5))-4)%>%</font></b></td>
    </tr>
  </table>
  <p>
    <%
    end if
    onda.movenext
    loop
    %>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2">Total
  Geral de Cenários : <b> <%=TOTAL%></b></font>
  </form>
<p>&nbsp;</p>
</body>
</html>
