<%
server.scripttimeout=1200

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
db.cursorlocation=3

visual=1
visual=request("selvisual")

'set rs1=db.execute("SELECT DISTINCT CENA_TX_PACT_TESTE FROM CENARIO WHERE CENA_TX_PACT_TESTE<>'' ORDER BY CENA_TX_PACT_TESTE")
select case visual
case 1
	set rs1=db.execute("SELECT DISTINCT PCTE_TX_PACT_TESTE, PCTE_DT_INICIO_TESTES AS DATAINI FROM PACOTE_TESTES ORDER BY PCTE_TX_PACT_TESTE")
	ordena=1
case 2
	set rs1=db.execute("SELECT DISTINCT PCTE_TX_PACT_TESTE, PCTE_DT_INICIO_TESTES AS DATAINI FROM PACOTE_TESTES ORDER BY PCTE_TX_PACT_TESTE")
	ordena=2
case 3
	set rs1=db.execute("SELECT DISTINCT PCTE_TX_PACT_TESTE, PCTE_DT_INICIO_TESTES AS DATAINI FROM PACOTE_TESTES ORDER BY PCTE_DT_INICIO_TESTES")
	ordena=1
case 4
	set rs1=db.execute("SELECT DISTINCT PCTE_TX_PACT_TESTE, PCTE_DT_INICIO_TESTES AS DATAINI FROM PACOTE_TESTES ORDER BY PCTE_DT_INICIO_TESTES")
	ordena=2
case else
	set rs1=db.execute("SELECT DISTINCT PCTE_TX_PACT_TESTE, PCTE_DT_INICIO_TESTES AS DATAINI FROM PACOTE_TESTES ORDER BY PCTE_TX_PACT_TESTE")
	ordena=1
end select
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<script language="JavaScript" src="pupdate.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#CC3300" vlink="#CC3300" alink="#CC3300">
<form name="frm1" method="POST" action="rel_desenv_atrasado.asp">
  <table width="903" height="86" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="margin-bottom: 0">
    <tr> 
      <td height="66" colspan="2">&nbsp;</td>
      <td height="66" colspan="2">&nbsp;</td>
      <td valign="top" colspan="2" height="66"> 
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
      <td height="20" width="1">&nbsp; </td>
      <td height="20" width="1">&nbsp;</td>
      <td height="20" width="625"><table width="342" border="0" align="center">
          <tr> 
            <td width="26">&nbsp;</td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28">&nbsp;</td>
            <td width="26">&nbsp;</td>
            <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          </tr>
        </table></td>
      <td colspan="2" height="20">&nbsp;
        
      </td>
      <td height="20" width="274">&nbsp;</td>
    </tr>
  </table>
  <table width="90%" border="0">
    <tr> 
      <td width="21%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif">Desenvolvimentos p&oacute;s Cen&aacute;rios</font></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center">&nbsp;</div></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <p><font color="#330099" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo 
    de Visualiza&ccedil;&atilde;o :</strong></font> 
    <select name="selvisual" onChange="document.frm1.submit()">
	<%select case visual
	case 1
	%>
	  <option selected value="1">Por Chamada de Desenvolvimento - Por Pacote de Testes</option>
      <option value="2">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio  - Por Pacote de Testes</option>
	  <option value="3">Por Chamada de Desenvolvimento - Por Data de Inicio de Testes</option>
      <option value="4">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio -  - Por Data de Inicio de Testes</option>
	<%case 2%>
	  <option value="1">Por Chamada de Desenvolvimento - Por Pacote de Testes</option>
      <option selected value="2">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio  - Por Pacote de Testes</option>
	  <option value="3">Por Chamada de Desenvolvimento - Por Data de Inicio de Testes</option>
      <option value="4">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio - Por Data de Inicio de Testes</option>
	<%case 3%>
	  <option value="1">Por Chamada de Desenvolvimento - Por Pacote de Testes</option>
      <option value="2">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio  - Por Pacote de Testes</option>
	  <option selected value="3">Por Chamada de Desenvolvimento - Por Data de Inicio de Testes</option>
      <option value="4">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio - Por Data de Inicio de Testes</option>
	<%case 4%>
	  <option value="1">Por Chamada de Desenvolvimento - Por Pacote de Testes</option>
      <option value="2">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio  - Por Pacote de Testes</option>
	  <option value="3">Por Chamada de Desenvolvimento - Por Data de Inicio de Testes</option>
      <option selected value="4">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio - Por Data de Inicio de Testes</option>
	<%case else%>
	  <option selected value="1">Por Chamada de Desenvolvimento - Por Pacote de Testes</option>
      <option value="2">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio  - Por Pacote de Testes</option>
	  <option value="3">Por Chamada de Desenvolvimento - Por Data de Inicio de Testes</option>
      <option value="4">Por Transa&ccedil;&atilde;o Associada ao Cen&aacute;rio - Por Data de Inicio de Testes</option>
	<%end select%>
    </select>
  </p>
  <table width="80%" border="0" cellpadding="0" cellspacing="0" bordercolor="#999999">
    <tr bgcolor="#330099">
      <td width="10%" height="27"> <div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Pacote</font></strong></div></td>
      <td width="18%"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
          de In&iacute;cio de Testes</font></strong></div></td>
      <td width="14%"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cen&aacute;rio</font></strong></div></td>
      <td width="15%"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Desenvolvimentos</font></strong></div></td>
      <td width="20%"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
          Prevista de T&eacute;rmino</font></strong></div></td>
      <td width="14%"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Atraso 
          em dias</font></strong></div></td>
    </tr>
    <%
	reg1=rs1.recordcount
	i1=0

	do until i1=reg1
	
	set rsversao=db.execute("SELECT MAX(PCTE_NM_VERSAO)AS VERSAO FROM PACOTE_TESTES WHERE PCTE_TX_PACT_TESTE='" & rs1("PCTE_TX_PACT_TESTE") & "'")
	
	if isnull(rsversao("versao")) then
		versao=1
	else
		versao=rsversao("versao")
	end if
	
	'set rsdata=db.execute("SELECT MAX(PCTE_DT_INICIO_TESTES)AS DATAINI FROM PACOTE_TESTES WHERE PCTE_TX_PACT_TESTE='" & rs1("CENA_TX_PACT_TESTE") & "' AND PCTE_NM_VERSAO=" & VERSAO)
	
	'data_comp= year(rsdata("dataini")) & "-" & right("000" & month(rsdata("dataini")),2) & "-" & right("000" & day(rsdata("dataini")),2)
	'data_comp2=rsdata("dataini")
	
	data_comp= year(rs1("dataini")) & "-" & right("000" & month(rs1("dataini")),2) & "-" & right("000" & day(rs1("dataini")),2)
	data_comp2=rs1("dataini")
	
	if data_comp<>"-00-00" then

	if ordena=2 then
		ssql=""
		ssql="SELECT  DISTINCT dbo.CENARIO.CENA_TX_PACT_TESTE, dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO, "
		ssql=ssql+"                      dbo.DESENVOLVIMENTO.DESE_CD_DESENVOLVIMENTO, dbo.DESENVOLVIMENTO.DESE_DT_PREVISTA_REALIZACAO "
		ssql=ssql+"FROM         dbo.CENARIO INNER JOIN "
		ssql=ssql+"                      dbo.CENARIO_TRANSACAO ON dbo.CENARIO.CENA_CD_CENARIO = dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO INNER JOIN "
		ssql=ssql+"                      dbo.TRANSACAO_DESENV ON dbo.CENARIO_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO_DESENV.TRAN_CD_TRANSACAO INNER JOIN "
		ssql=ssql+"                      dbo.DESENVOLVIMENTO ON "
		ssql=ssql+"                      dbo.TRANSACAO_DESENV.DESE_CD_DESENVOLVIMENTO = dbo.DESENVOLVIMENTO.DESE_CD_DESENVOLVIMENTO "
		ssql=ssql+"WHERE     (NOT (dbo.TRANSACAO_DESENV.TRAN_CD_TRANSACAO IS NULL)) AND (NOT (dbo.CENARIO.CENA_TX_PACT_TESTE IS NULL) AND "
		ssql=ssql+"                      dbo.CENARIO.CENA_TX_PACT_TESTE = '" & rs1("PCTE_TX_PACT_TESTE") & "') AND (dbo.DESENVOLVIMENTO.DESE_DT_PREVISTA_REALIZACAO > CONVERT(DATETIME, "
		ssql=ssql+"                      '" & data_comp &" 00:00:00', 102)) AND (dbo.DESENVOLVIMENTO.DESE_DT_CONCLUSAO IS NULL)"
		ssql=ssql+"ORDER BY dbo.CENARIO.CENA_TX_PACT_TESTE, dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO	"
	else
		ssql=""
		ssql="SELECT DISTINCT dbo.CENARIO.CENA_TX_PACT_TESTE, dbo.CENARIO.CENA_CD_CENARIO, "
		ssql=ssql+"                      dbo.CENARIO_TRANSACAO.DESE_CD_DESENVOLVIMENTO, dbo.DESENVOLVIMENTO.DESE_DT_PREVISTA_REALIZACAO "
		ssql=ssql+"FROM         dbo.CENARIO INNER JOIN "
		ssql=ssql+"                      dbo.CENARIO_TRANSACAO ON dbo.CENARIO.CENA_CD_CENARIO = dbo.CENARIO_TRANSACAO.CENA_CD_CENARIO INNER JOIN "
		ssql=ssql+"                      dbo.DESENVOLVIMENTO ON "
		ssql=ssql+"                      dbo.CENARIO_TRANSACAO.DESE_CD_DESENVOLVIMENTO = dbo.DESENVOLVIMENTO.DESE_CD_DESENVOLVIMENTO "
		ssql=ssql+"WHERE     (NOT (dbo.CENARIO.CENA_TX_PACT_TESTE IS NULL) AND dbo.CENARIO.CENA_TX_PACT_TESTE = '" & rs1("PCTE_TX_PACT_TESTE") & "') AND "
		ssql=ssql+"                      (dbo.DESENVOLVIMENTO.DESE_DT_PREVISTA_REALIZACAO > CONVERT(DATETIME, '" & data_comp &" 00:00:00', 102)) AND (dbo.DESENVOLVIMENTO.DESE_DT_CONCLUSAO IS NULL)"
		ssql=ssql+"ORDER BY dbo.CENARIO.CENA_TX_PACT_TESTE, dbo.CENARIO.CENA_CD_CENARIO "
	end if
	
	'response.write ssql & "<p>"
	
	err.clear
	on error resume next
	set fonte=db.execute(ssql)
	if err.number<>0 then
		do until err.number=0
			err.clear
			set fonte=db.execute(ssql)
		loop
	end if
	err.clear	
	
	pcte=rs1("PCTE_TX_PACT_TESTE")
	datai=day(data_comp2) & "/" & month(data_comp2) &"/"& year(data_comp2)
		
	anterior=""
	atual=""
	
	reg=fonte.recordcount
	i=0
	
	do until i=reg
	
		atual=fonte("CENA_CD_CENARIO")		

		if cor="#DDDDDD" then
			cor="white"
		else
			cor="#DDDDDD"
		end if
	%>
    <tr> 
      <td height="16" bgcolor="<%=cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=UCASE(pcte)%></b></font></td>
      <td bgcolor="<%=cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=datai%></b></font></td>
      <%if anterior<>atual then%>
      <td bgcolor="<%=cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=UCASE(fonte("CENA_CD_CENARIO"))%></b></font></td>
      <%else%>
      <td bgcolor="<%=cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b></b></font></td>
      <%end if%>
      <td bgcolor="<%=cor%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=UCASE(fonte("DESE_CD_DESENVOLVIMENTO"))%></font></td>
      <%
	  dataf=day(fonte("DESE_DT_PREVISTA_REALIZACAO")) & "/"& month(fonte("DESE_DT_PREVISTA_REALIZACAO")) & "/" & year(fonte("DESE_DT_PREVISTA_REALIZACAO"))
	  %>
      <td bgcolor="<%=cor%>"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=dataf%></font></div></td>
      <td width="9%" bgcolor="<%=cor%>"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=-(datediff("d",fonte("DESE_DT_PREVISTA_REALIZACAO"),data_comp2))%></font></div></td>
    </tr>
    <%
			pcte=" "
			datai=" "
			anterior=fonte("CENA_CD_CENARIO")
			tem=tem+1
			i=i+1
			fonte.movenext
		loop
		end if
		i1=i1+1
		rs1.movenext
	loop
	%>
  </table>
  </form>
<%if tem=0 then%>
<p><b><font color="#800000">Nenhum Registro Encontrado</font></b></p>
<%ELSE%>
<p><b><font color="#800000">Registros Encontrados: <%=tem%></font></b></p>
<%end if%>
</body>
</html>