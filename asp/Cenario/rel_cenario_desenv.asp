<%
response.buffer = false

server.scripttimeout=99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
db.cursorlocation=3

visual=1
visual=request("selvisual")

select case visual
case 1
	ssql=""
	ssql="select distinct "
	ssql=ssql+"dbo.cenario_transacao.CENA_CD_CENARIO, "
	ssql=ssql+"dbo.cenario_transacao.DESE_CD_DESENVOLVIMENTO, "
	ssql=ssql+"dbo.desenvolvimento.DESE_DT_PREVISTA_REALIZACAO, "
	ssql=ssql+"dbo.desenvolvimento.DESE_DT_CONCLUSAO "
	ssql=ssql+"from cenario_transacao "
	ssql=ssql+"inner join dbo.desenvolvimento on "
	ssql=ssql+"dbo.cenario_transacao.dese_cd_desenvolvimento = dbo.desenvolvimento.dese_cd_desenvolvimento "
	ssql=ssql+"order by 1,3,4,2"
	
	ordena=1

case 2
	ssql=""
	ssql="select distinct "
	ssql=ssql+"UPPER(dbo.cenario_transacao.CENA_CD_CENARIO,"
	ssql=ssql+"UPPER(dbo.transacao_desenv.DESE_CD_DESENVOLVIMENTO,"
	ssql=ssql+"dbo.desenvolvimento.DESE_DT_PREVISTA_REALIZACAO, "
	ssql=ssql+"dbo.desenvolvimento.DESE_DT_CONCLUSAO "
	ssql=ssql+"from cenario_transacao "
	ssql=ssql+"inner join transacao_desenv on "
	ssql=ssql+"dbo.cenario_transacao.tran_cd_transacao = dbo.transacao_desenv.tran_cd_transacao "
	ssql=ssql+"inner join dbo.desenvolvimento on "
	ssql=ssql+"dbo.transacao_desenv.dese_cd_desenvolvimento = dbo.desenvolvimento.dese_cd_desenvolvimento "
	ssql=ssql+"order by 1,3,4,2"

	ordena=2

case else
	ssql=""
	ssql="select distinct "
	ssql=ssql+"dbo.cenario_transacao.CENA_CD_CENARIO, "
	ssql=ssql+"dbo.cenario_transacao.DESE_CD_DESENVOLVIMENTO, "
	ssql=ssql+"dbo.desenvolvimento.DESE_DT_PREVISTA_REALIZACAO, "
	ssql=ssql+"dbo.desenvolvimento.DESE_DT_CONCLUSAO "
	ssql=ssql+"from cenario_transacao "
	ssql=ssql+"inner join dbo.desenvolvimento on "
	ssql=ssql+"dbo.cenario_transacao.dese_cd_desenvolvimento = dbo.desenvolvimento.dese_cd_desenvolvimento "
	ssql=ssql+"order by 1,3,4,2"

	ordena=1

end select

set rs1 = db.execute(ssql)
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
<form name="frm1" method="POST" action="rel_cenario_desenv.asp">
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
      <td><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Relação Cenário x Desenvolvimentos</font></div></td>
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
	  <option selected value="1">Por Chamada de Desenvolvimento</option>
      <option value="2">Por Transação Associada ao Cenário</option>
	<%case 2%>
	  <option value="1">Por Chamada de Desenvolvimento</option>
      <option selected value="2">Por Transação Associada ao Cenário</option>
	<%case else%>
	  <option selected value="1">Por Chamada de Desenvolvimento</option>
      <option value="2">Por Transação Associada ao Cenário</option>
	<%end select%>
    </select>
  </p>
  <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#999999" height="42">
    <tr bgcolor="#330099">
      <td width="200" height="21"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Cen&aacute;rio</font></strong></div></td>
      <td width="203" height="21"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Desenvolvimentos</font></strong></div></td>
      <td width="229" height="21"><div align="left"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
          Prevista de T&eacute;rmino</font></strong></div></td>
      <td width="208" height="21"><div align="left"><strong><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Data de Conclusão</font></strong></div></td>
    </tr>
	<%
	i = 0
	reg = rs1.recordcount
	cen_ant=""
	do until i = reg	
	
	if cor="white" then
		cor="#E5E5E5"
	else
		cor="white"
	end if			

	cen_atual = rs1("cena_cd_cenario")
	%>
    <tr> 
      <%if cen_atual<>cen_ant then%>
      <td bgcolor="<%=cor%>" width="200" height="21"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=UCASE(rs1("cena_cd_cenario"))%></b></font></td>
      <%else%>
      <td bgcolor="<%=cor%>" width="154" height="21"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> </b></font></td>
      <%end if%>
      <td bgcolor="<%=cor%>" width="231" height="21"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=UCASE(rs1("dese_cd_desenvolvimento"))%></b></font></td>
      <%
	  dataf=day(rs1("DESE_DT_PREVISTA_REALIZACAO")) & "/"& month(rs1("DESE_DT_PREVISTA_REALIZACAO")) & "/" & year(rs1("DESE_DT_PREVISTA_REALIZACAO"))
	  if dataf="//" then
	  	dataf=" "
	  end if
	  %>
      <td bgcolor="<%=cor%>" width="235" height="21"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=dataf%></font></td>
      <%
	  datas=day(rs1("DESE_DT_CONCLUSAO")) & "/"& month(rs1("DESE_DT_CONCLUSAO")) & "/" & year(rs1("DESE_DT_CONCLUSAO"))
	  if datas="//" then
	  	datas=" "
	  end if
	  %>
       <td bgcolor="<%=cor%>" width="113" height="21"><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=datas%></font></div></td>
    </tr>
    <%
    i = i + 1
    cen_ant =rs1("cena_cd_cenario")
	rs1.movenext
	loop
	%>
  </table>
  </form>
<%if reg < 1 then%>
<p><b><font color="#800000">Nenhum Registro Encontrado</font></b></p>
<%ELSE%>
<p><b><font color="#800000">Registros Encontrados: <%=reg%></font></b></p>
<%end if%>
</body>
</html>