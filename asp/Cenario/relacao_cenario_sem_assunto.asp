<%
if request("opt")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("selMegaProcesso") <> "" then
   str_MegaProcesso=request("selMegaProcesso")
else
   str_MegaProcesso = "0"
end if   

if request("selAssunto") <> "" then
   str_Assunto=request("selAssunto")
else
   str_Assunto= "0"
end if   

if request("selProcesso") <> "" then
   str_Processo=request("selProcesso")
else
   str_Processo = "0"
end if

if request("selSubProcesso") <> "" then
   str_SubProcesso=request("selSubProcesso")
else
   str_SubProcesso = "0"
end if   

if request("selOnda") <> "" then
   str_Onda=request("selOnda")
else
   str_Onda = "0"
end if

if request("selStatus") <> "" then
   str_Status=request("selStatus")
else
   str_Status= "0"
end if

if request("selCenario") <> "" then
   str_Cenario=request("selCenario")
else
   str_Cenario = "0"
end if   

str_SQL = ""
str_SQL = str_SQL & " SELECT * "
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "CENARIO" 
str_SQL = str_SQL & " where  CENA_NR_SEQUENCIA > 0 " 
str_SQL = str_SQL & " and SUMO_NR_CD_SEQUENCIA IS NULL "

if str_MegaProcesso<>"0" then
	compl=compl+" AND MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso 
end if

if str_Assunto<>"0" then
	compl=" AND SUMO_NR_CD_SEQUENCIA =" & str_Assunto 
end if

if str_Processo<>"0" then
	compl=compl+" AND PROC_CD_PROCESSO=" & str_Processo 
end if

if str_SubProcesso<>"0" then
	compl=compl+" AND SUPR_CD_SUB_PROCESSO=" & str_SubProcesso 
end if

if str_Onda<>"0" then
	compl=compl+" AND ONDA_CD_ONDA=" & str_Onda 
end if

if str_Cenario<>"0" then
	compl=compl+" AND CENA_CD_CENARIO ='" & str_Cenario  & "'"
end if

str_SQL = str_SQL & compl

'response.write str_SQL
'response.write compl

set rs=db.execute(str_SQL)

IF RS.EOF=TRUE THEN
	TEM=0
ELSE
	TEM=1
END IF

str_SQL = ""
str_SQL = str_SQL & " SELECT MEPR_TX_DESC_MEGA_PROCESSO, "
str_SQL = str_SQL & " MEPR_CD_MEGA_PROCESSO"
str_SQL = str_SQL & " FROM MEGA_PROCESSO"
str_SQL = str_SQL & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
set rs_MegaProcesso=db.execute(str_SQL)
str_DsMegaProcesso = rs_MegaProcesso("MEPR_TX_DESC_MEGA_PROCESSO")

SQL_Assunto=""
SQL_Assunto = SQL_Assunto & " SELECT SUMO_NR_CD_SEQUENCIA"
SQL_Assunto = SQL_Assunto & " ,SUMO_TX_DESC_SUB_MODULO"
SQL_Assunto = SQL_Assunto & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
SQL_Assunto = SQL_Assunto & " FROM " & Session("PREFIXO") & "SUB_MODULO"
if str_MegaProcesso <> "0" then
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
else
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS = '9999'"
end if
SQL_Assunto=SQL_Assunto + " ORDER BY SUMO_TX_DESC_SUB_MODULO"

set rs_assunto=db.execute(SQL_Assunto)

%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--

function Confirma() 
{ 
  document.frm1.submit();
 }

function Limpa(){
	document.frm1.reset();
}
//  End -->
</script>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#CC3300" vlink="#CC3300" alink="#CC3300">
<form name="frm1" method="POST" action="grava_alteracao_assunto.asp">
  <table width="100%" height="86" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="margin-bottom: 0">
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
      <td height="20" width="14%">&nbsp; </td>
      <td height="20" width="4%">&nbsp;</td>
      <td height="20" width="28%">&nbsp;</td>
      <td colspan="2" height="20">&nbsp;
        
      </td>
      <td height="20" width="43%">&nbsp;</td>
    </tr>
  </table>
  <table width="76%" border="0">
    <tr> 
      <td width="21%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center"><font color="#330099" size="3" face="Verdana">Cen&aacute;rio 
          sem Assunto</font></div></td>
      <td><img src="../../imagens/carregando01.gif" name="imagem1" width="120" height="18" id="imagem1"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center"><font color="#330099" size="3" face="Verdana">Mega 
          : </font><font face="Verdana" size="3"><%=str_DsMegaProcesso%></font></div></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="866" border="0" cellspacing="3">
  <% str_Contador = 0
     do until rs.eof=true
       str_Contador = str_Contador + 1
  %>
    <tr> 
      <td width="46"><b></b></td>
      <td width="807"><font face="Verdana" size="2"><a href="gera_rel_geral.asp?id=<%=rs("CENA_CD_CENARIO")%>"><%=rs("CENA_CD_CENARIO")%></a>- <%=rs("CENA_TX_TITULO_CENARIO")%></font></td>
    </tr>   
    <%
    rs.movenext
    loop
    %>
	  </table>
    <%if tem=0 then%>
    <font size="2" color="#800000" face="Verdana"><b>Nenhum Registro encontrado 
    para a Seleção</b></font> </p> 
    <%end if%>
    <input name="txtQtdObj" type="hidden" id="txtQtdObj" value="<%=str_Contador%>">
  </p>
  <p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total registro = <%=str_Contador%></font> </p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  </form>
<p></p>
</body>
<script language="JavaScript" type="text/JavaScript">
document.imagem1.src = "../../imagens/carregando_limpa.gif"
</script>
</html>
