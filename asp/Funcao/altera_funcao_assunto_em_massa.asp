<%
if request("opt")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")

if str_Uso = 1 and str_Desuso = 1 then
   str_usoDesuso =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	end if        	  
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
db.cursorlocation=3

if request("selMegaProcesso") <> "" then
   str_MegaProcesso=request("selMegaProcesso")
else
   str_MegaProcesso = "0"
end if   

'response.Write("<p>" & request("selSubModulo") & "<p>")

if request("selSubModulo") <> "" then
   str_Assunto=request("selSubModulo")
else
   str_Assunto= "0"
end if   

if request("selFuncao") <> "" then
   str_Funcao=request("selFuncao")
else
   str_Funcao = "0"
end if

if str_MegaProcesso<>"0" then
	compl=compl+" AND FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso 
end if

if str_Assunto<>"0" then
	compl=" AND FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA =" & str_Assunto 
end if

if str_Funcao<>"0" then
	compl=compl+" AND FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO='" & str_Funcao &"'"
end if

str_SQL = ""
str_SQL = "SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
str_SQL = str_SQL & " FROM FUNCAO_NEGOCIO_SUB_MODULO "
str_SQL = str_SQL & " INNER JOIN FUNCAO_NEGOCIO ON "
str_SQL = str_SQL & " FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "

str_SQL = str_SQL & compl & str_usoDesuso

'response.write str_sql

set rs=db.execute(str_SQL)

qtos = rs.recordcount

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
<%
conta=1
do until conta = qtos+1
%>
function carrega_txt<%=conta%>(fbox) 
{
document.frm1.txtAssunto<%=conta%>.value = "";
for(var i=0; i<fbox.options.length; i++) 
{
document.frm1.txtAssunto<%=conta%>.value = document.frm1.txtAssunto<%=conta%>.value + "," + fbox.options[i].value;
}
}
<%
conta=conta+1
loop
%>

function Confirma() 
{
<%
conta=1
do until conta = qtos+1
%>
if(document.frm1.list<%=conta%>.options.length == 0 )
     {
     alert("É obrigatória a seleção de, pelo menos, um ASSUNTO!");
	 document.frm1.selAssunto<%=conta%>.focus();
return;
}
<%
conta=conta+1
loop
%>
else
{
<%
conta=1
do until conta = qtos+1
%>
carrega_txt<%=conta%>(document.frm1.list<%=conta%>)
<%
conta=conta+1
loop
%>
document.frm1.submit();
}
}

function Limpa()
{
	document.frm1.reset();
}
//  End -->
</script>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>

<script language="javascript" src="../js/troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#CC3300" vlink="#CC3300" alink="#CC3300">
<form name="frm1" method="POST" action="grava_alteracao_assunto_massa.asp">
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
      <td height="20" width="0%">&nbsp; </td>
      <td height="20" width="0%">&nbsp;</td>
      <td height="20" width="80%"> 
        <table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img src="../Cenario/confirma_f02.gif" width="24" height="24" border="0"></a></td>
            <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
            <td width="26">&nbsp;</td>
            <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28">&nbsp;</td>
            <td width="26">&nbsp;</td>
            <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          </tr>
        </table></td>
      <td colspan="2" height="20">&nbsp;
        
      </td>
      <td height="20" width="20%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0">
    <tr> 
      <td width="21%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="center"><font color="#330099" size="3" face="Verdana"> Assunto 
          de Fun&ccedil;&otilde;es</font></div></td>
      <td align="left"><img src="../../imagens/carregando01.gif" name="imagem1" width="120" height="18" id="imagem1"></td>
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
  <table width="85%" height="147" border="0" cellspacing="3">
  <% str_Contador = 0
     do until rs.eof=true
       str_Contador = str_Contador + 1
  %>
    <tr> 
      <td width="7%"><b></b></td>
      <td colspan="3"><font face="Verdana" size="2"><a href="../funcao/exibe_dados_funcao.asp?selMegaProcesso=<%=str_MegaProcesso%>&selFuncao=<%=rs("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs("FUNE_CD_FUNCAO_NEGOCIO")%></a>- <%=rs("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
    </tr>
    <tr> 
      <td width="7%" rowspan="5"><b><font face="Verdana" size="1" color="#FFFFFF">Descrição 
        do Cenário</font></b></td>
      <td width="30%" rowspan="5"><font face="Verdana" size="1"> 
        <input name="txtFuncao<%=str_Contador%>" type="hidden" value="<%=rs("FUNE_CD_FUNCAO_NEGOCIO")%>">
        <select name="selAssunto<%=str_Contador%>" size="5" multiple>
        <% rs_assunto.movefirst
		  do while not rs_assunto.EOF 
		  set temp1=db.execute("SELECT * FROM FUNCAO_NEGOCIO_SUB_MODULO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' AND SUMO_NR_CD_SEQUENCIA=" & rs_assunto("SUMO_NR_CD_SEQUENCIA"))
		  if temp1.eof=true then
		  %>
          <option value="<%=rs_assunto("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_assunto("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
		  end if
		  rs_assunto.movenext 
		  Loop %>
        </select>
        </font></td>
      <td width="9%" height="21"><div align="center"></div></td>
	  
      <td width="54%" rowspan="5">
	  <input name="txtAssunto<%=str_Contador%>" type="hidden" value="">
	  <select name="list<%=str_Contador%>" size="5" multiple>
	  <%
		ssql=""
		ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA, SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
		ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
		ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
		ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
		ssql=ssql+"INNER JOIN SUB_MODULO ON "		
		ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA= SUB_MODULO.SUMO_NR_CD_SEQUENCIA "
		ssql=ssql+"WHERE FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' "		
		ssql=ssql+"ORDER BY SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
		
		set temp=db.execute(ssql)
		
		do until temp.eof=true
		%>
        <option value="<%=temp("SUMO_NR_CD_SEQUENCIA")%>"><%=temp("SUMO_TX_DESC_SUB_MODULO")%></option>
 	    <%
		temp.movenext		
		loop
		%>
      </select></td>
    </tr>
    <tr>
      <td><div align="center"><img src="continua_F01.gif" width="24" height="24" onClick="move(document.frm1.selAssunto<%=str_Contador%>,document.frm1.list<%=str_Contador%>,1)"></div></td>
    </tr>
    <tr>
      <td><div align="center"><img src="continua2_F01.gif" width="24" height="24" onClick="move(document.frm1.list<%=str_Contador%>,document.frm1.selAssunto<%=str_Contador%>,1)"></div></td>
    </tr>
    <tr>
      <td><div align="center"></div></td>
    </tr>
    <tr>
      <td><div align="center"></div></td>
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
