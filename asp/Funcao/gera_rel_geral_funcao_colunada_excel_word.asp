<%@LANGUAGE="VBSCRIPT"%>  
<%
server.ScriptTimeout=99999999
str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")

if str_Uso = "" then
   str_Uso = 0
end if   
if str_Desuso = "" then
   str_Desuso = 0
end if   

if request("opt") = 1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
elseif request("opt") = 2 then
	Response.Buffer = TRUE
	Response.ContentType = "application/msword"
end if

RESPONSE.Write("<p>" & str_Uso)
RESPONSE.Write("<p>" & str_Desuso)

if str_Uso = 1 and str_Desuso = 1 then
   str_usoDesuso =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	end if        	  
end if

RESPONSE.Write("<p>" & str_usoDesuso)

str_Opc = Request("txtOpc")
str_MegaProcesso = request("selMegaProcesso")
str_Modulo=request("selSubModulo")
'response.Write(str_Modulo)
selG=request("selG")

if str_modulo<>"0" then
	compl1=" AND FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA=" & str_modulo 
end if

'response.write selG
if selG=1 then
	compl1=compl1 + " AND FUNCAO_NEGOCIO.FUNE_TX_TP_FUN_NEG ='G'"
else
   selG = 0	
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso)

if str_MegaProcesso=0 then
	IF selG=1 then
		ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_TX_TP_FUN_NEG ='G' " & str_usoDesuso  '& " ORDER BY MEPR_CD_MEGA_PROCESSO, FUNE_TX_TITULO_FUNCAO_NEGOCIO"
	else
		ssql="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO  where MEPR_CD_MEGA_PROCESSO > 0 " & str_usoDesuso ' & " ORDER BY MEPR_CD_MEGA_PROCESSO, FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	end if
	ssql = ssql & " order by FUNE_TX_TITULO_FUNCAO_NEGOCIO "
else
	sql=""
	ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_DESC_FUNCAO_NEGOCIO "
	ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
	ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
	ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
	ssql=ssql+"WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & compl1 & str_usoDesuso
	if request("selFuncao")<>"0" then
		ssql=ssql+" AND FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO='" & request("selFuncao") & "' ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "
	else
		ssql=ssql+" ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "	
	end if	
end if

response.write ssql

set rs=db.execute(ssql)
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
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

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
 <input type="hidden" name="txtOpc" value="1">
  <table width="74%" height="62" border="1" cellpadding="1" cellspacing="1">
    <tr valign="top">
      <td width="14%"><strong><font face="Verdana" size="2">C&oacute;digo</font></strong></td>
      <td width="16%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="2">Fun&ccedil;&atilde;o</font></strong></td>
      <td width="27%" height="19"> <p style="margin-top: 0; margin-bottom: 0"><strong><font face="Verdana" size="2">Descrição</font></strong></td>
      <td width="24%"><strong><font face="Verdana" size="2">&Aacute;rea de Abrang&ecirc;ncia</font></strong></td>
      <td width="28%"><strong><font face="Verdana" size="2">Observa&ccedil;&atilde;o 
        Espec&iacute;fica para Fun&ccedil;&atilde;o</font></strong></td>
    </tr>
    <%
        conta=0
        DO UNTIL RS.EOF=TRUE
		%>
    <tr valign="top">
      <td width="14%"><font face="Verdana" size="2"><%=RS("FUNE_CD_FUNCAO_NEGOCIO")%></font></td>
      <td width="16%" height="19"><font face="Verdana" size="1">&nbsp;<%=RS("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
      <td width="27%" height="19"><font face="Verdana" size="1"><%=RS("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font></td>
      <%
		str_SQL = ""
		str_SQL = str_SQL & " SELECT dbo.ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO "
        str_SQL = str_SQL & " FROM  dbo.FUN_NEG_ORG_AGLU INNER JOIN "
        str_SQL = str_SQL & " dbo.ORGAO_AGLUTINADOR ON dbo.FUN_NEG_ORG_AGLU.AGLU_CD_AGLUTINADO = dbo.ORGAO_AGLUTINADOR.AGLU_CD_AGLUTINADO "
        str_SQL = str_SQL & " WHERE (dbo.FUN_NEG_ORG_AGLU.FUNE_CD_FUNCAO_NEGOCIO = '" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "')"
		set rdsAreaAbran = db.Execute(str_SQL)
		str_AreaAbrang = ""
		if not rdsAreaAbran.EOF then
		   str_AreaAbrang = rdsAreaAbran("AGLU_SG_AGLUTINADO")
		   rdsAreaAbran.movenext
		end if   
		do while not rdsAreaAbran.EOF
		   str_AreaAbrang = str_AreaAbrang & ", " & rdsAreaAbran("AGLU_SG_AGLUTINADO")
		   rdsAreaAbran.movenext
		loop
		rdsAreaAbran.close
		set rdsAreaAbran = Nothing
		%>
      <td width="24%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=str_AreaAbrang%></font></td>
      <%
	  set temp=db.execute("SELECT FUNE_TX_OBS_ESPECIFICA FROM FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "'")
	  	  if temp.eof=false then
	  %>
	  <td width="28%"><font face="Verdana" size="1"><%=temp("FUNE_TX_OBS_ESPECIFICA")%></font></td>
    </tr>
    <%
	end if
        conta=conta+1
			RS.MOVENEXT        
        	LOOP
        %>
  </table>
</body>
</html>