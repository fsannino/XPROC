<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")
proc=request("selProcesso")
subp=request("selSubProcesso")
onda=request("selOnda")
status=request("selStatus")
str_Assunto=0
str_Assunto=request("selAssunto")
str_Escopo = request("selEscopo")

cenario1=request("ID")
cenario2=request("ID2")

if cenario1="0" and cenario2="0" then

if mega<>0 then
	compl=compl+"MEPR_CD_MEGA_PROCESSO=" & mega & " AND "
end if
if proc<>0 then
	compl=compl+"PROC_CD_PROCESSO=" & proc & " AND "
end if
if subp<>0 then
	compl=compl+"SUPR_CD_SUB_PROCESSO=" & subp & " AND "
end if
if onda<>0 then
	compl=compl+"ONDA_CD_ONDA=" & onda & " AND "
ELSE
	compl=compl+"ONDA_CD_ONDA<>4 AND "
end if

if status<>"0" then
	compl=compl+"CENA_TX_SITUACAO='" & status & "' AND "
end if

if str_Escopo<>2 then
		compl=compl+ "CENA_TX_SITUACAO_VALIDACAO=" & str_Escopo & " AND "
end if

if str_Assunto<>0 then
		compl=compl+ "SUMO_NR_CD_SEQUENCIA =" & str_Assunto & " AND "
end if

tamanho=len(compl)
tamanho=tamanho-5
compl=left(compl,tamanho)

else

if cenario1<>"0" then
	compl="CENA_CD_CENARIO='" & cenario1& "'"
	cenario=cenario1
else
if cenario2<>"0" then
	compl="CENA_CD_CENARIO='" & cenario2& "'"
	cenario=cenario2
end if
end if
end if

if len(compl)>0 then
	compl=" WHERE " & compl
end if

ordem=request("ORDER")

if ordem="" then
	ordem="MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO, CENA_CD_CENARIO"
end if

ordem=" ORDER BY " & ordem

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO" & compl & ordem

SSQL1=SSQL

response.write ssql

if request("excel")=1 then
	ssql=request("ssql")
end if

'set rs=db.execute(ssql)

'IF RS.EOF=TRUE THEN
'	TEM=0
'ELSE
'	TEM=1
'END IF

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

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="GERA.ASP">
  <table width="773" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
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
      <td height="20" width="111">&nbsp; </td>
      <td height="20" width="30">&nbsp;</td>
      <td colspan="2" height="20">&nbsp;
      </td>
      <td height="20" width="334">&nbsp;</td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  <p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="3">Relatório
  Geral de Cenário</font> </p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  </form>
<p></p>
</body>
</html>
