<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

SERVER.SCRIPTTIMEOUT = 99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO")
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#INCLUDE file="ADOVBS.INC" -->
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
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../index.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%">&nbsp;</td>
          <td width="64%"><a href="exibe_dono.asp?excel=1" target="_blank"><img border="0" src="../imagens/exp_excel.gif" width="78" height="29"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%" bgcolor="white">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%">
      <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es x Donos</font></p>
    </td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
</table>
<table border="0" cellpadding="0" width="792">
  <tr> 
    <td width="373" bgcolor="#FFFFFF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Transa&ccedil;&atilde;o</font></b></td>
    <td width="55" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>SUP</b></font></div>
    </td>
    <td width="58" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>MES</b></font></div>
    </td>
    <td width="73" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>VEN</b></font></div>
    </td>
    <td width="73" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>EMP</b></font></div>
    </td>
    <td width="59" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>MAN</b></font></div>
    </td>
    <td width="57" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>PO&Ccedil;</b></font></div>
    </td>
    <td width="55" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>PRO</b></font></div>
    </td>
    <td width="57" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>QUA</b></font></div>
    </td>
    <td width="54" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>LOG</b></font></div>
    </td>
    <td width="58" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>PLA</b></font></div>
    </td>
    <td width="54" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>FIN</b></font></div>
    </td>
    <td width="34" bgcolor="#FFFFCC" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>RHU</b> 
      </font></td>
    <td width="59" bgcolor="#FFFFCC" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>GER</b> 
      </font></td>
    <td width="64" bgcolor="#FFFFCC" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>TI</b> 
      </font></td>
  </tr>
  <%
  DO UNTIL RS.EOF=TRUE
  
  SET TMP=DB.EXECUTE("select * from " & Session("PREFIXO") & "transacao_mega where tran_cd_transacao='" & rs("tran_cd_transacao") & "'")

  if tmp.eof=false then

  valor1=""
  valor2=""
  valor3=""
  valor4=""
  valor5=""
  valor6=""
  valor7=""
  valor8=""
  valor9=""
  valor10=""
  valor11=""
  valor12=""
  valor13=""
  valor14=""
 
  do until tmp.eof=true
	  
	select case tmp("MEPR_CD_MEGA_PROCESSO")
		CASE 1
			VALOR1="X"
		CASE 2
			VALOR2="X"
		CASE 3
			VALOR3="X"
		CASE 4
			VALOR4="X"
		CASE 5
			VALOR5="X"
		CASE 6
			VALOR6="X"
		CASE 7
			VALOR7="X"
		CASE 8
			VALOR8="X"
		CASE 9
			VALOR9="X"
		CASE 10
			VALOR10="X"
		CASE 11
			VALOR11="X"
		CASE 12
			VALOR12="X"
		CASE 13
			VALOR13="X"
		CASE 14
			VALOR14="X"

	END SELECT  
  tmp.movenext
  loop
  
  if cor="white" then
  	cor="#D8D8D8"
  else
  	cor="white"
  end if
  
  %>
  <tr bgcolor="<%=cor%>"> 
    <td width="373"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("TRAN_CD_TRANSACAO")%>-<%=rs("TRAN_TX_DESC_TRANSACAO")%></font></td>
    <td width="55" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor1%></font></b> </font></td>
    <td width="58" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor2%></font></b> </font></td>
    <td width="73" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor3%></font></b> </font></td>
    <td width="73" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor4%></font></b> </font></td>
    <td width="59" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor5%></font></b> </font></td>
    <td width="57" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor6%></font></b> </font></td>
    <td width="55" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor7%></font></b> </font></td>
    <td width="57" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor8%></font></b> </font></td>
    <td width="54" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor9%></font></b> </font></td>
    <td width="58" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor10%></font></b> </font></td>
    <td width="54" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor11%></font></b> </font></td>
    <td width="34" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor12%></font></b> </font></td>
    <td width="59" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor13%></font></b> </font></td>
    <td width="64" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor14%></font></b> </font></td>
  </tr>
  <%
  end if
  RS.MOVENEXT
  LOOP
  %>
</table>   
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<%
if request("excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

SERVER.SCRIPTTIMEOUT = 99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO")
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!--#INCLUDE file="ADOVBS.INC" -->
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
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../index.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="36%">&nbsp;</td>
          <td width="64%"><a href="exibe_dono.asp?excel=1" target="_blank"><img border="0" src="../imagens/exp_excel.gif" width="78" height="29"></a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20%" bgcolor="white">&nbsp;</td>
    <td width="62%">&nbsp;</td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%">&nbsp;</td>
    <td width="62%">
      <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Rela&ccedil;&atilde;o 
      de Transa&ccedil;&otilde;es x Donos</font></p>
    </td>
    <td width="18%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
  <tr> 
    <td width="20%"></td>
    <td width="62%"></td>
    <td width="18%"></td>
  </tr>
</table>
<table border="0" cellpadding="0" width="792">
  <tr> 
    <td width="373" bgcolor="#FFFFFF"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Transa&ccedil;&atilde;o</font></b></td>
    <td width="55" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>SUP</b></font></div>
    </td>
    <td width="58" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>MES</b></font></div>
    </td>
    <td width="73" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>VEN</b></font></div>
    </td>
    <td width="73" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>EMP</b></font></div>
    </td>
    <td width="59" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>MAN</b></font></div>
    </td>
    <td width="57" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>PO&Ccedil;</b></font></div>
    </td>
    <td width="55" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>PRO</b></font></div>
    </td>
    <td width="57" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>QUA</b></font></div>
    </td>
    <td width="54" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>LOG</b></font></div>
    </td>
    <td width="58" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>PLA</b></font></div>
    </td>
    <td width="54" bgcolor="#FFFFCC" align="center"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>FIN</b></font></div>
    </td>
    <td width="34" bgcolor="#FFFFCC" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>RHU</b> 
      </font></td>
    <td width="59" bgcolor="#FFFFCC" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>GER</b> 
      </font></td>
    <td width="64" bgcolor="#FFFFCC" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>TI</b> 
      </font></td>
  </tr>
  <%
  DO UNTIL RS.EOF=TRUE
  
  SET TMP=DB.EXECUTE("select * from " & Session("PREFIXO") & "transacao_mega where tran_cd_transacao='" & rs("tran_cd_transacao") & "'")

  if tmp.eof=false then

  valor1=""
  valor2=""
  valor3=""
  valor4=""
  valor5=""
  valor6=""
  valor7=""
  valor8=""
  valor9=""
  valor10=""
  valor11=""
  valor12=""
  valor13=""
  valor14=""
 
  do until tmp.eof=true
	  
	select case tmp("MEPR_CD_MEGA_PROCESSO")
		CASE 1
			VALOR1="X"
		CASE 2
			VALOR2="X"
		CASE 3
			VALOR3="X"
		CASE 4
			VALOR4="X"
		CASE 5
			VALOR5="X"
		CASE 6
			VALOR6="X"
		CASE 7
			VALOR7="X"
		CASE 8
			VALOR8="X"
		CASE 9
			VALOR9="X"
		CASE 10
			VALOR10="X"
		CASE 11
			VALOR11="X"
		CASE 12
			VALOR12="X"
		CASE 13
			VALOR13="X"
		CASE 14
			VALOR14="X"

	END SELECT  
  tmp.movenext
  loop
  
  if cor="white" then
  	cor="#D8D8D8"
  else
  	cor="white"
  end if
  
  %>
  <tr bgcolor="<%=cor%>"> 
    <td width="373"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs("TRAN_CD_TRANSACAO")%>-<%=rs("TRAN_TX_DESC_TRANSACAO")%></font></td>
    <td width="55" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor1%></font></b> </font></td>
    <td width="58" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor2%></font></b> </font></td>
    <td width="73" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor3%></font></b> </font></td>
    <td width="73" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor4%></font></b> </font></td>
    <td width="59" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor5%></font></b> </font></td>
    <td width="57" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor6%></font></b> </font></td>
    <td width="55" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor7%></font></b> </font></td>
    <td width="57" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor8%></font></b> </font></td>
    <td width="54" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor9%></font></b> </font></td>
    <td width="58" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor10%></font></b> </font></td>
    <td width="54" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor11%></font></b> </font></td>
    <td width="34" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor12%></font></b> </font></td>
    <td width="59" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor13%></font></b> </font></td>
    <td width="64" align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#800000"><%=valor14%></font></b> </font></td>
  </tr>
  <%
  end if
  RS.MOVENEXT
  LOOP
  %>
</table>   
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
