<%
Response.Buffer=False
Server.ScriptTimeOut=99999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM FUN_NEG_TRANSACAO ORDER BY FUNE_CD_FUNCAO_NEGOCIO")
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form>
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
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>

<table border="0" width="804">
  <tr>
    <td width="249" bgcolor="#000080">
      <p>
      <b><font face="Verdana" size="2" color="#FFFFFF">Código da Função</font></b>
    </td>
    <td width="741" bgcolor="#000080">
      <b><font face="Verdana" size="2" color="#FFFFFF">Título
      da Função</font></b>
    </td>
    <td width="303" bgcolor="#000080">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2" color="#FFFFFF">&nbsp;</font></b><font color="#FFFFFF" size="2" face="Verdana"><b>Transação</b></font></p>
    </td>
  </tr>
  <%
  do until rs.eof=true
  set fonte=db.execute("SELECT DISTINCT TRAN_CD_TRANSACAO FROM FUN_NEG_TRANSACAO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")
  valor=rs("FUNE_CD_FUNCAO_NEGOCIO")
  SET TEMP=DB.EXECUTE("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")    
  
  On Error Resume next
	  VALOR2=TEMP("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
  Err.Clear	
  
  COR1="#E1E1E1"
  
  do until fonte.eof=true
  
  IF COR="#E1E1E1" THEN
  	COR="WHITE"
  ELSE
   COR="#E1E1E1"
  END IF
  
  %>
  <tr>
    <td width="249" bgcolor="<%=COR1%>">
      <p><font face="Verdana" size="2"><B><%=valor%></B></font></td>
    <td width="741" bgcolor="<%=COR1%>">  
      <font size="2" face="Verdana"><b><%=VALOR2%></b></font></td>
    <td width="303" bgcolor="<%=COR%>">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font size="2" face="Verdana"><%=fonte("TRAN_CD_TRANSACAO")%></font></td>
  </tr>
  <%
  valor=" "
  valor2=" "
  COR1="WHITE"
  fonte.movenext
  loop
  rs.movenext
  loop
  %>
</table>
</form>
</body>
</html>