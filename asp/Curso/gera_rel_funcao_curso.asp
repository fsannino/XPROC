 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MEGA=REQUEST("selMegaProcesso")
FUNCAO=REQUEST("SelFuncao")

ON ERROR RESUME NEXT

IF FUNCAO=0  THEN
	IF ERR.NUMBER=0 THEN
	COMPL1=" WHERE MEPR_CD_MEGA_PROCESSO=" & MEGA
ELSE
	COMPL1=" WHERE FUNE_CD_FUNCAO_NEGOCIO='" & FUNCAO & "'"
END IF
END IF

RESPONSE.WRITE COMPL1

SSQL="SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO" & COMPL1 & " ORDER BY MEPR_CD_MEGA_PROCESSO, FUNE_CD_FUNCAO_NEGOCIO"

set rs=DB.EXECUTE(SSQL)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
          <td width="26"></td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>
      </td>
    </tr>
    <tr>
      <td>
        <div align="center">
          <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório
          de Fun&ccedil;&atilde;o R/3 x Curso</font></div>
      </td>
    </tr>
  </table>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;
<table border="0" width="100%">
  <tr>
    <td width="33%" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Mega-Processo</b></font></td>
    <td width="33%" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Fun&ccedil;&atilde;o R/3</b></font></td>
    <td width="34%" bgcolor="#330099"><font face="Verdana" size="2" color="#FFFFFF"><b>Curso</b></font></td>
  </tr>
  <%
  DO UNTIL RS.EOF=TRUE
  
  set mega_=db.execute("SELECT * FROM COGEST.MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & MEGA)
 
  FUNCAO_=RS("FUNE_CD_FUNCAO_NEGOCIO")
  NOME_=RS("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
  
  %>
  <tr>
    <td width="33%"><font face="Verdana" size="1"><%=RS("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
    <td width="33%"><font face="Verdana" size="1"><%=NOME_%></font></td>
    <td width="34%"><font face="Verdana" size="1">sss</font></td>
  </tr>
  <%
	RS.MOVENEXT  
  	LOOP
  %>
</table>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;
<b>
  </form>

</body>

</html>
