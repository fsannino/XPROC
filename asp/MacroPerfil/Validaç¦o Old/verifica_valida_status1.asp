<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso")  & " AND MCPE_TX_SITUACAO='EE' OR MCPE_TX_SITUACAO='NA' OR MCPE_TX_SITUACAO='EA'")
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

function Confirma()
{
document.frm1.submit();
}

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
<form name="frm1" method="post" action="valida_status1.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="39" valign="middle" align="center">
              <div align="center">
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Cenario/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Cenario/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Cenario/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../Cenario/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Cenario/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../Cenario/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
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
      </table>
    </td>
  </tr>
</table>
        <p align="center"><font color="#330099" face="Verdana" size="3">Encaminhamento
        de Status :&nbsp; Em Elaboração -&gt; Em Aprovação</font></p>
        <%SET TEMP=CONN_DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso"))%>
        <p align="left"><font color="#330099" face="Verdana" size="2"><b>Mega-Processo
        Selecionado : </b><%=request("selMegaProcesso")%>  - <input type="hidden" name="mega" size="20" value="<%=request("selMegaProcesso")%>"><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font></p>
        <table border="0" width="88%">
          <tr>
            <td width="16%" bgcolor="#330099" align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em
              Elaboração</font></b></td>
            <td width="16%" bgcolor="#330099" align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Não
              Aprovado</font></b></td>
            <td width="16%" bgcolor="#330099" align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em
              Aprovação</font></b></td>
            <td width="21%" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Macro
              - Perfil</font></b></td>
            <td width="84%" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Descrição</font></b></td>
          </tr>
          <%
          valor1=""
          valor2=""
          valor3=""
          
          tem=0
          
          DO UNTIL RS.EOF=TRUE
			
			select case rs("MCPE_TX_SITUACAO")
			
			case "EE"
				VALOR1="checked"
			case "NA"
				VALOR2="checked"
			case "EA"
				VALOR3="checked"
			end select

          %>
          <tr>
            <td width="16%" align="center">
              <input type="radio" value="1" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" <%=valor1%>></td>
            <td width="16%" align="center">
              <input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="2" <%=valor2%>></td>
            <td width="16%" align="center">
              <p align="center"><input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="3" <%=valor3%>></td>
            <td width="21%"><font color="#330099" face="Verdana" size="1"><%=RS("MCPE_TX_NOME_TECNICO")%></font></td>
            <td width="84%"><font color="#330099" face="Verdana" size="1"><%=RS("MCPE_TX_DESC_MACRO_PERFIL")%></font></td>
            </tr>
            <%
            tem=tem+1
            
            valor1=""
            valor2=""
            VALOR3=""
            
            RS.MOVENEXT
            LOOP
            %>
        </table>
        <%if tem=0 then%>
        <font color="#800000"><b>
        Nenhum Registro Encontrado!</b></font>
        <%end if%>
  </form>
</body>
</html>
