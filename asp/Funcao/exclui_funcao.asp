 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=0
str_mega=request("selMegaProcesso")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

if str_mega<>0 then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " ORDER BY FUNE_TX_TITULO_FUNCAO_NEGOCIO")
else
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY FUNE_CD_FUNCAO_NEGOCIO")
 	str_mega=0
end if

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " MEPR_CD_MEGA_PROCESSO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_mega
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "
'response.write str_Sub_Modulo
set rs_SubModulo=db.execute(str_Sub_Modulo)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script>
function manda()
{
window.location.href='exclui_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value
}

function Confirma()
{
if(document.frm1.selMegaProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MEGA-PROCESSO!");
document.frm1.selMegaProcesso.focus();
return;
}
if(document.frm1.selFuncao.selectedIndex == 0)
{
alert("É obrigatória a seleção de uma FUNÇÃO DE NEGÓCIO!");
document.frm1.selFuncao.focus();
return;
}
else
{
if(confirm("Confirma Exclusão(TODOS os registros relacionados serão excluídos - Função de Negócio, Transações da Função, Qualificações da Função, Público Principal da Função)?"))
{ 
document.frm1.submit();
}
}
}
</script>
<body topmargin="0" leftmargin="0">
<form method="POST" action="valida_exclui_funcao.asp" name="frm1">
<input type="hidden" name="txtpub" size="20"><input type="hidden" name="txtQua" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
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
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
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
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Exclusão
        de Fun&ccedil;&atilde;o R/3</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <table border="0" width="829">
    <tr> 
      <td width="76"></td>
      <td width="150"> 
        <div align="right"><b><font face="Verdana" color="#330099" size="2">Mega-Processo 
          : </font></b></div>
      </td>
      <td height="41" width="589"> 
        <select size="1" name="selMegaProcesso" onchange="javascript:manda()">
          <option value="0">== Selecione o Mega-Processo ==</option>
          <%do until rs.eof=true
                	if trim(str_mega)=trim(rs("MEPR_CD_MEGA_PROCESSO")) then
                	%>
          <option selected value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%else%>
          <option value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
					end if
					rs.movenext
					loop
					%>
        </select>
      </td>
    </tr>
	<% If str_mega = 11 then	 
	'if rs_mega("MEPR_CD_MEGA_PROCESSO") = 11 then
	%>
    <tr> 
      <td width="76"></td>
      <td width="150"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b>Sub-Modulo 
          : </b></font></div>
      </td>
      <td width="589"> 
        <select size="1" name="selSubModulo">
          <option value="0">== Selecione o Sub Módulo ==</option>
          <%do until rs_SubModulo.eof=true%>
          <option value="<%=rs_SubModulo("SUMO_NR_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
					rs_SubModulo.movenext
					loop
					%>
        </select>
      </td>
    </tr>
	<% end if %>
    <tr> 
      <td width="76"></td>
      <td width="150"> 
        <div align="right"><b><font face="Verdana" color="#330099" size="2">Fun&ccedil;&atilde;o R/3 : </font></b></div>
      </td>
      <td height="41" width="589"> 
        <select size="1" name="selFuncao">
          <option value="0">== Selecione a Fun&ccedil;&atilde;o R/3 ==</option>
          <%do until rs1.eof=true%>
          <option value="<%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs1("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
					rs1.movenext
					loop
					%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="76"></td>
      <td width="150"></td>
      <td width="589"></td>
    </tr>
  </table>
  </form>

<p>&nbsp;</p>

</body>

</html>
