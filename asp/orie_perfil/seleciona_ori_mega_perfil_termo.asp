<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Opt = Request("pOpt")
if Request("selMegaProcesso") = "" then
   str_MegaProcesso = "0"
else
	str_MegaProcesso = Request("selMegaProcesso")
end if
'response.Write("  ainicio ")
'response.Write(str_Opt)
'response.Write(" fim ")
'if str_MegaProcesso <> "0" then
'   Session("MegaProcesso") = str_MegaProcesso
'else
'    if Session("MegaProcesso") <> "" then
'       str_MegaProcesso = Session("MegaProcesso") 
'	end if   
'end if

'RESPONSE.Write(Session("Conn_String_Cogest_Gravacao"))
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " MEPR_CD_MEGA_PROCESSO"
str_SQL = str_SQL & " , ORTE_NR_SEQUENCIAL"
str_SQL = str_SQL & " , ORTE_TX_TERMO"
str_SQL = str_SQL & " FROM  PERFIL_ORIEN_MEGA_TERMOS"
str_SQL = str_SQL & " WHERE  MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso

%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}

function manda1()
{
//alert(" entrei")
//alert(document.frm1.txtOpt.value)
window.location.href='seleciona_ori_mega_perfil_termo.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&pOpt='+document.frm1.txtOpt.value
}

function Confirma() 
{ 
  if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
  if (document.frm1.selTermo.selectedIndex == 0)
     { 
	 alert("A seleção de um Termo é obrigatório!");
     document.frm1.selTermo.focus();
     return;	 
	 }	 
  //else
	// {
	 //alert(document.frm1.txtOpt.value)
     if(document.frm1.txtOpt.value == "AT")
       {
       document.frm1.action="alterar_excluir_ori_mega_mape_perfil_termos.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "ET")
       {
       document.frm1.action="alterar_excluir_ori_mega_mape_perfil_termos.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     //}  
 }

function Limpa(){
	document.frm1.reset();
}

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
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
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"><a href="javascript:Limpa()"><img src="../../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">limpa</font></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="96%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="15%">&nbsp;</td>
      <td width="75%">&nbsp;</td>
    </tr>
	<% str_Acao = ""
	   If str_Opt = "AT" then
	      str_Acao = "ALTERAÇÃO"
	   elseIf str_Opt = "ET" then
	      str_Acao = "EXCLUSÃO"
	   elseIf str_Opt = "ZT" then
	      str_Acao = "XXXXX"	   
	   end if
	%>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="15%">&nbsp;</td>
      <td width="75%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Selecione 
        o Mega-Processo - Termo - Perfil - <%=str_Acao%></font></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="15%">&nbsp;</td>
      <td width="75%"><div align="right"><a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>"><img src="../../imagens/conteudo_01.gif" width="18" height="22" border="0"></a><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>">Relat&oacute;rio 
          completo</a> </font> </div></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="15%"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Mega-Processo 
          : </b></font></div></td>
      <td width="75%"> <select name="selMegaProcesso"  onChange="javascript:manda1()">
          <option value="0">== Selecione um Mega Processo ==</option>
          <% Set rdsMegaProcesso= Conn_db.Execute(str_SQL_MegaProc)
While (NOT rdsMegaProcesso.EOF)
           if (Trim(str_MegaProcesso) = Trim(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)) then %>
               <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>" selected ><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
       <% else %>
               <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>"><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
       <% end if %>
          <%
  rdsMegaProcesso.MoveNext()
Wend
If (rdsMegaProcesso.CursorType > 0) Then
  rdsMegaProcesso.MoveFirst
Else
  rdsMegaProcesso.Requery
End If

rdsMegaProcesso.Close
set rdsMegaProcesso = Nothing
%>
        </select></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="15%"> <input type="hidden" name="txtOpt" value="<%=str_Opt%>"> 
      </td>
      <td width="75%"> <table width="89%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%">&nbsp;</td>
            <td width="39%"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"></font></div></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="15%"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Selecione 
          Termo :</b></font></div></td>
      <td width="75%"><select name="selTermo" size="1" id="selTermo">
          <option value="0">== Selecione um Termo ==</option>
          <% set rs=Conn_db.execute(str_SQL)
		     do until rs.eof=true %>
          <option value="<%=rs("ORTE_NR_SEQUENCIAL")%>"><%=rs("ORTE_NR_SEQUENCIAL")%> 
          - <%=Left(rs("ORTE_TX_TERMO"),90)%></option>
          <% rs.movenext
			loop %>
        </select></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
