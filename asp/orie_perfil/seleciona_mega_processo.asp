<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Opt = Request("pOpt")
str_Opt2 = Request("pOpt2")
'response.Write(" opt1  ")
'response.Write(str_Opt)
'response.Write(" fimopt1  ")
'response.Write(" opt2  ")
'response.Write(str_Opt2)
'response.Write(" fimopt2  ")

if Request("selMegaProcesso") = "" then
   str_MegaProcesso = "0"
else
	str_MegaProcesso = Request("selMegaProcesso")
end if

if Request("selSubModulo") = "" then
   str_SubModulo = "0"
else
	str_SubModulo = Request("selSubModulo")
end if
'response.Write(" =====submodu  ")
'response.Write(str_SubModulo)
'response.Write(" fimsubmodu  ")

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
'if str_Opt = "IM" or str_Opt = "AM" or str_Opt = "EM"  then
'  str_SQL_MegaProc = str_SQL_MegaProc & " and MEPR_TX_INDICA_SUB_MODULO = '1' "
'end if   
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

SQL_Assunto=""
SQL_Assunto = SQL_Assunto & " SELECT SUMO_NR_CD_SEQUENCIA"
SQL_Assunto = SQL_Assunto & " ,SUMO_TX_DESC_SUB_MODULO"
SQL_Assunto = SQL_Assunto & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
SQL_Assunto = SQL_Assunto & " FROM " & Session("PREFIXO") & "SUB_MODULO"
if str_MegaProcesso <> 0 then
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
else
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS = '9999'"
end if
SQL_Assunto=SQL_Assunto + " ORDER BY SUMO_TX_DESC_SUB_MODULO"

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
window.location.href='seleciona_mega_processo.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOpt='+document.frm1.txtOpt.value+'&pOpt2='+document.frm1.txtOpt2.value     
}

function Confirma() 
{ 
  if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatória!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
  //alert(document.frm1.selSubModulo.selectedIndex)
  //alert(document.frm1.txtOpt2.value)	 
  if (document.frm1.txtOpt2.value == "M")
     { 
     if (document.frm1.selSubModulo.selectedIndex == 0)
        { 
	    alert("A seleção de um Sub-Módulo é obrigatória!");
        document.frm1.selSubModulo.focus();
        return;
        }
     }
	 
  //else
	// {
	 //alert(document.frm1.txtOpt.value)
     if(document.frm1.txtOpt.value == "IO")
       {
       document.frm1.action="inclui_ori_mega_mape_perfil.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "IT")
       {
       document.frm1.action="inclui_ori_mega_mape_perfil_termos.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "RM")
       {
       document.frm1.action="relat_ori_gerais_mega_mapeamento.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "IM")
       {
       document.frm1.action="inclui_ori_mega_mape_perfil_submodulo.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "AM")
       {
       document.frm1.action="alterar_excluir_ori_mega_mape_perfil_submodulo.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "EM")
       {
       document.frm1.action="alterar_excluir_ori_mega_mape_perfil_submodulo.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
 }

function Limpa(){
	document.frm1.reset();
}

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="inclui_ori_mega_mape_perfil.asp">
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
  <table width="88%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="77%">&nbsp;</td>
    </tr>
	<%
	if str_Opt = "IO" then
	   str_Titulo = " Mega-Processo - Incluir Orientações por Mega "
	elseif str_Opt = "IT" then
	   str_Titulo = " Mega-Processo - Incluir Termos por Mega "	
	elseif str_Opt = "RM" then
	   str_Titulo = " Mega-Processo - Relatório por Mega "
	elseif str_Opt = "IM" then
	   str_Titulo = " Mega-Processo/Assunto - Incluir Assunto por Mega "	   		
	elseif str_Opt = "AM" then
	   str_Titulo = " Mega-Processo/Assunto - Alterar Assunto por Mega "	   		
	elseif str_Opt = "EM" then
	   str_Titulo = " Mega-Processo/Assunto - Excluir Assunto por Mega "	   		
	end if
	%>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="18%">&nbsp;</td>
      <td width="77%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Selecione -<%=str_Titulo%>- Perfil </font></td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="18%"><input type="hidden" name="txtOpt" value="<%=str_Opt%>">
        <input type="hidden" name="txtOpt2" value="<%=str_Opt2%>"></td>
      <td width="77%"><div align="right"><a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>"><img src="../../imagens/conteudo_01.gif" width="18" height="22" border="0"></a><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
          <a href="relat_ori_gerais_mega_mapeamento.asp?txtMegaProcesso=<%=str_MegaProcesso%>">Relat&oacute;rio 
          completo</a> </font> </div></td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="18%"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Mega-Processo 
          : </b></font></div></td>
      <td width="77%"> <select name="selMegaProcesso"  onChange="javascript:manda1()">
          <option value="0">== Selecione um Mega Processo ==</option>
          <%Set rdsMegaProcesso= Conn_db.Execute(str_SQL_MegaProc)
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
        </select> </td>
    </tr>
    <tr> 
      <td width="5%">&nbsp;</td>
      <td width="18%">&nbsp; </td>
      <td width="77%"> </td>
    </tr>
    <% 'response.write str_MegaFuncao
	   if InStrRev("11/10", Right("00" & str_MegaProcesso, 2)) = 0 then
	%>	
    <tr> 
      <td>&nbsp;</td>
      <td></td>
      <td>&nbsp;</td>
    </tr>
	<%end if%>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Assunto 
          :</b></font></div></td>
      <td><select size="1" name="selSubModulo">
          <option value="0">== Selecione o Assunto ==</option>
          <%set rs_SubModulo=conn_db.execute(SQL_Assunto)
		  do until rs_SubModulo.eof=true
		  if trim(str_SubModulo)=trim(rs_SubModulo("SUMO_NR_CD_SEQUENCIA")) then
		  %>
          <option selected value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <% else %>
          <option value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
		     end if
					rs_SubModulo.movenext
					loop
					%>
        </select></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  </form>
</body>
</html>
