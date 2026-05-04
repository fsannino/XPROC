<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Opt = Request("pOpt")
str_Opt2 = Request("pOpt2")

str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")
str_usoDesuso = ""
if str_Uso = "" and str_Desuso = "" then
   str_Uso = "true" 
   str_Desuso = "false"
end if   
if str_Uso = "true" and str_Desuso = "true" then
   checado01 = "checked"
   checado02 = "checked"
   str_usoDesuso =  " (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = "false" and str_Desuso = "false" then
      checado01 = ""
      checado02 = ""   
      str_usoDesuso =  " FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '3' "
   else
      if str_Uso = "true" then
         checado01 = "checked"
         checado02 = ""   
         str_usoDesuso =  " FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
      else
         checado01 = ""
         checado02 = "checked"   
         str_usoDesuso =  " FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	  end if        	     
   end if
end if


if Request("selMegaProcesso") = "" then
   str_MegaProcesso = "0"
else
	str_MegaProcesso = Request("selMegaProcesso")
end if

if Request("selSubModulo") = "" or Request("selSubModulo") = "undefined" then
   str_SubModulo = "0"
else
	str_SubModulo = Request("selSubModulo")
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
'if str_Opt <> "RM" then
   str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
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

if str_MegaProcesso<>0 then
	if str_SubModulo <> "" and str_SubModulo <> "undefined" and str_SubModulo <> "0"  then
	   str_SQL_SubModulo = " FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA = " &  str_SubModulo
	else
	   str_SQL_SubModulo = " "
	end if   

	'str_Sql = "SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " " & str_SQL_SubModulo & str_usoDesuso & "ORDER BY MEPR_CD_MEGA_PROCESSO,FUNE_TX_TITULO_FUNCAO_NEGOCIO"

	ssql=""	
	ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
	ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
	ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO WHERE "
	
	if str_SQL_SubModulo<>" " then
		ssql=ssql+ str_SQL_SubModulo & " AND " & str_usoDesuso 
	else
		ssql=ssql+ str_usoDesuso 	
	end if
	
	ssql=ssql+" AND FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " ORDER BY FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	
	'response.write ssql
	
	set rs1=conn_db.execute(ssql)
	
else
    str_Sql = "SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO > 99 ORDER BY MEPR_CD_MEGA_PROCESSO,FUNE_TX_TITULO_FUNCAO_NEGOCIO"
'    response.write(" sem mega")
'	response.write str_Sql
	set rs1=conn_db.execute(str_Sql)
	str_MegaProcesso=0
end if

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
//window.location.href='seleciona_mega_processo.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOpt='+document.frm1.txtOpt.value+'&pOpt2='+document.frm1.txtOpt2.value     
//alert(document.frm1.selMegaProcesso.value)
//alert(document.frm1.selSubModulo.value)
//alert(document.frm1.txtOpt.value)
//alert(document.frm1.chkEmUso.checked)
//alert(document.frm1.chkEmDesuso.checked)
window.location.href='seleciona_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOpt='+document.frm1.txtOpt.value+"&chkEmUso="+document.frm1.chkEmUso.checked+"&chkEmDesuso="+document.frm1.chkEmDesuso.checked

}

function Confirma() 
{ 
  if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatória!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
  if (document.frm1.txtOpt.value == "CF")
     {
     document.frm1.action="alterar_excluir_ori_mega_mape_funcao.asp";
     //document.frm1.target="corpo";
     document.frm1.submit();
     }
     if(document.frm1.txtOpt.value == "RM")
       {
       document.frm1.action="relat_ori_gerais_mega_mapeamento.asp";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "XM")
       {
       document.frm1.action="";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "IM")
       {
       document.frm1.action="";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "AM")
       {
       document.frm1.action="";
       //document.frm1.target="corpo";
       document.frm1.submit();
       }
     if(document.frm1.txtOpt.value == "EM")
       {
       document.frm1.action="";
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
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%">&nbsp;</td>
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
		str_Titulo = "Mega-Processo/Assunto ou a Funcao"
	%>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Selecione 
        <%=str_Titulo%></font></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%"><input type="hidden" name="txtOpt" value="<%=str_Opt%>">
      </td>
      <td width="70%">&nbsp; </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Mega-Processo 
          : </b></font></div></td>
      <td width="70%"> <select name="selMegaProcesso"  onChange="javascript:manda1()">
          <option value="0">== Selecione um Mega Processo ==</option>
          <%Set rdsMegaProcesso= Conn_db.Execute(str_SQL_MegaProc)
         While (NOT rdsMegaProcesso.EOF)
           if (Trim(str_MegaProcesso) = Trim(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)) then 
		      if rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value = 1 then
			     str_Opt2 = "M"
			  else
		     	 str_Opt2 = ""
			  end if
		   %>
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
    <% 'response.write str_MegaFuncao
	   if InStrRev("11/10", Right("00" & str_MegaProcesso, 2)) = 0 then
	%>
    <tr> 
      <td>&nbsp;</td>
      <td></td>
      <td>&nbsp;</td>
    </tr>
    <% end if
	  %>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Assunto 
          :</b></font></div></td>
      <td><select size="1" name="selSubModulo" onChange="javascript:manda1()">
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
      <td><div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Uso 
          :</font></b></font></div></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        uso </font></b></font> <font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmUso" type="checkbox" value="1" OnClick="manda1()" <%=checado01%>>
        </b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        desuso </font></b></font><font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmDesuso" type="checkbox" value="1"  OnClick="manda1()" <%=checado02%>>
        </b></font></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><div align="right"><b><font face="Verdana" color="#330099" size="2">Fun&ccedil;&atilde;o 
          R/3 : </font></b></div></td>
      <td><select size="1" name="selFuncao">
          <option value="0">== Selecione a Fun&ccedil;&atilde;o R/3 ==</option>
          <%do until rs1.eof=true%>
          <option value="<%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%> 
          - <%=rs1("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
					rs1.movenext
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
