<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

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

str_mega=0
str_mega=request("selMegaProcesso")
str_OPT = request("pOPT") 

if request("selSubModulo") <> "" then
   str_SubModulo = request("selSubModulo") 
else
   str_SubModulo = ""
end if
if request("selFuncao") <> "" then
   str_Funcao = request("selFuncao") 
else
   str_Funcao = 0
end if

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"   
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

set rs=db.execute(str_SQL_MegaProc)

if str_mega<>0 then
	if str_SubModulo <> "" and str_SubModulo <> "0"  then
	   str_SQL_SubModulo = "FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA = " &  str_SubModulo
	else
	   str_SQL_SubModulo = " "
	end if   
	
	if str_SQL_SubModulo<>" " then
		ssql=""	
		ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
		ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
		ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
		ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO"
		ssql=ssql+  " WHERE " + str_SQL_SubModulo & " AND " & str_usoDesuso 
		ssql=ssql+" AND FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & str_mega & " ORDER BY FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "	
	else
		ssql=""	
		ssql="SELECT DISTINCT FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
		ssql=ssql+"FROM FUNCAO_NEGOCIO "
		ssql=ssql+  " WHERE " & str_usoDesuso 
		ssql=ssql+" AND FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & str_mega & " ORDER BY FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "	
	end if
	
	set rs1=db.execute(ssql)
else
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY FUNE_CD_FUNCAO_NEGOCIO")
	str_mega=0
end if

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_CD_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_mega,2) & "%'" 
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "

set rs_SubModulo=db.execute(str_Sub_Modulo)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<script>

function manda()
{
window.location.href='seleciona_funcao_macro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+'&pOPT='+document.frm1.txtOPT.value
}

function manda1()
{
window.location.href='seleciona_funcao_macro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+"&chkEmUso="+document.frm1.chkEmUso.checked+"&chkEmDesuso="+document.frm1.chkEmDesuso.checked
}

function manda2()
{
window.location.href='seleciona_funcao_macro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+"&chkEmUso="+document.frm1.chkEmUso.checked+"&chkEmDesuso="+document.frm1.chkEmDesuso.checked+"&selFuncao="+document.frm1.selFuncao.value
}

function Confirma2()
{
	//alert(document.frm1.txtOPT.value);

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
	document.frm1.action="alterar_excluir_ori_mega_mape_perfil.asp";
	//document.frm1.target="corpo";
	document.frm1.submit();
	  
}

function Confirma()
	{
	//alert(document.frm1.txtOPT.value)

	if((document.frm1.txtOPT.value != 6)&&(document.frm1.txtOPT.value != 7)&&(document.frm1.txtOPT.value != 8))
		{ 
      	if(document.frm1.selFuncao.selectedIndex == 0)
        	{
        	alert("É obrigatória a seleção de uma FUNÇÃO DE NEGÓCIO!");
        	document.frm1.selFuncao.focus();
        	return;
        	}
      	else
        	{
         	if(document.frm1.txtOPT.value == 1)
           		{
           		document.frm1.action="alterar_excluir_ori_mega_mape_perfil.asp";
           		//document.frm1.target="corpo";
           		document.frm1.submit();
           		}
			}   
	 }  
   else
      {
	  //alert(document.frm1.txtOPT.value);
      if(document.frm1.txtOPT.value == 6)
        {
        document.frm1.action="relat_ori_gerais_mega_mape_perfil.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
     }

}
</script>

<body topmargin="0" leftmargin="0">
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

    <%
if str_OPT  = 1 then
   str_Titulo = "Seleção de Função para orientação de Perfil"
elseif str_OPT  = 6 then
   str_Titulo = "Seleção para orientação de Perfil"
else
   str_Titulo = "OUTRO DE FUNÇÃO"
end if
%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%"> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=str_Titulo%></font></div>
      </td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="829" height="132">
    <tr> 
      <td width="29"> <% If str_mega <> 11 and str_mega <> 10 then %> <input type="hidden" name="selSubModulo111" value="0"> <% end if %> </td>
      <td width="120"> <div align="right"><b><font face="Verdana" color="#330099" size="2">Mega-Processo 
          : </font></b></div></td>
      <td height="41" width="666"> <select size="1" name="selMegaProcesso" onChange="javascript:manda()">
          <%if str_OPT <> 8 then%>
          <option value="0">== Selecione o Mega-Processo ==</option>
          <%else%>
          <option value="0">== TODOS ==</option>
          <%end if%>
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
        </select> <% 'if InStrRev("11/10", Right("00" & str_mega, 2)) = 0 then %> <input type="hidden" name="txtSubModulo1" value="<%=str_txt_SubModulo%>"> 
        <% 'end if %> </td>
    </tr>
    <% 
	   'if InStrRev("11/10", Right("00" & str_mega, 2)) <> 0 then
	%>
    <tr> 
      <td width="29">&nbsp;</td>
      <td width="120"> <div align="right"><font face="Verdana" size="2" color="#330099"><b>Assunto 
          : </b></font></div></td>
      <td width="666"> <select size="1" name="selSubModulo" onChange="javascript:manda1()">
          <%if str_OPT <> 8 then%>
          <option value="0">== Selecione o Assunto ==</option>
          <%else%>
          <option value="0">== TODOS ==</option>
          <%end if%>
          <%do until rs_SubModulo.eof=true
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
        </select> </td>
    </tr>
    <% 'end if %>
    <tr>
      <td></td>
      <td><div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Uso 
          : </font></b></font></div></td>
      <td height="41"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        uso </font></b></font> <font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmUso" type="checkbox" value="1" OnClick="manda1()" <%=checado01%>>
        </b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        desuso </font></b></font><font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmDesuso" type="checkbox" value="1"  OnClick="manda1()" <%=checado02%>>
        </b></font></td>
    </tr>
    <tr> 
      <td width="29"></td>
      <td width="120"> <div align="right"><b><font face="Verdana" color="#330099" size="2">Fun&ccedil;&atilde;o 
          R/3 : </font></b></div></td>
      <td height="41" width="666"> <select size="1" name="selFuncao">
          <%if str_OPT <> 8 then%>
          <option value="0">== Selecione a Função ==</option>
          <%else%>
          <option value="0">== TODAS ==</option>
          <%end if%>
          <%do until rs1.eof=true%>
          <option value="<%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%> - <%=rs1("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
					rs1.movenext
					loop
					%>
        </select> </td>
    </tr>
    <tr> 
      <td width="29"></td>
      <td width="120"> <div align="right"> 
          <input type="hidden" name="txtOPT" value="<%=str_OPT%>">
        </div></td>
      <td height="41" width="666">&nbsp; </td>
    </tr>
    <tr> 
      <td width="29" height="2"></td>
      <td width="120" height="2"></td>
      <td width="666" height="2"></td>
    </tr>
  </table>
</form>

<p>&nbsp;</p>

</body>

</html>
<%
rs.close
set rs = nothing
rs1.close
set rs1 = nothing
rs_SubModulo.close
set rs_SubModulo = nothing
db.close
set db = nothing

%>