 
<%

if request("selMacroPerfil") <> 0 then
   str_MacroPerfil = request("selMacroPerfil")
else
   str_MacroPerfil = "0"
end if

if request("txtNomeTecnico") <> "" then
   str_NomeTecnico = request("txtNomeTecnico")
else
   str_NomeTecnico = ""
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQl = str_SQL & " SELECT  "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_SITUACAO, "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL, "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_DESC_MACRO_PERFIL, "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_NOME_TECNICO, "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.MEPR_CD_MEGA_PROCESSO, "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_SITUACAO, "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO, "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO,"
str_SQl = str_SQL & "  " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO"
str_SQl = str_SQL & "  FROM " & Session("PREFIXO") & "MACRO_PERFIL INNER JOIN"
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MEGA_PROCESSO ON "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQl = str_SQL & "   INNER JOIN"
str_SQl = str_SQL & "  " & Session("PREFIXO") & "FUNCAO_NEGOCIO ON "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO"
str_SQl = str_SQL & "   AND "
str_SQl = str_SQL & "  " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO"
str_SQl = str_SQL & " WHERE "
if str_NomeTecnico <> "0" then
   str_SQl = str_SQL & " " & Session("PREFIXO") & "MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
else
   str_SQl = str_SQL & " " & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_NOME_TECNICO  = '" & str_NomeTecnico & "'"
end if	  
str_SQl = str_SQL & " and MCPE_TX_SITUACAO <> 'ER' and MCPE_TX_SITUACAO <> 'EP'" 

set rs_MacroPerfil=db.execute(str_SQL)   
'response.write str_SQL
str_MegaProcesso = rs_MacroPerfil("MEPR_CD_MEGA_PROCESSO")
str_DescMegaProcesso = rs_MacroPerfil("MEPR_TX_DESC_MEGA_PROCESSO")
str_MacroPerfil = rs_MacroPerfil("MCPR_NR_SEQ_MACRO_PERFIL")
if str_MegaProcesso <> 15 then
   int_Max_Nome_Tecnico = 19
   valor1 = 11
   valor2 = 19
else
   int_Max_Nome_Tecnico = 16
   valor1 = 14
   valor2 = 16
end if
str_PrefixoNomeTecnico = Left(rs_MacroPerfil("MCPE_TX_NOME_TECNICO"),valor1)
'str_NomeTecnico2 = Trim(Right(rs_MacroPerfil("MCPE_TX_NOME_TECNICO"),19))
str_NomeTecnico2 = Trim(Mid(rs_MacroPerfil("MCPE_TX_NOME_TECNICO"),valor1+1,valor2))
str_DescMacroPerfil = rs_MacroPerfil("MCPE_TX_DESC_MACRO_PERFIL")
str_FuncPrinc = rs_MacroPerfil("FUNE_CD_FUNCAO_NEGOCIO")
str_TituFuncPrinc = rs_MacroPerfil("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
'response.Write("            aaaaaaaaaaa            ")
'response.Write(str_NomeTecnico2)

        str_Situacao = rs_MacroPerfil("MCPE_TX_SITUACAO")
		If str_Situacao = "EE" then
			str_Situacao2 = "Em elaboração"
		 elseIf str_Situacao = "AT" then
			str_Situacao2 = "Alterado transação"
		 elseIf str_Situacao = "EA" then
			str_Situacao2 = "Em aprovação"			  
		 elseIf str_Situacao = "NA" then
			str_Situacao2 = "Não aprovado"			  
		 elseIf str_Situacao = "EC" then
			str_Situacao2 = "Em criação no R/3"			  
		 elseIf str_Situacao = "RE" then
			str_Situacao2 = "Recusado no R/3"			  
		 elseIf str_Situacao = "EX" then
			str_Situacao2 = "Excluída a função"			  
		 elseIf str_Situacao = "MR" then
			str_Situacao2 = "Mudado para referência"			  
		 elseIf str_Situacao = "EL" then
			str_Situacao2 = "Excluído"			  
		 elseIf str_Situacao = "CR" then
			str_Situacao2 = "Criado no R3"			  
		 elseIf str_Situacao = "AR" then
			str_Situacao2 = "Em alteração no R/3"			  
		 elseIf str_Situacao = "ER" then
			str_Situacao2 = "Em exclusão no R/3"			  
		 elseIf str_Situacao = "AP" then
			str_Situacao2 = "Alterado no R/3"			  
		 elseIf str_Situacao = "EP" then
			str_Situacao2 = "Excluído no R/3"			  
         end if

'***********************************
'set rs=db.execute("SELECT MEPR_TX_ABREVIA, MEPR_TX_DESC_MEGA_PROCESSO FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso )
'if not rs.eof then
'   str_PrefixoNomeTecnico = "Z:" & Trim(rs("MEPR_TX_ABREVIA")) & "_PB000_"
'else
'   str_PrefixoNomeTecnico = ""
'end if

'rs.CLOSE
'SET rs = NOTHING

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_changePropOO(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  var obj2 = MM_findObj(theValue);
  //alert("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
}
//-->
</script>
<script language="javascript" src="../js/troca_lista.js"></script>
<script>

function carrega_txt1(fbox) 
   {
   document.frm1.txtFuncSelec.value = "";
   for(var i=0; i<fbox.options.length; i++) {
      document.frm1.txtFuncSelec.value = document.frm1.txtFuncSelec.value + "," + fbox.options[i].value;
      }
   }

function Confirma()
{

if(document.frm1.txtNomeTecnico.value == "")
{
alert("É obrigatória a especificação do NOME TÉCNICO!");
document.frm1.txtNomeTecnico.focus();
return;
}

if(document.frm1.txtDescMacroPerfil.value == "")
{
alert("É obrigatória a especificação da DESCRIÇÃO DO MACROPERFIL!");
document.frm1.txtDescMacroPerfil.focus();
return;
}
else
{
//carrega_txt1(document.frm1.list2)
document.frm1.submit();
}
}

function pega_tamanho()
{
valor=document.frm1.txtDescMacroPerfil.value.length;
document.frm1.txttamanho.value=valor
if (valor > 61) {
	str1=document.frm1.txtDescMacroPerfil.value;
	str2=str1.slice(0,61);
	document.frm1.txtDescMacroPerfil.value=str2;
	valor=str2.length;
	document.frm1.txttamanho.value=valor;
}
}

</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" onLoad=";pega_tamanho()">
<form method="POST" action="grava_macro_perfil.asp" name="frm1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a>
              </div>
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
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
            <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
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
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <div align="center"><font face="Verdana" color="#330099" size="3">Altera&ccedil;&atilde;o 
          de Macro Perfil</font></div>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="764" height="164">
    <tr> 
      <td width="66" height="25"></td>
      <td width="156" height="25" valign="top">&nbsp;</td>
      <td width="341" height="25">&nbsp; </td>
      <td width="287" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="66" height="24"></td>
      <td width="156" height="24" valign="top"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          : </b></font></div>
      </td>
      <td width="341" height="24"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_MegaProcesso%> - <%=str_DescMegaProcesso%>
        <input type="hidden" name="selMegaProcesso" value="<%=str_MegaProcesso%>">
        </font></td>
      <td width="287" height="24"> 
        <p align="left">&nbsp; 
      </td>
    </tr>
    <tr> 
      <td width="66" height="23"></td>
      <td width="156" height="23"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Nome 
          T&eacute;cnico : </b></font><font face="Verdana" size="2" color="#330099"></font></div>
      </td>
      <td height="23" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099"><%=str_PrefixoNomeTecnico%></font> 
        <input type="hidden" name="txtPrefixoNomeTecnico" value="<%=str_PrefixoNomeTecnico%>">
        <input type="text" name="txtNomeTecnico" size="20" maxlength="<%=int_Max_Nome_Tecnico%>" value="<%=str_NomeTecnico2%>">
        <input type="hidden" name="txtMacroPerfil" value="<%=str_MacroPerfil%>">
        <input type="hidden" name="txtAcao" value="M">
        <input type="hidden" name="txtNomeTecnico_Original" value="<%=str_NomeTecnico2%>">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Situa&ccedil;&atilde;o 
        : </b><%=str_Situacao2%></font> <input type="hidden" name="txtSituacao" value="<%=str_Situacao%>"></td>
    </tr>
    <tr> 
      <td width="66" height="25"></td>
      <td width="156" height="25" valign="top">
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          : </b></font></div>
      </td>
      <td height="25" valign="top" colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=str_FuncPrinc%> - <%=str_TituFuncPrinc%>
        <input type="hidden" name="selFuncPrinc" value="<%=str_FuncPrinc%>">
        </font></td>
    </tr>
    <tr> 
      <td width="66" height="83"></td>
      <td width="156" height="83" valign="top"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b> </b></font> 
          <font face="Verdana" size="2" color="#330099"><b>Descrição : </b></font> 
          <input type="hidden" name="txtFuncSelec" size="20">
        </div>
      </td>
      <td height="83" valign="top" colspan="2"> <p align="left" style="margin-top: 0; margin-bottom: 0"> 
          <textarea rows="3" name="txtDescMacroPerfil" cols="49" onkeyup="javascript:pega_tamanho()"><%=str_DescMacroPerfil%></textarea>
          <input type="hidden" name="txtDescMacroPerfil_Original" size="20" value="<%=str_DescMacroPerfil%>">
        <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>Caracteres 
          digitados&nbsp; 
          <input type="text" name="txttamanho" size="5" value="0" maxlength="50">
          </b></font><font face="Verdana" color="#330099" size="1">(Máximo 61 
          caracteres)</font> 
      </td>
    </tr>
  </table>
  <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="2">
    <tr> 
      <td width="351" height="1" bgcolor="#FFFFFF"></td>
      <td width="315" height="1" bgcolor="#FFFFFF"></td>
    </tr>
  </table>
</form>
</body>

</html>
