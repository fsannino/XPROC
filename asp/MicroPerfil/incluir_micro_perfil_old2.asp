 
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = 0
end if

if request("selFuncPrinc") <> "" then
   str_funcao = UCase(Trim(request("selFuncPrinc")))
else
   str_funcao = "0"
end if

if request("selMacroPerfil") <> "" then
   str_MacroPerfil = UCase(Trim(request("selMacroPerfil")))
else
   str_MacroPerfil = "0"
end if
'response.Write(str_MacroPerfil)
str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
set rs_mega=db.execute(str_SQL_MegaProc)

str_SQL_Fun_Neg = ""
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " SELECT DISTINCT " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO, " & Session("PREFIXO") & "FUN_NEG_TRANSACAO " 
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso  
'If str_Opt = 1 or str_Opt = 2 then ' 1-alterar 2-excluir
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND FUNE_TX_INDICA_REFERENCIADA = '0' "
'end if
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ORDER BY " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
RESPONSE.WRITE str_SQL_Fun_Neg
set rs_funcao=db.execute(str_SQL_Fun_Neg)

str_SQL_Mac_Per = str_SQL_Mac_Per & " SELECT DISTINCT "
str_SQL_Mac_Per = str_SQL_Mac_Per & " MCPE_TX_NOME_TECNICO, "
str_SQL_Mac_Per = str_SQL_Mac_Per & " FUNE_CD_FUNCAO_NEGOCIO, MCPR_NR_SEQ_MACRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " FROM " & Session("PREFIXO") & "MACRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 
if str_funcao <> "0" Then
   str_SQL_Mac_Per = str_SQL_Mac_Per & " and FUNE_CD_FUNCAO_NEGOCIO ='" & str_funcao & "'"
end if
'response.write str_SQL_Mac_Per
set rs_MacPer=db.execute(str_SQL_Mac_Per)

' --------- DETERMINA O NOME DO MICRO PERFIL ----------------------
set rs=db.execute("SELECT MEPR_TX_ABREVIA, MEPR_TX_DESC_MEGA_PROCESSO FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso )
if not rs.eof then
   str_PrefixoNomeTecnico = "Z:" & Trim(rs("MEPR_TX_ABREVIA")) & "_PB???" & Mid(str_MacroPerfil,12,16)
else
   str_PrefixoNomeTecnico = ""
end if


'set rs=db.execute("SELECT MEPR_TX_ABREVIA, MEPR_TX_DESC_MEGA_PROCESSO FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso )
'if not rs.eof then
'   str_PrefixoNomeTecnico = "Z:" & Trim(rs("MEPR_TX_ABREVIA")) & "_PB"
'else
'   str_PrefixoNomeTecnico = ""
'end if

rs.CLOSE
SET rs = NOTHING

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

function manda1()
{
alert("aqui")
//alert (document.frm1.selMegaProcesso.value)
//alert (document.frm1.selFuncPrinc.value)
//alert (document.frm1.txtOPT.value)
//window.location.href='incluir_micro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncPrinc='+document.frm1.selFuncPrinc.value+'&pOPT='+document.frm1.txtOPT.value
}

function Confirma()
{
   if((document.frm1.selMacroPerfil.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um Macro Perfil !");
      document.frm1.selMacroPerfil.focus();
      return;
      }
   else
      {
	  //alert(document.frm1.txtOPT.value);
      if(document.frm1.txtOPT.value == 1)
        {
        document.frm1.action="altera_macro_perfil.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 2)
        {
        document.frm1.action="grava_exclusao_macro_perfil.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 3)
        {
        document.frm1.action="rel_funcao_transacao.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 4)
        {
        document.frm1.action="cad_funcao_transacao2_outro.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 5)
        {
        document.frm1.action="gera_rel_mega_funcao.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 6)
        {
        document.frm1.action="../MacroPerfil/inclui_macro_perfil.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 7)
        {
        document.frm1.action="../MacroPerfil/edita_objetos_macro_perfil.asp.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		

     }
}

function pega_tamanho()
{
valor=document.frm1.txtDesc.value.length;
document.frm1.txttamanho.value=valor
if (valor > 61) {
	str1=document.frm1.txtDesc.value;
	str2=str1.slice(0,61);
	document.frm1.txtDesc.value=str2;
	valor=str2.length;
	document.frm1.txttamanho.value=valor;
}
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="POST" action="grava_micro_perfil.asp" name="frm1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a>
            </div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
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
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="33">
    <tr> 
    <td> 
        <div align="center"><font face="Verdana" color="#330099" size="3">Inclus&atilde;o 
          de Micro Perfil</font></div>
    </td>
  </tr>
  <tr> 
      <td> 
        <div align="center">opt : <%=str_OPT%><b>
          <input type="hidden" name="txtOpt" value="<%=str_OPT%>">
          </b></div>
      </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="73">
    <tr> 
      <td width="6%" height="35">&nbsp;</td>
      <td width="4%" height="35">&nbsp; </td>
      <td width="23%" height="35"> <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          :</b></font></div></td>
      <td width="67%" height="35"><b> 
        <select size="1" name="selMegaProcesso" onChange="javascript:Alert("aqui")">
          <option value="0">== Selecione o Mega-Processo ==</option>
          <%do until rs_mega.eof=true
       if trim(str_MegaProcesso)=trim(rs_mega("MEPR_CD_MEGA_PROCESSO")) then
       %>
          <option selected value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%ELSE%>
          <option value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
		end if
		rs_mega.movenext
		loop
		%>
        </select>
        </b></td>
    </tr>
    <tr> 
      <td width="6%" height="35">&nbsp;</td>
      <td width="4%" height="35">&nbsp;</td>
      <td width="23%" height="35"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          :</b></font></div></td>
      <td width="67%" height="35"><b> 
        <select size="1" name="selFuncPrinc" onChange="javascript:manda1()">
          <option value="0"><b>== Selecione uma Funcao de Negocio ==</b></option>
          <%do until rs_funcao.eof=true
       if trim(str_Funcao)=trim(rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")) then
       %>
          <option selected value=<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>><b><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs_funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></b></option>
          <%ELSE%>
          <option value=<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>><b><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs_funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></b></option>
          <%
		end if
		rs_funcao.movenext
		loop
		%>
        </select>
        </b></td>
    </tr>
    <tr> 
      <td width="6%" height="35">&nbsp;</td>
      <td width="4%" height="35">&nbsp;</td>
      <td width="23%" height="35"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Nome 
          T&eacute;cnico Macro :</b></font></div></td>
      <td width="67%" height="35"><b> 
        <select size="1" name="selMacroPerfil" onChange="javascript:manda1()">
          <option value="0"><b>== Selecione um Macro Perfil ==</b></option>
          <%do until rs_MacPer.eof=true
         if trim(str_MacroPerfil)=trim(rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")) then %>
          <option value="<%=rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")%>" selected><b><%=rs_MacPer("MCPE_TX_NOME_TECNICO")%></b></option>
          <%ELSE%>
          <option value="<%=rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")%>"><b><%=rs_MacPer("MCPE_TX_NOME_TECNICO")%></b></option>
          <%
		  end if
        rs_MacPer.movenext
        loop
        %>
        </select>
        <input type="hidden" name="txtAcao" value="C">
        </b></td>
    </tr>
    <tr> 
      <td width="6%" height="83">&nbsp;</td>
      <td width="4%" height="83">&nbsp;</td>
      <td width="23%" height="83" valign="top"> <div align="right"><font face="Verdana" size="2" color="#330099"><b>Descrição 
          : </b></font> 
          <input type="hidden" name="txtFuncSelec" size="20">
          <input type="hidden" name="txtImp" size="20">
        </div></td>
      <td width="67%" height="83" valign="top"> <p> 
          <textarea rows="3" name="txtDesc" cols="49" ></textarea>
        </p>
        <p> <font face="Verdana" size="1" color="#330099">Caracteres digitados</font><font face="Verdana" size="2" color="#330099"><b> 
          <input type="text" name="txttamanho" size="5" value="0" maxlength="50">
          </b></font><font face="Verdana" color="#330099" size="1">(Máximo 61 
          caracteres)</font> </p></td>
    </tr>
    <tr> 
      <td width="6%" height="35">&nbsp;</td>
      <td width="4%" height="35">&nbsp;</td>
      <td width="23%" height="35">&nbsp;</td>
      <td width="67%" height="35"> <input type="hidden" name="txtOPT" value="<%=str_OPT%>"> 
      </td>
    </tr>
  </table>
<p>&nbsp;</p>
</form>
</body>
</html>
