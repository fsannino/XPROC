<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Desc_Macro = ""

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
   str_MacroPerfil = 0
end if

if request("pOPT") <> "" then
   str_Opt = request("pOPT")
else
   str_Opt = 1
end if
if request("pOPT2") <> "" then
   str_Opt2 = request("pOPT2")
else
   str_Opt2 = 0
end if
str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

set rs_mega=db.execute(str_SQL_MegaProc)

str_SQL_Fun_Neg = ""
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " SELECT DISTINCT "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
'str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO INNER JOIN "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO ON "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " INNER JOIN"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "MACRO_PERFIL ON "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = " & Session("PREFIXO") & "MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso  
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " and " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_INDICA_REFERENCIADA = '0'"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND (dbo.MACRO_PERFIL.MCPE_TX_SITUACAO = 'CR' or dbo.MACRO_PERFIL.MCPE_TX_SITUACAO = 'AP') "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ORDER BY " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
set rs_funcao=db.execute(str_SQL_Fun_Neg)

IF str_Opt2 = 2 Then
   str_SQL_Fun_Neg = ""
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " SELECT DISTINCT "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_DESC_MACRO_PERFIL"
'   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_DESC_DETA_MACRO_PERFIL"
'   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_ESPECIFICACAO"   
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO INNER JOIN "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO ON "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " INNER JOIN"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "MACRO_PERFIL ON "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = " & Session("PREFIXO") & "MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_funcao & "'"
   'response.Write(str_SQL_Fun_Neg)   
   set rs_funcao2=db.execute(str_SQL_Fun_Neg)
   if not rs_funcao2.EOF then
      str_MacroPerfil = rs_funcao2("MCPR_NR_SEQ_MACRO_PERFIL")
      str_Desc_Macro = rs_funcao2("MCPE_TX_DESC_MACRO_PERFIL")
'      str_desc_detal = rs_funcao2("MCPE_TX_DESC_DETA_MACRO_PERFIL")
'      str_espec = rs_funcao2("MCPE_TX_ESPECIFICACAO")	  
   else	  
   end if
end if

str_SQL_Mac_Per = str_SQL_Mac_Per & " SELECT DISTINCT "
str_SQL_Mac_Per = str_SQL_Mac_Per & " MCPE_TX_NOME_TECNICO, "
str_SQL_Mac_Per = str_SQL_Mac_Per & " FUNE_CD_FUNCAO_NEGOCIO, MCPR_NR_SEQ_MACRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " FROM " & Session("PREFIXO") & "MACRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " WHERE "
str_SQL_Mac_Per = str_SQL_Mac_Per & " (MCPE_TX_SITUACAO = 'CR' OR MCPE_TX_SITUACAO = 'AP') "
str_SQL_Mac_Per = str_SQL_Mac_Per & " and MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 

set rs_MacPer=db.execute(str_SQL_Mac_Per)

IF str_Opt2 = 3  OR str_Opt2 = 2 Then
   str_SQL_Mac_Per = ""
   str_SQL_Mac_Per = str_SQL_Mac_Per & " SELECT "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " MCPE_TX_NOME_TECNICO "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPE_TX_DESC_MACRO_PERFIL "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,FUNE_CD_FUNCAO_NEGOCIO"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPR_NR_SEQ_MACRO_PERFIL"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPE_TX_DESC_DETA_MACRO_PERFIL"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPE_TX_ESPECIFICACAO"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " FROM " & Session("PREFIXO") & "MACRO_PERFIL"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " WHERE "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " (MCPE_TX_SITUACAO = 'CR' OR MCPE_TX_SITUACAO = 'AP') "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " and MCPR_NR_SEQ_MACRO_PERFIL =" & str_MacroPerfil
   set rs_MacPer2=db.execute(str_SQL_Mac_Per)
   if not rs_MacPer2.EOF then
      str_funcao = rs_MacPer2("FUNE_CD_FUNCAO_NEGOCIO")
      str_Desc_Macro = rs_MacPer2("MCPE_TX_DESC_MACRO_PERFIL")
      str_desc_detal = rs_MacPer2("MCPE_TX_DESC_DETA_MACRO_PERFIL")
      str_espec = rs_MacPer2("MCPE_TX_ESPECIFICACAO")
   else
   end if	  
end if

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function pega_tamanho()
{
document.frm1.txttamanho.value=document.frm1.txtDescM.value.length
if (document.frm1.txtDescM.value.length > 61) {
	str1=document.frm1.txtDescM.value;
	document.frm1.txtDescM.value=str1.slice(0,61);
	document.frm1.txttamanho.value=str2.length;
}
}

function manda1()
{
window.location.href='incluir_micro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncPrinc='+document.frm1.selFuncPrinc.value+'&selMacroPerfil='+document.frm1.selMacroPerfil.value+'&pOPT2='+document.frm1.txtOPT2.value
}

function Confirma()
{
   if(document.frm1.selMegaProcesso.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um Mega Processo !");
      document.frm1.selMegaProcesso.focus();
      return;
      }
   if(document.frm1.selFuncPrinc.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de uma Função !");
      document.frm1.selFuncPrinc.focus();
      return;
      }
   if(document.frm1.selMacroPerfil.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um Macro Perfil !");
      document.frm1.selMacroPerfil.focus();
      return;
      }
   if(document.frm1.txtDescM.value == "")
      {
      alert("É obrigatória o preenchimento do campo Descrição.");
      document.frm1.txtDescM.focus();
      return;
      }
   if(document.frm1.txtdetalM.value == "")
      {
      alert("É obrigatória o preenchimento do campo Descrição Detalhada.");
      document.frm1.txtdetalM.focus();
      return;
      }	        
	if(document.frm1.txtespecM.value == "")
      {
      alert("É obrigatória o preenchimento do campo Especificação.");
      document.frm1.txtespecM.focus();
      return;
      }	        
   else
      {
	   if(document.frm1.txtOPT.value == 1)
        {
        document.frm1.action="grava_micro_perfil.asp";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 2)
        {
        document.frm1.action="grava_exclusao_macro_perfil.asp";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 3)
        {
        document.frm1.action="rel_funcao_transacao.asp";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 4)
        {
        document.frm1.action="cad_funcao_transacao2_outro.asp";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 5)
        {
        document.frm1.action="gera_rel_mega_funcao.asp";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 6)
        {
        document.frm1.action="../MacroPerfil/inclui_macro_perfil.asp";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 7)
        {
        document.frm1.action="../MacroPerfil/edita_objetos_macro_perfil.asp.asp";
        document.frm1.submit();
        }		

     }
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form method="POST" action="grava_micro_perfil.asp" name="frm1">
        <input type="hidden" name="txtOPT2" value="<%=str_OPT%>"><input type="hidden" name="txtOPT" value="<%=str_OPT%>">
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
      <%
if str_OPT  = 1 then
   str_Titulo = "ALTERAÇÃO DE MACRO PERFIL"
elseif str_OPT  = 2 then
   str_Titulo = "EXCLUSÃO DE MACRO PERFIL"
elseif str_OPT  = 3 then
   str_Titulo = "EDIÇÃO DE OBJETOS"
elseif str_OPT  = 4 then
   str_Titulo = "MUDAR 1"
elseif str_OPT  = 5 then
   str_Titulo = "MUDAR 2"
elseif str_OPT  = 6 then
   str_Titulo = "INCLUSÃO DE MACRO-PERFIL"
elseif str_OPT  = 7 then
   str_Titulo = "EDIÇÃO DE OBJETOS - MACRO-PERFIL"

else
   str_Titulo = "OUTRO DE FUNÇÃO"
end if
%>

</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="33">
    <tr> 
    <td> 
        <div align="center"><font face="Verdana" color="#330099" size="3">SOLICITA&Ccedil;&Atilde;O 
          PARA CRIA&Ccedil;&Atilde;O DE MICRO PERFIL</font></div>
    </td>
  </tr>
  <tr> 
      <td> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%'=str_Titulo%> </font></div>
      </td>
  </tr>
</table>
  <table width="968" border="0" cellspacing="0" cellpadding="0" height="45">
    <tr> 
      <td width="8" height="7">&nbsp;</td>
      <td width="217" height="7">&nbsp;</td>
      <td width="12" height="7"></td>
      <td width="724" height="7">&nbsp;</td>
    </tr>
    <tr> 
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          :</b></font></div>
      </td>
      <td width="12" height="35"></td>
      <td width="724" height="35"><b> 
        <select size="1" name="selMegaProcesso" onChange="javascript:document.frm1.txtOPT2.value = 1;manda1()">
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
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          R/3 :</b></font></div>
      </td>
      <td width="12" height="35"></td>
      <td width="724" height="35"><b> 
        <select size="1" name="selFuncPrinc" onChange="javascript:document.frm1.txtOPT2.value = 2;manda1()">
          <option value="0"><b>== Selecione uma  Funcao R/3 ==</b></option>
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
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>ou 
          Macro Perfil :</b></font></div>
      </td>
      <td width="12" height="35"></td>
      <td width="724" height="35"><b> 
        <select size="1" name="selMacroPerfil" onChange="javascript:document.frm1.txtOPT2.value = 3;manda1()">
          <option value="0"><b>== Selecione um Macro Perfil ==</b></option>
          <%do until rs_MacPer.eof=true
           if trim(str_MacroPerfil)=trim(rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")) then
          %>		  
          <option value="<%=rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")%>" selected><b><%=rs_MacPer("MCPE_TX_NOME_TECNICO")%></b></option>
          <%ELSE%>
          <option value="<%=rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")%>"><b><%=rs_MacPer("MCPE_TX_NOME_TECNICO")%></b></option>
          <% end if
        rs_MacPer.movenext
        loop
        %>
        </select>
        </b></td>
    </tr>
    <tr> 
      <td width="8" height="35">&nbsp;</td>
      <td width="217" height="35"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descri&ccedil;&atilde;o 
          : </b></font></div></td>
      <td width="12" height="35"> </td>
      <td width="724" height="35"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=str_Desc_Macro%><input type="hidden" name="txtDesc" size="20" value="<%=str_Desc_Macro%>"></font></td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descri&ccedil;&atilde;o
        Detalhada: </b></font></td>
      <td width="12" height="35"> 
      </td>
      <td width="724" height="35"> 
      <font face="Verdana" color="#330099" size="1"><%=str_desc_detal%><input type="hidden" name="txtdetal" size="20" value="<%=str_desc_detal%>"></font> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Especificação
        :&nbsp;</b></font></td>
      <td width="12" height="35"> 
      </td>
      <td width="724" height="35"> 
        <font face="Verdana" color="#330099" size="1"><%=str_espec%><input type="hidden" name="txtespec" size="20" value="<%=str_espec%>"></font> 
      </td>
    </tr>
  </table>
        <p><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  &nbsp;<span style="background-color: #C0C0C0">&nbsp;&nbsp;
  Detalhes do Micro Perfil&nbsp;&nbsp;&nbsp;&nbsp;</span></font></b></p>
  <table width="968" border="0" cellspacing="0" cellpadding="0" height="31">
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35" valign="top">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descrição:&nbsp;</b></font></td>
      <td width="502" height="35"> 
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><textarea rows="3" name="txtDescM" cols="61" onKeyUp="pega_tamanho()"></textarea> 
        </p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="1">Tamanho
        Atual : <input type="text" name="txttamanho" size="5" maxlength="2" value="0">&nbsp;&nbsp;
        (Max
        61 Carateres)&nbsp;&nbsp;</font> 
      </td>
      <td width="222" height="35"> 
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> 
        <font face="Verdana" color="#330099" size="1"><a href="#" onClick="document.frm1.txtDescM.value=document.frm1.txtDesc.value;pega_tamanho()"><img border="0" src="../../imagens/copiar_inf.gif"></a></font> 
        </p>
      </td>
    </tr>
    <tr> 
      <td width="8" height="19"></td>
      <td width="217" height="19" valign="top">
      </td>
      <td width="502" height="19"> 
      </td>
      <td width="222" height="19"> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35" valign="top">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descri&ccedil;&atilde;o
        Detalhada: </b></font></td>
      <td width="502" height="35"> 
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><textarea rows="2" name="txtdetalM" cols="61"></textarea> 
      </td>
      <td width="222" height="35"> 
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="1"><a href="#" onClick="document.frm1.txtdetalM.value=document.frm1.txtdetal.value"><img border="0" src="../../imagens/copiar_inf.gif"></a></font> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="9"></td>
      <td width="217" height="9" valign="top">
      </td>
      <td width="502" height="9"> 
      </td>
      <td width="222" height="9"> 
      </td>
    </tr>
    <tr> 
      <td width="8" height="35"></td>
      <td width="217" height="35" valign="top">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Especificação
        : </b></font></td>
      <td width="502" height="35"> 
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><textarea rows="2" name="txtespecM" cols="61"></textarea> 
      </td>
      <td width="222" height="35"> 
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="1"><a href="#" onClick="document.frm1.txtespecM.value=document.frm1.txtespec.value"><img border="0" src="../../imagens/copiar_inf.gif"></a></font> 
      </td>
    </tr>
  </table>
</form>
</body>
</html>
