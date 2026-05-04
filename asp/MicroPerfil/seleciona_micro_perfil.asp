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
' 1-alterar 2-excluir
if request("pOPT") <> "" then
   str_Opt = request("pOPT")
else
   str_Opt = ""
end if
'response.write str_Opt

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
'RESPONSE.WRITE str_SQL_Fun_Neg
set rs_funcao=db.execute(str_SQL_Fun_Neg)

str_SQL_Mac_Per = str_SQL_Mac_Per & " SELECT"
str_SQL_Mac_Per = str_SQL_Mac_Per & " MICR_TX_DESC_MICRO_PERFIL, MICR_NR_SEQ_MICRO,"
str_SQL_Mac_Per = str_SQL_Mac_Per & " FUNE_CD_FUNCAO_NEGOCIO, MICR_TX_SEQ_MICRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " FROM " & Session("PREFIXO") & "MICRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " WHERE MICR_TX_SITUACAO <> 'ER' AND MICR_TX_SITUACAO <> 'EP' AND MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 
if str_funcao <> "0" Then
   str_SQL_Mac_Per = str_SQL_Mac_Per & " and FUNE_CD_FUNCAO_NEGOCIO ='" & str_funcao & "'"
end if
Set rs_MicPer=db.execute(str_SQL_Mac_Per)
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function manda()
{
window.location.href='seleciona_micro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+'&txtSubModulo='+document.frm1.txtSubModulo.value
}

function manda1()
{
window.location.href='seleciona_micro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncPrinc='+document.frm1.selFuncPrinc.value+'&pOPT='+document.frm1.txtOPT.value
}

function Confirma()
{
   if(document.frm1.selMicroPerfil.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um Micro Perfil !");
      document.frm1.selMicroPerfil.focus();
      return;
      }
   else
      {
        window.location='verifica_op.asp?opt='+ document.frm1.selOPT.value +'&selMicro=' + document.frm1.selMicroPerfil.value
      }		
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
if str_OPT  = 2 then
   str_Titulo = "ALTERAÇÃO DE MICRO PERFIL"
elseif str_OPT  = 3 then
   str_Titulo = "EXCLUSÃO DE MICRO PERFIL"
elseif str_OPT  = 4 then
   str_Titulo = "EDIÇÃO DE OBJETOS"
elseif str_OPT  = 5 then
   str_Titulo = "MUDAR 1"
elseif str_OPT  = 6 then
   str_Titulo = "MUDAR 2"
elseif str_OPT  = 7 then
   str_Titulo = "INCLUSÃO DE MACRO-PERFIL"
elseif str_OPT  = 8 then
   str_Titulo = "EDIÇÃO DE OBJETOS - MACRO-PERFIL"

else
   str_Titulo = "OUTRO DE FUNÇÃO"
end if
%>

</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="51">
    <tr> 
    <td height="30"> 
        <div align="center"></div>
    </td>
  </tr>
  <tr> 
      <td height="21"> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><%=str_Titulo%> </font></div>
      </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="73">
    <tr> 
      <td width="1%" height="35">&nbsp;</td>
      <td width="1%" height="35">&nbsp;</td>
      <td width="17%" height="35">&nbsp;</td>
      <td width="81%" height="35">&nbsp;
      <input type="hidden" name="selOPT" size="20" value="<%=trim(request("pOPT"))%>"></td>
    </tr>
    <tr> 
      <td width="1%" height="35">&nbsp;</td>
      <td width="1%" height="35">&nbsp; </td>
      <td width="17%" height="35"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          :&nbsp;&nbsp;</b></font></div>
      </td>
      <td width="81%" height="35"><b> 
        <select size="1" name="selMegaProcesso" onChange="javascript:manda1()">
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
      <td width="1%" height="35">&nbsp;</td>
      <td width="1%" height="35">&nbsp;</td>
      <td width="17%" height="35"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          :&nbsp;&nbsp;</b></font></div>
      </td>
      <td width="81%" height="35"><b> 
        <select size="1" name="selFuncPrinc" onChange="javascript:manda1()">
          <option value="0"><b>== Selecione uma  Funcao de Negocio ==</b></option>
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
      <td width="1%" height="35">&nbsp;</td>
      <td width="1%" height="35">&nbsp;</td>
      <td width="17%" height="35"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Micro 
          Perfil :&nbsp;&nbsp;</b></font></div>
      </td>
      <td width="81%" height="35"><b> 
        <select size="1" name="selMicroPerfil">
          <option value="0">== Selecione um Micro-Perfil ==</option>
          <%do until rs_MicPer.eof=true%>
          <option value="<%=rs_MicPer("MICR_TX_SEQ_MICRO_PERFIL")%>"><%=rs_MicPer("MICR_TX_SEQ_MICRO_PERFIL")%> - <%=rs_MicPer("MICR_TX_DESC_MICRO_PERFIL")%></option>
          <%
        rs_MicPer.movenext
        loop
        %>
        </select>
        </b></td>
    </tr>
    <tr> 
      <td width="1%" height="35">&nbsp;</td>
      <td width="1%" height="35">&nbsp;</td>
      <td width="17%" height="35">&nbsp;</td>
      <td width="81%" height="35"> 
        <input type="hidden" name="txtOPT" value="<%=str_OPT%>">
      </td>
    </tr>
  </table>
<p>&nbsp;</p>
</form>
</body>
</html>
