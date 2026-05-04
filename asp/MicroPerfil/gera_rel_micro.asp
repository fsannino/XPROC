<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

mega=request("selMegaProcesso")

if mega<>0 then
	COMPL=" WHERE MEPR_CD_MEGA_PROCESSO=" & mega
else
	compl=""
end if

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MICRO_PERFIL" & COMPL &" ORDER BY MEPR_CD_MEGA_PROCESSO, MICR_TX_SEQ_MICRO_PERFIL")
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
<form name="frm1" method="post" action="gera_rel_micro.asp">
  <input type="hidden" name="INC" size="20" value="1"> 
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
          <td width="26"></td>
          <td width="50"></td>
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
  <table width="87%" border="0" cellpadding="0" cellspacing="5" name="tblSubProcesso" height="47">
    <tr> 
      <td width="7%" height="1"></td>
      <td width="83%" height="1"> 
      </td>
      <td width="5%" height="1"> 
      </td>
      <td width="17%" height="1"> 
      </td>
    </tr>
    <tr> 
      <td width="7%" height="1">&nbsp;</td>
      <td width="83%" height="1"> 
        <input type="hidden" name="txtOpc" value="1">
        <p align="left"><font color="#330099" face="Verdana" size="3">Relatório
        de Micro-Perfil</font>
      </td>
      <td width="5%" height="1"> 
       </td>
      <td width="17%" height="1"> 
       </td>
    </tr>
  </table>
  &nbsp;
  <%
  tem=0
  do until rs.eof=true
  tem=1
  %>
  <table border="0" width="84%" cellpadding="2">
    <tr> 
      <td width="27%" bgcolor="#E0E0E0"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099"><b>Mega-Processo</b></font></td>
      <%
      SET TEMP=CONN_DB.EXECUTE("select * from " & session("prefixo") & "mega_processo where mepr_cd_mega_processo=" & rs("mepr_cd_mega_processo"))
      %>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><B><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></B></font></td>
    </tr>
    <tr> 
      <td width="27%" bgcolor="#E0E0E0"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099">&nbsp;<b>Função 
        R/3</b></font></td>
      <%
      SET TEMP3=CONN_DB.EXECUTE("select * from " & session("prefixo") & "funcao_negocio where FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")
      if temp3.eof=false then
      	valor2=temp3("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
      else
      	valor2=""
      end if
      %>
      <td width="73%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=RS("FUNE_CD_FUNCAO_NEGOCIO")%> - <%=VALOR2%></font></td>
    </tr>
    <tr> 
      <%
      SET TEMP2=CONN_DB.EXECUTE("select * from " & session("prefixo") & "macro_perfil where MCPR_NR_SEQ_MACRO_PERFIL=" & rs("MCPR_NR_SEQ_MACRO_PERFIL"))
      if temp2.eof=false then
      	valor = temp2("MCPE_TX_NOME_TECNICO")
		valor_Desc_Mac =  temp2("MCPE_TX_DESC_MACRO_PERFIL")
		valor_Desc_Det_Mac =  temp2("MCPE_TX_DESC_DETA_MACRO_PERFIL")
		valor_Espec_Mac =  temp2("MCPE_TX_ESPECIFICACAO")
      else
      	valor=""
      end if
      %>
      <td bgcolor="#E0E0E0"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#330099"><b><em>Macro-Perfil</em></b></font></div></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=valor%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#E0E0E0"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099"><b>Descrição</b></font></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=valor_Desc_Mac%></font></td>
    </tr>
    <tr>
      <td bgcolor="#E0E0E0"><font face="Verdana" size="2" color="#330099"><b>Descri&ccedil;&atilde;o 
        Detalhada </b></font><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099">&nbsp;</font></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=valor_Desc_Det_Mac%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#E0E0E0"><font face="Verdana" size="2" color="#330099"><b>Especifica&ccedil;&atilde;o</b></font><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099">&nbsp;</font></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=valor_Espec_Mac%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#330099"><b><em>Micro-Perfil</em></b></font></div></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <%'=rs("MICR_TX_SEQ_MICRO_PERFIL")%>
        </font></td>
    </tr>
    <tr> 
      <td width="27%" bgcolor="#CCCCCC"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099"><b>Descrição</b></font></td>
      <td width="73%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MICR_TX_DESC_MICRO_PERFIL")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC"><font face="Verdana" size="2" color="#330099"><b>Descri&ccedil;&atilde;o 
        Detalhada </b></font><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <% str_Situacao = RS("MICR_TX_SITUACAO")
		If str_Situacao = "EE" then
			str_Situacao = "Em elaboração"
		 elseIf str_Situacao = "AT" then
			str_Situacao = "Alterado transação"
		 elseIf str_Situacao = "EA" then
			str_Situacao = "Em aprovação"			  
		 elseIf str_Situacao = "NA" then
			str_Situacao = "Não aprovado"			  
		 elseIf str_Situacao = "EC" then
			str_Situacao = "Em criação no R/3"			  
		 elseIf str_Situacao = "RE" then
			str_Situacao = "Recusado no R/3"			  
		 elseIf str_Situacao = "EX" then
			str_Situacao = "Excluída a função"			  
		 elseIf str_Situacao = "MR" then
			str_Situacao = "Mudado para referência"			  
		 elseIf str_Situacao = "EL" then
			str_Situacao = "Excluído"			  
		 elseIf str_Situacao = "CR" then
			str_Situacao = "Criado no R3"			  
		 elseIf str_Situacao = "AR" then
			str_Situacao = "Em alteração no R/3"			  
		 elseIf str_Situacao = "ER" then
			str_Situacao = "Em exclusão no R/3"			  
		 elseIf str_Situacao = "AP" then
			str_Situacao = "Alterado no R/3"			  
		 elseIf str_Situacao = "EP" then
			str_Situacao = "Excluído no R/3"			  
         end if
	  %>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MICR_TX_DESC_DETA_MICRO_PERFIL")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC"><font face="Verdana" size="2" color="#330099"><b>Especifica&ccedil;&atilde;o</b></font><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=rs("MICR_TX_ESPECIFICACAO")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#CCCCCC"><font face="Verdana" size="2" color="#330099"><b>Status</b></font><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
      <td><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="2"><%=str_Situacao%></font></td>
    </tr>
  </table>
  <p>
  <%
  rs.movenext
  loop
  if tem=0 then
  %>
  <p><font color="#800000"><b>Nenhum Registro Encontrado para a Seleção!</b></font></p>
  <%end if%>
  </form>

</body>
</html>
