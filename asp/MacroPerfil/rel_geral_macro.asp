<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script>
function manda()
{
//alert('_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value)
//+'&selSubModulo='+
//alert(document.frm1.selSubModulo.value)

document.frm1.txtSubModulo.value = document.frm1.selSubModulo.value
//alert(document.frm1.txtSubModulo.value)
window.location.href='rel_geral_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+'&txtSubModulo='+document.frm1.txtSubModulo.value
}

function Confirma()
{
   if(document.frm1.selMegaProcesso.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um MEGA-PROCESSO!");
      document.frm1.selMegaProcesso.focus();
      return;
      }
  		else
        {
        document.frm1.submit();
        }		
     }
    
</script>
<body topmargin="0" leftmargin="0">
<form method="POST" action="gera_rel_geral_macro.asp" name="frm1">
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
            <div align="center"><a href="#"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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

 
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%"> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Relatório
          Geral de Macro-Perfil</font></div>
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
      <td width="70"> 
        <% If str_mega <> 11 and str_mega <> 10 then %>
        <input type="hidden" name="selSubModulo" value="0">
        <% end if %>
      </td>
      <td width="165"> 
        <div align="right"><b><font face="Verdana" color="#330099" size="2">Mega-Processo 
          : </font></b></div>
      </td>
      <td height="41" width="574"> 
        <select size="1" name="selMegaProcesso">
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
        <input type="hidden" name="txtSubModulo" value="<%=str_txt_SubModulo%>">
      </td>
    </tr>
    <tr> 
      <td width="70"></td>
      <td width="165"> 
      </td>
      <td height="41" width="574"> 
      </td>
    </tr>
    <tr>
      <td width="70"></td>
      <td width="165"> 
        <div align="right">
          <input type="hidden" name="txtOPT" value="<%=str_OPT%>">
        </div>
      </td>
      <td height="41" width="574">&nbsp; </td>
    </tr>
    <tr> 
      <td width="70" height="2"></td>
      <td width="165" height="2"></td>
      <td width="574" height="2"></td>
    </tr>
  </table>
  </form>

<p>&nbsp;</p>

</body>

</html>
