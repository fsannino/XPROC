<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=0
str_mega=request("selMegaProcesso")
str_OPT = request("pOpt") 
str_txt_SubModulo = request("txtSubModulo")
If str_OPT = 1 then
	str_Titulo = "Seleção de Mega para relatório de Transações não associadas a Função"
elseif str_OPT = 2 then
	str_Titulo = "Seleção de Mega para relatório de Cursos sem Função associada"
end if
str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
'str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO not IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

set rs=db.execute(str_SQL_MegaProc)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script>
function manda()
{
window.location.href='sel_Mega.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&pOpt='+document.frm1.txtOPT.value
}

function Confirma()
{
	//alert(document.frm1.txtOPT.value)
	 if(document.frm1.txtOPT.value == 1)
	   {
	   document.frm1.action="rel_mega_proc_sub_ativ_trans_sem_funcao.asp";
	   //document.frm1.target="corpo";
	   document.frm1.submit();
	   }
	 if(document.frm1.txtOPT.value == 2)
	   {
	   document.frm1.action="rel_mega_proc_sub_ativ_trans_de_funcao_sem_curso.asp";
	   //document.frm1.target="corpo";
	   document.frm1.submit();
	   }
}

function Confirma2()
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
<form method="POST" action="" name="frm1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="211" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="57" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="52" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="102" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="57" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="52" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="102" valign="middle" align="center"> 
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
      <td width="14%">&nbsp;</td>
      <td width="74%">&nbsp;</td>
      <td width="12%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="74%"> 
        <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Titulo%></font></td>
      <td width="12%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14%">&nbsp;</td>
      <td width="74%">&nbsp;</td>
      <td width="12%">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="829" height="132">
    <tr> 
      <td width="66">&nbsp; 
      </td>
      <td width="175"> 
        <div align="right"><b><font face="Verdana" color="#330099" size="2">Mega-Processo 
          : </font></b></div>
      </td>
      <td height="41" width="574"> 
        <select size="1" name="selMegaProcesso" onChange="javascript:manda()">
          <option value="0">== TODOS ==</option>
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
      </td>
    </tr>
    <% 
	   'if InStrRev("11/10", Right("00" & str_mega, 2)) <> 0 then
	%>
    <tr> 
      <td width="66">&nbsp;</td>
      <td width="175"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"></font></div>
      </td>
      <td width="574">&nbsp; 
      </td>
    </tr>
    <% 'end if %>
    <tr> 
      <td width="66"></td>
      <td width="175"> 
      </td>
      <td height="41" width="574"> 
      </td>
    </tr>
    <tr>
      <td width="66"></td>
      <td width="175"> 
        <div align="right">
          <input type="hidden" name="txtOPT" value="<%=str_OPT%>">
        </div>
      </td>
      <td height="41" width="574">&nbsp; </td>
    </tr>
    <tr> 
      <td width="66" height="2"></td>
      <td width="175" height="2"></td>
      <td width="574" height="2"></td>
    </tr>
  </table>
</form>

<p>&nbsp;</p>

</body>

</html>
<%
rs.close
set rs = nothing
db.close
set db = nothing

%>
