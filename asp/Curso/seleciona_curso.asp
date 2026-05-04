<%

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

resposta=request("RESP")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
if request("option") = 6 or request("option") = 5 or request("option") = 2 or request("option") = 1 or request("option") = 3 or request("option") = 4 then
   str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
end if
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

set rs=db.execute(str_SQL_MegaProc)

if request("mega")<>0 then
	compl1=" WHERE MEPR_CD_MEGA_PROCESSO=" & request("mega")
	'compl1 = compl1 & " AND CURS_TX_STATUS_CURSO = '1'"
ELSE
	compl1=" WHERE MEPR_CD_MEGA_PROCESSO=0"
	'compl1 = compl1 & " AND CURS_TX_STATUS_CURSO = '1'"
end if

set rscurso=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO" + compl1 + " ORDER BY CURS_CD_CURSO")
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="../Planilhas/js/troca_lista.js"></script>

<script>

function Confirma()
{
	if(document.frm1.selMegaProcesso.selectedIndex == 0)
		{
		alert("É obrigatória a seleção de um MEGA-PROCESSO!");
		document.frm1.selMegaProcesso.focus();
		return;
		}
	if(document.frm1.selCurso.selectedIndex == 0)
		{
		alert("É obrigatória a seleção de um CURSO!");
		document.frm1.selCurso.focus();
		return;
		}
	else
	{
	window.location.href="envia_curso.asp?option="+document.frm1.txtopt.value+"&mega="+document.frm1.selMegaProcesso.value+"&curso="+document.frm1.selCurso.value 
	}
}

function Confirma2()
{
if(document.frm1.selMega.value != "")
{
//alert(document.frm1.txtopt.value);
window.location.href="envia_curso.asp?option="+document.frm1.txtopt.value+"&mega="+document.frm1.selMega.value;
}
else
{
alert("É obrigatória a especificação de um CURSO!");
document.frm1.selMega.focus();
return;
}
}

function ver_conteudo(fbox)
{
valor=fbox.value;
tamanho=valor.length;
str1=valor.slice(tamanho-1,tamanho);
if (str1!=0 && str1!=1 && str1!=2 && str1!=3 && str1!=4 && str1!=5 && str1!=6 && str1!=7 && str1!=8 && str1!=9){
	fbox.value="";
	str2=valor.slice(0,tamanho-1)
	fbox.value=str2;
}
}

function envia1()
{
window.location.href="seleciona_curso.asp?option="+document.frm1.txtopt.value+"&mega="+document.frm1.selMegaProcesso.value
}

</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="" name="frm1">
        <input type="hidden" name="txtImp" size="20"><input type="hidden" name="txtQua" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
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
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif" width="24" height="24"></a></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
          <td width="26">&nbsp;</td>
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
      <td>
      </td>
    </tr>
	<%  if request("option") = 6 then
	       str_Titulo = "Alteração "
		elseif request("option") = 5 then
	       str_Titulo = "Exclusão --- "		
		elseif request("option") = 7 then
	       str_Titulo = "Não definido "
		else
	       str_Titulo = "Seleção "				
		end if %>
    <tr>
      <td>
        <div align="center"><font face="Verdana" color="#330099" size="3"><%=str_Titulo%>
          de Curso</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="87">
          <tr>
            
      <td width="162" height="29"></td>
            
      <td width="136" height="29" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo
        :</b></font></td>
            
      <td width="531" height="29" valign="middle" align="left" colspan="2"> 
        <select size="1" name="selMegaProcesso" onchange="javascript:envia1()">
                <option value="0">== Selecione o Mega-Processo ==</option>
                	<%do until rs.eof=true
                	if trim(request("mega"))=trim(rs("MEPR_CD_MEGA_PROCESSO")) then
                	%>
                <option selected value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
                	<%else%>
                <option value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
					<%
					end if
					rs.movenext
					loop
					%>
              </select></td>
            
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Curso :</b></font></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
        <select size="1" name="selCurso">
          <option value="0">== Selecione o Curso ==</option>
          <%DO UNTIL RSCURSO.EOF=TRUE%>
          <option value="<%=RSCURSO("CURS_CD_CURSO")%>"><%=RSCURSO("CURS_CD_CURSO")%>-<%=RSCURSO("CURS_TX_NOME_CURSO")%></option>
          <%
			RSCURSO.MOVENEXT          
          LOOP
          %>
        </select></td>
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="377" bgcolor="#CCCCCC"> 
      <font face="Verdana" size="2" color="#330099"><b>Se Preferir, Digite o
      Código do Curso</b></font> 
      </td>
            
      <td height="1" valign="middle" align="left" width="154"> 
      </td>
          </tr>
          <tr>
            
      <td width="162" height="1"><%'=request("option")%></td>
            
      <td width="136" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="377" bgcolor="#CCCCCC"> 
      <input type="hidden" name="txtopt" size="20" value="<%=request("option")%>"> 
      <font face="Verdana" size="2" color="#330099">
        <input type="text" name="selMega" size="20" maxlength="6">
        </font><font face="Verdana" color="#330099" size="1">( Ex.: MAN004 )&nbsp; 
        </font><a href="javascript:Confirma2()"><img border="0" src="../Funcao/confirma_f02.gif" align="absmiddle" width="24" height="24"></a><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font> 
      </td>
            
      <td height="1" valign="middle" align="left" width="154"> 
      </td>
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="377"> 
      <b>
      <%if resposta=1 then%>
      <font face="Verdana" size="2" color="#FF0000">O Curso Selecionado é
      inexistente</font></b> 
      <%end if%>
      </td>
            
      <td height="1" valign="middle" align="left" width="154"> 
      </td>
          </tr>
        </table>
  </form>

</body>

</html>
