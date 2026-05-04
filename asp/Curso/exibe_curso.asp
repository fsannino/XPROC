 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

curso=request("curso")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & curso & "'")

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

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

if(document.frm1.txtnomecurso.value == "")
{
alert("É obrigatória a definição do NOME DO CURSO!");
document.frm1.txtnomecurso.focus();
return;
}

if(document.frm1.txtcargacurso.value == "")
{
alert("É obrigatória a CARGA HORÁRIA DO CURSO!");
document.frm1.txtcargacurso.focus();
return;
}
if(document.frm1.selMetodo.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MÉTODO!");
document.frm1.selMetodo.focus();
return;
}

else
{
document.frm1.submit();
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

</script>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form method="POST" action="valida_altera_curso.asp" name="frm1">
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
            <td width="26"></td>
          <td width="50"></td>
          <td width="26"></td>
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
    <tr>
      <td>
        <div align="center"><font face="Verdana" color="#330099" size="3">Exibição
          de Curso - <b><%=curso%></b></font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="741" height="87">
    <tr> 
      <td width="113" height="29"> <input type="hidden" name="selMegaProcesso" size="20" value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"> 
      </td>
      <td width="185" height="29" valign="middle" align="left" bgcolor="#D8D8D8"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
        :</b></font></td>
      <td width="423" height="29" valign="middle" align="left" bgcolor="#D8D8D8"> 
        <%
      set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
      %> <font face="Verdana" size="2" color="#330099"><%=rs("MEPR_CD_MEGA_PROCESSO")%> - <%=rstemp("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
    </tr>
    <tr> 
      <td width="721" height="1" colspan="3"><input type="hidden" name="txtCurso" size="20" value="<%=curso%>"></td>
    </tr>
    <tr> 
      <td width="113" height="1"></td>
      <td width="185" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Nome 
        do Curso :</b></font></td>
      <td height="1" valign="middle" align="left" width="423"> <font face="Verdana" size="2" color="#000080"><%=RS("CURS_TX_NOME_CURSO")%></font></td>
    </tr>
    <tr> 
      <td width="721" height="1" colspan="3"></td>
    </tr>
    <tr>
      <td height="1"></td>
      <td height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Abrang&ecirc;ncia 
        do Curso :</b></font></td>
		<% if not IsNull(RS("ONDA_CD_ONDA")) then
		     str_SQl = ""
		   	 str_SQL = str_SQL & " SELECT ONDA_TX_DESC_ONDA, ONDA_CD_ONDA"
			 str_SQL = str_SQL & " FROM dbo.ABRANGENCIA_CURSO"
			 str_SQL = str_SQL & " WHERE ONDA_CD_ONDA = " & RS("ONDA_CD_ONDA")
			 set rdsOnda = db.Execute(str_SQL)
			 if not rdsOnda.EOF then
			    str_Ds_Onda = rdsOnda("ONDA_TX_DESC_ONDA")
			 else
			    str_Ds_Onda = "erro"
			 end if
		   else
             str_Ds_Onda = "não selecionado onda"
		   end if 
		%>
      <td height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#000080"><%=str_Ds_Onda%></font></td>
    </tr>
    <tr> 
      <td width="113" height="1"></td>
      <td width="185" height="1" valign="middle" align="left" bgcolor="#D8D8D8"><font face="Verdana" size="2" color="#330099"><b>Carga 
        Horária (h):</b></font></td>
      <td height="1" valign="middle" align="left" width="423" bgcolor="#D8D8D8"> 
        <font face="Verdana" size="2" color="#000080"><%=RS("CURS_NUM_CARGA_CURSO")%></font></td>
    </tr>
    <tr> 
      <td width="721" height="1" colspan="3"></td>
    </tr>
    <tr> 
      <td width="113" height="1"></td>
      <td width="185" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Método 
        de Ensino:</b></font></td>
      <td height="1" valign="middle" align="left" width="423"> <%
          SELECT CASE RS("CURS_TX_METODO_CURSO")
          CASE "À DISTÂNCIA"
          VAL="À DISTÂNCIA"
          CASE "PRESENCIAL"
          VAL= "PRESENCIAL"
          case else
          VAL=""
          end select
          %> <font face="Verdana" size="2" color="#000080"><%=VAL%></font></td>
    </tr>
    <tr> 
      <td width="721" height="1" colspan="3"></td>
    </tr>
    <tr> 
      <td width="67" height="1"></td>
      <td width="258" height="1" valign="top" align="left" bgcolor="#D8D8D8"><font face="Verdana" size="2" color="#330099"><b>Público 
        Alvo : </b></font></td>
      <td height="1" valign="middle" align="left" width="396" bgcolor="#D8D8D8"> 
        <font face="Verdana" size="2" color="#000080"><%=rs("CURS_TX_PUBLICO_ALVO")%></font> </td>
    </tr>
    <tr> 
      <td width="721" height="1" colspan="3"></td>
    </tr>
    <tr> 
      <td width="67" height="1"></td>
      <td width="258" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b><a href="relat_curso_pre_requisito.asp?curso=<%=curso%>">Requisitos 
        n&atilde;o R/3</a> :</b></font></td>
      <td height="1" valign="middle" align="left" width="396"> <font face="Verdana" size="2" color="#000080"><%=rs("CURS_TX_PRE_REQUISITOS")%></font> </td>
    </tr>
    <tr> 
      <td width="721" height="1" colspan="3"></td>
    </tr>
    <tr> 
      <td width="67" height="1"></td>
      <td width="258" height="1" valign="top" align="left" bgcolor="#D8D8D8"><font face="Verdana" size="2" color="#330099"><b>Conteúdo 
        Programático :</b></font></td>
      <td height="1" valign="middle" align="left" width="396" bgcolor="#D8D8D8"> 
        <font face="Verdana" size="2" color="#000080"><%=rs("CURS_TX_CONTEUDO_PROGRAM")%></font> </td>
    </tr>
    <tr> 
      <td width="721" height="1" colspan="3"></td>
    </tr>
    <tr> 
      <td width="67" height="1"></td>
      <td width="258" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Objetivo 
        :</b></font></td>
      <td height="1" valign="middle" align="left" width="396"> <font face="Verdana" size="2" color="#000080"><%=rs("CURS_TX_OBJETIVO")%></font> </td>
    </tr>
  </table>
  </form>

</body>

</html>
