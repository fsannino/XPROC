 
<!--#include file="../../asp/protege/protege.asp" -->
<%

if (Request("txtMegaProcesso") <> "") then 
    str_MegaProcesso = Request("txtMegaProcesso")
else
    str_MegaProcesso = "não passado"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
'str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

set rs=db.execute(str_SQL_MegaProc)

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

if(document.frm1.txtCodCurso.value == "")
{
alert("É obrigatória a definição do CÓDIGO DO CURSO!");
document.frm1.txtCodCurso.focus();
return;
}

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
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_cad_curso.asp" name="frm1">
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
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
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
    <tr>
      <td>
        <div align="center"><font face="Verdana" color="#330099" size="3">Dados 
          de Cursos</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="87">
          <tr>
            
      <td width="116" height="29"></td>
            
      <td width="209" height="29" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Código
        do Curso</b></font></td>
            
      <td width="504" height="29" valign="middle" align="left"> 
        <input type="text" name="txtCodCurso" size="14" maxlength="6"></td>
            
          </tr>
          <tr>
            
      <td width="116" height="29"></td>
            
      <td width="209" height="29" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo
        :</b></font></td>
            
      <td width="504" height="29" valign="middle" align="left"> 
        <select size="1" name="selMegaProcesso">
                <option value="0">== Selecione o Mega-Processo ==</option>
                	<%do until rs.eof=true%>
                <option value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
					<%
					rs.movenext
					loop
					%>
              </select></td>
            
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Nome
        do Curso :</b></font></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
        <input type="text" name="txtnomecurso" size="58" maxlength="100"></td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Carga
        Horária (h):</b></font></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
        <input type="text" name="txtcargacurso" size="14" onkeyup="javascript:ver_conteudo(txtcargacurso)"></td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Método
        :</b></font></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
        <select size="1" name="selMetodo">
          <option value="0">== Selecione o Método ==</option>
          <option value="À DISTÂNCIA">À DISTÂNCIA</option>
          <option value="Presencial">PRESENCIAL</option>
        </select></td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Público
        Alvo : </b></font></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      <textarea rows="4" name="txtPublicoAlvo" cols="50"></textarea> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="top" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Pré-Requisitos
        :</b></font></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      <textarea rows="4" name="txtPreRequisitos" cols="50"></textarea> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="top" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Conteúdo
        Programático :</b></font></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      <textarea rows="4" name="txtConteudo" cols="50"></textarea> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="top" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Objetivo
        :</b></font></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      <textarea rows="4" name="txtObjetivo" cols="50"></textarea> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
          <tr>
            
      <td width="116" height="1"></td>
            
      <td width="209" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="504"> 
      </td>
          </tr>
        </table>
  </form>

</body>

</html>
