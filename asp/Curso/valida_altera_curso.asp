 
<!--#include file="../../asp/protege/protege.asp" -->
<%

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

curso=request("txtCurso")
nome=request("txtnomecurso")
strStatusCurso = request("checkAtivo")
carga=request("txtcargacurso")
metodo=request("selMetodo")
str_onda=request("selOnda")

if strStatusCurso <> 1 then
	strStatusCurso = 0
end if

ssql=""
ssql="UPDATE " & Session("PREFIXO") & "CURSO "
ssql=ssql+"SET CURS_TX_NOME_CURSO='"& ucase(nome) & "', "
ssql=ssql+"CURS_NUM_CARGA_CURSO="& carga & ", "
ssql=ssql+"CURS_TX_METODO_CURSO='"& ucase(metodo) & "', "

if str_onda  <> 0 then
   ssql=ssql+ " ONDA_CD_ONDA = " & str_onda & ", "
else
   ssql=ssql+ " ONDA_CD_ONDA = null , "
end if

ssql=ssql+"CURS_TX_STATUS_CURSO='" & strStatusCurso & "', "

ssql=ssql+"CURS_TX_PUBLICO_ALVO='"& ucase(request("txtPublicoAlvo")) & "', "
ssql=ssql+"CURS_TX_PRE_REQUISITOS='"& ucase(request("txtPreRequisitos")) & "', "
ssql=ssql+"CURS_TX_CONTEUDO_PROGRAM='"& ucase(request("txtConteudo")) & "', "
ssql=ssql+"CURS_TX_OBJETIVO='"& ucase(request("txtObjetivo")) & "' "

ssql=ssql+"WHERE CURS_CD_CURSO='"& curso & "'"
'Response.write ssql
on error resume next
db.execute(ssql)
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="../Planilhas/js/troca_lista.js"></script>

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
        
  <table width="847" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="845">
      </td>
    </tr>
    <tr>
      <td width="845">
        <div align="center"><font face="Verdana" color="#330099" size="3">Alteração
          de Cursos</font></div>
      </td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="81">
          <tr>
            
      <td width="205" height="29"></td>
            
      <td width="93" height="29" valign="middle" align="left"></td>
            
      <td width="531" height="29" valign="middle" align="left" colspan="2"> 
      <%if err.number=0 then%>
      <b><font face="Verdana" color="#330099" size="2">O Curso foi Alterado com
      Sucesso</font></b> 
      </td>
            
          </tr>
      <%else%>    
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      <b><font face="Verdana" size="2" color="#800000">Houve um erro na
      alteração do curso - <%=err.description%></font></b> 
      </td>
          </tr>
          <%end if%>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
        <a href="../../indexA.asp"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
            
      <td height="1" valign="middle" align="left" width="439"> 
        <font face="Verdana" color="#330099" size="2">Retornar para Tela
        Principal</font></td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
        <a href="seleciona_curso.asp?option=6"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
            
      <td height="1" valign="middle" align="left" width="439"> 
        <font face="Verdana" color="#330099" size="2">Retornar para a Tela de
        alteração de Curso</font></td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
      </td>
            
      <td height="1" valign="middle" align="left" width="439"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
      </td>
            
      <td height="1" valign="middle" align="left" width="439"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
        </table>
  </form>

</body>

</html>