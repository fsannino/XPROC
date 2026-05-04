 
<!--#include file="../../asp/protege/protege.asp" -->
<%
curso=request("curso")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

on error resume next

DB.EXECUTE("DELETE FROM " & Session("PREFIXO") & "CURSO_PRE_REQUISITO WHERE CURS_CD_CURSO='" & curso & "'")
'call grava_log(codigo,"" & Session("PREFIXO") & "CURSO_PRE_REQUISITO","D",1)

DB.EXECUTE("DELETE FROM " & Session("PREFIXO") & "CURSO_CENARIO WHERE CURS_CD_CURSO='" & curso & "'")
'call grava_log(codigo,"" & Session("PREFIXO") & "CURSO_CENARIO","D",1)

DB.EXECUTE("DELETE FROM " & Session("PREFIXO") & "CURSO_FUNCAO WHERE CURS_CD_CURSO='" & curso & "'")
'call grava_log(codigo,"" & Session("PREFIXO") & "CURSO_FUNCAO","D",1)

DB.EXECUTE("DELETE FROM " & Session("PREFIXO") & "CURSO_FUNCAO_TODAS WHERE CURS_CD_CURSO='" & curso & "'")
'call grava_log(codigo,"" & Session("PREFIXO") & "CURSO_FUNCAO","D",1)

DB.EXECUTE("DELETE FROM " & Session("PREFIXO") & "CURSO_TRANSACAO WHERE CURS_CD_CURSO='" & curso & "'")
'call grava_log(codigo,"" & Session("PREFIXO") & "CURSO_TRANSACAO","D",1)

DB.EXECUTE("DELETE FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & curso & "'")
'call grava_log(codigo,"" & Session("PREFIXO") & "CURSO","D",1)

if err.number=0 then

	set correio = server.CreateObject("Persits.MailSender")

	correio.host = "164.85.62.165"
	correio.from = "cursos@S600146.petrobras.com.br"
	correio.FromName = "Suporte Sinergia"
	     			
   	correio.AddAddress "xd47@petrobras.com.br"	     			
	correio.AddAddress "xd83@petrobras.com.br"
	     			
   	correio.Subject="Exclusão de Curso"
				
	data_Atual=day(date) &"/"& month(date) &"/"& year(date)
        			
   	correio.Body=" O Curso '" & UCASE(curso) & "' foi EXCLUÍDO em  " & DATA_ATUAL & " / POR : " & Session("CdUsuario")					
	correio.send
	
end if

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
        <div align="center"><font face="Verdana" color="#330099" size="3">Exclusão
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
      <b><font face="Verdana" color="#330099" size="2">O Curso foi Excluído com
      Sucesso</font></b> 
      </td>
            
          </tr>
      <%else%>    
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      <b><font face="Verdana" size="2" color="#800000">Houve um erro na
      exclusão do curso - <%=err.description%></font></b> 
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
        <a href="seleciona_curso.asp?option=5"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
            
      <td height="1" valign="middle" align="left" width="439"> 
        <font face="Verdana" color="#330099" size="2">Retornar para Tela de
        Exclusão de Curso</font></td>
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