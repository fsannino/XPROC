<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conn_consulta.asp" -->
<%
server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

chave=request("txtchave")
atrib=request("txtApoio")
megas=request("txtmegas")

curso=request("txtcurso")

db.execute("SELECT * FROM APOIO_LOCAL_CURSO WHERE APLO_NR_ATRIBUICAO=2 AND USMA_CD_USUARIO='" & chave & "'")

Sub Grava_Curso(SChave, SAtribu, sCurso)

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_CURSO(USMA_CD_USUARIO, APLO_NR_ATRIBUICAO, CURS_CD_CURSO, ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, ATUA_DT_ATUALIZACAO)"
	ssql=ssql+"VALUES('" & SChave & "',"
	ssql=ssql+"" & SAtribu & ","
	ssql=ssql+"'" & SCurso & "',"	
	ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

	on error resume next
	db.execute(ssql)
	err.clear
	
end sub

str_valor = curso

if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if
tamanho = Len(str_valor)
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If
tamanho = Len(str_valor)
contador = 1
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
			call Grava_Curso(chave,atrib,str_atual)
	   	
			valor_total=valor_total+1
        quantos = 0
    End If
    contador = contador + 1
Loop
%>

<html>
<head>

<title>SINERGIA # XPROC # Processos de Negócio...Redirecionando...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF">
<p> 
  <input type="hidden" name="edita" size="11" value="<%=edita%>">
  <input type="hidden" name="chave" size="11" value="<%=chave%>">
  <input type="hidden" name="atrib" size="11" value="<%=atrib%>">
</p>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> <div align="center"> 
              <p align="center">&nbsp;</div></td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> <div align="center">&nbsp;</div></td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> <div align="center">&nbsp;</div></td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center">&nbsp;</div></td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center">&nbsp;</div></td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center">&nbsp;</div></td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> <table width="625" border="0" align="center">
        <tr> 
          <td width="26"></td>
          <td width="26"></td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table></td>
  </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<p align="center"><font size="3" face="Verdana" color="#000080">Apoiadores Locais 
  e Multiplicadores</font></p>
<p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2"><%=topico%></font><font face="Verdana" color="#330099" size="3"></font></b></p>
<p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="889" height="86">
  <tr> 
    <td width="287" height="28"></td>
    <td width="26" height="28"><a href="menu.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
    <td height="28" width="556"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
      para tela Principal</font></td>
  </tr>
  <tr> 
    <td width="287" height="23"></td>
    <td width="26" height="23"> <p align="right"><a href="cad_curso.asp?chave=<%=request("txtchave")%>&amp;attrib=<%=request("txtApoio")%>"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
    <td height="23" width="556"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar 
      para a tela de Associação de Multiplicadores x Curso</font></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>