<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conn_consulta.asp" -->
<%
server.scripttimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

chave=request("chave")
atrib=request("atribb")
orgao=request("txtorgao")

db.execute("DELETE FROM " & Session("prefixo") & "APOIO_LOCAL_ORGAO WHERE USMA_CD_USUARIO='" & REQUEST("CHAVE") & "' AND APLO_NR_ATRIBUICAO=" & REQUEST("ATRIBB"))

str_valor = orgao

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
        
		SOrgao=str_atual
		SChave=chave
		SAtribu=atrib
		
		tamanho2 = len(trim(SOrgao))
	
		if tamanho2=2 then		
		
			ssql=""
			ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_ORGAO(USMA_CD_USUARIO,APLO_NR_ATRIBUICAO,ORME_CD_ORG_MENOR,APOG_NR_CHAVE_GERENTE,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO)"
			ssql=ssql+"VALUES('" & SChave & "',"
			ssql=ssql+"" & SAtribu & ","
			ssql=ssql+"'" & SOrgao& "',"
			ssql=ssql+"'',"
			ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

			db.execute(ssql)
		
		end if

		if tamanho2=7 then
		
			ssql=""
			ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_ORGAO(USMA_CD_USUARIO,APLO_NR_ATRIBUICAO,ORME_CD_ORG_MENOR,APOG_NR_CHAVE_GERENTE,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO)"
			ssql=ssql+"VALUES('" & SChave & "',"
			ssql=ssql+"" & SAtribu & ","
			ssql=ssql+"'" & SOrgao& "00000000',"
			ssql=ssql+"'',"
			ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

			on error resume next
			db.execute(ssql)
		
		end if
		
		if tamanho2=10 then

			ssql=""
			ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_ORGAO(USMA_CD_USUARIO,APLO_NR_ATRIBUICAO,ORME_CD_ORG_MENOR,APOG_NR_CHAVE_GERENTE,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO)"
			ssql=ssql+"VALUES('" & SChave & "',"
			ssql=ssql+"" & SAtribu & ","
			ssql=ssql+"'" & SOrgao & "00000',"
			ssql=ssql+"'',"
			ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

			db.execute(ssql)

		
		end if

		if tamanho2=13 then
		
			ssql=""
			ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_ORGAO(USMA_CD_USUARIO,APLO_NR_ATRIBUICAO,ORME_CD_ORG_MENOR,APOG_NR_CHAVE_GERENTE,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO)"
			ssql=ssql+"VALUES('" & SChave & "',"
			ssql=ssql+"" & SAtribu & ","
			ssql=ssql+"'" & SOrgao& "00',"
			ssql=ssql+"'',"
			ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

			db.execute(ssql)

		end if
			
		if tamanho2 = 15 then
		
			ssql=""
			ssql="INSERT INTO " & Session("PREFIXO") & "APOIO_LOCAL_ORGAO(USMA_CD_USUARIO,APLO_NR_ATRIBUICAO,ORME_CD_ORG_MENOR,APOG_NR_CHAVE_GERENTE,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO)"
			ssql=ssql+"VALUES('" & SChave & "',"
			ssql=ssql+"" & SAtribu & ","
			ssql=ssql+"'" & SOrgao& "',"
			ssql=ssql+"'',"
			ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"

			db.execute(ssql)
		
		end if
	   	
		valor_total = valor_total + 1

       quantos = 0
       
    End If
    
    contador = contador + 1
    
Loop

topico="Os Registros Foram Associados com Sucesso!"

if atrib=2 then
	response.redirect "cad_curso.asp?chave=" & chave & "&attrib=2"
end if

%>

<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;

function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
alert("Tecla não permitida!");
return;
}
}

</script>

<body topmargin="0" leftmargin="0" onKeyDown="verifica_tecla()">
<form method="POST" action="" name="frm1">
<input type="hidden" name="txtpub" size="20"><input type="hidden" name="txtQua" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
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
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font size="3" face="Verdana" color="#000080">Apoiadores
  Locais e Multiplicadores</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
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
            <td width="26" height="23">
              <p align="right"><a href="cad_orgao.asp?chave=<%=request("chave")%>&amp;atrib=<%=request("atribb")%>"><img src="../../imagens/selecao_F02.gif" border="0"></a></td>
            <td height="23" width="556"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para a tela de Associação de Órgãos Apoiados</font></td>
  </tr>
</table>
  </form>

<p>&nbsp;</p>

</body>

</html>

