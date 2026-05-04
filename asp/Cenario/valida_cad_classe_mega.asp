 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_classe=request("txtDescClasse")
str_mega=request("txtmega")

set rs=db.execute("SELECT MAX(CLCE_CD_NR_CLASSE_CENARIO)AS CODIGO FROM " & Session("PREFIXO") & "CLASSE_CENARIO")
valor=rs("CODIGO")

if isnull(valor) then
		valor = 1
	ELSE
		valor = valor + 1
end if

ssql=""
ssql="INSERT INTO " & Session("PREFIXO") & "CLASSE_CENARIO "
ssql=ssql+"VALUES('"& ucase(str_classe) &"', "
ssql=ssql+""& valor &", "
ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE())"

db.execute(ssql)

ssql=""

Sub Grava_classe(strC, strM)
	
	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "CLASSE_CENARIO_MEGA_PROCESSO "
	ssql=ssql+"VALUES(" & strC & ","
	ssql=ssql+"" & strM & ","
	ssql=ssql+"'C','" & Session("CdUsuario") & "',GETDATE())"

	on error resume next
	db.execute(ssql)

end sub

str_valor = str_mega

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
        
	   'Aqui entra o que vc quer fazer com o caracter em questão!
	
			call Grava_classe(valor, str_atual)
	   		valor_total=valor_total+1
	   		
        quantos = 0
        
    End If
    
    contador = contador + 1
Loop
%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body topmargin="0" leftmargin="0">
<form method="POST" action="" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>

<p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Cadastro
de Classes</font></p>
<p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>

<div align="center">
  <center>
  <table border="0" width="790" height="128">
    <%if err.number=0 then%>
    <tr>
      <td width="233" height="21"></td>
      <td width="543" height="21" colspan="2"><font face="Verdana" color="#330099" size="2"><b>O
        Registro foi atualizado com Sucesso!</b></font></td>
    </tr>
    <%else%>
    <tr>
      <td width="233" height="21"></td>
      <td width="543" height="21" colspan="2"><b><font face="Verdana" size="2" color="#FF0000">Ocorreu
        um erro na gravação do Registro</font></b></td>
    </tr>
    <%end if%>
    <tr>
      <td width="233" height="21"></td>
      <td width="543" height="21" colspan="2"></td>
    </tr>
    <tr>
            <td width="233" align="right">
              <p align="left"></p>
            </td>
            <td width="48"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
    </tr>
  </center>
    <tr>
      <td width="233"></td>
      <td width="48">
            <a href="javascript:history.go(-1)"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a>
            </td>
      <td width="489" height="49">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
        para tela de Cadastro de Classes</font>
        </td>
    </tr>
  </table>
</div>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>








