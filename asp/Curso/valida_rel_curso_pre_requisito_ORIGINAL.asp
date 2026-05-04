<!--#include file="../../asp/protege/protege.asp" -->
<%
curso=request("curso")
mega=request("mega")
prer=request("txtTrans")

tem=0

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

db.execute("DELETE FROM " & Session("PREFIXO") & "CURSO_PRE_REQUISITO WHERE CURS_CD_CURSO='" & curso & "'")

set fonte_funcao=db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "CURSO_FUNCAO WHERE CURS_CD_CURSO='" & curso & "'")

Sub Grava_Pre(strP)

	tem=tem+1
	
	seq=0
	
	set rs_seq=db.execute("SELECT MAX(CUPR_NR_SEQUENCIA) AS CODIGO FROM " & Session("PREFIXO") & "CURSO_PRE_REQUISITO WHERE CURS_CD_CURSO='" & curso & "'")
	
	if not isnull(rs_seq("CODIGO")) then
		seq=rs_seq("CODIGO")+1
	else
		seq=1
	end if	

	ssql=""
	ssql="INSERT INTO " & Session("PREFIXO") & "CURSO_PRE_REQUISITO (CURS_CD_CURSO,CUPR_NR_SEQUENCIA,CURS_PRE_REQUISITO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO) "
	ssql=ssql+"VALUES('" & ucase(curso) & "'," & seq & ","
	ssql=ssql+"'" & strP & "','I','" & Session("CdUsuario") & "',GETDATE())"
	
	on error resume next	
	db.execute(ssql)
	
	set fonte_funcao_todas=db.execute("SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "CURSO_FUNCAO WHERE CURS_CD_CURSO='" & strP & "'")
	
	do until fonte_funcao_todas.eof=true
		
		ssql=""
		ssql="INSERT INTO " & Session("PREFIXO") & "CURSO_FUNCAO_TODAS (CURS_CD_CURSO, FUNE_CD_FUNCAO_NEGOCIO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO) "
		ssql=ssql+"VALUES('" & ucase(strP) & "','" & fonte_funcao_todas("FUNE_CD_FUNCAO_NEGOCIO") & "',"
		ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"
		
		on error resume next
		db.execute(ssql)
		err.clear		
	
		fonte_funcao_todas.movenext

	loop
	
	if fonte_funcao.eof=false then
	
	fonte_funcao.movefirst
	
	do until fonte_funcao.eof=true
		
		ssql=""
		ssql="INSERT INTO " & Session("PREFIXO") & "CURSO_FUNCAO_TODAS (CURS_CD_CURSO, FUNE_CD_FUNCAO_NEGOCIO,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO) "
		ssql=ssql+"VALUES('" & ucase(strP) & "','" & fonte_funcao("FUNE_CD_FUNCAO_NEGOCIO") & "',"
		ssql=ssql+"'I','" & Session("CdUsuario") & "',GETDATE())"
		
		on error resume next
		db.execute(ssql)
		err.clear		
	
		fonte_funcao.movenext
		
	loop
	
	end if

end sub

str_valor = prer

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
        
			db.execute("DELETE FROM " & Session("PREFIXO") & "CURSO_FUNCAO_TODAS WHERE CURS_CD_CURSO='" & str_atual & "'")
			call Grava_Pre(str_atual)
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
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif" width="30" height="30"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif" width="30" height="30"></a></div>
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
        <p align="center"><font face="Verdana" color="#330099" size="3">Relação
          Curso x Pré - Requisito</font>
      </td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="57">
    <tr> 
      <td width="124" height="29"></td>
      <td width="56" height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2"> <%if err.number=0 then%> <b><font face="Verdana" color="#330099" size="2">O Curso 
        e seus Pré-Requisitos foram relacionados com Sucesso</font></b> </td>
    </tr>
    <%else%>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> <b><font face="Verdana" size="2" color="#800000">Houve 
        um erro no cadastro do registro - <%=err.description%></font></b> </td>
    </tr>
    <%end if%>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>

    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="109"> <a href="../../indexA.asp"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right" width="22" height="20"></a></td>
      <td height="1" valign="middle" align="left" width="542"> <font face="Verdana" color="#330099" size="2">Retornar 
        para Tela Principal</font></td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="109"> <a href="rel_curso_pre_requisitos.asp?mega=<%=mega%>&amp;curso=<%=curso%>"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right" width="22" height="20"></a></td>
      <td height="1" valign="middle" align="left" width="542"> <font face="Verdana" color="#330099" size="2">Retornar 
        para Tela de Relacionar Curso x Pré-Requisitos</font></td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="109"> </td>
      <td height="1" valign="middle" align="left" width="542"> </td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
  </table>
  </form>

</body>

</html>
