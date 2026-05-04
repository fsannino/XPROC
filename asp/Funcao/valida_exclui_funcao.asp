<%
SERVER.SCRIPTTIMEOUT = 99999999
funcao=request("selFuncao")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
tem=0
set temp1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
set temp2=db.execute("SELECT * FROM " & Session("PREFIXO") & "MICRO_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")

if (temp1.eof=false) or (temp2.eof=false) then
	tem=1
	if (temp1.eof=false) then
	    if temp1("MCPE_TX_SITUACAO") = "ER" then
		   ls_Existe = ""
		else
           ls_Existe = " MACRO-PERFIL "
		end if   
    ELSE	
	    if temp2("MICR_TX_SITUACAO") = "ER" then
		   ls_Existe = ""
		else
           ls_Existe = " MICRO-PERFIL "	
		end if   
	END IF	
else
    ls_Existe = ""
end if    

if ls_Existe = "" then
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUN_NEG_TRANSACAO","D",1)
	
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUN_NEG_TP_QUA","D",1)
	
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO_SUB_MODULO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUN_NEG_TP_QUA","D",1)
	
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUN_NEG_ORG_AGLU","D",1)
	
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_CONFL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUN_NEG_CONFL","D",1)
	
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_CONFL WHERE FUNC_CD_FUNCAO_CONFL='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUN_NEG_CONFL","D",1)
	
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUNCAO_USUARIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUNCAO_USUARIO","D",1)
	
	'Verifica se a função tem filhas
	set filha=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE (FUNE_CD_FUNCAO_NEGOCIO_PAI='" & funcao & "') AND (FUNE_CD_FUNCAO_NEGOCIO <> FUNE_CD_FUNCAO_NEGOCIO_PAI)")
	
	'Exclui todas as filhas desta funcao, assim como os registros relacionados
	do until filha.eof=true
		funcao2=filha("FUNE_CD_FUNCAO_NEGOCIO")	
	
		db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao2 & "'")
		''call grava_log(funcao2,"" & Session("PREFIXO") & "FUN_NEG_TP_QUA","D",1)
	
		db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao2 & "'")
		''call grava_log(funcao2,"" & Session("PREFIXO") & "FUN_NEG_ORG_AGLU","D",1)
		
		db.execute("DELETE FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO_SUB_MODULO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao2 & "'")
		''call grava_log(funcao,"" & Session("PREFIXO") & "FUN_NEG_TP_QUA","D",1)
	
		db.execute("DELETE FROM " & Session("PREFIXO") & "FUNCAO_USUARIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao2 & "'")
		''call grava_log(funcao2,"" & Session("PREFIXO") & "FUNCAO_USUARIO","D",1)
		
		db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_CONFL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao2 & "'")
		''call grava_log(funcao2,"" & Session("PREFIXO") & "FUN_NEG_CONFL","D",1)
		
		db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_CONFL WHERE FUNC_CD_FUNCAO_CONFL='" & funcao2 & "'")
		''call grava_log(funcao2,"" & Session("PREFIXO") & "FUN_NEG_CONFL","D",1)

		db.execute("DELETE FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO_PAI='" & funcao2 & "'")
		''call grava_log(funcao2,"" & Session("PREFIXO") & "FUNCAO_NEGOCIO","D",1)

	filha.movenext
	loop
	
	db.execute("DELETE FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & funcao & "'")
	''call grava_log(funcao,"" & Session("PREFIXO") & "FUNCAO_NEGOCIO","D",1)
	
	tem=0
end if
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<body topmargin="0" leftmargin="0">
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
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
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
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Exclusão
        de Fun&ccedil;&atilde;o R/3</font></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<%if tem=0 then%>
  <p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" color="#330099" size="2">O 
    Registro </font><font face="Verdana" color="#330099" size="3"><%=funcao%></font><font face="Verdana" color="#330099" size="2"> foi excluído 
    com sucesso!</font></b></p>
	<%else%>
  <p align="center"><b><font color="#660000" size="2" face="Verdana">Ocorreu um 
    erro na exclus&atilde;o da Fun&ccedil;&atilde;o, pois a mesma possui <%=ls_Existe%> 
    agregado &agrave; ela</font></b></p>
	<%end if%>
        <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="100%">
  <tr>
    <td width="33%"></td>
            <td width="48"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" border="0" align="right"></a></td>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
              para tela Principal</font></td>
  </tr>
  <tr>
    <td width="33%"></td>
            <td width="48">
              <p align="right"><a href="seleciona_funcao.asp?pOPT=2"><img src="../../imagens/selecao_F02.gif" border="0" width="22" height="20"></a></td>
            <td height="41" width="489"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Voltar
              para a tela de Exclusão de Fun&ccedil;&atilde;o R/3</font></td>
  </tr>
  <tr>
    <td width="33%"></td>
    <td width="9%"></td>
    <td width="58%"></td>
  </tr>
</table>
  </form>

<p>&nbsp;</p>

</body>

</html>