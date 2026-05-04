<%@LANGUAGE="VBSCRIPT"%> 
 
<%

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

int_QtdObj=request("txtQtdObj")
int_contador = 1
int_Qtd_Alterado = 0
int_Qtd_Nulo = 0
'response.write int_QtdObj
'response.write "   -   "
do while int_contador <= (int_QtdObj+1)-1
	str_Funcao = request("txtFuncao" & int_contador) 
	str_Desc_Atual =  request("txtDescFunc" & int_contador)
	str_Desc_Anterior =  request("txtObsAnterior" & int_contador)
	if str_Desc_Atual = "" then
	   int_Qtd_Nulo = int_Qtd_Nulo + 1	
	end if
	if str_Desc_Atual <> str_Desc_Anterior then	
	   call f_Altera_Desc(str_Funcao,str_Desc_Atual)
	end if   
	int_contador = int_contador + 1
'    response.write int_contador
loop

Sub f_Altera_Desc(p_Funcao, p_Desc)
    int_Qtd_Alterado = int_Qtd_Alterado + 1
    ssql=" UPDATE " & Session("PREFIXO") & "FUNCAO_NEGOCIO SET FUNE_TX_OBS_ESPECIFICA = '" & p_Desc & "'"	
	ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'"   
	ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
	ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE()"
	ssql=ssql & " WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & p_Funcao & "'"
	'on error resume next
	'response.write ssql	
	conn_db.execute(ssql)

end sub

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="SHORTCUT ICON" href="http://regina/imagens/Wrench.ico">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="../valida_altera_processo.asp?mega=<%=str_MegaProcesso%>&Proc=<%=str_Processo%>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr>
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
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
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="3%">&nbsp;</td>
      <td height="20" width="43%">&nbsp; </td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="39%">&nbsp; </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"> <font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Grava 
        orienta&ccedil;&atilde;o de Fun&ccedil;&atilde;o</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_MegaProcesso%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <%if erro=0 then%>
    <tr> 
      <td width="3%"></td>
      <td width="24%"><%'=str_Processo%></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b>Opera&ccedil;&atilde;o 
        realizada com Sucesso!</b></font></td>
      <td width="14%"></td>
    </tr>
    <% else %>
    <tr> 
      <td width="3%"></td>
      <td width="24%">&nbsp;</td>
      <td width="59%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#800000">Não 
        foi possível realizar a opera&ccedil;&atilde;o. Avise o problema.</font></b></td>
      <td width="14%"></td>
    </tr>
    <% end if %>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_SubProcesso%></td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Quantidade 
        alterada = <%=int_Qtd_Alterado%></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%"><%'=str_Cenario%></td>
      <td width="59%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Quantidade 
        sem Orienta&ccedil;&atilde;o = <%=int_Qtd_Nulo%></font></td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%"> <table width="74%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td height="41"><a href="seleciona_funcao.asp?pOpt=CF"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela de Cadastro de Orienta&ccedil;&otilde;es das Fun&ccedil;&otilde;es</font></td>
          </tr>
          <tr> 
            <td height="41"><a href="../../indexA.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
            <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela Principal</font></td>
          </tr>
        </table></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
</form>
<p>&nbsp;</p>
</body>
</html>
