<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql=request("SQL")

set rs=db.execute(ssql)

tem=0

do until rs.eof=true

	atual=rs("CENA_TX_SITUACAO_VALIDACAO")
	
	proposta=request("situacao_" & rs("CENA_CD_CENARIO"))
	
	if proposta<>atual then

		str_cenario=rs("CENA_CD_CENARIO")
		str_atual=proposta
		str_coment=UCASE(request("coment_" & rs("CENA_CD_CENARIO")))
		
	if trim(str_coment)<>"" then
				
		ssql=""
		ssql="UPDATE " & Session("PREFIXO") & "CENARIO SET "
		ssql=ssql & " CENA_TX_SITUACAO_VALIDACAO='" & str_atual & "'"
		ssql=ssql & " ,ATUA_TX_OPERACAO = 'A'"  
	    ssql=ssql & " ,ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
	    ssql=ssql & " ,CENA_DT_VALIDACAO = GETDATE()"	    
	    ssql=ssql & " ,ATUA_DT_ATUALIZACAO = GETDATE()"
		ssql=ssql & " WHERE CENA_CD_CENARIO='" & str_Cenario & "'"

		db.execute(ssql)

		set ordem=db.execute("SELECT MAX(CEVA_NR_SEQUENCIA)AS CODIGO FROM " & Session("PREFIXO") & "CENARIO_VALIDACAO WHERE CENA_CD_CENARIO='" & str_cenario & "'")
		if not isnull(ordem("CODIGO")) then
			atual=ordem("codigo")+1
		else
			atual=1
		end if

		SSQL = ""
		SSQL = "INSERT INTO " & Session("PREFIXO") & "CENARIO_VALIDACAO(CENA_CD_CENARIO,CEVA_NR_SEQUENCIA,ATUA_TX_OPERACAO,ATUA_CD_NR_USUARIO,ATUA_DT_ATUALIZACAO,CEVA_TX_SITUACAO,CEVA_TX_COMENTARIO)"
		SSQL = SSQL + "VALUES('" & str_Cenario & "', " & ATUAL & ",'I','" & Session("CdUsuario") & "', GETDATE(), '" & str_atual & "','" & str_Coment& "')"
		
		db.execute(ssql)
		
		tem=tem+1
		
	else
		
		nada = nada & str_Cenario & ", "
		
	end if
	
	end if

RS.MOVENEXT

LOOP
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="valida_altera_escopo.asp">
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
          <td width="26"></td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="97%" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso" height="254">
    <tr>
      <td width="22%" height="21"></td>
      <td width="70%" height="21" colspan="2"> 
      </td>
    </tr>
    <tr>
      <td width="22%" height="21"></td>
      <td width="70%" height="21" colspan="2"> 
        <font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="3">Validação
        de Escopo de Cenário</font>
      </td>
    </tr>
    <tr>
      <td width="22%" height="1">&nbsp;</td>
      <td width="70%" height="1" colspan="2"> 
        <input type="hidden" name="txtOpc" value="1">
      </td>
    </tr>
    <tr> 
      <td width="22%" height="21"> 
      </td>
      <%
		if tem>0 then
			cor="#330099"
			valor="O Escopo dos Cenários foram alterados com Sucesso!"
		else
			cor="#800000"
			valor="Não foi efetuada nenhuma alteração no Escopo dos Cenários"
		end if
      %>
      <td width="70%" colspan="2" height="21"> <font color="<%=cor%>" face="Verdana, Arial, Helvetica, sans-serif" size="2">
 <b>
      <%=valor%>
              </b>
      </font></td>
    </tr>
    <tr> 
      <td width="22%" height="21"> 
      
      </td>
      <td width="70%" height="21" colspan="2"> 
 <%if tem<>0 then%>
      <font color="#666666"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Total
      de Cenários com Escopo Alterado : </font> </b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=tem%>
      &nbsp;&nbsp; </font><%end if%></font> 
      </td>
    </tr>
    <tr> 
      <td width="22%" height="21"> 
      
      </td>
      <td width="70%" height="21" colspan="2">
      <p style="margin-top: 0; margin-bottom: 0">
      <%
      if len(nada)<>0 then
      %>
      <font color="#FF0000" size="3" face="Arial"><b>
      Os Cenários Relacionados à seguir não sofreram&nbsp;</b></font></p>
      <p style="margin-top: 0; margin-bottom: 0"><font color="#FF0000" size="3" face="Arial"><b> alteração de Escopo por não terem nenhum comentário :</b></font></p>
      <%
      nada=left(nada,(len(nada))-2)
      %>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="3"><b><font color="black">&nbsp; <%=nada%>
      </font><%
      end if
      %></b></font></p>
      </td>
    </tr>
    <tr> 
      <td width="22%" height="21"> 
      
      </td>
      <td width="12%" height="21"> 
      </td>
      <td width="58%" height="21"> 
      </td>
    </tr>
    <tr> 
      <td width="22%" height="21">

      </td>
      <td width="12%" height="21"> 
        <p align="right"><a href="../../indexA.asp"><img border="0" src="../../imagens/selecao_F02_off.gif" width="22" height="20">
        </a>
      </td>
      <td width="58%" height="21"> 
        <font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2">Retornar
        para a Tela Principal</font>
      </td>
    </tr>
    <tr> 
      <td width="22%" height="21"> 
      
      </td>
      <td width="12%" height="21"> 
        <p align="right"><a href="altera_escopo.asp"><img border="0" src="../../imagens/selecao_F02_off.gif" width="22" height="20"></a>
      </td>
      <td width="58%" height="21"> 
        <font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2">Retornar
        para a Tela de Alteração de Escopo</font>
      </td>
    </tr>
    <tr> 
      <td width="22%" height="21">
      </td>
      <td width="12%" height="21"> </td>
      <td width="58%" height="21"> </td>
    </tr>
    <tr> 
      <td width="22%" height="21">&nbsp;</td>
      <td width="12%" height="21">&nbsp;<input type="hidden" name="INC" size="20" value="1"> </td>
      <td width="58%" height="21"> </td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>