 
<!--#include file="../../asp/protege/protege.asp" -->
<%
str_cenario=request("txtCenario")
str_tipo=request("txtTipo")
str_desenv=request("txtDesenv")
str_teste=request("txtTeste")
str_conf=request("txtConf")

str_MegaProcesso=request("txtmega")
str_MegaProcesso=request("txtproc")
str_SubProcesso=request("txtsub")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

desen=0
nmuda=0

select case str_tipo
case 1
	if str_desenv=1 and str_conf=1 and str_teste=1 and desen=0 and nmuda=0 then
		str_status="PT"
	else
		str_status="DS"
	end if
	
	if str_desenv=1 then
	set ver_data=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & str_cenario & "'")
	data1=ver_data("CENA_DT_PREV_TERMINO")
	if isnull(data1) then
		nmuda=1
		muda="NAO"
	else
		set ver_trans=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & str_cenario & "'")
		tem=0
		maior=0
		menor=0
		nulo2=0
		do until ver_trans.eof=true
			set temp1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO_DESENV WHERE TRAN_CD_TRANSACAO='" & ver_trans("TRAN_CD_TRANSACAO") & "'")
			do until temp1.eof=true	
				set temp2=db.execute("SELECT * FROM " & Session("PREFIXO") & "DESENVOLVIMENTO WHERE DESE_CD_DESENVOLVIMENTO='" & temp1("DESE_CD_DESENVOLVIMENTO") & "'")
				data2=temp2("DESE_DT_CONCLUSAO")
				if data2<>"" then
				if data1> data2 then
					maior=maior+1
				else
					menor=menor+1
				end if
			else
				nulo2=nulo2+1			
			end if
			temp1.movenext
		loop
		ver_trans.movenext
	loop
	if menor=0 then
		nmuda=0
		muda="SIM"
	else
		nmuda=1
		muda="NAO"	
	end if
	if nulo2>0 then
		nmuda=1
		muda="NAO"	
	end if
	end if

	if nmuda=0 then
	
	set ver_data=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & str_cenario & "'")
	
	data1=ver_data("CENA_DT_PREV_TERMINO")
	
	if isnull(data1) then
		nmuda=1
		muda="NAO"
	else
		set ver_trans=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & str_cenario & "' AND DESE_CD_DESENVOLVIMENTO<>''")
		tem=0
		maior=0
		menor=0
		nulo2=0
		do until ver_trans.eof=true
				set temp2=db.execute("SELECT * FROM " & Session("PREFIXO") & "DESENVOLVIMENTO WHERE DESE_CD_DESENVOLVIMENTO='" & ver_trans("DESE_CD_DESENVOLVIMENTO") & "'")
				data2=temp2("DESE_DT_CONCLUSAO")
				if data2<>"" then
				if data1> data2 then
					maior=maior+1
				else
					menor=menor+1
				end if
			else
				nulo2=nulo2+1			
			end if
		ver_trans.movenext
	loop
	if menor=0 then
		nmuda=0
		muda="SIM"
	else
		nmuda=1
		muda="NAO"	
	end if
	if nulo2>0 then
		nmuda=1
		muda="NAO"	
	end if
	end if
	end if
			
	if nmuda=1 then
		str_status="DS"
		str_desenv=0
	else
		str_status=str_status
	end if
	end if
	
	if nmuda=1 then
		if isnull(data1) then
			verifica="Não foi Possível alterar o Status do Cenário - Cenário sem data prevista para término!"
			cor="#663300"
		else
			verifica="Não foi Possível alterar o Status do Cenário - Desenvolvimentos não concluídos!"
			cor="#663300"
		end if
	else
		verifica="O Status do Cenário foi alterado com Sucesso!"
		cor="#330099"
	end if
	
case 2
	str_desenv=0	
	if str_conf=1 and str_teste=1 and desen=0 and nmuda=0 then
		str_status="PT"
	else
		str_status="DS"
	end if

end select

if request("option")=1 then
	str_status="EE"
	str_tipo=0
	str_desenv=0
	str_conf=0
end if

ssql=""
ssql=" UPDATE " & Session("PREFIXO") & "CENARIO"
ssql=ssql+" SET CENA_TX_SITUACAO='" & str_status & "', "
ssql=ssql+" CENA_TX_SITU_DESENHO_TIPO='" & str_tipo & "', "
ssql=ssql+" CENA_TX_SITU_DESENHO_DESE='" & str_desenv & "', "
ssql=ssql+" CENA_TX_SITU_DESENHO_TESTE='" & str_teste & "', "
ssql=ssql+" ATUA_CD_NR_USUARIO='" & Session("CDUsuario") & "', "
ssql=ssql+" ATUA_DT_ATUALIZACAO= GETDATE() , "
ssql=ssql+" CENA_TX_SITU_DESENHO_CONF='" & str_conf & "' "
ssql=ssql+" WHERE CENA_CD_CENARIO='" & str_cenario & "'"



DB.EXECUTE(SSQL)

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
<form name="frm1" method="post" action="alterar_cenario.asp">
  <input type="hidden" name="INC" size="20" value="1"><input type="hidden" name="txtOpc" value="1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
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
  <table border="0" width="100%">
    <tr>
      <td width="100%" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="100%" colspan="3">
        <p align="center"><font face="Verdana" color="#330099" size="3">Alteração 
          de Status de Cenário - <font color="#330099">Questionário</font> - <b><%=str_cenario%></b></font></p>
      </td>
    </tr>
    <tr>
      <td width="100%" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
        <b><font face="Verdana" color=<%=cor%> size="2"><%=verifica%></font></b>
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="12%">
        <p align="right"><a href="../../indexA.asp"><img src="selecao_F02.gif" width="22" height="20" border="0"></a>
      </td>
      <td width="60%">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
              para tela Principal</font>
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="12%">
      <p align="right"><a href="gerencia_cenario_transa.asp?selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>&selCenario=<%=str_Cenario%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" border="0"></a>
      </td>
      <td width="60%">
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta
      para a tela de Edição de Cenário</font>
      </td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="12%">
      </td>
      <td width="60%">
      </td>
    </tr>
    <tr>
      <td width="28%"></td>
      <td width="72%" colspan="2"></td>
    </tr>
    <tr>
      <td width="28%">
      </td>
      <td width="72%" colspan="2">
      </td>
    </tr>
    <tr>
      <td width="100%" colspan="3"></td>
    </tr>
  </table>
  </form>
<p></p>
</body>
</html>
