<%@LANGUAGE="VBSCRIPT"%> 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

ssql=""
ssql="SELECT DISTINCT dbo.MACRO_OBJ_AUTORIZA.MCPR_NR_SEQ_MACRO_PERFIL, dbo.MACRO_OBJ_AUTORIZA.MAOA_TX_AUTORIZADO, "
ssql=ssql+"dbo.MACRO_PERFIL.MCPE_TX_SITUACAO, dbo.MACRO_PERFIL.MEPR_CD_MEGA_PROCESSO "
ssql=ssql+"FROM dbo.MACRO_OBJ_AUTORIZA INNER JOIN "
ssql=ssql+"dbo.MACRO_PERFIL ON "
ssql=ssql+"dbo.MACRO_OBJ_AUTORIZA.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
ssql=ssql+"WHERE (dbo.MACRO_OBJ_AUTORIZA.MAOA_TX_AUTORIZADO <> '1') AND (dbo.MACRO_PERFIL.MCPE_TX_SITUACAO = 'EA') AND "
ssql=ssql+"(dbo.MACRO_OBJ_AUTORIZA.MEPR_CD_MEGA_PROCESSO = " & request("selMegaProcesso") & ")"

str_sql=ssql

set rs=conn_db.execute(str_SQL)

'response.write ssql

'SELECT DISTINCT dbo.MACRO_OBJ_AUTORIZA.MCPR_NR_SEQ_MACRO_PERFIL, dbo.MACRO_OBJ_AUTORIZA.MAOA_TX_AUTORIZADO, 
'dbo.MACRO_PERFIL.MCPE_TX_SITUACAO, dbo.MACRO_PERFIL.MEPR_CD_MEGA_PROCESSO
'FROM dbo.MACRO_OBJ_AUTORIZA INNER JOIN
'dbo.MACRO_PERFIL ON dbo.MACRO_OBJ_AUTORIZA.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL
'WHERE (dbo.MACRO_PERFIL.MCPE_TX_SITUACAO = 'EA') AND (dbo.MACRO_OBJ_AUTORIZA.MAOA_TX_AUTORIZADO <> '1') AND 
'(dbo.MACRO_PERFIL.MEPR_CD_MEGA_PROCESSO = 2) 
'ORDER BY dbo.MACRO_OBJ_AUTORIZA.MCPR_NR_SEQ_MACRO_PERFIL 

'str_SQL = "SELECT DISTINCT MCPR_NR_SEQ_MACRO_PERFIL, MAOA_TX_AUTORIZADO " 
'str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA " 
'str_SQL = str_SQL & " WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso") 
'str_SQL = str_SQL & " ORDER BY MCPR_NR_SEQ_MACRO_PERFIL"

'set rs=conn_db.execute("SELECT DISTINCT MCPR_NR_SEQ_MACRO_PERFIL, MAOA_TX_AUTORIZADO FROM MACRO_OBJ_AUTORIZA WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso") & " ORDER BY MCPR_NR_SEQ_MACRO_PERFIL")

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
function Confirma()
{
document.frm1.submit();
}
function addbookmark()
{
bookmarkurl="#"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->

function pega_caminho()
{
	var a = document.URL;
	var n=0;

	for (var i = 1 ; i < 1000; i++)
	{
	var final=a.slice(0,i)
	var t=a.slice(i-1,i);
	if (t=='/')
	{
	n = n + 1;
	}
	if(n == 4)
	{
	i = 1000;
	}
	}
	var tam=final.length;
	var caminho = final.slice(0,tam-1);
	
	document.frm1.txtcaminho.value=caminho;
	
}

function ver_historico(macro)
{
var a=macro;
window.open("ver_historico.asp?macro=" + a + "","_blank","width=600,height=260,history=0,scrollbars=1,titlebar=0,resizable=0")
}

function ver_validacao(mega,macro)
{
var a=mega;
var b=macro;
window.open("ver_validacao.asp?macro=" + b + "&mega=" + a + "","_blank","width=600,height=260,history=0,scrollbars=1,titlebar=0,resizable=0")
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="pega_caminho()" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="post" action="valida_status2.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="39" valign="middle" align="center">
              <div align="center">
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Cenario/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Cenario/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="#"><img border="0" src="../Cenario/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../Cenario/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Cenario/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../Cenario/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr> 
          <%IF RS.EOF=FALSE THEN%>
          <td width="26"><a href="javascript:Confirma()"><img src="../Cenario/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <%END IF%>
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
        <p align="center"><font color="#330099" face="Verdana" size="3">Encaminhamento
        de Status :&nbsp; Em Aprovação -&gt; Aprovação<input type="hidden" name="usuario" size="10" value="<%=Session("CdUsuario")%>"></font></p>
        <%SET TEMP=CONN_DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso"))%>
  <p align="left"><font color="#330099" face="Verdana" size="2"><b>Mega-Processo 
    Selecionado : </b><%=request("selMegaProcesso")%> - 
    <input type="hidden" name="mega" size="20" value="<%=request("selMegaProcesso")%>">
    <%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font></p>
  <table width="79%" border="0" cellpadding="0" cellspacing="4">
    <tr> 
      <td width="17%"><div align="right"></div></td>
      <td width="56%"><div align="right"><font face="Verdana" size="1"><img border="0" src="../../imagens/b061.gif" alt="Clique aqui para editar os objetos deste Macro-Perfil">Mostra 
          as transa&ccedil;&otilde;es que necessitam de valida&ccedil;&atilde;o</font></div></td>
      <td width="27%"><div align="right"><font face="Verdana" size="1"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico">Mostra 
          o hist&oacute;rico da valida&ccedil;&atilde;o</font></div></td>
    </tr>
  </table>
  <table border="0" width="95%">
    <tr> 
      <td width="11%" bgcolor="#330099" align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Aprovado</font></b></td>
      <td width="5%" bgcolor="#330099" align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Reprovado</font></b></td>
      <td width="10%" bgcolor="#330099"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Aprovado
          Diferenciado</font></b></div></td>
      <td width="8%" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Macro 
        - Perfil</font></b></td>
      <td width="45%" bgcolor="#330099"><b><font face="Verdana" size="1" color="#FFFFFF">Descrição</font></b></td>
      <td width="2%" bgcolor="#FFFFFF" align="center">&nbsp;</td>
      <td width="19%" bgcolor="#FFFFFF" align="center">&nbsp;</td>
    </tr> 
    <% 
          tem=0
          
          VALOR1=""
          VALOR2=""
          VALOR3=""
		   VALOR4=""
          
          DO UNTIL RS.EOF=TRUE
          
          IF RS("MCPE_TX_SITUACAO")="EA" THEN
          
          SELECT CASE RS("MAOA_TX_AUTORIZADO")
          
          CASE "0"
          		VALOR1="checked"
          CASE "1"
          		VALOR2="checked"
          CASE "2"
          		VALOR3="checked"
   		  CASE "3"
          		VALOR4="checked"
			end select
			
			ssql="SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL " 
			ssql= ssql & " WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & RS("MCPR_NR_SEQ_MACRO_PERFIL")
			SET RS1=CONN_DB.EXECUTE(ssql)
			%>
    <tr> 
      <td width="11%" align="center"> <input type="radio" value="2" name="macro_<%=trim(RS1("MCPE_TX_NOME_TECNICO"))%>" <%=valor2%>></td>
      <td width="5%" align="center"> <p align="center"> 
          <input type="radio" name="macro_<%=trim(RS1("MCPE_TX_NOME_TECNICO"))%>" value="3" <%=valor3%>>
      </td>
      <td width="10%"><div align="center"> 
          <input type="radio" name="macro_<%=trim(RS1("MCPE_TX_NOME_TECNICO"))%>" value="4" <%=valor4%>>
        </div></td>
      <td width="8%"><font color="#330099" face="Verdana" size="1"><a href="exibe_transacao_macro.asp?selMegaProcesso=<%=request("selMegaProcesso")%>&txtOPT=3&selMacroPerfil=<%=RS("MCPR_NR_SEQ_MACRO_PERFIL")%>"><b><%=RS1("MCPE_TX_NOME_TECNICO")%></a></b></font></td>
      <td width="45%"><font color="#330099" face="Verdana" size="1"><%=RS1("MCPE_TX_DESC_MACRO_PERFIL")%></font></td>
      <td width="2%" bgcolor="#FFFFFF" align="center"><a href="#" onclick="ver_validacao('<%=trim(request("selMegaProcesso"))%>','<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>')"><img border="0" src="../../imagens/b04.gif" alt="Clique aqui para Visualizar o andamento das validações">
        </a>
      </td>
      <td width="19%" bgcolor="#FFFFFF" align="center"> <a href="#" onclick="ver_historico('<%=trim(RS1("MCPR_NR_SEQ_MACRO_PERFIL"))%>')"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico"></a> 
      </td>
    </tr>
    <tr> 
      <td height="56" align="center" valign="middle">
<p><font color="#000033" size="2"><strong>coment&aacute;rios : </strong></font></p>
        </td>
      <td colspan="12" align="center" valign="middle"><div align="left"> 
          <textarea name="coment_<%=trim(RS1("MCPE_TX_NOME_TECNICO"))%>" cols="90" rows="2"></textarea>
        </div></td>
    </tr>
    <%
				tem=tem+1
            
				VALOR1=""
				VALOR2=""
				VALOR3=""
				VALOR4=""
				
				END IF
				
				RS.MOVENEXT
	          LOOP
				%>
  </table>
        <%if tem=0 then%>
        <font color="#800000"><b>
        Nenhum Registro Encontrado!</b></font>
        <%end if%>
        <input type="hidden" name="txtcaminho" size="20">
  </form>
</body>
</html>
