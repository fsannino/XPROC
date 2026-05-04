<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso")  & " AND (MCPE_TX_SITUACAO = 'EE' OR MCPE_TX_SITUACAO = 'AT' OR  MCPE_TX_SITUACAO = 'EA' OR  MCPE_TX_SITUACAO = 'RD' OR  MCPE_TX_SITUACAO = 'RE' OR  MCPE_TX_SITUACAO = 'EC') ORDER BY MCPR_NR_SEQ_MACRO_PERFIL")
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

function ver_validacao(macro)
{
var a=macro;
window.open("ver_validacao.asp?macro=" + a + "","_blank","width=600,height=260,history=0,scrollbars=1,titlebar=0,resizable=0")
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="pega_caminho()">
<form name="frm1" method="post" action="valida_status1.asp">
        <input type="hidden" name="txtcaminho" size="20"><input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1">
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
        de Status :&nbsp; Em Elaboração -&gt; Em Aprovação <input type="hidden" name="usuario" size="10" value="<%=Session("CdUsuario")%>"></font></p>
        <%SET TEMP=CONN_DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMegaProcesso"))%>
  <p align="left"><font color="#330099" face="Verdana" size="2"><b>Mega-Processo 
    Selecionado : </b><%=request("selMegaProcesso")%> - 
    <input type="hidden" name="mega" size="20" value="<%=request("selMegaProcesso")%>">
    <%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></font></p>
  <table width="77%" border="0" cellpadding="0" cellspacing="4">
    <tr>
      <td width="46%"><div align="right"><font face="Verdana" size="1"><img border="0" src="../../imagens/b04.gif" alt="Clique aqui para Visualizar o andamento das validações">Mostra 
          as transa&ccedil;&otilde;es que necessitam de valida&ccedil;&atilde;o</font></div></td>
      <td width="27%"><div align="right"><font face="Verdana" size="1"><img border="0" src="../../imagens/b061.gif" alt="Clique aqui para editar os objetos deste Macro-Perfil">Mostra 
          os objetos e os valores </font></div></td>
      <td width="27%"><div align="right"><font face="Verdana" size="1"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico">Mostra 
          o hist&oacute;rico da valida&ccedil;&atilde;o</font></div></td>
    </tr>
  </table>
  <table border="0" width="949">
    <tr> 
      <td width="97" bgcolor="#330099" align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em 
        Elaboração</font></b></td>
      <td width="73" bgcolor="#330099" align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em 
        Aprovação</font></b></td>
      <td width="111" bgcolor="#330099" align="center"><div align="center"><b><font color="#FFFFFF" size="1" face="Verdana">Altera&ccedil;&atilde;o 
          de Transa&ccedil;&atilde;o</font></b></div></td>
      <td width="73" bgcolor="#330099" align="center"><b><font color="#FFFFFF" size="1" face="Verdana">Em 
        cria&ccedil;&atilde;o R/3</font></b></td>
      <td width="90" bgcolor="#330099" align="center"><b><font color="#FFFFFF" size="1" face="Verdana">Recusado 
        R/3</font></b></td>
      <td width="130" bgcolor="#330099" align="center" valign="middle"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Macro 
          - Perfil</font></b></div></td>
      <td width="187" bgcolor="#330099" align="center" valign="middle"><b><font face="Verdana" size="1" color="#FFFFFF">Descrição</font></b></td>
      <td width="5" bgcolor="#FFFFFF" valign="middle" align="center">&nbsp;</td>
    </tr>
    <%
          valor1=""
          valor2=""
          valor3=""
          
          tem=0
          
          DO UNTIL RS.EOF=TRUE
			
			select case rs("MCPE_TX_SITUACAO")
			
			case "EE"
				VALOR1="checked"
			case "NA"
				VALOR2="X"
			case "EA"
				VALOR3="checked"
			case "AT"
				VALOR4="X"
			case "RD"
				VALOR5="X"
			case "EC"
				VALOR6="checked"
			case "RE"
				VALOR7="X"
			end select

          %>
    <tr> 
      <td width="97" align="center" valign="middle"> <font face="Verdana" size="1"> 
        <input type="radio" value="1" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" <%=valor1%>>
        </font></td>
       <td width="73" align="center" valign="middle"> <p align="center"> <font face="Verdana" size="1"> 
          <%
          exibe=0
			SSQL=""
			SSQL="SELECT * FROM " & Session("PREFIXO") & "MACRO_OBJ_AUTORIZA WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))
			Set tem_=CONN_db.execute(SSQL)
			if tem_.eof=false then
			exibe=1
		  %>
          <input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="3" <%=valor3%>>
          <%end if%>
          </font> </td>
      <td width="111" align="center" valign="middle"> <p align="center"> <font face="Verdana" size="1"> 
          <%if valor4="X" then%>
          <img border="0" src="../../imagens/marcado.gif"> 
          <%end if%>
          </font></p></td>
      <td align="center" width="73" valign="middle"><font face="Verdana" size="1">&nbsp; 
        <input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="6" <%=valor6%>>
        </font> 
      <td align="center" width="90" valign="middle"><font face="Verdana" size="1"> 
        <%if valor7="X" then%>
        <img border="0" src="../../imagens/marcado.gif"> 
        <%end if%>
        </font></td>
      <td width="130" align="center" valign="middle"> <font color="#330099" face="Verdana" size="1"><%=trim(RS("MCPE_TX_NOME_TECNICO"))%></font> 
      <td width="187" align="center" valign="middle"><font color="#330099" face="Verdana" size="1"><%=RS("MCPE_TX_DESC_MACRO_PERFIL")%></font></td>
      <td width="5" align="center" valign="middle"></td>
      <%IF VALOR2="X" THEN%>
      <td width="18" align="center" valign="middle" bgcolor="#FFFFFF"> <p align="left"><a href="rel_funcao_transacao.asp?selmacroperfil=<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>&amp;txtopt=3"><font face="Verdana" size="1"><img border="0" src="../../imagens/b061.gif" alt="Clique aqui para editar os objetos deste Macro-Perfil"></font></a></p></td>
      <td width="34" align="center" valign="middle" bgcolor="#FFFFFF"> <div align="left"><a href="#" onclick="ver_historico('<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>')"><font face="Verdana" size="1"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico"></font></a> 
        </div></td>
      <%ELSE%>
      <td width="26" align="center" valign="middle" bgcolor="#FFFFFF"> <p align="center"> 
        <%if exibe=1 then%>
        <a href="#" onclick="ver_validacao('<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>')">
        <img border="0" src="../../imagens/b04.gif" alt="Clique aqui para Visualizar o andamento das validações">
        </a>
        <%end if%>
        </p></td>
      <td width="31" align="center" valign="middle" bgcolor="#FFFFFF"> <div align="left"><a href="#" onclick="ver_historico('<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>')"><font face="Verdana" size="1"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico"></font></a> 
        </div></td>
      <%END IF%>
    </tr>
    <tr> 
      <td height="53" colspan="2" align="center" valign="middle" width="176"><font color="#000033" size="2"><strong>Coment&aacute;rios/Motivo 
        :</strong></font></td>
      <td colspan="10" align="center" valign="middle" width="741"><div align="left">
          <textarea name="coment_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" cols="90" rows="2"></textarea>
        </div></td>
    </tr>
    <%
            tem=tem+1
            
			valor1=""
			valor2=""
			valor3=""
			valor4=""
			valor5=""
			valor6=""
			valor7=""
            
            RS.MOVENEXT
            LOOP
            %>
  </table>
        
  
  <p style="margin-bottom: 0">
    <%if tem=0 then%>
    <font color="#800000"><b> Nenhum Registro Encontrado!</b></font> 
    <%end if%>
    &nbsp; </p>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <%IF RS.EOF=FALSE THEN%>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr bgcolor="#00FF99">
    <td height="20">
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26">
            <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><a href="javascript:Confirma()"><img src="../Cenario/confirma_f02.gif" width="24" height="24" border="0"></a></p>
          </td>
          <td width="50">
            <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></p>
          </td>
          <td width="26">
            <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
          </td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%END IF%>         
  </form>
</body>
</html>
