<%@LANGUAGE="VBSCRIPT"%> 
<%
server.ScriptTimeout=99999999

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MACRO_PERFIL WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ") AND (MCPE_TX_SITUACAO='EC' OR MCPE_TX_SITUACAO = 'ER' OR MCPE_TX_SITUACAO = 'AR') ORDER BY MEPR_CD_MEGA_PROCESSO, MCPR_NR_SEQ_MACRO_PERFIL")
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
window.open("ver_historico.asp?macro=" + a + "","_blank","width=600,height=260,history=0,scrollbars=1,titlebar=0,resizable=0,top=500,left=100")
}

</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="pega_caminho()" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="post" action="valida_status5.asp">
        <font color="#330099" face="Verdana" size="3">
        <input type="hidden" name="usuario" size="10" value="<%=Session("CdUsuario")%>">
        <input type="hidden" name="AcessoUsuario" size="10" value="<%=Session("AcessoUsuario")%>">
        </font><font color="#330099" face="Verdana" size="2">
        </font><input type="hidden" name="txtOpc" value="1">
        <input type="hidden" name="INC" size="20" value="1"> 
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
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>
            <input type="hidden" name="cduser" value="<%=Session("CdUsuario")%>">
          </b></font></td>
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
        <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="3">Encaminhamento
        de Status :&nbsp; Em Criação -&gt; Criado no R/3</font></p>
        <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table width="849" border="0" cellpadding="0" cellspacing="4" style="margin-bottom: 0">
    <tr> 
      <td width="167"> <div align="right"></div></td>
      <td width="209"> <div align="right"></div></td>
      <td width="457"> <div align="right"><font face="Verdana" size="1"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico">Mostra 
          o hist&oacute;rico da valida&ccedil;&atilde;o</font></div></td>
    </tr>
    <tr> 
      <td colspan="3"><div align="left"><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&#8226; 
          Em todas as situa&ccedil;&otilde;es voc&ecirc; dever&aacute; consultar 
          sempre as transa&ccedil;&otilde;es do Macro Perfil verificando se houve 
          inclus&atilde;o ou exclus&atilde;o de transa&ccedil;&atilde;o.</font></div></td>
    </tr>
    <tr> 
      <td colspan="3"><div align="left"><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&#8226; 
          Quando voc&ecirc; registrar &quot;Criado R/3&quot; ou &quot;Alterado 
          R/3&quot;, as trasa&ccedil;&otilde;es marcadas para inclus&atilde;o 
          ou exclus&atilde;o ser&atilde;o marcadas como &quot;processadas&quot;. 
          Isto facilitar&aacute; para saber quais as novas transa&ccedil;&otilde;es 
          inclu&iacute;das ou exclu&iacute;das.</font></div></td>
    </tr>
  </table>
  <table border="0" width="1213" height="23">
    <tr> 
      <td width="137" bgcolor="#330099" align="center" height="1"><b><font face="Verdana" size="1" color="#FFFFFF">Mega-Processo</font></b></td>
      <td width="72" bgcolor="#330099" align="center" height="1"><b><font face="Verdana" size="1" color="#FFFFFF">Criado 
        R/3 </font></b></td>
      <td width="98" bgcolor="#330099" height="1"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Recusado</font></b></div></td>
      <td width="90" bgcolor="#330099"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em 
          Altera&ccedil;&atilde;o R/3</font></b></div></td>
      <td width="85" bgcolor="#330099"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Em 
          Exclus&atilde;o R/3</font></b></div></td>
      <td width="104" bgcolor="#330099"><div align="center"><b><font color="#FFFFFF" size="1" face="Verdana">Alterado 
          R/3 </font></b></div></td>
      <td width="104" bgcolor="#330099"><div align="center"><b><font face="Verdana" size="1" color="#FFFFFF">Exclu&iacute;do 
          R/3 </font></b></div></td>
      <td width="103" bgcolor="#330099" height="1"><b><font face="Verdana" size="1" color="#FFFFFF">Macro 
        - Perfil</font></b></td>
      <td width="224" bgcolor="#330099" height="1"><b><font face="Verdana" size="1" color="#FFFFFF">Descrição</font></b></td>
      <td width="46" bgcolor="#FFFFFF" height="1">&nbsp;</td>
    </tr>
    <%
          VALOR_2=""
		  VALOR_3=""
		  VALOR_10=""
		  VALOR_11=""
		  
		  tem=0
          DO UNTIL RS.EOF=TRUE
		  
		  select case rs("MCPE_TX_SITUACAO")
		  case "EL"
			valor_3="X"
		  case "RE"
			VALOR_2="checked"
			case "AR"
			VALOR_10="X"
			case "ER"
			valor_11="X"
			case "AP"
			valor_12="X"
			case "EP"
			valor_12="X"
		  end select
		  %>
    <tr> 
      <%
    	SET TEMP=CONN_DB.EXECUTE("SELECT * FROM " & Session("Prefixo") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
    	%>
      <td width="137" height="13" align="center"> <font face="Verdana" size="1"><b><%=TEMP("MEPR_TX_DESC_MEGA_PROCESSO")%></b></font><input type="hidden" name="mega_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" size="20" value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"></td>
      <td width="72" height="13" align="center"> <p> 
          
          <%IF VALOR_3="" AND rs("MCPE_TX_SITUACAO") <> "AR" AND rs("MCPE_TX_SITUACAO") <> "ER" THEN%>
          <input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="1">
          <%END IF%>
          
          <font color="#330099" face="Verdana" size="1"> </font></p></td>
      		<td width="98" height="13"><div align="center"> 
          <p> 
          
			  <%IF VALOR_3="" AND rs("MCPE_TX_SITUACAO") <> "ER" THEN%>
            <input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="2" <%=VALOR_2%>>
            <%END IF%>
            
            <font color="#330099" face="Verdana" size="1"> </font></p>
        </div></td>
      <td width="90"><div align="center"><font color="#0000CC" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=valor_10%></strong></font></div></td>
      <td width="85"><div align="center"><font color="#0000CC" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=valor_11%></strong></font></div></td>
      <%IF rs("MCPE_TX_SITUACAO") = "AR" THEN%>
      <td width="104"><div align="center"><font color="#0000CC" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="10">
          </strong></font></div></td>
		<%ELSE%>       
      <td width="104"><div align="center"><font color="#0000CC" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></div></td>
      <%END IF%>    
      <%IF rs("MCPE_TX_SITUACAO") = "ER" THEN%>
      <td width="104"><div align="center"><font color="#0000CC" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong><input type="radio" name="macro_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" value="11">
          </strong></font></div></td>
      <%ELSE%>
      <td width="104"><div align="center"><font color="#0000CC" size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong></strong></font></div></td>
      <%END IF%>
      <td width="103" height="13"><font color="#330099" face="Verdana" size="1"><a href="exibe_transacao_macro.asp?txtOPT=3&selMacroPerfil=<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>"><b><%=RS("MCPE_TX_NOME_TECNICO")%></b></a></font></td>
      <td width="224" height="13"><font color="#330099" face="Verdana" size="1"><%=RS("MCPE_TX_DESC_MACRO_PERFIL")%></font></td>
      <td width="46" bgcolor="#FFFFFF" height="13"> <p align="center"> <a href="#" onclick="ver_historico('<%=trim(RS("MCPR_NR_SEQ_MACRO_PERFIL"))%>')"><img border="0" src="../../imagens/icon_empresa.gif" alt="Visualizar Histórico"></a> 
      </td>
    </tr>
    <tr> 
      <td height="1" align="center" valign="middle" width="137"></td>
      <td height="1" colspan="2" align="center" valign="middle"><font color="#000033" size="2"><strong>Coment&aacute;rios/Motivo 
        :</strong></font></td>
      <td colspan="7" align="center" valign="middle" height="1"> <div align="left"> 
          <textarea name="coment_<%=trim(RS("MCPE_TX_NOME_TECNICO"))%>" cols="62" rows="1"></textarea>
        </div></td>
    </tr>
    <%
            tem=tem+1
            RS.MOVENEXT
			
			VALOR_2=""
			VALOR_3=""
			VALOR_10=""
		    VALOR_11=""
		    VALOR_12=""
		    VALOR_13=""
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
