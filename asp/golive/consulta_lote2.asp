<%@LANGUAGE="VBSCRIPT"%>
<%

if request("str_Tipo_Saida")="Excel" then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

if request("pAcao") = "C" then
	str_Acao = "C"
else
	str_Acao = "I"
end if

set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
conn_Cogest.CursorLocation=3

str_SQL = ""
str_SQL = str_SQL & " SELECT  "
str_SQL = str_SQL & " LOTE_NR_SEQ_LOTE"
str_SQL = str_SQL & " ,LOTE_TX_DESCRICAO"
str_SQL = str_SQL & " , LOTE_DT_ENVIO"
str_SQL = str_SQL & " , LOTE_NR_QTD_EXPORTACAO"
str_SQL = str_SQL & " , LOTE_TX_ORGAO_SELEC"
str_SQL = str_SQL & " , LOTE_TX_FUNCAO_SELEC"
str_SQL = str_SQL & " , LOTE_TX_POSSUI_VISAO"
str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
str_SQL = str_SQL & " FROM dbo.GOLI_LOTE"
str_SQL = str_SQL & " order by LOTE_NR_SEQ_LOTE desc"
set rds_Lote = conn_Cogest.Execute(str_SQL)

'response.Write(str_SQL)

function formatadata(str_Formato_Original,str_Formato_Saida,str_Data)

	strDia = ""		
	strMes = ""
	strAno = ""
	
	int_PosicaoDiaOrig = InStrRev(str_Formato_Original,"D" )
	int_PosicaoMesOrig = InStrRev(str_Formato_Original,"M" )
	int_PosicaoAnoOrig = InStrRev(str_Formato_Original,"A")

	int_PosicaoDiaSaida = InStrRev(str_Formato_Saida,"D")
	int_PosicaoMesSaida = InStrRev(str_Formato_Saida,"M")
	int_PosicaoAnoSaida = InStrRev(str_Formato_Saida,"A")

	vetDataOriginal = split(Trim(str_Data),"/")							
	vetDataSaida = split(Trim(str_Data),"/")		
	
	vetDataSaida(int_PosicaoDiaSaida-1) = Right("00" & trim(vetDataOriginal(int_PosicaoDiaOrig-1)),2)
	vetDataSaida(int_PosicaoMesSaida-1) = Right("00" & trim(vetDataOriginal(int_PosicaoMesOrig-1)),2)	
	vetDataSaida(int_PosicaoAnoSaida-1) = trim(vetDataOriginal(int_PosicaoAnoOrig-1))

	if int_PosicaoDiaSaida = 1 and int_PosicaoMesSaida = 2 then
		formatadata = vetDataSaida(int_PosicaoDiaSaida-1) & "/" & vetDataSaida(int_PosicaoMesSaida-1) & "/" & vetDataSaida(int_PosicaoAnoSaida-1) 			
	end if
	if int_PosicaoDiaSaida = 2 and int_PosicaoMesSaida = 1 then
		formatadata = vetDataSaida(int_PosicaoMesSaida) & "/" & vetDataSaida(int_PosicaoDiaSaida) & "/" & vetDataSaida(int_PosicaoAnoSaida-1) 				
	end if
	
end function
%>
<html>
<head>
<script>
function importacao() 
	{
	// selOnda=<%=int_Onda%>&selFases=<%=int_Fase%>&selPlano=<%=int_Plano1%>&selPlano2=<%=int_Plano2%>&selTask1=<%=int_Atividade%>
	window.open('importacao.asp?par_PaginaPrint=importa_usuario1.asp','jan1','toolbar=no, location=no, scrollbars=no, status=no, directories=no, resizable=no, menubar=no, fullscreen=no, height=50, width=250, status=no, top=200, left=260');
	}

function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href.href='altera_Atividade1.asp?selAtiv='+document.frm1.selAtividade.value
	 }
 }
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

</SCRIPT>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="">
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
            </div></td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div></td>
            <td bgcolor="#330099" width="27" valign="middle" align="center">
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div></td>
          </tr>
          <tr>
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center">
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div></td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div></td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center">
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
  <table width="100%"  border="0" cellspacing="0" cellpadding="1">
    <tr>
      <td width="21%">&nbsp;</td>
      <td width="46%">&nbsp;</td>
      <td width="33%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Rela&ccedil;&atilde;o de Lotes</font></td>
      <td><img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="87%"  border="0" cellspacing="5" cellpadding="1">
    <tr bgcolor="#000099">
      <td width="9%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Usu-Func (Ok)</font></strong></div></td>
      <td width="9%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Usu-Func (n&atilde;o Ok)</font></strong></div></td>
      <td width="32%"><div align="left"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Descri&ccedil;&atilde;o</font></strong></div></td>
      <td width="10%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Criado por</font></strong></div></td>
      <td width="12%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif"> Cria&ccedil;&atilde;o</font></strong></div></td>
      <td width="11%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Vezes Exportadas </font></strong></div></td>
      <td width="17%"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Uacute;ltima Exporta&ccedil;&atilde;o </font></strong></div></td>
    </tr>
	<% do while not rds_Lote.Eof %>
    <tr>
      <td><div align="center">
        <% 
	  if Trim(rds_Lote("LOTE_TX_POSSUI_VISAO")) = "N" then %>
        <a href="consulta_lote_mapeados.asp?str_Tipo_Saida=Tela&pLote=<%=rds_Lote("LOTE_NR_SEQ_LOTE")%>&pDescLote=<%=rds_Lote("LOTE_TX_DESCRICAO")%>&pVezesImp=<%=rds_Lote("LOTE_NR_QTD_EXPORTACAO")%>&pOrdem=1"><img src="../../imagens/b04.gif" width="16" height="16" border="0">
        <%
		else 
	%>
        </a><a href="consulta_lote_mapeados_rh.asp?str_Tipo_Saida=Tela&pLote=<%=rds_Lote("LOTE_NR_SEQ_LOTE")%>&pDescLote=<%=rds_Lote("LOTE_TX_DESCRICAO")%>&pVezesImp=<%=rds_Lote("LOTE_NR_QTD_EXPORTACAO")%>&pOrdem=1"><img src="../../imagens/b04.gif" width="16" height="16" border="0">
        <%
		end if
	%>
      </a></div></td>
      <td><div align="center"><a href="consulta_lote_nao_mapeados.asp?str_Tipo_Saida=Tela&pLote=<%=rds_Lote("LOTE_NR_SEQ_LOTE")%>&pDescLote=<%=rds_Lote("LOTE_TX_DESCRICAO")%>&pVezesImp=<%=rds_Lote("LOTE_NR_QTD_EXPORTACAO")%>&pOrdem=1"><img src="../../imagens/b04.gif" width="16" height="16" border="0"></a></div></td>
      <td><div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_NR_SEQ_LOTE")%> - <%=rds_Lote("LOTE_TX_DESCRICAO")%> - <%=rds_Lote("LOTE_TX_POSSUI_VISAO")%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("ATUA_CD_NR_USUARIO")%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=formatadata("MDA","DMA",rds_Lote("LOTE_DT_ENVIO"))%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_NR_QTD_EXPORTACAO")%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=formatadata("MDA","DMA",rds_Lote("ATUA_DT_ATUALIZACAO"))%></font></div></td>
   	</tr>
    <tr>
      <td height="25" colspan="7"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&Oacute;rg&atilde;o:</font></strong> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_TX_ORGAO_SELEC")%></font></td>
    </tr>
    <tr>
      <td height="25" colspan="7"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Fun&ccedil;&atilde;o</font></strong>: <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_TX_FUNCAO_SELEC")%></font></td>
    </tr>
    <tr bgcolor="#CCCCCC">
      <td height="1" colspan="7"></td>
    </tr>
	<% rds_Lote.movenext
	Loop %>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td height="25">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>
<%
rds_Lote.close
set rds_Lote = Nothing
conn_Cogest.Close
set conn_Cogest = Nothing
%>