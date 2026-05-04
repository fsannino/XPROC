<%@LANGUAGE="VBSCRIPT"%>
<%

if request("str_Tipo_Saida")="Excel" then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

if request("pLote") <> 0 then
   str_Lote = request("pLote")
else
   str_Lote = 0
end if

if request("pOrdem") <> 0 then
   str_Ordem = request("pOrdem")
else
   str_Ordem = 0
end if

set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
conn_Cogest.CursorLocation=3

str_SQL = ""
str_SQL = str_SQL & " SELECT  "
str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO.LOTE_NR_SEQ_LOTE"
str_SQL = str_SQL & " , dbo.GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO"
str_SQL = str_SQL & " , dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"
str_SQL = str_SQL & " , dbo.GOLI_FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL = str_SQL & " , dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO"
str_SQL = str_SQL & " FROM dbo.GOLI_FUNCAO_USUARIO INNER JOIN"
str_SQL = str_SQL & " dbo.USUARIO_MAPEAMENTO ON dbo.GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN"
str_SQL = str_SQL & " dbo.FUNCAO_NEGOCIO ON dbo.GOLI_FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL = str_SQL & " WHERE dbo.GOLI_FUNCAO_USUARIO.LOTE_NR_SEQ_LOTE = " & str_Lote
if str_Ordem = 1 then
	str_SQL = str_SQL & " order by dbo.GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO  "
else
	str_SQL = str_SQL & " order by dbo.GOLI_FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO "
end if	
set rds_Lote = conn_Cogest.Execute(str_SQL)

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
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Rela&ccedil;&atilde;o de usu&aacute;rios aprovados em fun&ccedil;&atilde;o </font></div></td>
      <td><img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Lote n&uacute;mero : </font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("LOTE_NR_SEQ_LOTE")%></font></div></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="97%"  border="0" cellspacing="5" cellpadding="1">
    <tr>
      <td width="38%" bgcolor="#00FFFF"><div align="left"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif"> <a href="consulta_conteudo_lote.asp?pLote=<%=str_Lote%>&str_Tipo_Saida=Tela&pOrdem=1">Usu&aacute;rio</a></font></strong></div></td>
      <td width="49%" bgcolor="#00FFFF"><div align="left"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="consulta_conteudo_lote.asp?pLote=<%=str_Lote%>&str_Tipo_Saida=Tela&pOrdem=2">Fun&ccedil;&atilde;o</a></font></strong></div></td>
      <td width="13%" bgcolor="#00FFFF"><div align="center"><strong><font color="#0000CE" size="2" face="Verdana, Arial, Helvetica, sans-serif">Perfil</font></strong></div></td>
    </tr>
	<% int_Tot_Nao_Mapeado = 0
	   int_Tot_Mapeado = 0
	   int_Tot_Geral = 0
	do while not rds_Lote.Eof 
		int_Tot_Geral = int_Tot_Geral + 1
	%>
    <tr>
      <td><div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("USMA_CD_USUARIO")%> - <%=rds_Lote("USMA_TX_NOME_USUARIO")%></font></div></td>
      <td><div align="left"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote("FUNE_CD_FUNCAO_NEGOCIO")%> - <%=rds_Lote("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></div></td>
<%
str_SQL = " SELECT DISTINCT "
str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUNCAO_USUARIO_PERFIL.USMA_CD_USUARIO"
str_SQL = str_SQL & " FROM dbo.FUNCAO_USUARIO_PERFIL INNER JOIN"
str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3 ON "
str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL AND "
str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL = dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL"
str_SQL = str_SQL & " WHERE dbo.FUNCAO_USUARIO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = '" & rds_Lote("FUNE_CD_FUNCAO_NEGOCIO") & "'"
str_SQL = str_SQL & " AND dbo.FUNCAO_USUARIO_PERFIL.USMA_CD_USUARIO = '" & rds_Lote("USMA_CD_USUARIO") & "'"
str_SQL = str_SQL & " AND FUUP_IN_VALIDADO = 'S'"
'response.Write(str_SQL)
'response.End()
set rds_Perfil = conn_Cogest.Execute(str_SQL)
%>	  
      <td>
	    <div align="center">
          <% if not rds_Perfil.EOF then %>
          <img src="../../imagens/aprova_02.gif" width="14" height="12"></div></td>
<% int_Tot_Mapeado = int_Tot_Mapeado + 1
else %>
		<strong><font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sem perfil</font></strong>
<% int_Tot_Nao_Mapeado = int_Tot_Nao_Mapeado + 1
end if %>
    </tr>
	<% rds_Lote.movenext
	Loop %>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <p align="center">Total mapeado : <%=int_Tot_Mapeado%></p>
  <p align="center">&nbsp;</p>
  <p align="center">&nbsp;</p>
  <p align="center">Total n&atilde;o mapeado : <%=int_Tot_Nao_Mapeado%></p>
  <p align="center">Total Geral : <%=int_Tot_Geral%></p>
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