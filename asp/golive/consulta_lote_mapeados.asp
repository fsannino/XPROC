<%@LANGUAGE="VBSCRIPT"%>
<%
response.Buffer=false

set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
conn_Cogest.CursorLocation=3

if request("pLote") <> 0 then
   str_Lote = request("pLote")
else
   str_Lote = 0
end if

if request("pTipo_Saida")="Excel" then
	str_SQL = ""
    str_SQL = str_SQL & " Update "
	str_SQL = str_SQL & " GOLI_LOTE set "
	str_SQL = str_SQL & " LOTE_NR_QTD_EXPORTACAO = " & (Request("pVezesImp") + 1)
	str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO = GETDATE() "
	str_SQL = str_SQL & " where LOTE_NR_SEQ_LOTE = " & str_Lote
	'response.Write(str_SQL)
	'response.End()
	conn_Cogest.Execute str_SQL 

	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"

end if

if request("pAcao") = "C" then
	str_Acao = "C"
else
	str_Acao = "I"
end if

if request("pDescLote") <> "" then
   str_DescLote = request("pDescLote")
else
   str_DescLote = ""
end if

'str_SQL = ""
'str_SQL = str_SQL & " SELECT DISTINCT "
'str_SQL = str_SQL & " GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO"
'str_SQL = str_SQL & " , MICRO_PERFIL_R3.MIPE_TX_NOME_TECNICO"
'str_SQL = str_SQL & " , 'Z:BC_USO_GERAL' AS Expr1 "
'str_SQL = str_SQL & " FROM GOLI_FUNCAO_USUARIO INNER JOIN "
'str_SQL = str_SQL & " dbo.MACRO_PERFIL ON dbo.GOLI_FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN"
'str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3 ON dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL"
'str_SQL = str_SQL & " WHERE dbo.GOLI_FUNCAO_USUARIO.LOTE_NR_SEQ_LOTE = " & str_Lote
'str_SQL = str_SQL & " order by USMA_CD_USUARIO"

str_SQL = ""
str_SQL = str_SQL & " SELECT DISTINCT "
str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO, dbo.MICRO_PERFIL_R3.MIPE_TX_NOME_TECNICO, "
str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO.LOTE_NR_SEQ_LOTE, dbo.FUNCAO_USUARIO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO"
'str_SQL = str_SQL & " FROM dbo.GOLI_FUNCAO_USUARIO INNER JOIN"
'str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL ON "
'str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_USUARIO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO AND "
'str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO = dbo.FUNCAO_USUARIO_PERFIL.USMA_CD_USUARIO INNER JOIN"
'str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3 ON "
'str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL"
str_SQL = str_SQL & " FROM dbo.GOLI_FUNCAO_USUARIO INNER JOIN"
str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL ON "
str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_USUARIO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO AND "
str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO = dbo.FUNCAO_USUARIO_PERFIL.USMA_CD_USUARIO INNER JOIN"
str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3 ON "
str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL AND "
str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL = dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL"
str_SQL = str_SQL & " WHERE dbo.GOLI_FUNCAO_USUARIO.LOTE_NR_SEQ_LOTE = " & str_Lote
str_SQL = str_SQL & " AND FUUP_IN_VALIDADO = 'S'"
str_SQL = str_SQL & " order by dbo.GOLI_FUNCAO_USUARIO.USMA_CD_USUARIO"

str_SQL = ""
str_SQL = str_SQL & " SELECT LOTE_NR_SEQ_LOTE"
str_SQL = str_SQL & " , dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL.USMA_CD_USUARIO"
str_SQL = str_SQL & " , dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL = str_SQL & " , dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL"
str_SQL = str_SQL & " , dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL"
str_SQL = str_SQL & " , dbo.MICRO_PERFIL_R3.MIPE_TX_NOME_TECNICO"
str_SQL = str_SQL & " FROM dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL INNER JOIN"
str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3 ON "
str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL AND "
str_SQL = str_SQL & " dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL = dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL"
str_SQL = str_SQL & " WHERE LOTE_NR_SEQ_LOTE = " &  str_Lote
str_SQL = str_SQL & " order by dbo.GOLI_FUNCAO_USUARIO_COM_PERFIL.USMA_CD_USUARIO"

'response.Write(str_SQL)
'response.End()
set rds_Lote_Usu_Fun = conn_Cogest.Execute(str_SQL)
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
<% 
if request("str_Tipo_Saida")<> "Excel" then 
%>
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
      <td colspan="3" height="20"><table width="625" border="0" align="center">
        <tr>
          <td width="26"></td>
          <td width="50"><a href="javascript:print()"><img border="0" src="../../imagens/print.gif"></a></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50">
	  <% if str_Acao = "I" then %>
		  <a href="consulta_lote_mapeados.asp?pTipo_Saida=Excel&pLote=<%=str_Lote%>&pDescLote=<%=str_DescLote%>&pVezesImp=<%=Request("pVezesImp")%>" target="blank"> <img border="0" src="../../imagens/exp_excel.gif"></a>
		<% end if %>		  
		  </td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table> </td>
    </tr>
  </table>
  <table width="100%"  border="0" cellspacing="0" cellpadding="1">
    <tr>
      <td width="13%">&nbsp;</td>
      <td width="62%">&nbsp;</td>
      <td width="25%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><font size="3" face="Verdana, Arial, Helvetica, sans-serif">Rela&ccedil;&atilde;o de Usu&aacute;rio x Perfil </font></td>
      <td><img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Lote :<strong> <%=str_Lote%> - <%=str_DescLote%></strong></font></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Vezes exportado: <strong><%=Request("pVezesImp")%></strong></font></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <% end if %>
  <table width="58%"  border="0" cellspacing="5" cellpadding="1">
    <tr bgcolor="#000099">
      <td width="19%" height="30"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Usuario</font></strong></div></td>
      <td width="48%" height="30"><div align="center"><strong><font color="#FFFF00" size="2" face="Verdana, Arial, Helvetica, sans-serif">Perfil</font></strong></div></td>
    </tr>
	<% 
	str_Cd_Usu_anterior = ""
	if not rds_Lote_Usu_Fun.Eof  then
		str_Cd_Usu_anterior =  rds_Lote_Usu_Fun("USMA_CD_USUARIO")
		str_Funcao_Anterior = rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO")
		do while not rds_Lote_Usu_Fun.Eof 
	%>
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote_Usu_Fun("USMA_CD_USUARIO")%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rds_Lote_Usu_Fun("MIPE_TX_NOME_TECNICO")%></font></div></td>
    </tr>
	<%  rds_Lote_Usu_Fun.movenext
		if rds_Lote_Usu_Fun.Eof then
		   exit do 
		end if
		if str_Cd_Usu_anterior <> rds_Lote_Usu_Fun("USMA_CD_USUARIO") then
		%>
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Cd_Usu_anterior%></font></div></td>
	  <% if Left(rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO"),2) = "BW" then
	  		str_Uso_Geral = "Z:BC_TRAN_GERAIS"
		 else
	  		str_Uso_Geral = "Z:BC_USO_GERAL"
		 end if	
	  %>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Uso_Geral%></font></div></td>
    </tr>
	<% 		str_Funcao_Anterior = rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO")
	        'response.Write("left = " &  Left(rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO"),2))
			'response.Write("1 = " &  rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO"))
			
			if Left(rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO"),2) = "MM" AND _
				 rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO") <> "MM.100" AND _
				 rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO") <> "MM.66" AND _
				 rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO") <> "MM.36"  then
	%>
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Cd_Usu_anterior%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Z:MM_PB001_CON_INF_MES</font></div></td>
    </tr>
<%			end if
			if Left(rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO"),2) = "PM" then
	%>
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Cd_Usu_anterior%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Z:PM_PB001_EXIBICAO_GERAL</font></div></td>
    </tr>
<%			end if
			str_Cd_Usu_anterior = rds_Lote_Usu_Fun("USMA_CD_USUARIO")
		end if
		str_Mega = Left(rds_Lote_Usu_Fun("FUNE_CD_FUNCAO_NEGOCIO"),2)
	Loop %>
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Cd_Usu_anterior%></font></div></td>
	  <% if str_Mega = "BW" then
	  		str_Uso_Geral = "Z:BC_TRAN_GERAIS"
		 else
	  		str_Uso_Geral = "Z:BC_USO_GERAL"
		 end if	
	  %>	  
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Uso_Geral%></font></div></td>
    </tr>	
	<% 		
	        'response.Write("left = " &  Left(str_Funcao_Anterior,2))
			'response.Write("1  = " &  str_Funcao_Anterior)
	
	if Left(str_Funcao_Anterior,2) = "MM" AND _
		 str_Funcao_Anterior <> "MM.100" AND _
		 str_Funcao_Anterior <> "MM.66" AND _
		 str_Funcao_Anterior <> "MM.36"  then
	%>
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Cd_Usu_anterior%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Z:MM_PB001_CON_INF_MES</font></div></td>
    </tr>
<%
	end if
	if Left(str_Funcao_Anterior,2) = "PM" then
	%>
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Cd_Usu_anterior%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Z:PM_PB001_EXIBICAO_GERAL</font></div></td>
    </tr>
<%
	end if	
%>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <% 
  	str_SQL = ""
  else
    str_SQL = "Não existem dados a serem exportados"	%>
<table>	
    <tr>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_SQL%></font></div></td>
      <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
    </tr>
</table>	
<%	
  end if
%>
</form>
</body>
<% 
if request("str_Tipo_Saida")<> "Excel" then 
%>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
<% END IF %>
</html>
<%
rds_Lote_Usu_Fun.close
set rds_Lote_Usu_Fun = Nothing
conn_Cogest.Close
set conn_Cogest = Nothing
%>