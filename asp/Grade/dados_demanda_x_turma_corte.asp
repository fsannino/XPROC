<%
Session("Corte") 		= request("selCorte")

strUnidade 				= request("selUnidade")
strDiretoria 			= request("selDiretoria")
strCurso 				= request("selCurso")

strDescentralizado 		= request("rdDescentralizado")
strEaD 					= request("rdEad")
strInLoco 				= request("rdInLoco")

strNumRel				= request("pNumRel")
strTituloRel			= request("pTituloRel")

'Response.write "strNumRel - " & strNumRel & "<br>"
'Response.write "strTituloRel - " & strTituloRel & "<br>"

'set db_banco = Server.CreateObject("AdoDB.Connection")
'db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("Petrobras 2004_v2.mdb")
'db_banco.open Session("Conn_String_Cogest_Gravacao")
'db_banco.CursorLocation = 3

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3


if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

strSQLTurma = ""
strSQLTurma = strSQLTurma & " SELECT "
strSQLTurma = strSQLTurma & " dbo.GRADE_UNIDADE.UNID_CD_UNIDADE"
strSQLTurma = strSQLTurma & " , dbo.GRADE_UNIDADE.UNID_TX_DESC_UNIDADE"
strSQLTurma = strSQLTurma & " , dbo.GRADE_DEMANDA_ORIGINAL_SEM.CURS_CD_CURSO"
strSQLTurma = strSQLTurma & " , SUM(dbo.GRADE_DEMANDA_ORIGINAL_SEM.DEMA_NR_TOTAL) AS TOTAL_DEMANDA"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_NOME_CURSO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_NUM_CARGA_CURSO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_METODO_CURSO "
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.MEPR_CD_MEGA_PROCESSO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_CENTRALIZADO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_IN_LOCO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_DT_FIM_MATERIAL_DIDATICO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.ONDA_CD_ONDA_ABRANGENCIA"
strSQLTurma = strSQLTurma & " , dbo.GRADE_UNIDADE.CTRO_CD_CENTRO_TREINAMENTO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_DIRETORIA.DIRE_TX_DESC_DIRETORIA"
strSQLTurma = strSQLTurma & " , dbo.GRADE_DIRETORIA.ORLO_CD_ORG_LOT"
strSQLTurma = strSQLTurma & " FROM dbo.GRADE_UNIDADE_ORGAO_MENOR "
strSQLTurma = strSQLTurma & " INNER JOIN dbo.GRADE_DEMANDA_ORIGINAL_SEM ON dbo.GRADE_UNIDADE_ORGAO_MENOR.ORME_CD_ORG_MENOR = dbo.GRADE_DEMANDA_ORIGINAL_SEM.ORME_CD_ORG_MENOR "
strSQLTurma = strSQLTurma & " INNER JOIN dbo.GRADE_UNIDADE ON dbo.GRADE_UNIDADE_ORGAO_MENOR.UNID_CD_UNIDADE = dbo.GRADE_UNIDADE.UNID_CD_UNIDADE "
strSQLTurma = strSQLTurma & " INNER JOIN dbo.GRADE_CURSO ON dbo.GRADE_DEMANDA_ORIGINAL_SEM.CORT_CD_CORTE = dbo.GRADE_CURSO.CORT_CD_CORTE "
strSQLTurma = strSQLTurma & " AND dbo.GRADE_DEMANDA_ORIGINAL_SEM.CURS_CD_CURSO = dbo.GRADE_CURSO.CURS_CD_CURSO"
strSQLTurma = strSQLTurma & " INNER JOIN dbo.GRADE_DIRETORIA ON dbo.GRADE_UNIDADE.ORLO_CD_ORG_LOT_DIR = dbo.GRADE_DIRETORIA.ORLO_CD_ORG_LOT "
strSQLTurma = strSQLTurma & " WHERE dbo.GRADE_DEMANDA_ORIGINAL_SEM.CORT_CD_CORTE > 0 "
strSQLTurma = strSQLTurma & " AND dbo.GRADE_CURSO.CURS_CD_CURSO not in ("

int_Cd_Corte = 2

strSQLTurma = strSQLTurma & " 		Select CURS_CD_CURSO "
strSQLTurma = strSQLTurma & " 		from GRADE_CURSO_UNIDADE "
strSQLTurma = strSQLTurma & " 		WHERE CORT_CD_CORTE = " &  int_Cd_Corte & ") "

if int_Cd_Corte <> "" then
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_DEMANDA_ORIGINAL_SEM.CORT_CD_CORTE = " & int_Cd_Corte
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_UNIDADE.CORT_CD_CORTE = " & int_Cd_Corte
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_CURSO.CORT_CD_CORTE = " & int_Cd_Corte
end if
strSQLTurma = strSQLTurma & " GROUP BY dbo.GRADE_UNIDADE.UNID_TX_DESC_UNIDADE"
strSQLTurma = strSQLTurma & " , dbo.GRADE_DEMANDA_ORIGINAL_SEM.CURS_CD_CURSO" 
strSQLTurma = strSQLTurma & " , dbo.GRADE_UNIDADE.UNID_CD_UNIDADE"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_NOME_CURSO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_NUM_CARGA_CURSO" 
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_METODO_CURSO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.MEPR_CD_MEGA_PROCESSO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_CENTRALIZADO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_TX_IN_LOCO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.CURS_DT_FIM_MATERIAL_DIDATICO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_CURSO.ONDA_CD_ONDA_ABRANGENCIA"
strSQLTurma = strSQLTurma & " , dbo.GRADE_UNIDADE.CTRO_CD_CENTRO_TREINAMENTO"
strSQLTurma = strSQLTurma & " , dbo.GRADE_DIRETORIA.DIRE_TX_DESC_DIRETORIA"
strSQLTurma = strSQLTurma & " , dbo.GRADE_DIRETORIA.ORLO_CD_ORG_LOT"
strSQLTurma = strSQLTurma & " HAVING dbo.GRADE_CURSO.MEPR_CD_MEGA_PROCESSO > 0 "
if int_Cd_Diretoria <> "" then
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_DIRETORIA.ORLO_CD_ORG_LOT = " & int_Cd_Diretoria
end if
if int_Cd_CT <> "" then
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_UNIDADE.CTRO_CD_CENTRO_TREINAMENTO  = " & int_Cd_CT
end if
if int_Cd_Mega_Processo <> "" then	
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_CURSO.MEPR_CD_MEGA_PROCESSO = " & int_Cd_Mega_Processo 
end if	
if str_Cd_Curso <> "" then
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_CURSO.CURS_CD_CURSO = '" & str_Cd_Curso & "'"
end if
if str_Cd_Metodo_Curso <> "" then
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_CURSO.CURS_TX_METODO_CURSO = '" & str_Cd_Metodo_Curso & "'"
end if
if str_Centralizado <> "" then
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_CURSO.CURS_TX_CENTRALIZADO = '" & str_Centralizado & "'"
end if
if str_InLoco <> "" then
	strSQLTurma = strSQLTurma & " AND dbo.GRADE_CURSO.CURS_TX_IN_LOCO = '" & str_InLoco & "'" 
end if

strSQLTurma = strSQLTurma & " order by dbo.GRADE_DIRETORIA.DIRE_TX_DESC_DIRETORIA, dbo.GRADE_UNIDADE.UNID_TX_DESC_UNIDADE, dbo.GRADE_DEMANDA_ORIGINAL_SEM.CURS_CD_CURSO  "
'Response.write strSQLTurma
'Response.end		
		
set rstTurmas = db_banco.execute(strSQLTurma)				
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
        <style type="text/css">
<!--
.style5 {font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; font-size: 12px; }
.style8 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
-->
        </style>
<script>
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
</script>			
</head>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">		
		<%		
		if request("excel") <> 1 then
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
					  <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" width="27" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
				  </td>
				</tr>
				<tr> 
				  <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
					<div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
					<div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
				  </td>
				  <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
					<div align="center"><a href="../../indexA_grade.asp?selCorte=<%=Session("Corte")%>"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
				  </td>
				</tr>
			  </table>
			</td>
		  </tr>
		  <tr bgcolor="#F1F1F1">
			<td colspan="3" height="20">
			  <table width="625" border="0" align="center">
				<tr>
					<td width="26"></td>
				  <td width="50"></td>
				  <td width="26">&nbsp;</td>
				  <td width="195"></td>
					<td width="27"></td>  
					<td width="50"></td>
				  <td width="28"></td>
				  <td width="26">&nbsp;</td>
				  <td width="159"></td>
				</tr>
			  </table>
			</td>
		  </tr>
		</table>	
		
			<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td width="561"></td>	
					<td width="237">
						<div align="center">	
							 <a href="gera_consulta_turma.asp?excel=1&amp;selDiretoria=<%=strDiretoria%>&amp;selUnidade=<%=strUnidade%>&amp;rdDescentralizado=<%=strDescentralizado%>&amp;rdEad=<%=strEaD%>&amp;rdInLoco=<%=strInLoco%>&amp;selCurso=<%=strCurso%>&amp;pTituloRel=<%=strTituloRel%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
						</div>
					</td>						
				    <td width="190"><img src="../../Flash/preloader.gif" name="loader" width="190" height="50" id="loader"></td>
				</tr>
			</table>
		<%end if%>		
		
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
			<tr>
			  <td height="10">
			  </td>
			</tr>
			<tr>
			  <td>
				<div align="center"><font face="Verdana" color="#330099" size="3"><b>Relatório de <%=strTituloRel%> - Grade de Treinamento</b></font></div>
			  </td>
			</tr>
			<tr>
			  <td>&nbsp;</td>
			</tr>
    </table>		

    <table width="956" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr bgcolor="#CCCCCC">
            <td width="92"><span class="style5">Diretoria</span></td>
            <td width="102"><span class="style5">Unidade</span></td>
            <td width="102"><div align="center"><span class="style5">Curso</span></div></td>
            <td width="109"><span class="style5">M&eacute;todo</span></td>
            <td width="136"><span class="style5">Tipo</span></td>
            <td width="92"><span class="style5">C.T.</span></td>
            <td width="70"><div align="right"><span class="style5">Qtd demanda ant </span></div></td>
            <td width="90"><div align="right"><span class="style5">Qtd turmas ant </span></div></td>
            <td width="80"><div align="right"></div>              
              <div align="right"><span class="style5">Qtd demanda atu</span></div></td>
          </tr>
		  <%	int_Tot_Registro = 0 
		  	do while not rstTurmas.eof 
				if str_Cor_Linha = "#FFFFFF" then 
				   str_Cor_Linha = "#F1F1F1"
				else
				   str_Cor_Linha = "#FFFFFF"
				end if				
			%>
          <tr bgcolor="<%=str_Cor_Linha%>">
		  <% 
		  %>
            <td><span class="style8"><%=rstTurmas("DIRE_TX_DESC_DIRETORIA")%></span></td>
            <td><span class="style8"><%=rstTurmas("UNID_TX_DESC_UNIDADE")%></span></td>
            <td height="19"><div align="center"><span class="style8"><%=rstTurmas("CURS_CD_CURSO")%></span></div></td>
            <td><span class="style8"><%=rstTurmas("CURS_TX_METODO_CURSO")%></span></td>
			<%	if rstTurmas("CURS_TX_CENTRALIZADO") = "S" then
					str_Desc_Centralizado = "Centralizado" 
				else
					str_Desc_Centralizado = "Descentralizado"
				end if	
				if rstTurmas("CURS_TX_IN_LOCO") = "S" then
					str_Desc_InLoco = " - In Loco" 
				else
					str_Desc_InLoco = ""
				end if					
			 %>
            <td><span class="style8"><%=str_Desc_Centralizado%><%=str_Desc_InLoco%> </span></td>
			<% 	str_Sql = ""
				str_Sql = str_Sql & " SELECT "
				str_Sql = str_Sql & " CORT_CD_CORTE"
				str_Sql = str_Sql & " , CTRO_CD_CENTRO_TREINAMENTO"
				str_Sql = str_Sql & " , CTRO_TX_NOME_CENTRO_TREINAMENTO"
				str_Sql = str_Sql & " FROM  dbo.GRADE_CENTRO_TREINAMENTO"
				str_Sql = str_Sql & " WHERE CTRO_CD_CENTRO_TREINAMENTO = " &  rstTurmas("CTRO_CD_CENTRO_TREINAMENTO")
				str_Sql = str_Sql & " AND CORT_CD_CORTE = " & int_Cd_Corte
				set rds_Geral = db_banco.execute(str_Sql)			
				if not	rds_Geral.Eof then
					str_Desc_CT = rds_Geral("CTRO_TX_NOME_CENTRO_TREINAMENTO")
				else
					str_Desc_CT = ""
				end if
				rds_Geral.close
			 %>
            <td><span class="style8"><%=str_Desc_CT%></span></td>
			<% 	str_Sql = ""
				str_Sql = str_Sql & " SELECT "     
				str_Sql = str_Sql & " Unidade"
				str_Sql = str_Sql & " , CodCurso"
				str_Sql = str_Sql & " , tot_usu"
				str_Sql = str_Sql & " FROM GRADE_DEMANDA_ANTERIOR"
				str_Sql = str_Sql & " WHERE "
				str_Sql = str_Sql & " Unidade = '" & rstTurmas("UNID_TX_DESC_UNIDADE") & "'"
				str_Sql = str_Sql & " AND CodCurso = '" & rstTurmas("CURS_CD_CURSO") & "'"
				set rds_Geral = db_banco.execute(str_Sql)			
				if not	rds_Geral.Eof then
					ind_Demanda_Anterior = rds_Geral("tot_usu")
				else
					ind_Demanda_Anterior = ""
				end if
				rds_Geral.close				
			%>
            <td><div align="right"><span class="style8"><%=ind_Demanda_Anterior%></span></div></td>
			<%
				str_Sql = "" 
				str_Sql = str_Sql & " SELECT "
				str_Sql = str_Sql & " COUNT(TURM_TX_DESC_TURMA) AS QTD_TURMA"
				str_Sql = str_Sql & " , TURM_TX_UNI_DIR_ANTERIOR"
				str_Sql = str_Sql & " , CURS_CD_CURSO"
				str_Sql = str_Sql & " FROM "
				str_Sql = str_Sql & " dbo.GRADE_TURMA"
				str_Sql = str_Sql & " WHERE "    
				str_Sql = str_Sql & " CORT_CD_CORTE = " & int_Cd_Corte - 1
				str_Sql = str_Sql & " GROUP BY "
				str_Sql = str_Sql & " TURM_TX_UNI_DIR_ANTERIOR "
				str_Sql = str_Sql & " , CURS_CD_CURSO"
				str_Sql = str_Sql & " HAVING TURM_TX_UNI_DIR_ANTERIOR = '" & rstTurmas("UNID_TX_DESC_UNIDADE") & "'" 
				str_Sql = str_Sql & " AND CURS_CD_CURSO = '" & rstTurmas("CURS_CD_CURSO") & "'"
				'response.Write(str_Sql)
				'response.End()
				set rds_Geral = db_banco.execute(str_Sql)			
				if not	rds_Geral.Eof then
					int_Qtd_Turma_Anterior = rds_Geral("QTD_TURMA")
				else
					int_Qtd_Turma_Anterior = ""
				end if
				rds_Geral.close								
			%>
            <td><div align="right"><span class="style8"><%=int_Qtd_Turma_Anterior%></span></div></td>
            <td><div align="right"><span class="style8"><%=rstTurmas("TOTAL_DEMANDA")%></span></div></td>
          </tr>
		  <%	int_Tot_Registro = int_Tot_Registro + 1 
		  	rstTurmas.movenext
		  loop %>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
</table>
	    <p align="center" class="style5">Total de registros = <%=int_Tot_Registro%></p>
	    <table width="98%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><span class="style5">Unidade</span></td>
            <td><span class="style5">Curso</span></td>
            <td><span class="style5">Qtd demanda</span></td>
            <td><span class="style5">Unidade</span></td>
            <td><span class="style5">Unidade</span></td>
            <td><span class="style5">Unidade</span></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>
	    <p align="center" class="style5">&nbsp;</p>
	    <p align="center" class="style5">&nbsp;</p>
</body>	
	<%
	rstTurmas.close
	set rstTurmas = nothing
	db_banco.close
	set db_banco = nothing
	%>		
<script>
MM_swapImage('loader','','../../Flash/branco.gif',1);
</script>	
</html>
