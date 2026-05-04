
<!--#include file="../../asp/protege/protege.asp" -->
<%
if request("Excel")=1 then
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql=""
ssql="SELECT dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, dbo.CURSO.CURS_CD_CURSO, dbo.CURSO.ONDA_CD_ONDA, (SELECT DISTINCT dbo.ABRANGENCIA_CURSO.ONDA_TX_DESC_ONDA FROM dbo.ABRANGENCIA_CURSO WHERE dbo.ABRANGENCIA_CURSO.ONDA_CD_ONDA=dbo.CURSO.ONDA_CD_ONDA) AS TX_ONDA,"
ssql=ssql+"dbo.CURSO.CURS_TX_NOME_CURSO, dbo.CURSO.CURS_NUM_CARGA_CURSO, dbo.CURSO.CURS_TX_METODO_CURSO, dbo.CURSO.CURS_TX_STATUS_CURSO, "
ssql=ssql+"                      dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO, dbo.CURSO_PRE_REQUISITO.CURS_PRE_REQUISITO, dbo.CURSO.CURS_TX_PUBLICO_ALVO,  "
ssql=ssql+"                      dbo.CURSO.CURS_TX_OBJETIVO, dbo.CURSO.CURS_TX_CONTEUDO_PROGRAM, dbo.CURSO.CURS_TX_PRE_REQUISITOS "
ssql=ssql+"FROM         dbo.CURSO LEFT OUTER JOIN "
ssql=ssql+"                      dbo.CURSO_FUNCAO ON dbo.CURSO.CURS_CD_CURSO = dbo.CURSO_FUNCAO.CURS_CD_CURSO LEFT OUTER JOIN "
ssql=ssql+"                      dbo.CURSO_PRE_REQUISITO ON dbo.CURSO.CURS_CD_CURSO = dbo.CURSO_PRE_REQUISITO.CURS_CD_CURSO LEFT OUTER JOIN "
ssql=ssql+"                      dbo.MEGA_PROCESSO ON dbo.CURSO.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
ssql=ssql+"ORDER BY dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, dbo.CURSO.CURS_CD_CURSO "

SET RS=DB.EXECUTE(SSQL)
%>
<html>
	<head>
		<title>SINERGIA # XPROC # Processos de Negócio</title>
		
		<style type="text/css">
			<!--
			.style7 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11; color: #FFFFFF; }
			.style9 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 11; }
			-->
		</style>
	</head>

	<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
		<form method="POST" action="" name="frm1">
		
			<%if request("excel") <> 1 then%>
			
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
						<div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
					  <td width="195"></td>
						<td width="27"></td>  <td width="50"><a href="curso_prerequisito.asp?excel=1" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
					  <td width="28"></td>
					  <td width="26">&nbsp;</td>
					  <td width="159"></td>
					</tr>
				  </table>
				</td>
			  </tr>
			</table>
			
			<%end if%>
			
			  <table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr><td height="10"></td></tr>
				<tr>
				  <td>
					<div align="center">
					  <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório Geral de Cursos</font>           
					  <table width="983" border="0" bordercolor="#333333">
						<tr><td height="10"></td></tr>
						<tr bgcolor="#31009C">
						  <td width="89"><span class="style7">Mega-Processo</span></td>
						  <td width="43"><span class="style7">Curso</span></td>
						  <td width="45"><span class="style7">Nome</span></td>
						  <td width="45"><span class="style7">Status</span></td>
						  <td width="46"><span class="style7">Carga Hor&aacute;ria</span></td>
						  <td width="56"><span class="style7">M&eacute;todo</span></td>
						  <td width="88"><span class="style7">Abrang&ecirc;ncia</span></td>
						  <td width="84"><span class="style7">P&uacute;blico Alvo </span></td>
						  <td width="75"><span class="style7">Requisitos n&atilde;o R/3 </span></td>
						  <td width="88"><span class="style7">Objetivo</span></td>
						  <td width="89"><span class="style7">Conte&uacute;do</span></td>
						  <td width="77"><span class="style7">Fun&ccedil;&atilde;o</span></td>
						  <td width="127"><span class="style7">Curso Pr&eacute;-Requisito</span></td>
						</tr>
						<%
						anterior=""
						atual=""
						DO UNTIL RS.EOF=TRUE
							atual=rs("CURS_CD_CURSO")
										
							if COR="WHITE" then
								COR="#E4E4E4"
							else
								COR="WHITE"
							end if
							
							if trim(rs("CURS_TX_STATUS_CURSO")) = "1" then
								strStatus = "Ativo"
							else
								strStatus = "Inativo"
							end if							
							%>
							<tr>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_CD_CURSO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_TX_NOME_CURSO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=strStatus%></span></td>
							  <td bgcolor="<%=COR%>" align="center"><span class="style9"><%=rs("CURS_NUM_CARGA_CURSO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_TX_METODO_CURSO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("TX_ONDA")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_TX_PUBLICO_ALVO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_TX_PRE_REQUISITOS")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_TX_OBJETIVO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_TX_CONTEUDO_PROGRAM")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("FUNE_CD_FUNCAO_NEGOCIO")%></span></td>
							  <td bgcolor="<%=COR%>"><span class="style9"><%=rs("CURS_PRE_REQUISITO")%></span></td>
							</tr>
							<%
							if anterior<>atual then
								tem = tem + 1
							end if
							anterior=rs("CURS_CD_CURSO")
							RS.MOVENEXT
						LOOP
						%>
					  </table>
					  <p align="left" class="style9">Total de Cursos Dispon&iacute;veis : <strong><%=tem%></strong> 
					  <p align="left">          
					</div>
				  </td>
				</tr>
			  </table>
				<p style="margin-top: 0; margin-bottom: 0">
				<b>
		</form>		
		<%
		RS.close
		set RS = nothing
		%>
	</body>
</html>
