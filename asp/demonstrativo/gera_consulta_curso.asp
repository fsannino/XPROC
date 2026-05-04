<%
set fs = server.createobject("Scripting.FileSystemObject")
set arquivo = fs.GetFile(server.mappath("banco.mdb"))

atualiza = arquivo.DateLastModified

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("banco.mdb")
db_banco.CursorLocation = 3

set db_cogest = Server.CreateObject("AdoDB.Connection")
db_cogest.Open "Provider=SQLOLEDB.1;server=S6000db21;pwd=cogest00;uid=cogest;database=cogest"
db_cogest.cursorlocation = 3

if request("excel") = 1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

situacao = request("selStatus")
'Response.write "situacao = " & situacao
'Response.end

v1=0
v2=0
t1=0
a1=0
pr=0
re=0

select case request("selAbrangencia")

case "6,8,9"
	tit = "TODAS"
case "6,9"
	tit = "PETROBRAS"
case "8,9"
	tit = "REFAP"
end select

strMega = request("selMega")

strSQLCurso = ""
strSQLCurso = strSQLCurso & "SELECT  * FROM DEMONSTRATIVO "

if strMega <> "0" then
	strSQLCurso = strSQLCurso & "WHERE DEMO_CD_CURSO LIKE '" & strMega & "%' "	
end if

frente = request("selFrente")

nome_curso = ""

select case frente
case 1
	strSQLCurso = strSQLCurso & "WHERE DEMO_CD_CURSO LIKE 'BW%' "
	nome_curso = "BW"
case 2
	strSQLCurso = strSQLCurso & "WHERE DEMO_CD_CURSO LIKE 'FIN%' "
	nome_curso = "FIN"
case 3
	strSQLCurso = strSQLCurso & "WHERE DEMO_CD_CURSO LIKE 'PLC%' OR DEMO_CD_CURSO LIKE 'SUP%' OR DEMO_CD_CURSO LIKE 'COM%' "
	nome_curso = "OIL & CO"
case 4
	strSQLCurso = strSQLCurso & "WHERE DEMO_CD_CURSO LIKE 'EMP%' OR DEMO_CD_CURSO LIKE 'LTE%' OR DEMO_CD_CURSO LIKE 'MAN%' OR DEMO_CD_CURSO LIKE 'MES%' OR DEMO_CD_CURSO LIKE 'POS%' OR DEMO_CD_CURSO LIKE 'PRD%' OR DEMO_CD_CURSO LIKE 'QUA%' "
	nome_curso = "P* & MES"
case 5
	strSQLCurso = strSQLCurso & "WHERE DEMO_CD_CURSO LIKE 'RHU%' "
	nome_curso = "RH"	
end select

strSQLCurso = strSQLCurso & "ORDER BY DEMO_CD_CURSO"

set rs_curso = db_banco.execute(strSQLCurso)	
%>
<html>
	<head>
		<title>:: Demostrativo de Cursos</title>
		<style type="text/css">
			<!--
			.style2 
			{
				font-family: Verdana, Arial, Helvetica, sans-serif;
				font-weight: bold;
				color: ##000080;
			}		
			-->
		</style>
		
		<script language="JavaScript">		
			function impressao() 
			{
				window.open('impressao.asp?selMega=<%=strMega%>&selAbrangencia=<%=request("selAbrangencia")%>&selStatus=<%=situacao%>','jan1','toolbar=no, location=no, scrollbars=no, status=no, directories=no, resizable=no, menubar=no, fullscreen=no, height=50, width=250, status=no, top=200, left=260');
			}
		</script>
	</head>
	
	<body>	
		<%	
		if nome_curso = "" then
		
		Select Case strMega
			Case "SUP"
				nome_curso = "SUPRIMENTOS DE PETRÓLEO"
			Case "MES"
				nome_curso = "MATERIAIS"
			Case "COM"
				nome_curso = "VENDAS E DISTRIBUIÇÃO"
			Case "EMP"
				nome_curso = "EMPREENDIMENTOS"
			Case "MAN"
				nome_curso = "MANUTENÇÃO"
			Case "POS"
				nome_curso = "POCOS"
			Case "PRD"
				nome_curso = "PRODUÇÃO"
			Case "QUA"
				nome_curso = "QUALIDADE"
			Case "LTE"
				nome_curso = "LOGÍSTICA"
			Case "PLC"
				nome_curso = "PLANEJAMENTO E CONTROLE"
			Case "FIN"
				nome_curso = "FINANÇAS"
			Case "RHU"
				nome_curso = "RECURSOS HUMANOS"			
			Case "BW","BWA","BWC","BWF","BWG","BWJ","BWL","BWM","BWP","BWQ","BWR","BWS","BWT","BWU"
				nome_curso = "BW"
			Case Else
				nome_curso = "TODOS OS MEGA-PROCESSOS"
        End Select
		
		end if
		
		if request("excel") <> 1 then
		%>
			<table cellspacing="0" cellpadding="0" border="0">
				<tr>
					<td width="720"></td>	
					<td width="50">
						<div align="center">	 
							 <a href="gera_consulta_curso.asp?excel=1&amp;selMega=<%=strMega%>&selAbrangencia=<%=request("selAbrangencia")%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif" title="Exportar para o Excel"></a>
						</div>
					</td>
					<!--<td width="20"></td>							
					<td width="50">			 
						<div align="center">
							<a href="javascript:print()"><img border="0" src="../../imagens/print.gif"></a>
						</div>
					</td>-->					
					<td width="20"></td>							
					<td width="50">			 
						<div align="center">
							<a href="#" onclick="impressao();"><img border="0" src="../../imagens/print.gif" title="Imprimir Consulta"></a>
						</div>
					</td>
					
					<td width="20"></td>					
					<td width="50">						
						<div align="center">
							<p align="center"><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/demonstrativo/grafico.asp?Abrag=<%=request("selAbrangencia")%>" target="blank"><img border="0" src="../../imagens/grafico.gif" title="Gerar Gráfico"></a>
						</div>	
					</td>
					<td width="20"></td>										
					<td width="50">						
						<div align="center">
							<p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/link_retornar.gif" title="Voltar"></a>
						</div>	
					</td>
				</tr>
			</table>
		<%end if%>		
		
		
<table cellspacing="0" cellpadding="0" border="0" width="100%" height="188">
  <tr> 
    <td height="20" colspan="2"></td>
  </tr>
  <tr> 
    <td align="left" colspan="10" class="style2">RELATÓRIO - ACOMPANHAMENTO DE 
      MATERIAL DIDÁTICO - <%=nome_curso%></td>
  </tr>
  <tr> 
    <td height="31" colspan="2"> 
      <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><b>ABRANGÊNCIA: <%=tit%></b></font></p>
    </td>
  </tr>
  <tr>
    <td height="10"></td>
	<td height="10"></td>
	<td height="10"></td>
  </tr>
  <tr> 
    <td height="30" width="20" valign="middle" align="center"><img src="marcador.gif" border="0" title="Indicação de Curso em atraso."></td>
    <td height="30"><font color="#FF3401" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Legenda 
      - Possíveis problemas:</b><br>
      </font> <font color="#FF3401" face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
      (Data de Término do curso é menor do que a Data Atual e Status do Curso é diferente<br> de "Em Aprovação Coordenador", "Em Aprovação 
      Treinameto" ou "Publicado em Produção")<br> ou (Data Início do curso
      é menor ou igual a Data Atual e o Status do Curso é igual a "Não Inicializado") </font></td>
    <td align="right" valign="top"><font color="#FF3401" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Informações 
      Atualizadas em:<br>
      <%=atualiza%> Hs</b></font></td>
  </tr>
  <tr> 
    <td height="31" width="20" valign="middle" align="center"><img src="amarelo.gif" width="17" height="17"></td>
    <td height="31"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#C6C600">Curso 
      em Aprova&ccedil;&atilde;o com Treinamento</font></td>
    <td align="right" valign="top" height="31">&nbsp;</td>
  </tr>
  <tr> 
    <td height="28" width="20" valign="middle" align="center"><img src="verde.gif" width="19" height="18"></td>
    <td height="28"> <font color="#009900" face="Verdana, Arial, Helvetica, sans-serif" size="1">Curso 
      Publicado em Produ&ccedil;&atilde;o</font> </td>
    <td align="right" valign="top" height="28"> <font color="#FF3401" face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
      </font> </td>
  </tr>
  
   <tr> 
    <td height="28" width="20" valign="middle" align="center"><img src="azul.gif" width="17" height="17"></td>
    <td height="28"> <font color="#4F29F3" face="Verdana, Arial, Helvetica, sans-serif" size="1">Aprovação do Coordenador em atraso</font> </td>
    <td align="right" valign="top" height="28"> <font color="#FF3401" face="Verdana, Arial, Helvetica, sans-serif" size="1"> 
      </font> </td>
  </tr>
 
  <tr> 
    <td height="10" colspan="2"></td>
  </tr>
</table>		
				
		
<table width="1378" border="0" cellpadding="2" cellspacing="2">
  <tr bgcolor="#ECE9CF"> 
    <td colspan="10" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Informa&ccedil;&otilde;es 
      Sobre o Curso</b></font> </td>
    <td colspan="2" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Informa&ccedil;&otilde;es 
      sobre BPPs</b></font> </td>
    <td colspan="7" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Informa&ccedil;&otilde;es 
      Sobre o Material Did&aacute;tico</b></font> </td>
  </tr>
  <tr bgcolor="#ECE9CF"> 
    <td width="65" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Situa&ccedil;&atilde;o 
      do curso</b></font> </td>
    <td width="43" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Curso</b></font> 
    </td>
    <td width="165" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>T&iacute;tulo</b></font> 
    </td>
    <td width="58" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>N&uacute;m. 
      de usu&aacute;rios priorizados</b></font></td>
    <td width="52" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Carga 
      Hor&aacute;ria</b></font> </td>
    <td width="64" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Pr&eacute; 
      Requisito</b></font> </td>
    <td width="65" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>M&eacute;todo 
      de Aplica&ccedil;&atilde;o</b></font> </td>
    <td width="46" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Status 
      do Curso</b></font> </td>
    <td colspan="2" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Previs&atilde;o</b></font> 
    </td>
    <td width="62" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Quant. 
      BPPs relativos ao Curso</b></font> </td>
    <td width="76" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Quant. 
      BPPs associados ao Curso</b></font> </td>
    <td width="84" align="center" rowspan="2"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Quant. 
      Unidades Cadastradas</b></font> </td>
    <td colspan="6" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Status</b></font> 
    </td>
  </tr>
  <tr bgcolor="#ECE9CF"> 
    <td width="37" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Data 
      de In&iacute;cio</b></font> </td>
    <td width="56" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Data 
      de T&eacute;rmino</b></font> </td>
    <td width="75" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Em 
      Elabora&ccedil;&atilde;o</b></font> </td>
    <td width="64" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Liberado 
      para Procwork</b></font> </td>
    <td width="43" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Em 
      Ajuste</b></font> </td>
    <td width="66" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Em alida&ccedil;&atilde;o</b></font> </td>
    <td width="65" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Aprovado</b></font> </td>
    <td width="76" align="center"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Liberado 
      para a Equipe Pedag&oacute;gica</b></font> </td>
  </tr>
  <%
			
			'*** INICIALIZAÇÕES ***	
			strCor = "#FFFFFF"
			intTotal = 0
			
			if not rs_curso.eof then			
				
				do until rs_curso.eof = true
					
					pinta1 = false
					pinta2 = false
					pinta3 = false
					pinta4 = false
					
					exibe_curso = 0
					
					onda = request("selAbrangencia")
					
					ssql1 = "SELECT * FROM CURSO WHERE CURS_CD_CURSO = '" & rs_curso("demo_cd_curso") & "' AND ONDA_CD_ONDA IN (" & onda & ")"
					
					set tem_curso = db_cogest.execute(ssql1)

					if tem_curso.eof=false then

					'*** INICIALIZAÇÕES DO LOOP ***		
					int_usuarios 	= 0
					pinta 			= False   
					sit1 			= False
					sit2 			= False
				
					'*** REFERENTE AO TOTAL DE USUÁRIOS ***
					strCurso_Atual = rs_curso("demo_cd_curso")

					'*** REFERENTE AO CURSO COM PLOBLEMA ***								
					str_estado 	= rs_curso("demo_status")
					data_inicio = rs_curso("demo_previsao_inicio")
					data_fim 	= rs_curso("demo_previsao_termino")

					If ((data_fim < Date) And ((str_estado <> "Em Aprovação Coordenador") And (str_estado <> "Publicado em Produção") And (str_estado <> "Em Aprovação Treinamento"))) OR ((data_inicio <= Date) And (str_estado = "Não Inicializado")) Then
						sit1 = True
					Else
						sit1 = False
					End If

					If sit1 = True Then
						pinta1 = True
					Else
						pinta1 = False
					End If
					
					if pinta1 = false then											
						if str_estado = "Publicado em Produção" then
							pinta2 = true					
						elseif str_estado = "Em Aprovação Coordenador" and data_fim < Date then
							pinta4 = true
						else 
							if str_estado = "Em Aprovação Treinamento" then
								pinta3 = true	
							end if							
						end if
					end if
					
					if situacao = 0 then
						exibe_curso = 1
					else					
						if pinta1 = true and situacao = 1 then
							exibe_curso=1
						else
							if pinta3 = true and situacao = 2 then
								exibe_curso=1
							else
								if pinta2 = true and situacao = 3 then
									exibe_curso = 1
								else
									exibe_curso = 0
								end if
							end if
						end if
					end if
					
					end if

					if exibe_curso = 1 then
					
					if ((data_fim)<=Date) And (str_estado = "Publicado em Produção") then
						re = re + 1
					else
						if (str_estado = "Publicado em Produção") then
							re = re + 1	
						else
							if (data_fim<=Date) then
								pr = pr + 1
							end if
						end if
					end if
					
					ssql = ""
					ssql = "SELECT COUNT(USMA_CD_USUARIO) AS CONTA, CURS_CD_CURSO FROM USU_CUR_FUN WHERE FUUS_IN_PRIORITARIO=1 "
					ssql = ssql + "AND CURS_CD_CURSO ='" & strCurso_Atual & "' "
					ssql = ssql + "GROUP BY CURS_CD_CURSO ORDER BY CURS_CD_CURSO"
					
					Set rsTotalUsuarios = db_cogest.Execute(ssql)
					
					if not rsTotalUsuarios.EOF then
						int_usuarios = rsTotalUsuarios("CONTA")						
					End If							
				
					rsTotalUsuarios.close
					set rsTotalUsuarios = nothing
					
					if strCor = "#FFFFFF" then
						strCor = "#F4F3EA"
					else
						strCor = "#FFFFFF"
					end if												
					%>
				  <tr bgcolor="<%=strCor%>"> 
				  <td width="65" align="center" title="Situação do Curso"> 
			      <%
							if pinta1 = true then
								t1 = t1 + 1
					   		%>
					      <img src="marcador.gif" border="0" title="Curso em Atraso"> 
	  
					      <%
							end if		
							%>
					    
						  <%
							if pinta2 = true then
							  v1 = v1 + 1 
					   		%>
					      <img src="verde.gif" border="0" title="Publicado em Produção">
					      <%
							end if		
							%>
						    <%
							if pinta3 = true then
						    a1 = a1 + 1
					   		%>
					      <img src="amarelo.gif" border="0" title="Em Aprovação com Eq. Treinamento"> 

						   <%
							end if		
							%>
							
							<%
							if pinta4 = true then
								v2 = v2 + 1
								%>
							    <img src="azul.gif" border="0" title="Aprovação do Coordenador em atraso">	
							   <%
							end if		
							%>
</td>
    <td width="43" align="center" title="Curso"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_cd_curso")%> 
      <!--Curso-->
      </font> </td>
    <td width="165" align="center" title="Título"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_nome_curso")%> 
      <!--T&iacute;tulo-->
      </font> </td>
    <td width="58" align="center" title="Núm. de usuários priorizados"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=int_usuarios%></font> 
    </td>
    <td width="52" align="center" title="Carga Horária"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_carga_horaria")%> 
      <!--Carga Hor&aacute;ria-->
      </font> </td>
    <td width="64" align="center" title="Pré-Requisito"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_pre_requisito")%> 
      <!--Pr&eacute; Requisito-->
      </font> </td>
    <td width="65" align="center" title="Método de Aplicação"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_metodo")%> 
      <!--M&eacute;todo de Aplica&ccedil;&atilde;o-->
      </font> </td>
    <td width="46" align="center" title="Status do Curso"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_status")%> 
      <!--Status do Curso-->
      </font> </td>
    <% 	
						if Trim(rs_curso("demo_previsao_inicio")) <> "" then	
							strDia = ""		
							strMes = ""
							strAno = ""										
							vetDtAprov = split(Trim(rs_curso("demo_previsao_inicio")),"/")						
							strDia = trim(vetDtAprov(0))
							if cint(strDia) < 10 then
								strDia = "0" & strDia
							end if			
							strMes = trim(vetDtAprov(1))			
							if cint(strMes) < 10 then
								strMes = "0" & strMes
							end if
							strAno = trim(vetDtAprov(2))
							dat_DtInicio = strDia & "/" & strMes & "/" & strAno 
						else
							dat_DtInicio = ""
						end if
						%>
    <td width="37" align="center" title="Data de Início do Curso"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=dat_DtInicio%> 
      <!--Data de In&iacute;cio-->
      </font> </td>
    <% 	
						if Trim(rs_curso("demo_previsao_termino")) <> "" then						
							strDia = ""		
							strMes = ""
							strAno = ""						
							vetDtAprov = split(Trim(rs_curso("demo_previsao_termino")),"/")						
							strDia = trim(vetDtAprov(0))
							if cint(strDia) < 10 then
								strDia = "0" & strDia
							end if			
							strMes = trim(vetDtAprov(1))			
							if cint(strMes) < 10 then
								strMes = "0" & strMes
							end if
							strAno = trim(vetDtAprov(2))
							dat_DtFim = strDia & "/" & strMes & "/" & strAno 
						else
							dat_DtFim = ""
						end if
						%>
    <td width="56" align="center" title="Data de Término do Curso"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=dat_DtFim%> 
      <!--Data de T&eacute;rmino-->
      </font> </td>
    <td width="62" align="center" title="Quant. BPPs relativos ao Curso"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_qtd_BPP_rel_curso")%> 
      <!--QTD BPP rel Curso-->
      </font> </td>
    <td width="76" align="center" title="Quant. BPPs associados ao Curso "> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_qtd_BPP_ass_curso")%> 
      <!--QTD BPP ass Curso-->
      </font> </td>
    <td width="84" align="center" title="Quant. Unidades Cadastradas "> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_qtd_unidades")%> 
      <!--QTD Unidades-->
      </font> </td>
    <td width="75" align="center" title="Em Elaboração "> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_em_elaboracao")%> 
      <!--Em Elabora&ccedil;&atilde;o-->
      </font> </td>
    <td width="64" align="center" title="Liberado para Procwork "> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_liberado_procwork")%> 
      <!--Liberado para Procwork-->
      </font> </td>
    <td width="43" align="center" title="Em Ajuste "> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_em_ajuste")%> 
      <!--Em Ajuste-->
      </font> </td>
    <td width="66" align="center" title="Em Validação "> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_em_validacao")%> 
      <!--Em Valida&ccedil;&atilde;o-->
      </font> </td>
    <td width="65" align="center" title="Aprovado"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_aprovado")%> 
      <!--Aprovado-->
      </font> </td>
    <td width="76" align="center" title="Liberado para a Equipe Pedagógica "> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=rs_curso("demo_liberados_eq_pedagogica")%> 
      <!--Liberado para a Equipe Pedag&oacute;gica-->
      </font> </td>
  </tr>
  <%
					intTotal = intTotal + 1
					
					end if
						
					rs_curso.movenext					
				loop
				
				%>
  <tr> 
    <td height="20"></td>
  </tr>
  <tr> 
    <td colspan="10"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Total 
      de Cursos:</b></font>&nbsp;&nbsp; <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><%=intTotal%></b></font>	
    </td>
  </tr>
  
  			<%
			else
			%>
			  <tr> 
				<td colspan="10"> <font color="#000080" face="Verdana, Arial, Helvetica, sans-serif" size="1">Não 
				  existe registros para esta consulta.</font> </td>
			  </tr>
			  <%
			end if
			
			rs_curso.close
			set rs_curso = nothing			
			%>
</table>		

  <%
  if intTotal>0 then
  %>
<table width="63%" border="0">
  <tr> 
    <td colspan="4"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#000080">Total 
      por Situa&ccedil;&atilde;o</font></b></font></td>
  </tr>
  <tr> 
    <td height="12" colspan="2">&nbsp;</td>
    <td height="12" colspan="2">&nbsp;</td>
  </tr>
  <tr> 
    <td width="18%"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="marcador.gif" border="0" title="Indicação de Curso em atraso."></font></b></div>
    </td>
    <td width="11%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><%=t1%></font></td>
    <td width="36%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000080">Total 
      de Cursos Previstos : </font></b></td>
    <td width="35%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><%=pr%></font></b></td>
  </tr>
  <tr> 
    <td width="18%"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="amarelo.gif" width="17" height="17"></font></b></div>
    </td>
    <td width="11%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><%=a1%></font></td>
    <td width="36%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000080">Total 
      de Cursos Prontos : </font></b></td>
    <td width="35%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><%=re%></font></b></td>
  </tr>
  <tr> 
    <td width="18%"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="verde.gif" width="19" height="18"></font></b></div>
    </td>
    <td width="11%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><%=v1%></font></td>
    <td colspan="2">&nbsp;</td>
  </tr>
  
   <tr> 
    <td width="18%"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><img src="azul.gif" width="17" height="17"></font></b></div>
    </td>
    <td width="11%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000080"><%=v2%></font></td>
    <td colspan="2">&nbsp;</td>
  </tr>
</table>
<%end if%>
</body>
	
	<%
	db_Cogest.close
	set db_Cogest = nothing
	
	db_banco.close
	set db_banco = nothing
	%>	
	
</html>
