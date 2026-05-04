<%
ct = request("selCT")
corte = request("selCorte")
strNomeTirulo = Request("pTituloRel")

ano = 2004

dim cor1(32)
dim cor2(32)
dim periodo1(32)
dim periodo2(32)
dim datas_atual(32)

set db_banco = Server.CreateObject("AdoDB.Connection")
db_banco.open Session("Conn_String_Cogest_Gravacao")
db_banco.CursorLocation = 3

set rs = db_banco.execute("SELECT * FROM GRADE_UNIDADE WHERE CTRO_CD_CENTRO_TREINAMENTO=" & ct & " AND CORT_CD_CORTE=" & corte)
set rs2 = db_banco.execute("SELECT * FROM GRADE_CENTRO_TREINAMENTO WHERE CTRO_CD_CENTRO_TREINAMENTO=" & ct & " AND CORT_CD_CORTE=" & corte)
set rs3 = db_banco.execute("SELECT * FROM GRADE_SALA WHERE CTRO_CD_CENTRO_TREINAMENTO=" & ct & " AND CORT_CD_CORTE=" & corte)
set rs4 = db_banco.execute("SELECT * FROM GRADE_DIRETORIA WHERE ORLO_CD_ORG_LOT=" & rs("ORLO_CD_ORG_LOT_DIR"))
set rs5 = db_banco.execute("SELECT * FROM GRADE_CORTE WHERE CORT_CD_CORTE=" & corte)

set rscorte = db_banco.execute("SELECT * FROM GRADE_CORTE WHERE CORT_CD_CORTE=" & corte)

i = 0

do until i = 32
	cor1(i)="white"
	cor2(i)="white"
	periodo1(i)="<font color=""white""> - </font>"
	periodo2(i)="<font color=""white""> - </font>"	
	datas_atual(i)=""
	i = i + 1
loop

i = 1

dim cor(15)

	cor(1)="#CCFFFF"
	cor(2)="#FFFFCC"
	cor(3)="#CAFFCA"
	cor(4)="#FFCC99"
	cor(5)="#9999FF"
	cor(6)="#009900"
	cor(7)="#FF9900"
	cor(8)="#BF338E"
	cor(9)="#0099FF"
	cor(10)="#FF66FF"
	cor(11)="#7AC0B6"
	cor(12)="#FFFFCC"
	cor(13)="#CCCCCC"
	cor(14)="#DFCCE3"
	cor(15)="#FFD7D7"

%>
<html>
<head>
<title>Geração de Grade De Treinamento</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">

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
				 <td width="28"></td>  
					<td width="250"></td>
			  <td width="28"></td>
			  <td width="26">&nbsp;</td>
			  <td width="159"></td>
			</tr>
		  </table>
		</td>
	  </tr>
	</table>

<table width="100%" border="0">
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
  <tr>
    <td height="20">
		<div align="center"><font face="Verdana" color="#330099" size="4"><b><%=strNomeTirulo%></b></font></div>
	</td>
  </tr>
  <tr>
    <td height="20">&nbsp;</td>
  </tr>
  <tr> 
    <td width="50%" height="57"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000066"> 
      Centro de Treinamento : <b><%=rs2.fields(3).value%></b><br>
      Corte : <b> <%=rs5.fields(1).value%> - <%=rscorte("CORT_DT_DATA_CORTE")%></b></font></td>
  </tr>
</table>
<%
'=================== INÍCIO DE TRATAMENTO DE GRADE POR MESES ==================================================

meses = 4
	
do until meses = 7

%>
<p><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000066">	
  Mês : <b><%=ucase(monthname(meses))%> / <%=ano%></b></font></p>
<table border="1" height="120" bordercolor="#E8E8E8" cellspacing="0" cellpadding="5" width="294">
  <tr> 
    <td width="129" height="38"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sala</font></div>
    </td>
    <%
	valida=1
	datas = 1
	
	do until valida=0
	
		x = isdate(datas & "/" & meses & "/" & ano)
		
		if x = true then	
			atual = formatdatetime( datas & "/" & meses & "/" & ano)
		
			dia = weekday(atual)
			if dia = 1 or dia = 7 then
				cor_dia = "#FF0000"
				if dia = 1 then
					atual = "DOMINGO"
				else
					atual = "SÁBADO"
				end if
			else
				cor_dia = "BLACK"
			end if
			
			datas_atual(datas)=atual
			
			'================ Verifica se a data atual é feriado Nacional / Regional =====================
			
			ver_feriado = right("000" & datas,2) & "/" & right("000" & meses,2)
			
			ssql="SELECT * FROM GRADE_FERIADO WHERE FERI_TX_TIPO_FERIADO='0' AND FERI_DT_DATA_FERIADO='" & ver_feriado & "'"
			
			set temp = db_banco.execute(ssql)
			
			if temp.eof=false then
				atual = temp("FERI_TX_NOME_FERIADO")
				compl="*"
				cor_dia = "#4530FF"
				datas_atual(datas)=temp("FERI_TX_NOME_FERIADO")
			
			else
			
			do until rs3.eof=true

				ssql=""
				ssql="SELECT GRADE_FERIADO_SALA.SALA_CD_SALA, "
				ssql = ssql + "GRADE_FERIADO_SALA.FERI_CD_FERIADO, GRADE_FERIADO.FERI_TX_NOME_FERIADO, GRADE_FERIADO.FERI_DT_DATA_FERIADO "
				ssql = ssql + "FROM  GRADE_FERIADO_SALA INNER JOIN GRADE_FERIADO ON "
				ssql = ssql + "GRADE_FERIADO_SALA.FERI_CD_FERIADO = GRADE_FERIADO.FERI_CD_FERIADO "
				ssql = ssql + "WHERE GRADE_FERIADO_SALA.SALA_CD_SALA = " & rs3.fields(1).value & " "
				ssql = ssql + "AND GRADE_FERIADO.FERI_DT_DATA_FERIADO = '" & ver_feriado & "'"
				
				set temp2 = db_banco.execute(ssql)
				
				if temp2.eof=false then
					atual = temp2("FERI_TX_NOME_FERIADO")
					compl="*"
					cor_dia = "#4530FF"
					datas_atual(datas)=temp2("FERI_TX_NOME_FERIADO")
				end if
				
				rs3.movenext
			
			loop
			
			rs3.movefirst

			end if
		%>
    <td width="139" height="38" title="<%=atual%><%=compl%>"> 
      <div align="center" title="<%=atual%>"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" font color="<%=cor_dia%>"><%=atual%><%=compl%></font></b></div>
    </td>
    <%
		title=""
		compl=""
		datas=datas + 1	
	else
		valida=0
	end if
	
	loop
	
	datas = datas - 1
	
	%>
  </tr>
  <%
	rs3.movefirst  
	
	do until rs3.eof=true
	
	i = 1 
	
  %>
  <tr> 
    <td width="129" rowspan="2" bgcolor="#ECFFFF"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=ucase(rs3.fields(3).value)%></b></font></div>
      <%
 
	  num_sala = rs3.fields(1).value
	  
	  data_atual = 1
	  qte_periodos = 0
      cor_atual = "white"
	  curso = "<font color=""white""> - </font>" 
	  
	  do until data_atual => datas
	  
			monta_data = data_atual & "/" & meses & "/" & ano
			
			data_montada = ano & "-" & right("000" & meses,2) & "-" & right("000" & data_atual,2)
			
			ssql="SELECT * FROM GRADE_TURMA WHERE SALA_CD_SALA = " & num_sala & " AND TURM_DT_INICIO = CONVERT(DATETIME, '" & data_montada & " 00:00:00', 102)"
			
			if corte<>0 then
				ssql=ssql+" AND CORT_CD_CORTE=" & corte
			end if
			
			set turma = db_banco.execute(ssql)
			
			if turma.eof=false then
			
				curso = turma("CURS_CD_CURSO")
					
				Select Case left(turma("CURS_CD_CURSO"),3)
				Case "SUP"
					cor_atual=cor(1)
				Case "MES"
					cor_atual=cor(2)
				Case "COM"
					cor_atual=cor(3)
				Case "EMP"
					cor_atual=cor(4)
				Case "MAN"
					cor_atual=cor(5)
				Case "POS"
					cor_atual=cor(6)
				Case "PRD"
					cor_atual=cor(7)
				Case "QUA"
					cor_atual=cor(8)
				Case "LTE"
					cor_atual=cor(9)
				Case "PLC"
					cor_atual=cor(10)
				Case "FIN"
					cor_atual=cor(11)
				Case "RHU"
					cor_atual=cor(12)
				Case "PRE"
					cor_atual=cor(13)				
				Case "BW","BWA","BWC","BWF","BWG","BWJ","BWL","BWM","BWP","BWQ","BWR","BWS","BWT","BWU"
					cor_atual=cor(15)
        		End Select
				
				qte_periodos = turma.fields(11).value
				
				pinta = 1
				
				do until pinta > qte_periodos 
				
					x = pinta mod 2
					
					if x <> 0 then
						periodo1(i) = curso
						cor1(i) = cor_atual
					else
						periodo2(i) = curso
						cor2(i) = cor_atual
						data_atual = data_atual + 1
						i = i + 1
					end if
					
					pinta = pinta + 1
					
				loop
				
			end if
			
			data_atual = data_atual + 1
			i = i + 1
			
			curso= " - "
			cor_atual = "white"

	  loop
	  %>
    </td>
    <%
	f = 1
	do until f = datas + 1
	%>
    <td width="139" height="45" bgcolor="<%=cor1(f)%>" title="<%=datas_atual(f)%>"> 
      <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=periodo1(f)%></font></div>
    </td>
    <%
	f = f + 1
	loop
	f = 1
	%>
  </tr>
  <tr> 
    <%
	do until f = datas + 1
	%>
    <td width="139" height="45" bgcolor="<%=cor2(f)%>" title="<%=datas_atual(f)%>"> 
      <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=periodo2(f)%></font></div>
    </td>
    <%
	f = f + 1	
	loop
	%>
  </tr>
  <%
  
  	i = 0

	do until i = 32
		cor1(i)="white"
		cor2(i)="white"
		periodo1(i)="<font color=""white""> - </font>"
		periodo2(i)="<font color=""white""> - </font>"	
		i = i + 1
	loop
	
	i = 1

  	rs3.movenext
	loop
  %>
</table>
	<%
	i = 1
	rs3.movefirst
	meses = meses + 1
	loop
%>

<%
'=================== TÉRMINO DE TRATAMENTO DE GRADE POR MESES ==================================================
%>

<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
