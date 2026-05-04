<%
server.ScriptTimeOut = 99999999

set db = Server.CreateObject("AdoDB.Connection")
db.open "Provider = Microsoft.Jet.Oledb.4.0;Data Source = " & Server.Mappath("db.mdb")
db.CursorLocation = 3

set db2=server.createobject("AdoDB.Connection")
db2.Open "Provider=SQLOLEDB.1;server=S6000db21;pwd=cogest00;uid=cogest;database=cogest"
db2.cursorlocation=3

set db3=server.createobject("AdoDB.Connection")
db3.Open "Provider=SQLOLEDB.1;server=S5200DB01\DB01;pwd=sinergiacogest;uid=usr_cogest;database=IntranetSinergia"
db3.cursorlocation=3

f_gerente = "MM.56"
f_fiscal = "MM.61"

v = 0
a = 0
m = 0

if request("CBI")<>"X" THEN
	set rs = db.execute("SELECT DISTINCT CONTRATO FROM CONTRATO WHERE CONTRATO LIKE '" & request("CBI") & "%' ORDER BY CONTRATO")
else
	set rs = db.execute("SELECT DISTINCT CONTRATO FROM CONTRATO ORDER BY CONTRATO")
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relatório de Situação de Contratos</title>
</head>

<body vlink="#0000FF" alink="#0000FF">

<hr>
<%
i = 0
reg = rs.RecordCount
Do until i = reg

tg = 0
tf = 0

set v1 = db.execute("SELECT * FROM GERENTE WHERE CONTRATO='" & trim(rs("CONTRATO")) & "' ORDER BY CHAVE")
set v2 = db.execute("SELECT * FROM FISCAL WHERE CONTRATO='" & trim(rs("CONTRATO")) & "' ORDER BY CHAVE")

if v1.eof=false or v2.eof=false then
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="99%" id="AutoNumber1" height="24">
           <tr>
                      <td width="13%" height="24" bgcolor="#D7D5CC"><p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana">Contrato</font></b></td>
                      <td width="17%" height="24"><p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" size="2"><%=rs("CONTRATO")%></font></b></td>
                      
                      <%
                      set r3 = db.execute("SELECT DISTINCT R3 FROM CONTRATO WHERE CONTRATO='" & rs("CONTRATO") & "'")
                      
                      contr = ""
                      
                      do until r3.eof=true
                      	contr = contr & r3("R3") & ", "
                      	r3.movenext
                      loop
                      
                      %>
                      
                      <td width="125%" height="24"><font face="Verdana" size="1"><b>Número(s) R/3: </b><%=left(contr, len(contr)-2)%></font></td>
           </tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="85%" id="AutoNumber2">
           <tr>
                      <td width="20%" bgcolor="#C4D0C6"><p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" size="2">Gerentes</font></b></td>
                      <td width="16%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="17%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="23%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="27%">&nbsp;</td>
                      <td width="74%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
           </tr>
           <tr>
                      <td width="20%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="16%" bgcolor="#FFFFD9"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Chave</font></b></td>
                      <td width="17%" bgcolor="#FFFFD9"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Treinamento</font></b></td>
                      <td width="23%" bgcolor="#FFFFD9" align="left"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Mapeado em Perfil</font></b></td>
                      <td width="27%" bgcolor="#FFFFD9" align="left"><b><font face="Verdana" size="2">Validado em Perfil</font></b></td>
                      <td width="74%" bgcolor="#FFFFD9"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Situação</font></b></td>
           </tr>
           <%
           set gerente = db.execute("SELECT * FROM GERENTE WHERE CONTRATO='" & trim(rs("CONTRATO")) & "' ORDER BY CHAVE")
           g = 0
           regg = gerente.RecordCount
           
           do until g = regg
           
        	treinamento_g=""
			perfil_g=""
			validado_g=""
			situacao_g=""
           
           set g1 = db2.execute("SELECT * FROM FUNCAO_USUARIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & f_gerente & "' AND USMA_CD_USUARIO='" & TRIM(gerente("CHAVE")) & "'") 
           
           if g1.eof=false then
           
           		ap = 0
           		
           		set curso = db2.execute("SELECT DISTINCT CURS_CD_CURSO FROM CURSO_FUNCAO WHERE FUNE_CD_FUNCAO_NEGOCIO = '" & f_gerente & "'")
           		
           		do until curso.eof=true
           			set temp = db3.execute("SELECT * FROM TABINSCRITOS WHERE CHAVE='" & TRIM(gerente("CHAVE")) & "' AND COD_DISCIPLINA='" & curso("CURS_CD_CURSO") & "' AND APROV='AP'")
           			if temp.eof=false then
           				ap = ap + 1
           			else
						set temp2 = db2.execute("SELECT * FROM USUARIO_APROVADO WHERE USAP_CD_USUARIO='" & TRIM(gerente("CHAVE")) & "' AND CURS_CD_CURSO='" & curso("CURS_CD_CURSO") & "' AND (USAP_TX_APROVEITAMENTO='AP' OR USAP_TX_APROVEITAMENTO='LM')")
						if temp2.eof=false then
							ap = ap + 1
						end if
           			end if
	           		curso.movenext
           		loop
           		
           		if ap = curso.recordcount then
           			treinamento_g = "OK"
           		else
           			treinamento_g = "NÃO OK"
           		end if           		
           
				set perfilg1 = db2.execute("SELECT * FROM FUNCAO_USUARIO_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & f_gerente & "' AND USMA_CD_USUARIO='" & trim(gerente("chave")) & "'")
								
				if perfilg1.eof=false then
           			perfil_g = "OK"				
				else
           			perfil_g = "NÃO OK"				
				end if
				
				if perfil_g="OK" then
					set perfilg1 = db2.execute("SELECT * FROM FUNCAO_USUARIO_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & f_gerente & "' AND USMA_CD_USUARIO='" & trim(gerente("chave")) & "' AND FUUP_IN_VALIDADO='N'")
					if perfilg1.eof=true then
	           			validado_g = "OK"				
					else
	           			validado_g = "NÃO OK"				
					end if
				else
					validado_g="NÃO OK"				
				end if
				
				if treinamento_g = "OK" and perfil_g="OK" and validado_g="OK" then
					situacao_g = "OK"
					tg = tg + 1
				else
					situacao_g = "NÃO OK"					
				end if
				%>
           <tr>
                      <td width="20%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="16%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=gerente("chave")%></font></td>
                      <td width="17%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=treinamento_g%></font></td>
                      <td width="23%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=perfil_g%></font></td>
                      <td width="16%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=validado_g%></font></td>
                      <td width="74%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=situacao_g%></font></td>
           </tr>
           <%
           else
           %>
           <tr>
                      <td width="20%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"></font></td>
                      <td width="16%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=gerente("chave")%></font></td>
                      <td width="17%"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2" color="#FF0000">NÃO MAPEADO</font></b></td>
           </tr>
           <%
           end if
           g = g + 1
           gerente.movenext
           loop
           %>
           
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="85%" id="AutoNumber2">
           <tr>
                      <td width="20%" bgcolor="#C4D0C6"><p style="margin-top: 0; margin-bottom: 0" align="center"><b><font face="Verdana" size="2">Fiscais</font></b></td>
                      <td width="16%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="17%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="23%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="54%">&nbsp;</td>
                      <td width="47%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
           </tr>
           <tr>
                      <td width="20%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="16%" bgcolor="#FFFFD9"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Chave</font></b></td>
                      <td width="17%" bgcolor="#FFFFD9"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Treinamento</font></b></td>
                      <td width="23%" bgcolor="#FFFFD9" align="left"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Mapeado em Perfil</font></b></td>
                      <td width="54%" bgcolor="#FFFFD9" align="left"><b><font face="Verdana" size="2">Validado em Perfil</font></b></td>
                      <td width="47%" bgcolor="#FFFFD9"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Situação</font></b></td>
           </tr>
           <%
           set fiscal = db.execute("SELECT * FROM FISCAL WHERE CONTRATO='" & trim(rs("CONTRATO")) & "' ORDER BY CHAVE")
           f = 0
           regf = fiscal.RecordCount
           
           do until f = regf
           
           	treinamento_f=""
			perfil_f=""
			validado_f=""
			situacao_f=""
			
           set f1 = db2.execute("SELECT * FROM FUNCAO_USUARIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & f_fiscal & "' AND USMA_CD_USUARIO='" & TRIM(FISCAL("CHAVE")) & "'") 
           
           if f1.eof=false then

           		ap = 0
           		
           		set cursof = db2.execute("SELECT DISTINCT CURS_CD_CURSO FROM CURSO_FUNCAO WHERE FUNE_CD_FUNCAO_NEGOCIO ='" & f_fiscal & "'")
           		
           		do until cursof.eof=true
           			set temp = db3.execute("SELECT * FROM TABINSCRITOS WHERE CHAVE='" & TRIM(fiscal("CHAVE")) & "' AND COD_DISCIPLINA='" & cursof("CURS_CD_CURSO") & "' AND APROV='AP'")
           			if temp.eof=false then
           				ap = ap + 1
           			else
						set temp2 = db2.execute("SELECT * FROM USUARIO_APROVADO WHERE USAP_CD_USUARIO='" & TRIM(fiscal("CHAVE")) & "' AND CURS_CD_CURSO='" & cursof("CURS_CD_CURSO") & "' AND (USAP_TX_APROVEITAMENTO='AP' OR USAP_TX_APROVEITAMENTO='LM')")
						if temp2.eof=false then
							ap = ap + 1
						end if
           			end if
	           		cursof.movenext
           		loop
           		
           		if ap = cursof.recordcount then
           			treinamento_f = "OK"
           		else
           			treinamento_f = "NÃO OK"
           		end if           		
           
				set perfilf1 = db2.execute("SELECT * FROM FUNCAO_USUARIO_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & f_fiscal & "' AND USMA_CD_USUARIO='" & trim(fiscal("chave")) & "'")
				
				if perfilf1.eof=false then
           			perfil_f = "OK"				
				else
           			perfil_f = "NÃO OK"				
				end if
				
				if perfil_f = "OK" then
					set perfilf1 = db2.execute("SELECT * FROM FUNCAO_USUARIO_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & f_fiscal & "' AND USMA_CD_USUARIO='" & trim(fiscal("chave")) & "' AND FUUP_IN_VALIDADO='N'")
					if perfilf1.eof=true then
	           			validado_f = "OK"				
					else
	           			validado_f = "NÃO OK"				
					end if
				else
           			validado_f = "NÃO OK"								
				end if
				
				if treinamento_f = "OK" and perfil_f="OK" and validado_f="OK" then
					situacao_f = "OK"
					tf = tf + 1
				else
					situacao_f = "NÃO OK"					
				end if
           %>
           <tr>
                      <td width="20%"><p style="margin-top: 0; margin-bottom: 0">&nbsp;</td>
                      <td width="16%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=fiscal("chave")%></font></td>
                      <td width="17%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=treinamento_f%></font></td>
                      <td width="23%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=perfil_f%></font></td>
                      <td width="21%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=validado_f%></font></td>
                      <td width="47%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=situacao_f%></font></td>
           </tr>
           <%
           else
           %>
           <tr>
                      <td width="20%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"></font></td>
                      <td width="16%"><p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=fiscal("chave")%></font></td>
                      <td width="17%"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2" color="#FF0000">NÃO MAPEADO</font></b></td>
           </tr>
           <%
           end if
           f = f + 1
           fiscal.movenext
           loop
           
           imagem=""
           
           	if regg = tg and regf = tf and tg>0 and tf>0 then
           		imagem="azul.jpg"
           		v = v + 1
           	else
           		if tg>0 and tf>0 then
           			imagem="amarelo.jpg"
           			a = a + 1
           		else
           			imagem="vermelho.jpg"           		
           			m = m + 1
           		end if
           	end if
           
           %>
</table>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="27%" id="AutoNumber3" height="25">
           <tr>
                      <td width="70%" height="25" bgcolor="#FFFFD9"><b><font face="Verdana" size="2">Situação do Contrato </font></b></td>
                      <td width="30%" height="25"><p align="center"><img border="0" src="Imagens/<%=imagem%>"></td>
           </tr>
</table>

<hr>
<%
user = user + 1
end if
i = i + 1
rs.movenext
loop
if user>0 then
%> 
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="47%" id="AutoNumber4" height="81">
           <tr>
                      <td width="50%" height="27"><b><font face="Verdana" size="1">Todos os Gerentes e Fiscais</font></b></td>
                      <td width="8%" height="27"><font face="Verdana" size="2"><img border="0" src="Imagens/azul.jpg"></font></td>
                      <td width="121%" height="27"><font face="Verdana" size="2"><%=v%></font></td>
           </tr>
           <tr>
                      <td width="50%" height="25"><b><font face="Verdana" size="1">Pelo menos um Gerente e Fiscal</font></b></td>
                      <td width="8%" height="25"><font face="Verdana" size="2"><img border="0" src="Imagens/amarelo.jpg"></font></td>
                      <td width="121%" height="25"><font face="Verdana" size="2"><%=a%></font></td>
           </tr>
           <tr>
                      <td width="50%" height="29"><b><font face="Verdana" size="1">Nenhum Gerente ou Fiscal</font></b></td>
                      <td width="8%" height="29"><font face="Verdana" size="2"><img border="0" src="Imagens/vermelho.JPG"></font></td>
                      <td width="121%" height="29"><font face="Verdana" size="2"><%=m%></font></td>
           </tr>
</table>
<%
else
%>
<b><font color="#800000">Nenhum Gerente / Fiscal Encontrado para esta CBI
<hr>
<%
end if
%> </font></b>
</body>

</html>