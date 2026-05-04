<!--#include file="conn_consulta.asp" -->
<html>
<%
if request("excel")=1 then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

Session.LCID = 1046
chave = Ucase(request("txtchave"))

server.ScriptTimeout=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.cursorlocation=3

set db2=server.createobject("ADODB.CONNECTION")
db2.Open "Provider=SQLOLEDB.1;server=S5200DB01\DB01;pwd=sinergiacogest;uid=usr_cogest;database=IntranetSinergia"
db2.cursorlocation=3
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Concessão de Perfil de Acesso</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
</head>

<body link="#000099" vlink="#000099" alink="#000099" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">
<form name="frm1">
<%
if request("excel")<>1 then
%>
<table width="81%" height="26" border="0">
           <tr>
                      <td width="5%">
                         <div align="right">
                                   <a href="javascript:history.go(-1)"><img src="seta_esquerda_01.jpg" width="21" height="18" border="0" alt="Voltar para a Página anterior"></a></div>
                      </td>
                      <td width="5%"><strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Voltar</font></strong></td>
                      <td width="7%">
                         <div align="right">
                                   <a href="javascript:print()"><img src="impressão.jpg" width="27" height="21" border="0" alt="Imprimir Consulta Atual">
                                   </a>
                         </div>
                      </td>
                      <td width="10%">
                      <strong><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif">Imprimir</font></strong>
                      </td>
                      <td width="3%"><p align="center"><a href="gera_consulta_perfil.asp?txtchave=<%=chave%>&excel=1" target="blank"><img border="0" src="../excel.jpg" width="23" height="21" align="right" alt="Exportar Consulta Atual para o Excel"></a></td>
                      <td width="22%"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><font color="#000066">Exportar para Excel</font></font></strong></td>
           </tr>
</table>
<table width="93%" border="0">
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<%
end if
set rs = db.execute("SELECT * FROM FUNCAO_USUARIO WHERE USMA_CD_USUARIO='" & CHAVE & "' ORDER BY FUNE_CD_FUNCAO_NEGOCIO")

reg = rs.Recordcount

on error resume next

set usuario = db.execute("SELECT USMA_TX_NOME_USUARIO AS NOME, ORME_CD_ORG_MENOR FROM USUARIO_MAPEAMENTO WHERE USMA_CD_USUARIO='" & CHAVE & "'")
set orgao = db.execute("SELECT ORME_SG_ORG_MENOR AS LOTACAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & usuario("ORME_CD_ORG_MENOR") & "'")

if err.number=0 then
%>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#000080">Chave Consultada : <b><%=chave%> - <%=usuario("nome")%> - <%=orgao("lotacao")%></b></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<hr>
<%else%>
<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#000080"><b>Usuário não encontrado</b></font></p>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<hr>
<%
end if
i = 0
do until i = reg

tf=0
tfv=0

tt=0

tm=0
tmv=0

tl=0
tp=0

resto=0

set func = db.execute("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")

tf = 1

IF FUNC("FUNE_CD_FUNCAO_NEGOCIO") = FUNC("FUNE_CD_FUNCAO_NEGOCIO_PAI") THEN
	FUNCAO = FUNC("FUNE_CD_FUNCAO_NEGOCIO")
	FUNCAO_PAI = FUNC("FUNE_CD_FUNCAO_NEGOCIO")	
ELSE
	FUNCAO = FUNC("FUNE_CD_FUNCAO_NEGOCIO")
	FUNCAO_PAI = FUNC("FUNE_CD_FUNCAO_NEGOCIO_PAI")
END IF

if rs("FUUS_IN_VALIDADO")="S" then
	valida="VALIDADO"
	tfv = 1
else
	valida="NÃO VALIDADO"
end if
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="76%" id="AutoNumber2" height="28">
           <tr>
                      <td width="1%" height="28">&nbsp;</td>
                      <td width="20%" bgcolor="#D7D5CC" height="28"><font face="Verdana" size="2" color="#000080">Função de Negócio :</font></td>
                      <td width="124%" bgcolor="#D7D5CC" height="28"><font face="Verdana" size="2" color="#000080">&nbsp;<b><%=FUNCAO%> - <%=FUNC("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%> - <font color="red"><%=valida%></font></b></font></td>
           </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="78%" id="AutoNumber1">
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="28%" bgcolor="#D7D5CC"><b><font face="Verdana" size="2" color="#000080">Treinamento</font></b></td>
                      <td width="31%" bgcolor="#D7D5CC"><b><font face="Verdana" size="2" color="#000080">Situação</font></b></td>
                      <td width="77%" bgcolor="#D7D5CC"><b><font size="2" color="#000080"><font face="Verdana">Inscrito para</font></font></b></td>
           </tr>
           <%
           set curso = db.execute("SELECT * FROM CURSO_FUNCAO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & FUNCAO & "'")
           
           do until curso.eof=true           
           %>
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="28%"><font face="Verdana" size="2" color="#000080"><%=curso("CURS_CD_CURSO")%></font></td>
			<%
			ap = 0
						
			set temp = db2.execute("SELECT * FROM TABINSCRITOS WHERE CHAVE='" & CHAVE & "' AND COD_DISCIPLINA='" & curso("CURS_CD_CURSO") & "' AND APROV='AP'")

			if temp.eof=false then
				ap = ap + 1
			end if
			
			set temp = db.execute("SELECT * FROM USUARIO_APROVADO WHERE USAP_CD_USUARIO='" & CHAVE & "' AND CURS_CD_CURSO='" & curso("CURS_CD_CURSO") & "' AND (USAP_TX_APROVEITAMENTO='AP' OR USAP_TX_APROVEITAMENTO='LM')")
			
			if temp.eof=false then
				ap = ap + 1
			end if

			if ap > 0 then
				cor="#000080"
				situ="OK"
				tt = tt + 1
			else
				
				cor="red"
				situ="NÃO EXECUTADO"
				
				set temp3 = db2.execute("SELECT * FROM TABINSCRITOS WHERE CHAVE='" & CHAVE & "' AND COD_DISCIPLINA='" & curso("CURS_CD_CURSO") & "'")
				on error resume next
				set temp4 = db2.execute("SELECT * FROM TABTURMA WHERE EVENTO=" & temp3("EVENTO") & " AND ANO=" & temp3("ANO") & " AND PROJETO=" & temp3("PROJETO"))
				
				if err.number=0 and temp4("DATA_FIM") > date then
					inscrito = temp4("DATA_INICIO") & "-" & temp4("DATA_FIM")
				else
					inscrito = ""
					err.clear
				end if
				
			end if			
			%>
			<td width="31%"><font face="Verdana" size="2" color="<%=cor%>"><b><%=situ%></b></font></td>
			<td width="77%"><font face="Verdana" size="2" color="#000080"><b><%=inscrito%></b></font></td>
           </tr>
           <%
           inscrito = ""
           curso.movenext
           loop
           
           if tt > 0 then
	           resto = (tt mod 2)
           else
    	       resto=1
    	   end if
           %>
</table>
&nbsp;<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="61%" id="AutoNumber1">
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="46%" bgcolor="#D7D5CC"><b><font face="Verdana" size="2" color="#000080">Perfil</font></b></td>
                      <td width="45%" bgcolor="#D7D5CC"><b><font face="Verdana" size="2" color="#000080">Situação</font></b></td>
           </tr>
		   <%	
	       if left(funcao, 2)="HR" then           
		    	set perfil = db.execute("SELECT * FROM FUNCAO_USUARIO_PERFIL_VISAO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & FUNCAO & "' AND USMA_CD_USUARIO='" & CHAVE & "'")
           else
	           set perfil = db.execute("SELECT * FROM FUNCAO_USUARIO_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & FUNCAO & "' AND USMA_CD_USUARIO='" & CHAVE & "'")           
           end if

            do until perfil.eof=true    

			if perfil("FUUP_IN_VALIDADO")="S" then
				validap="VALIDADO"
				tmv = tmv + 1
			else
				validap="NÃO VALIDADO"
			end if
           
			set nomes = db.execute("SELECT * FROM MICRO_PERFIL_R3 WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & perfil("MCPR_NR_SEQ_MACRO_PERFIL") & " AND MIPE_NR_SEQ_MICRO_PERFIL=" & perfil("MIPE_NR_SEQ_MICRO_PERFIL"))
           %>
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="46%"><font face="Verdana" size="2" color="#000080"><%=nomes("MIPE_TX_NOME_TECNICO")%></font></td>
			          <td width="45%"><font face="Verdana" size="2" color="#000080"><b><%=validap%></b></font></td>
           </tr>
           <%
           tm = tm + 1
           perfil.movenext
           loop
           %>
</table>
<BR><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="70%" id="AutoNumber1">
           <tr>
                      <td width="8%">&nbsp;</td>
                      <td width="59%" bgcolor="#D7D5CC"><b><font face="Verdana" size="2" color="#000080">Lote</font></b></td>
                      <td width="33%" bgcolor="#FFFFFF">&nbsp;</td>
           </tr>
           <%
	       if left(funcao, 2)="HR" then
		       set lote = db.execute("SELECT * FROM GOLI_FUNCAO_USUARIO_COM_PERFIL_RH WHERE FUNE_CD_FUNCAO_NEGOCIO='" & FUNCAO & "' AND USMA_CD_USUARIO='" & CHAVE & "'")           
		   else
		       set lote = db.execute("SELECT * FROM GOLI_FUNCAO_USUARIO_COM_PERFIL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & FUNCAO & "' AND USMA_CD_USUARIO='" & CHAVE & "'")           
		   end if
          
           if lote.eof=false then
           tl = tl + 1
           set temp = db.execute("SELECT * FROM GOLI_LOTE WHERE LOTE_NR_SEQ_LOTE=" & LOTE("LOTE_NR_SEQ_LOTE"))
           %>
           <tr>
                      <td width="8%">&nbsp;</td>
                      <td width="59%"><font face="Verdana" size="2" color="#000080"><%=LOTE("LOTE_NR_SEQ_LOTE")%>-<%=TEMP("LOTE_TX_DESCRICAO")%></font></td>
			          <td width="33%" bgcolor="#FFFFFF">&nbsp;</td>
           </tr>
           <%
           end if
           %>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="61%" id="AutoNumber1">
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="73%" bgcolor="#D7D5CC"><b><font face="Verdana" size="2" color="#000080">Perfil no R/3</font></b></td>
                      <td width="18%" bgcolor="#D7D5CC">&nbsp;</td>
           </tr>
<%
if left(funcao, 2)="HR" then
		ssql=""
		ssql="SELECT DISTINCT dbo.USMA_MICRO_R3_VISAO_R3.USMA_CD_USUARIO, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, "
		ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.USMA_MICRO_R3_VISAO_R3.MCPR_NR_SEQ_MACRO_PERFIL, dbo.USMA_MICRO_R3_VISAO_R3.MIPE_NR_SEQ_MICRO_PERFIL  "
		ssql=ssql+"FROM dbo.USMA_MICRO_R3_VISAO_R3 "
		ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
		ssql=ssql+"dbo.USMA_MICRO_R3_VISAO_R3.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
		ssql=ssql+"INNER JOIN dbo.ORGAO_MENOR ON "
		ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR "
		ssql=ssql+"INNER JOIN dbo.MACRO_PERFIL ON "
		ssql=ssql+"dbo.USMA_MICRO_R3_VISAO_R3.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL WHERE "
		ssql=ssql+"dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = '" & FUNCAO_PAI & "' " 
		ssql=ssql+"AND dbo.USMA_MICRO_R3_VISAO_R3.USMA_CD_USUARIO = '" & CHAVE & "' "
		ssql=ssql+" order by dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USMA_MICRO_R3_VISAO_R3.MCPR_NR_SEQ_MACRO_PERFIL "
else
		ssql=""
		ssql="SELECT DISTINCT dbo.USUARIO_PERFIL.USPE_CD_USUARIO, dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, "
		ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR, dbo.USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL, dbo.USUARIO_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL  "
		ssql=ssql+"FROM dbo.USUARIO_PERFIL "
		ssql=ssql+"INNER JOIN dbo.USUARIO_MAPEAMENTO ON "
		ssql=ssql+"dbo.USUARIO_PERFIL.USPE_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO "
		ssql=ssql+"INNER JOIN dbo.ORGAO_MENOR ON "
		ssql=ssql+"dbo.USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR "
		ssql=ssql+"INNER JOIN dbo.MACRO_PERFIL ON "
		ssql=ssql+"dbo.USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL WHERE "
		ssql=ssql+"dbo.MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = '" & FUNCAO_PAI & "' " 
		ssql=ssql+"AND dbo.USUARIO_PERFIL.USPE_CD_USUARIO = '" & CHAVE & "' "
		ssql=ssql+" order by dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
end if

set perfilr3 = db.execute(ssql)

do until perfilr3.eof=true

set nome = db.execute("SELECT * FROM MICRO_PERFIL_R3 WHERE MCPR_NR_SEQ_MACRO_PERFIL=" & perfilr3("MCPR_NR_SEQ_MACRO_PERFIL") & " AND MIPE_NR_SEQ_MICRO_PERFIL=" & perfilr3("MIPE_NR_SEQ_MICRO_PERFIL"))
%>
           <tr>
                      <td width="9%">&nbsp;</td>
                      <td width="91%" colspan="2"><font face="Verdana" size="2" color="#000080"><%=nome("MIPE_TX_NOME_TECNICO")%></font></td>
           </tr>
<%
tp = tp + 1
perfilr3.movenext
loop

if tf>0 and tfv>0 and tm=0 and tmv=0 and tm=tmv and resto=0 and tl=0 and tp=0 then
	mensagem = "Service-Desk SAP : Repassar o chamado para o Sinergia"
end if

if tf>0 and tfv>0 and tm>0 and tmv>0 and tm=tmv and resto<>0 and tl=0 and tp=0 then
	mensagem = "Service-Desk SAP : Repassar o chamado para o Sinergia"
end if

if tf>0 and tfv>0 and tm>0 and tmv>0 and tm=tmv and resto=0 and tl=0 and tp=0 then
	mensagem = "Service-Desk SAP : Repassar o chamado para o Sinergia"
end if

if tf>0 and tfv>0 and tm>0 and tmv>0 and tm=tmv and resto=0 and tl=0 and tp=0 then
	mensagem = "Service-Desk SAP : Repassar o chamado para o Sinergia"
end if

if tf>0 and tfv>0 and tm>0 and tmv>0 and tm=tmv and resto=0 and tl>0 and tp=0 then
	mensagem = "Service-Desk SAP : Repassar o chamado para o Sinergia"
end if

%>          
</table>
&nbsp;<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="86%" id="AutoNumber3">
           <tr>
                      <td width="100%"><b><font face="Verdana" size="2" color="#FF0000"><%'=mensagem%></font></b></td>
           </tr>
</table>
<hr><b><br>
<%
I = I + 1
rs.movenext
loop
if reg=0 then
mensagem = "Service-Desk SAP : Orientar usuário a entra em contato com o Coordenador Local de Implantação"
%>
<font color="#800000"><b>Nenhum Registro Encontrado para a Seleção</b></font>
</table>
&nbsp;</b><b><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="86%" id="AutoNumber3" height="23">
           <tr>
                      <td width="100%" height="23"><b><font face="Verdana" size="2" color="#FF0000"><%'=mensagem%></font></b></td>
           </tr>
</table>
<%
end if
%>
</font></b>
</form>
</body>

</html>