<%@ Language=VBScript %>
<%
if request("opt") = 1 then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if
%>
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

id=""

id=request("id")
mega=request("selMegaProcesso")
proc=request("selProcesso")
subproc=request("selSubProcesso")
onda=request("selOnda")
str_assunto=request("selAssunto")
str_escopo=request("selEscopo")

if str_escopo=0 then
	val_escopo="FORA DO ESCOPO"
else
	if str_escopo=1 then
		val_escopo="DENTRO DO ESCOPO"
	end if
end if

if mega<>0 or proc<>0 or subproc<>0 or onda<>0 then

	if mega<>0 then
		compl1=" MEPR_CD_MEGA_PROCESSO=" & mega
	end if

	if proc<>0 then
	if len(compl1)=0 then
		compl1=" PROC_CD_PROCESSO=" & proc
	else
		compl1=compl1+" AND PROC_CD_PROCESSO=" & proc
	end if
	end if

	if subproc<>0 then
		if len(compl1)=0 then
			compl1=" SUPR_CD_SUB_PROCESSO=" & subproc
		else
			compl1=compl1+" AND SUPR_CD_SUB_PROCESSO=" & subproc
		end if
	end if
end if

if onda <> 0 then
	if len(compl1)=0 then
		compl1=" ONDA_CD_ONDA =" & onda
	else
		compl1=compl1+" AND ONDA_CD_ONDA=" & onda
	end if
else
	if len(compl1)=0 then
		compl1=" ONDA_CD_ONDA <>4"
		onda=0
	else
		compl1=compl1+" AND ONDA_CD_ONDA<>4"
		onda=0
	end if
end if

if str_Assunto<>0 then
	if len(compl1)=0 then
		compl1= " SUMO_NR_CD_SEQUENCIA =" & str_Assunto 
	else
		compl1=compl1+ " AND SUMO_NR_CD_SEQUENCIA=" & str_Assunto
	end if
end if

if str_Escopo<>2 then
	if len(compl1)=0 then
		compl1= " AND CENA_TX_SITUACAO_VALIDACAO =" & str_Escopo
	else
		compl1=compl1+ " AND CENA_TX_SITUACAO_VALIDACAO=" & str_Escopo
	end if
end if


IF ID<>"0" THEN
	compl1="CENA_CD_CENARIO='"& ID & "'"
END IF

if len(trim(compl1))>0 then
	compl1="WHERE " & compl1
end if

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO " & compl1

'response.write ssql

set rs=db.execute(ssql)

%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="">
<% if request("opt") <> 1 then %>
  <table width="773" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td height="66" colspan="2">&nbsp;</td>
      <td height="66" colspan="2">&nbsp;</td>
      <td valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="111">&nbsp; </td>
      <td height="20" width="30">&nbsp;</td>
      <td height="20" width="213"><a href="gera_rel_geral.asp?selMegaProcesso=<%=mega%>&amp;selProcesso=<%=proc%>&amp;selSubProcesso=<%=subproc%>&amp;ID=<%=ID%>&selOnda=<%=onda%>&opt=1&selEscopo=<%=str_escopo%>" target="blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
      <td colspan="2" height="20">&nbsp; </td>
      <td height="20" width="334">&nbsp;</td>
    </tr>
  </table>
  <% end if %>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
  <p style="margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="3">&nbsp;
  Relatório
  Geral de Cenários</font> </p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
        <%
        if rs.eof=false then
        SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & rs("CENA_CD_CENARIO") & "' ORDER BY CENA_NR_SEQUENCIA_TRANS")
        if rs1.eof=true then
        %>

  &nbsp;&nbsp;&nbsp;<font color="#800000" face="Verdana" size="2"><b> Não existem Transações para o cenário selecionado
  </b></font>

 		<%
 		end if       
        DO UNTIL RS.EOF=TRUE
        
        SET tem=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & rs("CENA_CD_CENARIO") & "' ORDER BY CENA_NR_SEQUENCIA_TRANS")

        if tem.eof=false then
        
        set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& RS("CENA_CD_CENARIO")& "'")
        %>
        
  <table border="0" width="94%">
    <%
          SET DADOS=DB.EXECUTE("SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO, " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO, " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO FROM (" & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN " & Session("PREFIXO") & "PROCESSO ON " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO) INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON (" & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO) AND (" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO)WHERE (((" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO)=" & atual("MEPR_CD_MEGA_PROCESSO")& ") AND ((" & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO)=" & atual("PROC_CD_PROCESSO")& ") AND ((" & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO)=" & atual("SUPR_CD_SUB_PROCESSO")& "))")
          
            If rs("CENA_TX_SITUACAO") = "DF" Then
			      ls_Situacao_Cenario = "DEFINIDO"
			   elseIf rs("CENA_TX_SITUACAO") = "EE" Then
			      ls_Situacao_Cenario = "EM ELABORAÇÃO"
		      elseIf rs("CENA_TX_SITUACAO") = "DS" Then
				      ls_Situacao_Cenario = "DESENHADO"
			   elseIf rs("CENA_TX_SITUACAO") = "PT" Then
				      ls_Situacao_Cenario = "PRONTO PARA TESTE"
				elseIf rs("CENA_TX_SITUACAO") = "TD" Then
				      ls_Situacao_Cenario = "TESTADO NO PED"
				elseIf rs("CENA_TX_SITUACAO") = "TQ" Then
				      ls_Situacao_Cenario = "TESTADO NO PEQ"
			   end if
          %>
    <tr> 
      <td width="17%"><font face="Verdana" color="#330099" size="1">Mega-Processo</font></td>
      <td width="32%"><font face="Verdana" color="#330099" size="1"><%=UCASE(DADOS("MEPR_TX_DESC_MEGA_PROCESSO"))%></font></td>
      <%SET ONDA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA ="& rs("ONDA_CD_ONDA"))%>
      <td width="51%"><font face="Verdana" size="2" color="#330099"><b>Onda - 
        </b></font><b><font size="2"><font face="Verdana" color="#330099"><%=ONDA("ONDA_TX_DESC_ONDA")%></font></font></b></td>
    </tr>
    <tr> 
      <td width="17%"><font face="Verdana" color="#330099" size="1">Assunto
        :&nbsp;</font></td>
        <%
        val_assunto=""      
        on error resume next
        set temp=db.execute("SELECT * FROM SUB_MODULO WHERE SUMO_NR_CD_SEQUENCIA=" & RS("SUMO_NR_CD_SEQUENCIA"))
        if temp.eof=false then
	        val_assunto=temp("SUMO_TX_DESC_SUB_MODULO")     
	     end if
	     err.clear
        %>
      <td width="32%"><font face="Verdana" color="#330099" size="1"><%=val_assunto%></font></td>
      <td width="42%"><font face="Verdana" color="#330099" size="1">Status - <%=ls_Situacao_Cenario%></font></td>
    </tr>
    <tr> 
      <td width="17%"><font face="Verdana" color="#330099" size="1">Processo</font></td>
      <td width="41%"><font face="Verdana" color="#330099" size="1"><%=UCASE(DADOS("PROC_TX_DESC_PROCESSO"))%></font></td>
      <%SET CLASSE=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CLASSE_CENARIO WHERE CLCE_CD_NR_CLASSE_CENARIO="& rs("CLCE_CD_NR_CLASSE_CENARIO"))%>
      <td width="51%"><font face="Verdana" color="#330099" size="1">Classe do 
        Cenário -<%=CLASSE("CLCE_TX_DESC_CLASSE_CENARIO")%></font></td>
    </tr>
    <tr> 
      <td width="17%"><font face="Verdana" color="#330099" size="1">Sub-Processo</font></td>
      <td width="32%"><font face="Verdana" color="#330099" size="1"><%=UCASE(DADOS("SUPR_TX_DESC_SUB_PROCESSO"))%></font></td>
      <td width="51%"><font face="Verdana" color="#330099" size="1">Data Prevista 
        T&eacute;rmino : <%=rs("CENA_DT_PREV_TERMINO")%></font></td>
    </tr>
    <tr>
      <td width="17%"><font face="Verdana" color="#330099" size="1">Responsável 
        : </font></td>
      <td width="32%"><font face="Verdana" color="#330099" size="1"><%=rs("CENA_TX_RESPONSAVEL")%></font></td>
	  <% if rs("CENA_TX_SITUACAO_VALIDACAO")=0 then
	val_escopo="FORA DO ESCOPO"
else
	if rs("CENA_TX_SITUACAO_VALIDACAO")=1 then
		val_escopo="DENTRO DO ESCOPO"
	end if
end if %>
      <td width="51%"><font face="Verdana" color="#330099" size="1">Escopo : <%=val_escopo%></font></td>
    </tr>
    <tr>
      <td width="17%"><font face="Verdana" color="#330099" size="1">Empresa
        :&nbsp;</font></td>
      <td width="32%"><font face="Verdana" color="#330099" size="1"><%=rs("CENA_TX_EMPRESA_RELAC")%></font></td>
      <td width="51%"></td>
    </tr>
  </table>
        <table border="0" width="82%" height="75">
          <tr>
            
      <td width="28%" height="24"></td>
            <td width="72%" height="24"></td>
          </tr>
          <tr>
            
      <td width="28%" height="18"><font size="2" face="Verdana" color="#330099">Cenário</font></td>
            <td width="72%" height="18"><font size="2" face="Verdana" color="#330099"><b><%=rs("CENA_CD_CENARIO")%></b></font></td>
          </tr>
          <tr>
            <td width="28%" height="21"></td>
            <td width="72%" height="21"><b><font size="2" face="Verdana" color="#330099"><%=ATUAL("CENA_TX_TITULO_CENARIO")%></font></b></td>
          </tr>
        </table>
        
  <table border="0" width="763" style="padding: 0" cellspacing="0" bordercolordark="#000000" bordercolor="#000000" height="95">
    <tr>
            
      <td width="27" bgcolor="#330099" height="34" align="center"><b><font size="2" face="Verdana" color="#FFFFFF">Seq</font></b></td>
            
      <td width="39" bgcolor="#330099" height="34" align="center"><b><font size="2" face="Verdana" color="#FFFFFF">Mega</font></b></td>
            
      <td width="370" bgcolor="#330099" height="34"><b><font size="2" face="Verdana" color="#FFFFFF">Desc 
        Transação</font></b></td>
            
      <td width="134" bgcolor="#330099" height="34"> 
        <p align="center"><b><font size="2" face="Verdana" color="#FFFFFF">Cód
              Transação</font></b></td>
            
      <td width="102" bgcolor="#330099" height="34" align="center"><b><font size="2" face="Verdana" color="#FFFFFF">Oper 
        Especial</font></b></td>
            
      <td width="132" bgcolor="#330099" height="34" align="center"> 
        <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#FFFFFF" size="1">Cenário/</font></b></p>
              <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#FFFFFF" size="1">Desenvolvimento&nbsp;</font></b></p>
            </td>
            
      <td width="34" bgcolor="#330099" height="34" align="center"><b><font size="2" face="Verdana" color="#FFFFFF">Bpp</font></b></td>
          </tr>
          <%
          SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & rs("CENA_CD_CENARIO") & "' ORDER BY CENA_NR_SEQUENCIA_TRANS")
          'response.write rs1.eof
          if rs1.eof=true then
          %>
          
      <td width="27" height="36"> 
        <p style="margin-top: 0; margin-bottom: 0"><b><font size="2" color="#800000">&nbsp;</font></b> </p>
		  <p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
			<%
			end if
          DO UNTIL RS1.EOF=TRUE
          %>
          <tr>
            <%
            VALOR=""
            conecta=""
            IF NOT ISNULL(rs1("MEPR_CD_MEGA_PROCESSO"))THEN
            SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs1("MEPR_CD_MEGA_PROCESSO"))
            VALOR=TEMP("MEPR_TX_ABREVIA")
            END IF
            %>
            <%
            VALOR2=""
            IF NOT ISNULL(RS1("OPES_CD_OPERACAO_ESP"))THEN
            SET TEMP2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "OPERACOES_ESPEC WHERE OPES_CD_OPERACAO_ESP=" & RS1("OPES_CD_OPERACAO_ESP"))
            VALOR2=TEMP2("OPES_TX_DESC_OPERACAO_ESP")
            END IF
            %>
            
      <td width="27" bgcolor="#FFFFFF" bordercolordark="#000000" height="15" align="center"><font face="Verdana" size="1" color="#330099"><%=RS1("CENA_NR_SEQUENCIA_TRANS")%></font></td>
            
      <td width="39" bgcolor="#FFFFFF" bordercolordark="#000000" height="15" align="center"><font face="Verdana" size="1" color="#330099"><%=VALOR%></font></td>
            
      <td width="370" bgcolor="#FFFFFF" bordercolordark="#000000" height="15"> 
        <div align="left"><font face="Verdana" size="1" color="#330099"><%=UCASE(RS1("CETR_TX_DESC_TRANSACAO"))%></font></div>
      </td>
            
      <td width="134" bgcolor="#FFFFFF" bordercolordark="#000000" height="15"><font face="Verdana" size="1" color="#330099"> 
        <p align="center"><%=RS1("TRAN_CD_TRANSACAO")%></font></td>
            
      <td width="102" bgcolor="#FFFFFF" bordercolordark="#000000" height="15" align="center"><font face="Verdana" size="1" color="#330099"><%=valor2%></font></td>
            
      <td width="132" bgcolor="#FFFFFF" bordercolordark="#000000" height="15" align="center"><font face="Verdana" size="1" color="#330099"><%=RS1("CENA_CD_CENARIO_SEGUINTE")%></font></td>
            
      <td width="34" bgcolor="#FFFFFF" bordercolordark="#000000" height="15" align="center"><font face="Verdana" size="1" color="#330099"><%=RS1("BPPP_CD_BPP")%></font></td>
            
          </tr>
          <tr>
            
      <td bgcolor="#FFFFFF" bordercolordark="#000000" colspan="6" height="2" align="center"><img border="0" src="../../imagens/line.jpg" width="704" height="1"></td>
          </tr>
           <%
          RS1.MOVENEXT
          LOOP
          %>
        </table>
&nbsp;
  <p>
        <%
        end if
        RS.MOVENEXT
        LOOP
        else
        %>
        &nbsp;&nbsp;&nbsp;<font color="#800000" face="Verdana" size="2"><b> Não existem Cenários Cadastrados para a Seleção
        </b></font>
        <%end if%>
  </p>
 </form>
<p>&nbsp;</p>
</body>
</html>