<%
if request("excel")=1 then
	Response.Buffer = True
	Response.ContentType = "application/vnd.ms-excel"
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

id2=""
id=""

mega=request("selMegaProcesso")
proc=request("selProcesso")
subproc=request("selSubProcesso")
onda=request("selOnda")
ID=REQUEST("ID")
id2=request("ID2")
situ=request("selStatus")
str_Empresa=request("selEmpresa")
str_assunto=request("selAssunto")
str_classe=request("selClasse")

str_Escopo = request("selEscopo")

if len(id2)=11 then
	response.redirect "gera_rel_geral.asp?ID=" & id2
end if

if mega<>0 or proc<>0 or subproc<>0 then

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
	else
		compl1=compl1+" AND ONDA_CD_ONDA<>4"
	end if
end if

if str_Assunto<>0 then
	if len(compl1)=0 then
		compl1= " AND SUMO_NR_CD_SEQUENCIA =" & str_Assunto 
	else
		compl1=compl1+ " AND SUMO_NR_CD_SEQUENCIA=" & str_Assunto
	end if
end if

if str_Classe <> 0 then
	if len(compl1)=0 then
		compl1= " AND CLCE_CD_NR_CLASSE_CENARIO =" & str_Classe 
	else
		compl1=compl1+ " AND CLCE_CD_NR_CLASSE_CENARIO=" & str_Classe
	end if
end if


if str_Escopo <> 2 and len(str_Escopo)<>0 then
	if len(compl1)=0 then
		compl1= " AND CENA_TX_SITUACAO_VALIDACAO =" & str_Escopo
	else
		compl1=compl1+ " AND CENA_TX_SITUACAO_VALIDACAO=" & str_Escopo
	end if
end if

IF ID<>"0" THEN
	if len(compl1)=0 then
		compl1="CENA_CD_CENARIO='"& ID & "'"
	else
		compl1=compl1+" AND CENA_CD_CENARIO='"& ID & "'"
	END IF
end if

'response.Write(str_Empresa)
'response.Write(str_Empresa)

'response.Write(compl1)
	
if onda<>0 then
	set rsonda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA=" & onda & " ORDER BY ONDA_CD_ONDA")
ELSE
	set rsonda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_CD_ONDA")
END IF
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="">
<%if request("excel")<>1 then%>
 <table width="842" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="160" height="20" colspan="2">&nbsp;</td>
      <td width="346" height="60" colspan="3">&nbsp;</td>
      <td width="330" valign="top" colspan="2"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td height="20" width="155">&nbsp; 
      
    </td>
    <td colspan="2" height="20" width="31">&nbsp; 
      
    </td>
    <td height="20" width="244">&nbsp; 
      
    </td>
    <td colspan="2" height="20" width="112"><a href="gera_rel_cond.asp?excel=1&amp;selMegaProcesso=<%=mega%>&amp;selProcesso=<%=proc%>&amp;selsubProcesso=<%=subproc%>&amp;selOnda=<%=onda%>&amp;ID=<%=id%>&amp;ID2=<%=id2%>&amp;selStatus=<%=situ%>&amp;selEmpresa=<%=str-Empresa%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a> 
      
    </td>
    <td height="20" width="290">&nbsp; 
      
    </td>
  </tr>
</table>
<%end if%>
  <p align="center"> <font color="#330099" face="Verdana" size="3">Relatório Geral 
    de Cenários </font> 
    <%
  tem=0
  do until rsonda.eof=true

if len(compl1)>0 then
	pre_compl = "WHERE ONDA_CD_ONDA=" & rsonda("ONDA_CD_ONDA") & " AND"
ELSE
	pre_compl = "WHERE ONDA_CD_ONDA=" & rsonda("ONDA_CD_ONDA") 
END IF

if str_Empresa = "0" then
   str_SQl_1 = ""
elseif str_Empresa = "9" then
   str_SQl_1 = " and CENA_TX_EMPRESA_RELAC is NULL "
else
   str_SQl_1 = " and CENA_TX_EMPRESA_RELAC LIKE '%" & str_Empresa  & "%'"
end if
ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO " & pre_compl & compl1 & str_SQl_1 & " ORDER BY CENA_NR_SEQUENCIA_ORDEM"

'response.write ssql

set rs=db.execute(ssql)

if rs.eof=false then
   'CENA_TX_EMPRESA_RELAC
tem=tem+1
  %>
  <table border="0" width="100%">
    <tr>
      <td width="19%" bgcolor="#FFFFFF">
        <td width="81%">
        <p style="margin-top: 0; margin-bottom: 0">&nbsp; <font face="Verdana" size="2" color="#330099"><b>Onda
        :<%=RSONDA("ONDA_TX_ABREV_ONDA")%>  - <%=RSONDA("ONDA_TX_DESC_ONDA")%></b></font></p>
      </td>
    </tr>
  </table>
      
  <table border="0" width="800" cellspacing="1" cellpadding="0">
    <tr> 
      <td width="45">&nbsp;</td>
      <td colspan="4" width="583"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Status</b> 
          : EE - Em elabora&ccedil;&atilde;o / DF - Definido / DS - Desenhado</font></div></td>
    </tr>
    <tr> 
      <td width="45">&nbsp;</td>
      <td colspan="4" width="583"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
          PT - Pronto para teste / TD - Testado no PED / TQ - Testado no PEQ</font></div></td>
    </tr>
    <tr> 
      <td width="45" bgcolor="#330099" align="center"> <b><font size="1" face="Verdana" color="#FFFFFF">Código</font></b></td>
      <td width="80" bgcolor="#330099" align="center"> <b><font size="1" face="Verdana" color="#FFFFFF">Descrição</font></b></td>
      <td width="173" bgcolor="#330099" align="center"> <div align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Assunto</font></b></div></td>
      <td width="163" bgcolor="#330099" align="center"> <b><font color="#FFFFFF" size="1" face="Verdana">Escopo</font></b></td>
      <td width="155" bgcolor="#330099" align="center"> <div align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Empresa</font></b></div></td>
      <td width="107" bgcolor="#330099" align="center"> <div align="center"><b><font size="1" face="Verdana" color="#FFFFFF"> 
          Status </font></b></div></td>
    </tr>
    <%end if%>
    <%DO UNTIL RS.EOF=TRUE
    
   		if rs("CENA_TX_SITUACAO_VALIDACAO")=0 then
			val_escopo="FORA DO ESCOPO"
		else
			if rs("CENA_TX_SITUACAO_VALIDACAO")=1 then
				val_escopo="DENTRO DO ESCOPO"
			end if
		end if

    
        set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& RS("CENA_CD_CENARIO")& "'")
                   if not Isnull(atual("SUMO_NR_CD_SEQUENCIA")) then
      str_SQL = ""
      str_SQL = str_SQL & " SELECT SUMO_TX_DESC_SUB_MODULO, "
      str_SQL = str_SQL & "     SUMO_NR_CD_SEQUENCIA"
      str_SQL = str_SQL & " FROM SUB_MODULO"
      str_SQL = str_SQL & " WHERE SUMO_NR_CD_SEQUENCIA = " & atual("SUMO_NR_CD_SEQUENCIA")
      set rs_Modulo = db.Execute(str_SQL)
	  'RESPONSE.Write(str_SQL)
      if not rs_Modulo.EOF then
         str_DsAssunto = rs_Modulo("SUMO_TX_DESC_SUB_MODULO")
      else
         str_DsAssunto = " não enconttado o assunto "
      end if
      rs_Modulo.close
	else
	  str_DsAssunto = ""
	end if  

        IF COR="WHITE" THEN
        	COR="#E4E4E4"
        ELSE
        	COR="WHITE"
        END IF
		str_Imprimir = 1
		if situ <> "0" then
		   if atual("CENA_TX_SITUACAO") = situ then
		      str_Imprimir = 1
		    else 	  
              str_Imprimir = 0	  
		   end if		
		end if   
		IF str_Imprimir = 1 THEN			  
        %>
    <tr> 
      <%if request("excel")=1 then%>
      <td width="45" align="center" bgcolor="<%=COR%>"> 
        <font face="Verdana" size="1"><%=rs("CENA_CD_CENARIO")%></font> </td>
      <%else%>
      <td width="80" align="center" bgcolor="<%=COR%>"><a href="gera_rel_geral.asp?id=<%=rs("CENA_CD_CENARIO")%>&selMegaProcesso=<%=rs("MEPR_CD_MEGA_PROCESSO")%>&selProcesso=<%=rs("PROC_CD_PROCESSO")%>&selSubProcesso=<%=rs("SUPR_CD_SUB_PROCESSO")%>"> 
        <font face="Verdana" size="1"><%=rs("CENA_CD_CENARIO")%></font> </a></td>
      <%end if%>
      <td width="173" align="center" bgcolor="<%=COR%>"> 
        <div align="left"><font face="Verdana" size="1"><%=ATUAL("CENA_TX_TITULO_CENARIO")%></font> </div></td>
      <td width="163" align="center" bgcolor="<%=COR%>"><font size="1" face="Verdana"><%=str_DsAssunto%></font></td>
      <td width="155" align="center" bgcolor="<%=COR%>"><div align="center"><font size="1" face="Verdana"><%=val_escopo%></font></div></td>
      <td width="107" align="center" bgcolor="<%=COR%>"><div align="center"><font face="Verdana" size="1"><%=ATUAL("CENA_TX_EMPRESA_RELAC")%></font></div></td>
      <td width="47" align="center" bgcolor="<%=COR%>"> 
        <div align="center"><font face="Verdana" size="1"><%=ATUAL("CENA_TX_SITUACAO")%></font></div></td>
    </tr>
    <%
		'else
		'response.write str_Imprimir
		END IF
        RS.MOVENEXT
        %>
    <%
        LOOP
        RSONDA.MOVENEXT
        IF TEM<>0 THEN
        %>
  </table>
        <%
        ELSE
        %>
        <BR>
        <%
        END IF
        loop
        %>
                
 </form>
<%if tem=0 then%>
<p style="margin-top: 0; margin-bottom: 0" align="center"><font color="#800000" face="Verdana" size="2"><b>Não
existe nenhum cenário cadastrado para a seleção</b></font></p>
<%end if%>
</body>
</html>