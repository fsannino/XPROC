<%
if request("excel")=1 then
	Response.Buffer = TRUE
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

	IF ID<>"0" THEN
		if len(compl1)=0 then
			compl1="CENA_CD_CENARIO='"& ID & "'"
		else
			compl1=compl1+" AND CENA_CD_CENARIO='"& ID & "'"
	END IF
	end if

if onda<>0 then
	set rsonda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA_REFAP WHERE ONDA_CD_ONDA=" & onda & " ORDER BY ONDA_CD_ONDA")
ELSE
	set rsonda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA_REFAP ORDER BY ONDA_CD_ONDA")
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
    <td colspan="2" height="20" width="112"><a href="gera_rel_cond_refap.asp?excel=1&amp;selMegaProcesso=<%=mega%>&amp;selOnda=<%=onda%>&amp;ID=<%=id%>&amp;ID2=<%=id2%>&amp;selStatus=<%=situ%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a> 
      
    </td>
    <td height="20" width="290">&nbsp; 
      
    </td>
  </tr>
</table>
<%end if%>
  <p align="center">
  <font color="#330099" face="Verdana" size="3">Relatório Geral de Cenários</font>
  <%
  tem=0
  do until rsonda.eof=true

if len(compl1)>0 then
	pre_compl = "WHERE ONDA_CD_ONDA=" & rsonda("ONDA_CD_ONDA") & " AND"
ELSE
	pre_compl = "WHERE ONDA_CD_ONDA=" & rsonda("ONDA_CD_ONDA") 
END IF

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO_REFAP " & pre_compl & compl1 & " ORDER BY CENA_CD_CENARIO"

set rs=db.execute(ssql)
if rs.eof=false then
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
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <div align="center">
    <center>
      <table border="0" width="79%" cellspacing="1" cellpadding="0">
        <tr>
          <td width="7%">&nbsp;</td>
          <td colspan="2">
            <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Status</b> 
              : EE - Em elabora&ccedil;&atilde;o / DF - Definido / DS - Desenhado</font></div>
          </td>
        </tr>
        <tr> 
          <td width="7%">&nbsp;</td>
          <td colspan="2"> 
            <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              PT - Pronto para teste / TD - Testado no PED / TQ - Tetado no PEQ</font></div>
          </td>
        </tr>
        <tr> 
          <td width="7%" bgcolor="#330099"> 
            <p style="margin-top: 0; margin-bottom: 0"><b><font size="1" face="Verdana" color="#FFFFFF">Código</font></b></p>
          </td>
          <td width="74%" bgcolor="#330099"> 
            <p style="margin-top: 0; margin-bottom: 0"><b><font size="1" face="Verdana" color="#FFFFFF">Descrição</font></b></p>
          </td>
          <td width="19%" bgcolor="#330099"> 
            <div align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Status</font></b></div>
          </td>
        </tr>
        <%end if%>
        <%DO UNTIL RS.EOF=TRUE
        set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO_REFAP WHERE CENA_CD_CENARIO='"& RS("CENA_CD_CENARIO")& "'")
        
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
          <td width="7%" align="left" bgcolor="<%=COR%>">
			<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=rs("CENA_CD_CENARIO")%></font> 
          </td>          
          <%else%>
          <td width="7%" align="left" bgcolor="<%=COR%>"><a href="gera_rel_geral_refap.asp?id=<%=rs("CENA_CD_CENARIO")%>&selMegaProcesso=<%=rs("MEPR_CD_MEGA_PROCESSO")%>&selProcesso=<%=rs("PROC_CD_PROCESSO")%>&selSubProcesso=<%=rs("SUPR_CD_SUB_PROCESSO")%>"> 
			<p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=rs("CENA_CD_CENARIO")%></font> 
          </a></td>          
          <%end if%>
            <td width="74%" align="left" bgcolor="<%=COR%>"> 
            <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1"><%=ATUAL("CENA_TX_TITULO_CENARIO")%></font> 
          </td>
          <td width="19%" align="left" bgcolor="<%=COR%>"> 
            <div align="center"><font face="Verdana" size="1"><%=ATUAL("CENA_TX_SITUACAO")%></font></div>
          </td>
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
       </center>
  		</div>
  		<p style="margin-top: 0; margin-bottom: 0">&nbsp; </p>
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
