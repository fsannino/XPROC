<%
CONECTA="Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"
set db = Server.CreateObject("ADODB.Connection")
db.Open CONECTA
id2=""
id=""
int_sequencia = 0
mega=request("selMegaProcesso")
proc=request("selProcesso")
subproc=request("selSubProcesso")
onda=request("selOnda")
ID=REQUEST("ID")
id2=request("ID2")

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
	set rsonda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA=" & onda & " ORDER BY ONDA_CD_ONDA")
ELSE
	set rsonda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA ORDER BY ONDA_CD_ONDA")
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
<script language="JavaScript">
<!--
function Confirma()
     {
	  document.frm1.submit();
	 }

//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="grava_alteracao_sequencia.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
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
    <td colspan="3" height="20">
        <table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img src="../../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
            <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
            <td width="26">&nbsp;</td>
            <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28">&nbsp;</td>
            <td width="26">&nbsp;</td>
            <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          </tr>
        </table>
      </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="12%">&nbsp;</td>
      <td width="80%"> 
        <div align="center"><font color="#330099" face="Verdana" size="3">Altera 
          ordem de visualiza&ccedil;&atilde;o </font><font color="#330099" face="Verdana" size="3"> 
          de Cenários</font></div>
      </td>
      <td width="8%">&nbsp;</td>
    </tr>
    <tr>
      <td width="12%">&nbsp;</td>
      <td width="80%">&nbsp;</td>
      <td width="8%">&nbsp;</td>
    </tr>
  </table>
    <%
  tem=0
  do until rsonda.eof=true

if len(compl1)>0 then
	pre_compl = "WHERE ONDA_CD_ONDA=" & rsonda("ONDA_CD_ONDA") & " AND"
ELSE
	pre_compl = "WHERE ONDA_CD_ONDA=" & rsonda("ONDA_CD_ONDA") 
END IF
ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO " & pre_compl & compl1 & " ORDER BY CENA_NR_SEQUENCIA_ORDEM"
set rs=db.execute(ssql)
if rs.eof=false then
tem=tem+1
  %>
  <table border="0" width="100%">
    <tr>
      <td width="19%">
        <td width="81%">
        &nbsp; <font face="Verdana" size="2" color="#330099">Onda 
          : <b><%=RSONDA("ONDA_TX_ABREV_ONDA")%> - <%=RSONDA("ONDA_TX_DESC_ONDA")%></b></font>
      </td>
    </tr>
  </table>
  <div align="center">
    <center>
      <table border="0" width="695" cellspacing="1" cellpadding="0">
        <tr> 
          <td width="94" bgcolor="#330099"> 
            <p style="margin-top: 0; margin-bottom: 0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Código</font></b></font></p>
          </td>
          <td width="554" bgcolor="#330099"> 
            <p style="margin-top: 0; margin-bottom: 0"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Descrição</font></b></font></p>
          </td>
          <td width="43" bgcolor="#330099"><font size="2"></font></td>
        </tr>
        <%end if%>
        <%DO UNTIL RS.EOF=TRUE
        set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& RS("CENA_CD_CENARIO")& "'")
        int_sequencia = int_sequencia + 1
        IF COR="WHITE" THEN
        	COR="#CACACA"
			'CINZA CLARO
			COR="#F7F7F7"
        ELSE
        	COR="WHITE"
        END IF
        %>
        <tr> 
          <td width="94" align="left" bgcolor="<%=COR%>">
            <div align="center"><a href="gera_rel_geral.asp?id=<%=rs("CENA_CD_CENARIO")%>&selMegaProcesso=<%=rs("MEPR_CD_MEGA_PROCESSO")%>&selProcesso=<%=rs("PROC_CD_PROCESSO")%>&selSubProcesso=<%=rs("SUPR_CD_SUB_PROCESSO")%>"><font size="2" face="Verdana"> 
              <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <input type="hidden" name="txtCen<%=int_sequencia%>" value="<%=rs("CENA_CD_CENARIO")%>">
              <%=rs("CENA_CD_CENARIO")%> </font> </font></a> </div>
          </td>
          <td width="554" align="left" bgcolor="<%=COR%>"><font size="2" face="Verdana"> 
            <p style="margin-top: 0; margin-bottom: 0"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ATUAL("CENA_TX_TITULO_CENARIO")%> </font>
            </font></td>
          <td width="43" align="left" bgcolor="<%=COR%>"> 
            <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <input type="text" name="txtSeq<%=int_sequencia%>" value="<%=rs("CENA_NR_SEQUENCIA_ORDEM")%>" size="4" maxlength="4">
            </font></td>
        </tr>
        <%
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
        <%
        ELSE
        %>       
        <%
        END IF
        loop
        %>
  <input type="hidden" name="txtQtdObj" value="<%=int_sequencia%>">
</form>
<%if tem=0 then%>
<font color="#800000" face="Verdana" size="2"><b>Não
existe nenhum cenário cadastrado para a seleção</b></font>
<%end if%>
</body>
</html>
