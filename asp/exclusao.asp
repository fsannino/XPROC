<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
str_id = Request("ID")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

select case str_id
	case 1
		valor="Agrupamento ( Master List R/3 )"
		tit="Selecione o Agrupamento ( Master List R/3 ) que deseja excluir"
	case 2
		valor="Atividade"
		tit="Selecione a Atividade que deseja excluir"
	case 3
		valor="Transação"
		tit="Selecione a Transação que deseja excluir"
	case 4
		valor="Empresa"
		tit="Selecione a Empresa que deseja excluir"
end select
%>
<html>
<head>
<%
select case str_id
	case 1%>

<script>
function Confirma() 
{ 
if (document.frm1.selModulo.selectedIndex == 0)
     { 
	 alert("A seleção de um Master List R/3 é obrigatório!");
     document.frm1.selModulo.focus();
     return;
     }
     else
     {
     frm1.submit();
     }

 }
</SCRIPT>

	<%case 2%>

<script>
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
     else
     {
     frm1.submit();
     }

 }
</SCRIPT>

	<%case 3%>
	
<script>
function Confirma() 
{ 
if (document.frm1.selTransacao.selectedIndex == 0)
     { 
	 alert("A seleção de uma Transação é obrigatório!");
     document.frm1.selTransacao.focus();
     return;
     }
     else
     {
     frm1.submit();
     }

 }
</SCRIPT>

	<%case 4%>

<script>
function Confirma() 
{ 
if (document.frm1.selEmpresa.selectedIndex == 0)
     { 
	 alert("A seleção de uma Empresa é obrigatório!");
     document.frm1.selEmpresa.focus();
     return;
     }
     else
     {
     frm1.submit();
     }
 }
</SCRIPT>

<%end select%>

<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="exclusao2.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="2%"><a href="javascript:Confirma()"><img border="0" src="../imagens/confirma_f02.gif"></a></td>
      <td height="20" width="43%"><font color="#330099" size="2" face="Verdana"><b>Excluir</b></font></td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="40%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Exclusão
        de <%=valor%></font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b><%=tit%></b></font></td>
      <td width="14%"></td>
    </tr>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <%select case str_id
      case 1%>
      <td width="59%"><select size="1" name="selModulo">
        <option value="0">== Selecione o Agrupamento ( Master List R/3 )==</option>
          <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY MODU_TX_DESC_MODULO")
        do until rs.eof=true
        %>
		  <option value=<%=RS("MODU_CD_MODULO")%>><%=RS("MODU_TX_DESC_MODULO")%></option>
        <%
		rs.movenext
		loop        
        %>
        </select></td>
      <%case 2%>
      <td width="59%"><select size="1" name="selAtividade">
        <option value="0">== Selecione a Atividade ==</option>
        <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
        do until rs.eof=true
        %>
		  <option value=<%=RS("ATCA_CD_ATIVIDADE_CARGA")%>><%=RS("ATCA_TX_DESC_ATIVIDADE")%></option>
        <%
		rs.movenext
		loop        
        %>
        </select></td>
      <%case 3%>
      <td width="59%"><select size="1" name="selTransacao">
        <option value="0">== Selecione a Transação ==</option>
        <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO")
        do until rs.eof=true
        %>
		  <option value=<%=RS("TRAN_CD_TRANSACAO")%>><%=RS("TRAN_CD_TRANSACAO")%>-<%=RS("TRAN_TX_DESC_TRANSACAO")%></option>
        <%
       rs.movenext
		loop        
        %>
        </select></td>
      <%case 4%>
      <td width="59%"><select size="1" name="selEmpresa">
        <option value="0">== Selecione a Empresa</option>
        <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE ORDER BY EMPR_TX_NOME_EMPRESA")
        do until rs.eof=true
        %>
		  <option value=<%=RS("empr_CD_NR_EMPRESA")%>><%=RS("EMPR_TX_NOME_EMPRESA")%></option>
        <%
       rs.movenext
		loop        
        %>
        </select></td>
      <%end select%>
      <td width="14%"></td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
str_id = Request("ID")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

select case str_id
	case 1
		valor="Agrupamento ( Master List R/3 )"
		tit="Selecione o Agrupamento ( Master List R/3 ) que deseja excluir"
	case 2
		valor="Atividade"
		tit="Selecione a Atividade que deseja excluir"
	case 3
		valor="Transação"
		tit="Selecione a Transação que deseja excluir"
	case 4
		valor="Empresa"
		tit="Selecione a Empresa que deseja excluir"
end select
%>
<html>
<head>
<%
select case str_id
	case 1%>

<script>
function Confirma() 
{ 
if (document.frm1.selModulo.selectedIndex == 0)
     { 
	 alert("A seleção de um Master List R/3 é obrigatório!");
     document.frm1.selModulo.focus();
     return;
     }
     else
     {
     frm1.submit();
     }

 }
</SCRIPT>

	<%case 2%>

<script>
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
     else
     {
     frm1.submit();
     }

 }
</SCRIPT>

	<%case 3%>
	
<script>
function Confirma() 
{ 
if (document.frm1.selTransacao.selectedIndex == 0)
     { 
	 alert("A seleção de uma Transação é obrigatório!");
     document.frm1.selTransacao.focus();
     return;
     }
     else
     {
     frm1.submit();
     }

 }
</SCRIPT>

	<%case 4%>

<script>
function Confirma() 
{ 
if (document.frm1.selEmpresa.selectedIndex == 0)
     { 
	 alert("A seleção de uma Empresa é obrigatório!");
     document.frm1.selEmpresa.focus();
     return;
     }
     else
     {
     frm1.submit();
     }
 }
</SCRIPT>

<%end select%>

<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="exclusao2.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="15%">&nbsp;</td>
      <td height="20" width="2%"><a href="javascript:Confirma()"><img border="0" src="../imagens/confirma_f02.gif"></a></td>
      <td height="20" width="43%"><font color="#330099" size="2" face="Verdana"><b>Excluir</b></font></td>
      <td colspan="2" height="20" width="6">&nbsp;</td>
      <td height="20" width="40%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Exclusão
        de <%=valor%></font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <td width="59%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><b><%=tit%></b></font></td>
      <td width="14%"></td>
    </tr>
    <tr> 
      <td width="3%"></td>
      <td width="24%"></td>
      <%select case str_id
      case 1%>
      <td width="59%"><select size="1" name="selModulo">
        <option value="0">== Selecione o Agrupamento ( Master List R/3 )==</option>
          <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY MODU_TX_DESC_MODULO")
        do until rs.eof=true
        %>
		  <option value=<%=RS("MODU_CD_MODULO")%>><%=RS("MODU_TX_DESC_MODULO")%></option>
        <%
		rs.movenext
		loop        
        %>
        </select></td>
      <%case 2%>
      <td width="59%"><select size="1" name="selAtividade">
        <option value="0">== Selecione a Atividade ==</option>
        <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
        do until rs.eof=true
        %>
		  <option value=<%=RS("ATCA_CD_ATIVIDADE_CARGA")%>><%=RS("ATCA_TX_DESC_ATIVIDADE")%></option>
        <%
		rs.movenext
		loop        
        %>
        </select></td>
      <%case 3%>
      <td width="59%"><select size="1" name="selTransacao">
        <option value="0">== Selecione a Transação ==</option>
        <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO")
        do until rs.eof=true
        %>
		  <option value=<%=RS("TRAN_CD_TRANSACAO")%>><%=RS("TRAN_CD_TRANSACAO")%>-<%=RS("TRAN_TX_DESC_TRANSACAO")%></option>
        <%
       rs.movenext
		loop        
        %>
        </select></td>
      <%case 4%>
      <td width="59%"><select size="1" name="selEmpresa">
        <option value="0">== Selecione a Empresa</option>
        <%
        set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE ORDER BY EMPR_TX_NOME_EMPRESA")
        do until rs.eof=true
        %>
		  <option value=<%=RS("empr_CD_NR_EMPRESA")%>><%=RS("EMPR_TX_NOME_EMPRESA")%></option>
        <%
       rs.movenext
		loop        
        %>
        </select></td>
      <%end select%>
      <td width="14%"></td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="3%">&nbsp;</td>
      <td width="24%">&nbsp;</td>
      <td width="59%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
