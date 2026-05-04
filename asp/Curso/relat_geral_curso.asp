<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega = request("MEGA")
str_onda = request("selOnda")

if request("rdbStatus") <> "" then
	strStatus = request("rdbStatus")
else
	strStatus = "0"
end if

if str_mega > 0 then
	compl=" and  " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO=" + str_mega
else
	compl=""
end if

if str_onda >0 then
	compl2=" and  " & Session("PREFIXO") & "CURSO.ONDA_CD_ONDA = " + str_onda
else
	compl2=""
end if

if strStatus = "0" then     '*** TODOS
	COMPL3 = ""
elseif strStatus = "1" then '*** ATIVOS
	COMPL3 = " and " & Session("PREFIXO") & "CURSO.CURS_TX_STATUS_CURSO = '1'"
elseif strStatus = "2" then	'*** INATIVOS
	COMPL3 = " and " & Session("PREFIXO") & "CURSO.CURS_TX_STATUS_CURSO = '0'"
end if

SSQL1="SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO AS MEGA, " & Session("PREFIXO") & "CURSO.* FROM " & Session("PREFIXO") & "CURSO INNER JOIN " & Session("PREFIXO") & "MEGA_PROCESSO ON " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO where MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO > 0 " & COMPL & COMPL2 & COMPL3 & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "CURSO.CURS_CD_CURSO"
'RESPONSE.Write(SSQL1)
SET RS=DB.EXECUTE(SSQL1)

SET MEGA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
set rs_onda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ABRANGENCIA_CURSO WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_TX_DESC_ONDA")
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>
<SCRIPT>
function envia()
{
	var intTipo = 0;
	if (document.frm1.rdbStatus[0].checked)
	{ 
		intTipo = 0;
	}
	
	if (document.frm1.rdbStatus[1].checked)
	{
		intTipo = 1;
	}
	
	if (document.frm1.rdbStatus[2].checked)
	{
		intTipo = 2;
	}
	
	this.location.href='relat_geral_curso.asp?mega='+document.frm1.selMegaProcesso.value+'&selOnda='+document.frm1.selOnda.value+'&rdbStatus='+intTipo;
}
</SCRIPT>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" link="#800000" vlink="#800000" alink="#800000">
<form method="POST" action="" name="frm1">
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
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>
      </td>
    </tr>
    <tr>
      <td>
        <div align="center">
          <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório
          Geral de Cursos</font></div>
      </td>
    </tr>
  </table>
  <p><b><font face="Verdana" color="#330099" size="2"> </font></b></p>
  <table width="75%" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="33%" height="25" valign="middle"><b><font face="Verdana" color="#330099" size="2">Mega-Processo Selecionado :</font></b></td>
      <td width="65%" height="25" valign="middle"><b><font face="Verdana" color="#330099" size="2">
        <select size="1" name="selMegaProcesso" onChange="javascript:envia();">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL MEGA.EOF=TRUE
  if trim(str_mega)=trim(MEGA("MEPR_CD_MEGA_PROCESSO")) then
  %>
          <option selected value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%else%>
          <option value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
  end if
  MEGA.MOVENEXT
  LOOP
  %>
        </select>
        </font></b></td>
      <td width="2%">&nbsp;</td>
    </tr>
    <tr>
      <td height="25" valign="middle"><font face="Verdana" size="2" color="#330099"><b>Onda :</b></font></td>
      <td height="25" valign="middle"><select size="1" name="selOnda" onChange="javascript:envia();">
          <option value="0">== Selecione a Onda ==</option>
          <%DO UNTIL RS_ONDA.EOF=TRUE
      IF TRIM(str_onda)=trim(rs_onda("ONDA_CD_ONDA")) then
      %>
          <option selected value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> 
          - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%else%>
          <option value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> 
          - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%
		END IF
		RS_ONDA.MOVENEXT
		LOOP
		%>
        </select></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td height="25" valign="middle"><font face="Verdana" size="2" color="#330099"><b>Status Curso:</b></font></td>
      <td height="25" valign="middle">
	  
	  <%if strStatus = "0" then%>	
		<input name="rdbStatus" type="radio" value="0" checked onClick="javascript:envia();">&nbsp;<font face="Verdana" size="2" color="#330099">Todos</font>
	  <%else%>	
		<input name="rdbStatus" type="radio" value="0" onClick="javascript:envia();">&nbsp;<font face="Verdana" size="2" color="#330099">Todos</font>
	  <%end if%>
	  
	  <%if strStatus = "1" then%>	  
	  	<input name="rdbStatus" type="radio" value="0" checked onClick="javascript:envia();">&nbsp;<font face="Verdana" size="2" color="#330099">Ativo</font>&nbsp;&nbsp;
      <%else%>	
	  	<input name="rdbStatus" type="radio" value="0" onClick="javascript:envia();">&nbsp;<font face="Verdana" size="2" color="#330099">Ativo</font>&nbsp;&nbsp;
	  <%end if%>
	  
	  <%if strStatus = "2" then%>	
		<input name="rdbStatus" type="radio" value="0" checked onClick="javascript:envia();">&nbsp;<font face="Verdana" size="2" color="#330099">Inativo</font>&nbsp;&nbsp;
	  <%else%>	
		<input name="rdbStatus" type="radio" value="0" onClick="javascript:envia();">&nbsp;<font face="Verdana" size="2" color="#330099">Inativo</font>&nbsp;&nbsp;
	  <%end if%>	 	 
	 
	  </td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td height="5"></td>
      <td height="5"></td>
      <td height="5"></td>
    </tr>
  </table>
  <table border="0" width="81%">
          <tr>
            <td width="21%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Mega-Processo</font></b></td>
            <td width="45%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Curso</font></b></td>
          </tr>
          <%
          tem=0
          
			atual1=""
			ant1=""			

          do until rs.eof=true
          atual1=rs("mepr_cd_mega_processo")
          %>
          <tr>
				<%
             	SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
             	if atual1<>ant1 then
					NOME1=RS1("MEPR_TX_DESC_MEGA_PROCESSO")            
				else
					nome1=""
				end if
				
				if nome1="" then
					cor="white"
				else
					cor="#CCCCCC"	
				end if			
				%>
            <td width="21%" bgcolor="<%=cor%>"><font size="1" face="Verdana"><%=nome1%></font></td>
            <td width="45%" bgcolor="#FFFFEA"><font size="1" face="Verdana"><a href="exibe_curso.asp?curso=<%=rs("CURS_CD_CURSO")%>"><%=rs("CURS_TX_NOME_CURSO")%></a></font></td>
          </tr>
          <%
          tem=tem+1
          
          ant1=rs("mepr_cd_mega_processo")
          
          rs.movenext
          
          on error resume next
          atual1=rs("mepr_cd_mega_processo")
          
          loop
          %>
          
  </table>
<b>
<%if tem=0 then%>
<font face="Verdana" size="2" color="#800000">Nenhum Registro encontrado</font></b>
<%end if%>
</form>

</body>

</html>
