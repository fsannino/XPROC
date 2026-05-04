<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conecta.asp" -->
<%

strdata=""
strperiodo=""
strerro=""
strorgao=""

strdata = request("txtData")
strperiodo = request("selPeriodo")
strerro = request("selErro")
strOrgao = request("selOrgao")
strModo = request("selModo")

if strdata<>"" then

	Session("Data_inicio") = strdata
	Session("Periodo") = cint(strperiodo)
	Session("Modo") = strModo
	
	if strerro="TODOS" then
		Session("Erro")="TODOS"
		Session("Compl")=""
	else
		Session("Erro")=strErro
		Session("Compl") = " AND IDENTIFICADOR='" & strErro & "'"
	end if

	if strOrgao="TODOS" then
		Session("Orgao")="TODOS"
		Session("Compl") = Session("Compl")
	else
		Session("Orgao")=strOrgao
		Session("Compl") = Session("Compl") & " AND ORGAO LIKE '" & strOrgao & "%'"
	end if

	response.redirect "target.asp"

end if	

set rs1 = db.execute("SELECT DISTINCT IDENTIFICADOR FROM SINERGIA WHERE IDENTIFICADOR <> '' ORDER BY IDENTIFICADOR")

set rs2 = db.execute("SELECT DISTINCT ORGAO FROM SINERGIA WHERE ORGAO <> '' ORDER BY ORGAO ")

%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Configuração do Sistema</title>
</head>

<script>
function Salvar()
{
document.frm1.submit()
}
</script>

<body>
<form name="frm1" method="POST" action="configuracao.asp">

<p><b><font face="Verdana" size="2">Configuração do Sistema</font></b></p>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="71%" id="AutoNumber1" height="58">
    <tr> 
      <td width="61%" height="18"><font face="Verdana" size="2">Data Base</font></td>
      <td width="39%" height="18"> 
        <input type="text" name="txtData" size="20" value="<%=Session("Data_inicio")%>">
        <br> <b><font face="Verdana" size="1">(formato dd/mm/yyyy)</font></b>
      </td>
    </tr>
    <tr> 
      <td width="61%" height="19">&nbsp;</td>
      <td width="39%" height="19">&nbsp;</td>
    </tr>
    <tr> 
      <td width="61%" height="19"><font face="Verdana" size="2">Período para Consulta</font></td>
      <td width="39%" height="19"> 
        <select size="1" name="selPeriodo">
          <%
                         periodo=Session("Periodo")
                         
                         select case periodo
                         case 7
                         	C1="selected"
                         case 14
                         	C2="selected"
                         case 21
                         	C3="selected"
                         case 28
                         	C4="selected"
                         case 35
                         	C5="selected"
                         end select
                         %>
          <option <%=C1%> value="7">7</option>
          <option <%=C2%> value="14">14</option>
          <option <%=C3%> value="21">21</option>
          <option <%=C4%> value="28">28</option>
          <option <%=C5%> value="35">35</option>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="61%" height="19">&nbsp;</td>
      <td width="39%" height="19">&nbsp;</td>
    </tr>
    <tr> 
      <td width="61%" height="22"> 
        <p><font face="Verdana" size="2">Modo de Visualiza&ccedil;&atilde;o </font>
        <br><font face="Verdana" size="2"><b>(Somente Perfil de Atendimento)</b></font>
        </td>
      <td width="39%" height="22"> 
        <select size="1" name="selModo">
          <%
		  if Session("Modo")="P" then
		  %>
		  <option value="Q">QUANTITATIVO</option>
		  <option selected value="P">PERCENTUAL</option>
		  <%
		  else
		  %>
		  <option selected value="Q">QUANTITATIVO</option>
		  <option value="P">PERCENTUAL</option>
		  <%
		  end if
		  %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="61%" height="19">&nbsp;</td>
      <td width="39%" height="19">&nbsp;</td>
    </tr>
    <tr> 
      <td width="61%" height="19"><font face="Verdana" size="2">Tipo</font></td>
      <td width="39%" height="19"> 
        <select size="1" name="selErro">
          <option value="TODOS">== TODOS ==</option>
          <%
                         do until rs1.eof=true
                         if trim(Session("Erro"))= trim(rs1("IDENTIFICADOR")) then
                         	check1="selected"                         
                         else
                         	check1=""
                         end if
                         %>
          <option <%=check1%> value="<%=rs1("IDENTIFICADOR")%>"><%=rs1("IDENTIFICADOR")%></option>
          <%
                         rs1.movenext
                         loop
                         %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="61%" height="19">&nbsp;</td>
      <td width="39%" height="19">&nbsp;</td>
    </tr>
    <tr> 
      <td width="61%" height="19"><font face="Verdana" size="2">Órgão</font></td>
      <td width="39%" height="19"> 
        <select size="1" name="selOrgao">
          <option value="TODOS">== TODOS ==</option>
          <%
                         f = 0
                         reg = rs2.recordcount
                         
                         do until f = reg
                         
                         parte=""
                         parte2=""
                         
                         org_atual = rs2("ORGAO")
                         tamanho = len(org_atual)
                         
                         i=1
                         
                         tem=0
                         
                         do until i = tamanho + 1
                         	parte2=left(org_atual,i)
                         	parte=right(parte2,1)
                         	if parte="/" then
                         		tem = tem + 1
                         	end if
                         	i = i + 1
                         loop
                         
                         if tem < 2 then
                         if trim(Session("Orgao"))= trim(rs2("ORGAO")) then
                         	check2="selected"                         
                         else
                         	check2=""
                         end if
                         %>
          <option <%=check2%> value="<%=rs2("ORGAO")%>"><%=rs2("ORGAO")%></option>
          <%
                         end if
                         f = f + 1
                         rs2.movenext
                         loop
                         %>
        </select>
      </td>
    </tr>
  </table>

<p>
<input type="button" value="Gravar" name="B1" onClick="Salvar()">
<input type="button" value="Retornar" name="B2" onClick="window.location='target.asp'"></p>
</form>
</body>

</html>