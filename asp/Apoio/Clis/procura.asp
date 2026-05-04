<%
apoio=request("apoio")
opt=request("op")

chave=session("CdUsuario")
tipo=Session("Tipo")

if tipo=1 then
	apoio=2
else
	apoio=1
end if

if apoio=1 then
	valor="APOIADOR"
else
	valor="MULTIPLICADOR"
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>Consulta</title>
</head>

<script>
function Enviar()
{
if(this.strparam.value=="")
{
	alert('Você precisa especificar um parâmetro de busca!');
	this.strparam.focus()
	return;
}
else
{
	window.opener.location="cad_cli.asp?valor="+this.strparam.value +"&op="+this.op.value;
	window.close();
}
}
</script>

<body topmargin="0" leftmargin="0" onload="javascript:window.moveTo(400,250);this.strparam.focus()">

  &nbsp;

        
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center">
</p>

        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center"><font size="2" face="Verdana"><b>Digite a Matrícula ou a Chave&nbsp;</b></font>
        </p>
  <font size="2" face="Verdana">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center">
		<input type="text" name="strparam" size="22">
		<input type="hidden" name="str_tipo" size="5" value="<%=apoio%>"> 
 </p>
        
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
</font> 
<div align="center"></div>
  
<font size="2" face="Verdana"> 
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center">&nbsp; </p>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0" align="center"> 
  <input type="submit" value="Enviar" name="B1" onClick="Enviar()">
  <input name="op" type="hidden" id="op" value="<%=opt%>">
</font> 
</body>

</html>
