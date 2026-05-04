<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Exclusão de Micro Perfil</title>
</head>

<script>
function valida_exclusao()
{
if(confirm("Confirma Exclusão do Micro Perfil Selecionado?"))
{ 
window.location="valida_excluir_micro.asp?selMicro=" + this.selMicro.value
}
else
{
window.location="seleciona_micro_perfil.asp?pOPT=3"
}
}
</script>

<body onLoad="valida_exclusao()">

<p><input type="hidden" name="selMicro" size="20" value="<%=request("selMicro")%>"></p>

</body>
</html>
