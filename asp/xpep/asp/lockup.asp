<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function abrir()
{
window.open("texto.asp","_blank","width=200,height=100,history=0,sizeable=0,titlebar=0,scrollbars=0")
}
</script>

<body>
<form name="form1" method="post" action="">
  <p>
    <textarea name="textarea1" wrap="PHYSICAL"></textarea>
  </p>
  <p><a href="#" onclick=javascript:abrir()>Abrir P&aacute;gina</p></a>
</form>
</body>
</html>
