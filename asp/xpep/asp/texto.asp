<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function transfere()
{
window.opener.document.form1.textarea1.value=document.form1.select1.value;
window.close();
}
</script>
<body>
<form name="form1" method="post" action="">
  <p> 
    <select name="select1">
      <option value="Rio de Janeiro">Rio de Janeiro</option>
      <option value="S&atilde;o Paulo">S&atilde;o Paulo</option>
    </select>
  </p>
  <p><a href="#" onclick="javascript:transfere()">Fechar</a></p>
</form>
</body>
</html>
