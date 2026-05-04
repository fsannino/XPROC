<<<<<<< HEAD
<%
'str_Opc = Request("txtOpc")
str_MegaProcesso= Request.form("txtMegaProcesso")
str_Processo = Request.form("txtProcesso")
str_SubProcesso = Request.form("txtSubProcesso")
str_Atividade = Request.form("txtAtividade")

if Session("CatUsu") = "indexA.js" then
   ls_script = "<script language=""JavaScript"" src=""../Templates/js/indexA.js""> </script>"
end if

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<SCRIPT>
function fullWindow() { 
   alert(document.frm1.txtTranSelecionada.value);
} 
</script>

<body bgcolor="#FFFFFF" text="#000000" onLoad=fullWindow()>
<%=ls_script%>
<script type= "text/javascript" language= "JavaScript">
<!--
//acesso();
goMenus();
//-->
</script>

<%=Session("CatUsu")%> 
<form name="frm1" method="post" action="">
  <p>
    <input type="text" name="txtTranSelecionada" value="<%=Session("CatUsu")%> ">
  </p>
  <p>&nbsp;</p>
</form>
</body>
</html>
=======
<%
'str_Opc = Request("txtOpc")
str_MegaProcesso= Request.form("txtMegaProcesso")
str_Processo = Request.form("txtProcesso")
str_SubProcesso = Request.form("txtSubProcesso")
str_Atividade = Request.form("txtAtividade")

if Session("CatUsu") = "indexA.js" then
   ls_script = "<script language=""JavaScript"" src=""../Templates/js/indexA.js""> </script>"
end if

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<SCRIPT>
function fullWindow() { 
   alert(document.frm1.txtTranSelecionada.value);
} 
</script>

<body bgcolor="#FFFFFF" text="#000000" onLoad=fullWindow()>
<%=ls_script%>
<script type= "text/javascript" language= "JavaScript">
<!--
//acesso();
goMenus();
//-->
</script>

<%=Session("CatUsu")%> 
<form name="frm1" method="post" action="">
  <p>
    <input type="text" name="txtTranSelecionada" value="<%=Session("CatUsu")%> ">
  </p>
  <p>&nbsp;</p>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
