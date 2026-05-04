<%
str_Dia = request("SelDia")
str_Mes = request("SelMes")
str_Ano = request("SelAno")
str_DtPrevTermino = str_Mes & "/" & str_Dia & "/" & str_Ano
IF IsDate(str_DtPrevTermino) = false then
   response.write " Erro "
   response.write str_DtPrevTermino
else
   response.write " Acerto "
   response.write str_DtPrevTermino
end if   
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">

</body>
</html>
