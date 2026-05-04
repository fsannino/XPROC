<%
'id=request("ID")
'planilha=request("plan")

PLAN=1
ID="MM.PET.0001"

dim objFileSys

set objFileSys = Server.CreateObject("Scripting.FileSystemObject")

main = Server.MapPath("./")

if objFileSys.FileExists(main&"\plans\" & ID & ".xls") = false then
	ObjFileSys.CopyFile main&"\templates\" & plan & ".xls", main&"\plans\" & ID & ".xls"
end if

'response.redirect "exibe_planilha?ID=" & ID

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Nova pagina 1</title>
</head>

<body>

</body>

</html>
