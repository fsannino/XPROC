<%
chave=ucase(REQUEST("CHAVE"))

if chave="" then
	response.redirect "erro.asp"	
end if

Session("CdUsuario")=chave

SELECT CASE  chave

CASE "DCOC"
	opt=1
	Session("Tipo")=0
CASE "XD34"
	opt=1
	Session("Tipo")=0
CASE "XK45" 'JOÃO LUIZ
	opt=1
	Session("Tipo")=2
CASE "XD83" ' SERGIO
	opt=1
	Session("Tipo")=2
CASE "X939" ' KATIA
	opt=1
	Session("Tipo")=2
CASE "XD35" ' GUSTAVO
	opt=1
	Session("Tipo")=2
CASE "SD39" ' VELOSO
	opt=1
	Session("Tipo")=2
CASE "XD47" ' ROBSON
	opt=1
	Session("Tipo")=2
CASE "EADE" ' MÁRIA
	opt=1
	Session("Tipo")=2
CASE "DCX0" 'DINIZ
	opt=1
	Session("Tipo")=0
CASE "K069" ' MARIA AMALIA
	opt=1
	Session("Tipo")=0
CASE "WS04" ' RESTUM
	opt=1
	Session("Tipo")=0
CASE "SM23" ' ALENCAR
	opt=1
	Session("Tipo")=0
CASE "B511" ' LUIZ ANTONIO PEREIRA DE ARAUJO
	opt=1
	Session("Tipo")=0
CASE "RV61" ' SONIA TUAN
	opt=1
	Session("Tipo")=0	
CASE "SD02"
	opt=1
	Session("Tipo")=1
CASE "BE05" ' ANA LUCIA VALENTE
	opt=1
	Session("Tipo")=1
CASE "XH09"
	opt=1
	Session("Tipo")=1
CASE "X964"
	opt=1
	Session("Tipo")=1
CASE ELSE
	opt=0
END SELECT

if opt = 1 then
	valor="menu.asp?cli=" & opt
else
	valor="Template/index.asp?op=0"
end if
%>

<html>
<head>
<title>Aguarde, Carregando Base de Apoiadores Locais...</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;

</script>

<SCRIPT>
function fullWindow(url) { 
    
	var str = "left=0,screenX=0,top=0,screenY=0,resizable=no,scrollbars=yes,toolbar=no,location=no";
	
	var a = document.URL;
	var n=0;

	for (var i = 1 ; i < 1000; i++)
	{
	var final=a.slice(0,i)
	var t=a.slice(i-1,i);
	if (t=='/')
	{
	n = n + 1;
	}
	if(n == 4)
	{
	i = 1000;
	}
	}
	var tam=final.length;
	var caminho = final.slice(0,tam-1);

    if (window.screen) {
      var ah = screen.availHeight - 30;
      var aw = screen.availWidth - 10;
      str += ",height=" + ah;
      str += ",innerHeight=" + ah;
      str += ",width=" + aw;
      str += ",innerWidth=" + aw;
    }
    //win=window.open(caminho+url, "w", str);
    window.location = caminho+url
} 
</SCRIPT>

<body bgcolor="#FFFFFF" text="#000000" onLoad=fullWindow('/asp/apoio/<%=valor%>')>
</body>
</html>