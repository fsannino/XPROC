<%
chave=ucase(REQUEST("CHAVE"))

if chave="" then
	response.redirect "erro.asp"	
end if
%>
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
    win=window.open(caminho+url, "w", str);
	window.location="revert.asp"
} 
</SCRIPT>
<%if tem=0 then%>
<body bgcolor="#FFFFFF" text="#000000" vlink="#0000FF" alink="#0000FF" onLoad=fullWindow('/asp/apoio/itens.asp?chave=<%=chave%>')>
<%else%>
<body bgcolor="#FFFFFF" text="#000000" vlink="#0000FF" alink="#0000FF">
<%end if%>
 <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
 <p style="margin-top: 0; margin-bottom: 0" align="center">  </p>
 </body>
</html>