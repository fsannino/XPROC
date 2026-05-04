
<%
str_chave = request("chave")
str_Senha = Request("senha")
%>
<HTML>
<script language="JavaScript" type="text/JavaScript">
<!--
function jump_3()
{
    var sscWindow
    sscWindow= window.open('http://localhost/xproc/index.asp?senha=<%=str_Senha%>&chave=<%=str_chave%>', 'test3', 'left=0,screenX=0,top=0,screenY=0,resizable=no,scrollbars=yes,toolbar=no,location=no');

    if (window.focus)
    {
        sscWindow.focus()
    }
    return false;
}
//-->

<!--
    opener.opener = opener;
    opener.close();
//-->


</script>
<body onload="javascript:jump_3();">


</BODY>

</HTML>
