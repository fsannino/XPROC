<%

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("selMacroPerfil") <> 0 then
   str_MacroPerfil = request("selMacroPerfil")
else
   str_MacroPerfil = "0"
end if

'response.Write(" macro perfil " & str_MacroPerfil & "<p>")

if request("txtNomeTecnico") <> "" then
   str_NomeTecnico = request("txtNomeTecnico")
else
   str_NomeTecnico = ""
end if

'response.Write(" nome tecnico " & str_NomeTecnico & "<p>")

'response.Write(request("pOrdem") & "<p>")

if request("pOrdem") = 2 then
	str_Ordem = 2
else
	str_Ordem = 1
end if

str_SQL = ""
str_SQL = str_SQL & " SELECT dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
str_SQL = str_SQL & " , dbo.MACRO_PERFIL.MCPE_TX_NOME_TECNICO "
str_SQL = str_SQL & " , dbo.MACRO_PERFIL.MCPE_TX_DESC_MACRO_PERFIL "
str_SQL = str_SQL & " , dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL "
str_SQL = str_SQL & " , dbo.MICRO_PERFIL_R3.MIPE_TX_NOME_TECNICO "
str_SQL = str_SQL & " , dbo.MICRO_PERFIL_R3.MIPE_TX_DESC_MICRO_PERFIL "
str_SQL = str_SQL & " , dbo.MICRO_PERFIL_R3.MIPE_TX_DESC_DETALHADA "
str_SQL = str_SQL & " , dbo.MICRO_PERFIL_R3.MIPE_TX_LIBERADO_MANUT_PERF "
str_SQL = str_SQL & " FROM dbo.MICRO_PERFIL_R3 INNER JOIN "
str_SQL = str_SQL & " dbo.MACRO_PERFIL ON dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL "
str_SQL = str_SQL & " WHERE dbo.MACRO_PERFIL.MCPE_TX_NOME_TECNICO <> ''"
if str_MacroPerfil <> "" then
   str_SQL = str_SQL & " and dbo.MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = '" &  str_MacroPerfil  & "'"
else
	if str_NomeTecnico <> "" then
		str_SQL = str_SQL & " and dbo.MACRO_PERFIL.MCPE_TX_NOME_TECNICO = '" &  str_NomeTecnico  & "'"
	end if
end if
'str_SQL = str_SQL & " order by dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL"
str_SQL = str_SQL & " order by dbo.MICRO_PERFIL_R3.MIPE_TX_DESC_MICRO_PERFIL"
				
set rs=db.execute(str_SQL)


%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
.style2 {font-size: 12px}
</style>
</head>

<script>
function manda()
{
document.frm1.submit();
}
 
function CheckCheckAll()
	{
		TotalOn = 0;
		TotalBoxes = 0;
			
		for (var i=0;i<document.frm1.elements.length;i++)
		{
			var e = document.frm1.elements[i];
//			if ((e.name != 'chk_ListaSolicitacao') && (e.type=='checkbox'))
			if (e.type=='checkbox')

			{
				TotalBoxes++; 
				Check(e);
			}
				
			if (e.checked)
			{
				TotalOn++;
			}
		}
			
	}

function ClearCheckCheckAll()
	{
		TotalOn = 0;
		TotalBoxes = 0;
			
		for (var i=0;i<document.frm1.elements.length;i++)
		{
			var e = document.frm1.elements[i];
//			if ((e.name != 'chk_ListaSolicitacao') && (e.type=='checkbox'))
			if (e.type=='checkbox')

			{
				TotalBoxes++; 
				Clear(e);
			}
				
			if (e.checked)
			{
				TotalOn++;
			}
		}
			
	}
    
</script>
<script type="text/javascript">

    function ShowContextHelp(num, width, height)
    {
        var url = 'http://help.yahoo.com/help/us/mail/context/context-';
        url = url + num + '.html';

        document.domain="yahoo.com";

        remote = window.open(url, 'help'+num, "width="+width+",height="+height +",resizable=yes,scrollbars=no,status=0");
                        
        if (remote != null)
        {
          if (remote.opener == null)
          remote.opener = self;
        }
    }

    function ShowRDContextHelp(rdurl, num, width, height)
    {
        var helpurl = 'http://help.yahoo.com/help/us/mail/context/context-' + num + '.html';

        var url = rdurl + helpurl;
        

        document.domain="yahoo.com";

        remote = window.open(url, 'help'+num, "width="+width+",height="+height +",resizable=yes,scrollbars=no,status=0");
                        
        if (remote != null)
        {
          if (remote.opener == null)
          remote.opener = self;
        }
    }

    function Toggle(e)
    {
	if (e.checked) {
	    Highlight(e);
	    document.messageList.toggleAll.checked = AllChecked();
	}
	else {
	    Unhighlight(e);
	    document.messageList.toggleAll.checked = false;
	}
    }

    function ToggleAll(e)
    {
	if (e.checked) {
	    CheckAll();
	}
	else {
	    ClearAll();
	}
    }

    function Check(e)
    {
	e.checked = true;
	Highlight(e);
    }

    function Clear(e)
    {
	e.checked = false;
	Unhighlight(e);
    }

    function CheckAll()
    {
	var ml = document.messageList;
	var len = ml.elements.length;
	for (var i = 0; i < len; i++) {
	    var e = ml.elements[i];
	    if (e.name == "Mid") {
		Check(e);
	    }
	}
	ml.toggleAll.checked = true;
    }

    function ClearAll()
    {
	var ml = document.messageList;
	var len = ml.elements.length;
	for (var i = 0; i < len; i++) {
	    var e = ml.elements[i];
	    if (e.name == "Mid") {
		Clear(e);
	    }
	}
	ml.toggleAll.checked = false;
    }

    function Highlight(e)
    {
	var r = null;
	if (e.parentNode && e.parentNode.parentNode) {
	    r = e.parentNode.parentNode;
	}
	else if (e.parentElement && e.parentElement.parentElement) {
	    r = e.parentElement.parentElement;
	}
	if (r) {
	    if (r.className == "msgnew") {
		r.className = "msgnews";
	    }
	    else if (r.className == "msgold") {
		r.className = "msgolds";
	    }
	}
    }

    function Unhighlight(e)
    {
	var r = null;
	if (e.parentNode && e.parentNode.parentNode) {
	    r = e.parentNode.parentNode;
	}
	else if (e.parentElement && e.parentElement.parentElement) {
	    r = e.parentElement.parentElement;
	}
	if (r) {
	    if (r.className == "msgnews") {
		r.className = "msgnew";
	    }
	    else if (r.className == "msgolds") {
		r.className = "msgold";
	    }
	}
    }

    function AllChecked()
    {
	ml = document.messageList;
	len = ml.elements.length;
	for(var i = 0 ; i < len ; i++) {
	    if (ml.elements[i].name == "Mid" && !ml.elements[i].checked) {
		return false;
	    }
	}
	return true;
    }

    var noDelAllMsgWarning = false;
    function Delete()
    {
	if (!noDelAllMsgWarning && AllChecked()) {
	    if (!confirm("Tem certeza de que deseja apagar todas as mensagens?")) {
		return;
	    }
	}
	var ml=document.messageList;
	ml.DEL.value = "1"; 
	ml.submit();
    }
    function SynchMoves(which) {
	var ml=document.messageList;
	if(which==1) {
	    ml.destBox2.selectedIndex=ml.destBox.selectedIndex;
	}
	else {
	    ml.destBox.selectedIndex=ml.destBox2.selectedIndex;
	}
    }

    function Move() {
	var ml = document.messageList;
	var dbox = ml.destBox;
	if(dbox.options[dbox.selectedIndex].value == "@NEW") {
	    nn = window.prompt("Insira um nome para sua pasta.","");
	    if(nn == null || nn == "null" || nn == "") {
		dbox.selectedIndex = 0;
		ml.destBox2.selectedIndex = 0;
	    }
	    else {
		ml.NewFol.value = nn;
		ml.MOV.value = "1";
		ml.submit();
	    }
	}
	else {
	    ml.MOV.value = "1";
	    ml.submit();
	}
    }
    function SynchFlags(which)
    {
	var ml=document.messageList;
	if (which == 1) {
	    ml.flags2.selectedIndex = ml.flags.selectedIndex;
	}
	else {
	    ml.flags.selectedIndex = ml.flags2.selectedIndex;
	}
    }

    function SetFlags()
    {
	var ml = document.messageList;
	ml.FLG.value = "1";
	ml.submit();
    }


   function markSpam() {
        var ml = document.messageList;
        ml.FLG.value = "1";
        ml.action += "&flags=spam";
        ml.submit();
   }

   function markAdd() {
        var ml = document.messageList;
        ml.FLG.value = "1";
        ml.action += "&flags=add";
        ml.submit();
   }

</script>

<body topmargin="0" leftmargin="0" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<form method="POST" action="../macroperfil/valida_micro_libera_mapeamento.asp" name="frm1">          
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"><a href="javascript:manda()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="32%">&nbsp;</td>
      <td width="48%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="32%">&nbsp;</td>
      <td width="48%"> 
        <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Libera Perfil para mapeamento </font></td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="32%">&nbsp;</td>
      <td width="48%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
	<% if not rs.EOF then %>
    <tr>
      <td>&nbsp;</td>
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Macro Perfil : <%=rs("MCPE_TX_NOME_TECNICO")%>
          <input name="hidCdMacro" type="hidden" id="hidCdMacro" value="<%=rs("MCPR_NR_SEQ_MACRO_PERFIL")%>">
      </font></div></td>
      <td>&nbsp;</td>
    </tr>
	<% else %>
    <tr>
      <td>&nbsp;</td>
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">N&atilde;o existem micro perfil associado a este macro perfil </font></div></td>
      <td>&nbsp;</td>
    </tr>
	<% end if %>
    <tr>
      <td><table width="234" border="0" align="right" cellpadding="0" cellspacing="0">
        <tr>
          <td width="100"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:CheckCheckAll();">Marca tudo</a></font></div></td>
          <td width="134"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <a href="javascript:ClearCheckCheckAll();">Desmarca tudo</a> </font></div></td>
        </tr>
      </table></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
<p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
  <input type="hidden" name="TSql" size="127" value="<%=str_SQL%>">
<table border="0" width="85%" cellspacing="0" cellpadding="2">
  <tr>
    <td width="10%"></td>
    <td width="3%" bgcolor="#000080"></td>
    <td width="14%" bgcolor="#000080"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Nome T&eacute;cnico </font></b></td>
    <td width="73%" bgcolor="#000080"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="2">Descri&ccedil;&atilde;o</font></b></td>
  </tr>
  <% int_Tot_Registro = 0
  do until rs.eof=true
  if rs("MIPE_TX_LIBERADO_MANUT_PERF")="1" then
  	valor="checked"
  else
  	valor=""
  end if  
  if cor="#E4E4E4" then
  	cor="white"
  else
  	cor="#E4E4E4"
  end if  
  %>
  <tr>
    <td width="10%" height="33"></td>
    <td width="3%" height="33" bgcolor="<%=cor%>">
      <p align="center"><font size="1"><input type="checkbox" name="func_<%=rs("MIPE_NR_SEQ_MICRO_PERFIL")%>" value="1" <%=valor%>></font></td>
    
    <td width="14%" bgcolor="<%=cor%>"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b><%=rs("MIPE_TX_NOME_TECNICO")%></b></font></td>
    <td width="73%" height="33" bgcolor="<%=cor%>"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b></b> <%=rs("MIPE_TX_DESC_MICRO_PERFIL")%></font></td>
    </tr>
    <%
    rs.movenext
	int_Tot_Registro = int_Tot_Registro + 1
    loop
    %>
</table>
<p align="center"> <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total de Registros</font><span class="style2">:</span> <%=int_Tot_Registro%></p>
</form>
</body>
</html>
