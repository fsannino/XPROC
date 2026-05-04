<%

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_Mega = request("selMegaProcesso")
str_SubModulo = request("selSubModulo") 

'response.Write(request("pOrdem") & "<p>")

if request("pOrdem") = 2 then
	str_Ordem = 2
else
	str_Ordem = 1
end if

ssql=""
ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_LIBERADO_MANUT_PERF "
ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
ssql=ssql & " WHERE "
ssql=ssql & " FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & str_mega 
IF str_SubModulo <> 0 THEN
   ssql=ssql & " and FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA =" & str_SubModulo
END IF
IF str_Ordem = 2 then
	ssql=ssql+"ORDER BY FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
else
	ssql=ssql+"ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "
end if
'RESPONSE.Write(ssql)
'RESPONSE.End()
set rs=db.execute(ssql)
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
.style1 {color: #333333}
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
<form method="POST" action="valida_func_libera_mapeamento.asp" name="frm1">          
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
        <p align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Libera Fun&ccedil;&atilde;o para mapeamento </font></td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="32%">&nbsp;</td>
      <td width="48%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
	<% if rs.EOF then %>
    <tr>
      <td>&nbsp;</td>
      <td><div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">N&atilde;o existem micro perfil associado a este macro perfil </font></div></td>
      <td>&nbsp;</td>
    </tr>
	<% end if %>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
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
<input type="hidden" name="TSql" size="127" value="<%=ssql%>">
<table border="0" width="85%" cellspacing="0" cellpadding="2" height="75">
  <tr>
    <td width="10%" height="21"></td>
    <td width="3%" bgcolor="#000080" height="21">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"></td>
    <td width="7%" bgcolor="#000080"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="2">Fun&ccedil;&atilde;o</font></b></td>
    <td width="80%" bgcolor="#000080" height="21">
      <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="2">      de Negócio</font></b></td>
  </tr>
  <tr>
    <td height="21"></td>
    <td height="21"><span class="style1"></span></td>
    <td><span class="style1"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="relaciona_func_libera_mapeamento.asp?pOrdem=1&selMegaProcesso=<%=str_Mega%>&selSubModulo=<%=str_SubModulo%>">C&oacute;digo</a> </font></b></span></td>
    <td height="21"><span class="style1"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2">ou <a href="relaciona_func_libera_mapeamento.asp?pOrdem=2&selMegaProcesso=<%=str_Mega%>&selSubModulo=<%=str_SubModulo%>">T&iacute;tulo</a> - (escolha ordena&ccedil;&atilde;o) </font></b></span></td>
  </tr>
  <% int_Tot_Registro = 0
  do until rs.eof=true
  if rs("FUNE_TX_LIBERADO_MANUT_PERF")="1" then
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
      <p align="center"><font size="1"><input type="checkbox" name="func_<%=rs("fune_cd_funcao_negocio")%>" value="1" <%=valor%>></font></td>
    
    <td width="7%" bgcolor="<%=cor%>"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b><a href="gera_rel_mega_funcao.asp?selMegaProcesso=<%=str_mega%>&selFuncao=<%=rs("fune_cd_funcao_negocio")%>"><%=rs("fune_cd_funcao_negocio")%></a></b></font></td>
    <td width="80%" height="33" bgcolor="<%=cor%>"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1"><b></b> <%=rs("fune_tx_titulo_funcao_negocio")%></font></td>
    </tr>
    <%
    rs.movenext
	int_Tot_Registro = int_Tot_Registro + 1	
    loop
    %>
</table>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total de Registros</font><span class="style2">:</span> <%=int_Tot_Registro%></p>
</form>

<p>&nbsp;</p>

</body>

</html>
