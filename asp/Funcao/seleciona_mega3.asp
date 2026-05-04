<%@LANGUAGE="VBSCRIPT"%> 
 
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso

str_MegaFuncao=0
str_SubModulo = 0
str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"

str_Opc = Request("txtOpc")

str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")
str_usoDesuso = ""

'response.Write("    sddfdv  ")
'response.Write("    Uso = " & str_Uso)
'response.Write("    Desuso = " & str_Desuso)

if str_Uso = "" and str_Desuso = "" then
   str_Uso = "true" 
   str_Desuso = "false"
end if   
if str_Uso = "true" and str_Desuso = "true" then
   checado01 = "checked"
   checado02 = "checked"
   str_usoDesuso =  " and (FUNE_TX_INDICA_EM_USO = '1' or FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = "false" and str_Desuso = "false" then
      checado01 = ""
      checado02 = ""   
      str_usoDesuso =  " and FUNE_TX_INDICA_EM_USO = '3' "
   else
      if str_Uso = "true" then
         checado01 = "checked"
         checado02 = ""   
         str_usoDesuso =  " and FUNE_TX_INDICA_EM_USO = '1' "
      else
         checado01 = ""
         checado02 = "checked"   
         str_usoDesuso =  " and FUNE_TX_INDICA_EM_USO = '0' "
	  end if        	     
   end if
end if

if (Request("selMegaFuncao") <> "") then 
    str_MegaFuncao = Request("selMegaFuncao")
else
    str_MegaFuncao = 0
end if
'response.Write("-10-")
'response.Write(Request("selMegaFuncao"))
'response.Write("-20-")
'response.Write(str_MegaFuncao)

if (Request("selSubModulo") <> "") then 
    str_SubModulo = Request("selSubModulo")
else
    str_SubModulo = "0"
end if
'response.Write("-sel-")
'response.Write(Request("selSubModulo"))
'response.Write("-str-")
'response.Write(str_SubModulo)

if (Request("txtSubModulo") <> "") then 
    str_txt_SubModulo = Request("txtSubModulo")
else
    str_txt_SubModulo = "0"
end if
'response.Write("-txt-")
'response.Write(Request("txtSubModulo"))
'response.Write("-txt-txt-")
'response.Write(str_txt_SubModulo)

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if
'response.Write(" mega ")
'response.Write(str_MegaProcesso)
if (Request("selProcesso") <> "") then 
    str_Processo = Request("selProcesso")
else
    str_Processo = "0"
end if

if (Request("selSubProcesso") <> "") then 
    str_SubProcesso = Request("selSubProcesso")
else
    str_SubProcesso = "0"
end if

if (Request("selSubProcesso") <> "") then 
    str_SubProcesso = Request("selSubProcesso")
else
    str_SubProcesso = "0"
end if

'if str_MegaProcesso <> "0" then
'   Session("MegaProcesso") = str_MegaProcesso
'else
'    if Session("MegaProcesso") <> "" then
'       str_MegaProcesso = Session("MegaProcesso") 
'	end if   
'end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

'response.write Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc1 = ""
str_SQL_MegaProc1 = str_SQL_MegaProc1 & " SELECT * "
str_SQL_MegaProc1 = str_SQL_MegaProc1 & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO order by MEPR_TX_DESC_MEGA_PROCESSO"

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
'str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

str_SQL_Proc = ""
str_SQL_Proc = str_SQL_Proc & " SELECT "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " ," & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " ," & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " FROM " & Session("PREFIXO") & "PROCESSO INNER JOIN "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO ON "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Proc = str_SQL_Proc & " WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 
str_SQL_Proc = str_SQL_Proc & " order by  " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO "

str_SQL_Sub_Proc = ""
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " SELECT "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO"
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " FROM "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN"
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "PROCESSO ON "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " INNER JOIN "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO ON "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO"
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " WHERE "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " order by  " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO "

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " MEPR_CD_MEGA_PROCESSO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaFuncao
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_CD_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaFuncao,2) & "%'" 
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "

'response.write str_Sub_Modulo

set rs_SubModulo=conn_db.execute(str_Sub_Modulo)

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function manda()
{
//document.frm1.txtSubModulo.value = document.frm1.selSubModulo.value
//alert(document.frm1.txtSubModulo.value)
//document.frm1.txtSubModulo.value = document.frm1.selSubModulo.value
//alert(document.frm1.selMegaFuncao.value)
//alert(document.frm1.selSubModulo.value)
//alert(document.frm1.txtSubModulo.value)
//alert(document.frm1.txtOpc.value)
//alert('Passei')
//document.frm1.txtSubModulo.value = document.frm1.selSubModulo.value
window.location.href='seleciona_mega3.asp?selMegaFuncao='+document.frm1.selMegaFuncao.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&txtSubModulo='+document.frm1.txtSubModulo.value+'&pOPC='+document.frm1.txtOpc.value+'&chkEmUso='+document.frm1.chkEmUso.checked+'&chkEmDesuso='+document.frm1.chkEmDesuso.checked
}

function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaFuncao="+document.frm1.selMegaFuncao.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0&selSubModulo="+document.frm1.selSubModulo.value+"'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaFuncao="+document.frm1.selMegaFuncao.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaFuncao="+document.frm1.selMegaFuncao.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc=3&&selMegaFuncao="+document.frm1.selMegaFuncao.value+"selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0'");
}
function MM_goToURL5() { //v3.0
  var i, args=MM_goToURL5.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc=3&&selMegaFuncao="+document.frm1.selMegaFuncao.value+"selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}

function Confirma() 
{
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("Você deve selecionar um MEGA-PROCESSO para as funções de negócio");
     document.frm1.selMegaProcesso.focus();
     return;
     }

if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("Você deve selecionar um MEGA-PROCESSO para as transações");
     document.frm1.selMegaProcesso.focus();
     return;
     }
	 else
     {
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../imagens/novo_registro_02.gif','../imagens/atualiza_02.gif')">
<form name="frm1" method="post" action="rel_func_transacao.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="207"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Enviar
            Consulta</font></b></font></td>
          <td width="81"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font>
            <p align="right"></td>
          <td width="273"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="94%" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso" height="356">
    <tr> 
      <td width="19%" height="21"></td>
      <td width="52%" height="21"> <p align="center"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="3">Relatório 
          Função x Transação</font> </td>
      <td width="12%" height="21"> </td>
      <td width="17%" height="21"> </td>
    </tr>
    <tr> 
      <td width="19%" height="21"></td>
      <td width="52%" height="21"> </td>
      <td width="12%" height="21"> </td>
      <td width="17%" height="21"> </td>
    </tr>
    <tr> 
      <td width="19%" height="21"></td>
      <td width="52%" height="21"> <font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Selecione 
        o Mega-Processo das Funções de Negócio</b></font> </td>
      <td width="12%" height="21"> </td>
      <td width="17%" height="21"> </td>
    </tr>
    <tr> 
      <td width="19%" height="25">&nbsp; </td>
      <td width="52%" height="25"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaFuncao" onChange="javascript:manda()">
          <option value="0" >Selecione um Mega Processo</option>
          <%
          Set rsMega1 = Conn_db.Execute(str_SQL_MegaProc1)
			do until rsmega1.eof=true
          if Trim(str_megafuncao) = Trim(rsMega1("MEPR_CD_MEGA_PROCESSO")) then
          %>
          <option selected value="<%=(rsMega1("MEPR_CD_MEGA_PROCESSO"))%>"><%=(rsMega1("MEPR_TX_DESC_MEGA_PROCESSO"))%></option>
          <% else %>
          <option value="<%=(rsMega1("MEPR_CD_MEGA_PROCESSO"))%>"><%=(rsMega1("MEPR_TX_DESC_MEGA_PROCESSO"))%></option>
          <% end if
   			rsMega1.MoveNext
			loop
			%>
        </select>
        <input type="hidden" name="txtSubModulo" value="<%=str_txt_SubModulo%>">
        </font> </td>
      <td width="12%" height="25"> </td>
      <td width="17%" height="25"> </td>
    </tr>
    <tr> 
      <td width="19%" height="1"></td>
      <td width="52%" height="1"> </td>
      <td width="12%" height="1"> </td>
      <td width="17%" height="1"> </td>
    </tr>
    <% 'response.write str_MegaFuncao
	   'if InStrRev("11/10", Right("00" & str_MegaFuncao, 2)) = 0 then
	%>
    <tr> 
      <td height="7"><input type="hidden" name="selSubModulo22" value="0"></td>
      <td height="15"></td>
      <td height="7">&nbsp;</td>
      <td height="7">&nbsp;</td>
    </tr>
    <% 'else %>
    <tr> 
      <td height="7"><div align="right"><font face="Verdana" size="2" color="#330099"><b>Assunto 
          :</b></font></div></td>
      <td height="15"> <select size="1" name="selSubModulo" onChange="javascript:window.location.href='seleciona_mega3.asp?selMegaFuncao='+document.frm1.selMegaFuncao.value+'&selSubModulo='+this.value+'&txtSubModulo='+document.frm1.txtSubModulo.value+'&pOPC='+document.frm1.txtOpc.value+'&chkEmUso='+document.frm1.chkEmUso.checked+'&chkEmDesuso='+document.frm1.chkEmDesuso.checked">
          <option value="0">== Selecione o Assunto ==</option>
          <%do until rs_SubModulo.eof=true
		  if trim(str_SubModulo)=trim(rs_SubModulo("SUMO_NR_CD_SEQUENCIA")) then
		  %>
          <option selected value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <% else %>
          <option value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
		     end if
					rs_SubModulo.movenext
					loop
					%>
        </select> </td>
      <td height="7">&nbsp;</td>
      <td height="7">&nbsp;</td>
    </tr>
    <% 'end if %>
    <tr> 
      <td width="19%" height="7">&nbsp;</td>
      <td width="52%" height="15" bgcolor="#330099"> <p align="center"> 
          <input type="hidden" name="txtOpc" value="1">
          <b><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="2">Filtro 
          de Transações</font> </b></p></td>
      <td width="12%" height="7"> <%'=str_Opc%></td>
      <td width="17%" height="7"> <%'=str_MegaProcesso%> <%'=str_Processo%></td>
    </tr>
    <tr> 
      <td width="19%" height="25"> <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Mega-Processo</font></b></font></div></td>
      <td width="52%" height="25"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso" onChange="javascript:window.location.href='seleciona_mega3.asp?selMegaFuncao='+document.frm1.selMegaFuncao.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&txtSubModulo='+document.frm1.txtSubModulo.value+'&pOPC='+document.frm1.txtOpc.value+'&selMegaProcesso='+this.value+'&selProcesso=0&selSubProcesso=0&chkEmUso='+document.frm1.chkEmUso.checked+'&chkEmDesuso='+document.frm1.chkEmDesuso.checked"">
          <% 
		  'response.Write(" opc  ")
		  'response.Write(str_Opc)
		  'response.Write("  mega 2 ")
		  'response.Write(Trim(str_MegaProcesso))
		  if str_Opc <> "1" then %>
          <option value="0" selected>Selecione um Mega Processo</option>
          <% else %>
          <option value="0" >Selecione um Mega Processo</option>
          <% end if %>
          <%Set rdsMegaProcesso = Conn_db.Execute(str_SQL_MegaProc)
While (NOT rdsMegaProcesso.EOF)
         if (Trim(str_MegaProcesso) = Trim(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)) then %>
          <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>" selected ><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
          <% else %>
          <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>" ><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
          <% end if %>
          <%
  rdsMegaProcesso.MoveNext()
Wend
If (rdsMegaProcesso.CursorType > 0) Then
  rdsMegaProcesso.MoveFirst
Else
  rdsMegaProcesso.Requery
End If
rdsMegaProcesso.Close
set rdsMegaProcesso = Nothing
%>
        </select>
        </font></td>
      <td width="12%" height="25">&nbsp; </td>
      <td width="17%" height="25"> <%'=str_SQL_MegaProc%> </td>
    </tr>
    <tr> 
      <td width="19%" height="25"> <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Processo</font></b></font></div></td>
      <td width="52%" height="25"> <select name="selProcesso" onChange="javascript:window.location.href='seleciona_mega3.asp?selMegaFuncao='+document.frm1.selMegaFuncao.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&txtSubModulo='+document.frm1.txtSubModulo.value+'&pOPC='+document.frm1.txtOpc.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+this.value+'&selSubProcesso=0&chkEmUso='+document.frm1.chkEmUso.checked+'&chkEmDesuso='+document.frm1.chkEmDesuso.checked"">
          <% 
		  if str_Opc <> "1" then %>
          <option value="0" selected>Selecione um Processo</option>
          <% else %>
          <option value="0" >Selecione um Processo</option>
          <% end if %>
          <%Set rdsProcesso = Conn_db.Execute(str_SQL_Proc)
While (NOT rdsProcesso.EOF)
  
           if (Trim(str_Processo) = Trim(rdsProcesso.Fields.Item("PROC_CD_PROCESSO").Value)) then %>
          <option value="<%=(rdsProcesso.Fields.Item("PROC_CD_PROCESSO").Value)%>" selected ><%=(rdsProcesso.Fields.Item("PROC_TX_DESC_PROCESSO").Value)%></option>
          <% else %>
          <option value="<%=(rdsProcesso.Fields.Item("PROC_CD_PROCESSO").Value)%>"><%=(rdsProcesso.Fields.Item("PROC_TX_DESC_PROCESSO").Value)%></option>
          <% end if %>
          <%
  rdsProcesso.MoveNext()
Wend
If (rdsProcesso.CursorType > 0) Then
  rdsProcesso.MoveFirst
Else
  rdsProcesso.Requery
End If

rdsProcesso.Close
set rdsProcesso = Nothing
%>
        </select> </td>
      <td width="12%" height="25">&nbsp; </td>
      <td width="17%" height="25"> <%'=str_SQL_Proc%> </td>
    </tr>
    <tr> 
      <td width="19%" height="25"> <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Sub-Processo</font></b></div></td>
      <td width="52%" height="25"> <select name="selSubProcesso">
          <% 
		  if str_Opc <> "1" then %>
          <option value="0" selected>Selecione um Sub Processo</option>
          <% else %>
          <option value="0" >Selecione um Sub Processo</option>
          <% end if %>
          <%Set rdsSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
While (NOT rdsSubProcesso.EOF)
           if (Trim(str_SubProcesso) = Trim(rdsSubProcesso.Fields.Item("SUPR_CD_SUB_PROCESSO").Value)) then %>
          <option value="<%=rdsSubProcesso.Fields.Item("SUPR_CD_SUB_PROCESSO").Value%>" selected ><%=(rdsSubProcesso.Fields.Item("SUPR_TX_DESC_SUB_PROCESSO").Value)%></option>
          <% else %>
          <option value="<%=rdsSubProcesso.Fields.Item("SUPR_CD_SUB_PROCESSO").Value%>" ><%=(rdsSubProcesso.Fields.Item("SUPR_TX_DESC_SUB_PROCESSO").Value)%></option>
          <% end if %>
          <%
  rdsSubProcesso.MoveNext()
Wend
If (rdsSubProcesso.CursorType > 0) Then
  rdsSubProcesso.MoveFirst
Else
  rdsSubProcesso.Requery
End If
rdsSubProcesso.close
set rdsSubProcesso = Nothing
%>
        </select> </td>
      <td width="12%" height="25"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp; 
        </font></td>
      <td width="17%" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp; 
        </font></td>
    </tr>
    <tr>
      <td height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
      <td height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        uso </font></b></font> <font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmUso" type="checkbox" value="1" <%=checado01%>>
        </b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        desuso </font></b></font><font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmDesuso" type="checkbox" value="1"  <%=checado02%>>
        </b></font> </td>
      <td height="25">&nbsp;</td>
      <td height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="19%" height="1"> </td>
      <td width="52%" height="10" bgcolor="#330099"> </td>
      <td width="12%" height="1"></td>
      <td width="17%" height="1"></td>
    </tr>
    <tr> 
      <td width="19%" height="7">&nbsp;</td>
      <td width="52%" height="7"> <p align="center"><font color="#FF0000" size="2" face="Verdana"><b>Este 
          relatório demora, em média, 3 minutos para ser gerado, dependendo do 
          filtro usado nas Transações</b></font> 
          <input type="hidden" name="INC" size="20" value="1">
        </p></td>
      <td width="12%" height="7">&nbsp;</td>
      <td width="17%" height="7">&nbsp;</td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>
