<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../../asp/protege/protege.asp" -->
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso
Dim str_Cenario

str_MegaFuncao=0
str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Cenario = 0

str_Opc = Request("txtOpc")

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if

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

if str_MegaProcesso <> "0" then
   Session("MegaProcesso") = str_MegaProcesso
else
    if Session("MegaProcesso") <> "" then
       str_MegaProcesso = Session("MegaProcesso") 
	end if   
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc1 = ""
str_SQL_MegaProc1 = str_SQL_MegaProc1 & " SELECT * "
str_SQL_MegaProc1 = str_SQL_MegaProc1 & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO order by MEPR_TX_DESC_MEGA_PROCESSO"

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

str_SQL_Proc = ""
str_SQL_Proc = str_SQL_Proc & " SELECT "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " ," & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " ," & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " FROM " & Session("PREFIXO") & "PROCESSO INNER JOIN "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO ON "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
'str_SQL_Proc = str_SQL_Proc & " WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 
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

str_SQL_Cenario = ""
str_SQL_Cenario = str_SQL_Cenario & " SELECT "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.CENA_CD_CENARIO, " & Session("PREFIXO") & "CENARIO.CENA_TX_TITULO_CENARIO"
str_SQL_Cenario = str_SQL_Cenario & " FROM " & Session("PREFIXO") & "CENARIO"
str_SQL_Cenario = str_SQL_Cenario & " WHERE " & Session("PREFIXO") & "CENARIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Cenario = str_SQL_Cenario & " AND " & Session("PREFIXO") & "CENARIO.PROC_CD_PROCESSO = " & str_Processo 
str_SQL_Cenario = str_SQL_Cenario & " AND " & Session("PREFIXO") & "CENARIO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso 
str_SQL_Cenario = str_SQL_Cenario & " order by " & Session("PREFIXO") & "CENARIO.CENA_TX_TITULO_CENARIO "

str_SQL_Cenario_Tot = ""
str_SQL_Cenario_Tot = str_SQL_Cenario_Tot & " SELECT "
str_SQL_Cenario_Tot = str_SQL_Cenario_Tot & " " & Session("PREFIXO") & "CENA_CD_CENARIO, " & Session("PREFIXO") & "CENA_TX_TITULO_CENARIO"
str_SQL_Cenario_Tot = str_SQL_Cenario_Tot & " FROM " & Session("PREFIXO") & "CENARIO"
str_SQL_Cenario_Tot = str_SQL_Cenario_Tot & " order by " & Session("PREFIXO") & "CENA_TX_TITULO_CENARIO "

str_megafuncao=request("selMegaFuncao")
'response.write str_megafuncao
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
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaFuncao="+document.frm1.selMegaFuncao.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0'");
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

function Confirma2() 
{ 
	  document.frm1.submit();
}
function Confirma() 
{
	  document.frm1.submit();
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
<form name="frm1" method="post" action="gera_rel_func_confl.asp">
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
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27">&nbsp;</td>
            <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="94%" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso" height="139">
    <tr> 
      <td width="19%" height="21"></td>
      <td width="52%" height="21"> <p align="center"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="3">Relatório 
          de Função Conflitante</font></td>
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
        o Mega-Processo</b></font></td>
      <td width="12%" height="21"> </td>
      <td width="17%" height="21"> </td>
    </tr>
    <tr> 
      <td width="19%" height="25"></td>
      <td width="52%" height="25"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMega" id="selMega">
          <option value="0" selected >== TODOS ==</option>
          <%
          Set rsMega1 = Conn_db.Execute(str_SQL_MegaProc1)
			do until rsmega1.eof=true
          if Trim(str_megafuncao) = Trim(rsMega1("MEPR_CD_MEGA_PROCESSO")) then
          %>
          <option value="<%=(rsMega1("MEPR_CD_MEGA_PROCESSO"))%>"><%=(rsMega1("MEPR_TX_DESC_MEGA_PROCESSO"))%></option>
          <% else %>
          <option value="<%=(rsMega1("MEPR_CD_MEGA_PROCESSO"))%>"><%=(rsMega1("MEPR_TX_DESC_MEGA_PROCESSO"))%></option>
          <% end if
   			rsMega1.MoveNext
			loop
			%>
        </select>
        </font> </td>
      <td width="12%" height="25"> </td>
      <td width="17%" height="25"> </td>
    </tr>
    <tr> 
      <td width="19%" height="1"><div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Uso 
          :</font></b></font></div></td>
      <td width="52%" height="1"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        uso </font></b></font> <font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmUso" type="checkbox" OnClick="manda1()" value="1" checked>
        </b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        desuso </font></b></font><font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmDesuso" type="checkbox" value="1"  OnClick="manda1()" <%=checado02%>>
        </b></font> </td>
      <td width="12%" height="1"> </td>
      <td width="17%" height="1"> </td>
    </tr>
    <tr>
      <td height="1"></td>
      <td height="1"></td>
      <td height="1"></td>
      <td height="1"></td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>
