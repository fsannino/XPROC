<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../../asp/protege/protege.asp" -->
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso
Dim str_Cenario

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Cenario = 0

if request("Situacao1")="true" THEN
	checado01="checked"
	compl_01=compl_01+" OR CENA_TX_SITUACAO_VALIDACAO='0'"
END IF

if request("Situacao2")="true" THEN
	checado02="checked"
	compl_01=compl_01+" OR CENA_TX_SITUACAO_VALIDACAO='1'"
END IF

if request("Situacao3")="true" THEN
	checado03="checked"
	compl_01=compl_01+" OR CENA_TX_SITUACAO_VALIDACAO='2'"
END IF

if len(trim(compl_01))>0 then
	compl_01 = right(compl_01,((len(compl_01))-4))
	compl_01="(" + compl_01 + ")"
end if

'response.write compl_01

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

str_Cenario = request("ID")

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

str_SQL_Cenario = ""
str_SQL_Cenario = str_SQL_Cenario & " SELECT "
str_SQL_Cenario = str_SQL_Cenario & " *"
str_SQL_Cenario = str_SQL_Cenario & " FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_TX_SITUACAO_VALIDACAO<>'3'"

if str_MegaProcesso<>0 then
	SSQL1 = SSQL1 & " AND " & Session("PREFIXO") & "CENARIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
end if
if str_Processo<>0 then
	SSQL1 = SSQL1 & " AND " & Session("PREFIXO") & "CENARIO.PROC_CD_PROCESSO = " & str_Processo 
end if
if str_SubProcesso<>0 then
	SSQL1 = SSQL1 & " AND " & Session("PREFIXO") & "CENARIO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso 
end if

if len(trim(compl_01))>0 then
	SSQL1 = SSQL1 & " AND " + compl_01
end if

if len(trim(ssql1))>0 then
	Str_SQL_Cenario = str_SQL_Cenario + SSQL1
end if

if len(trim(ssql1))=0 and len(trim(compl_01))=0 then
	Str_SQL_Cenario = str_SQL_Cenario + " AND CENARIO.MEPR_CD_MEGA_PROCESSO = 999"
	inicial=1
end if

str_SQL_Cenario = str_SQL_Cenario & " order by " & Session("PREFIXO") & "CENARIO.CENA_CD_CENARIO "

'RESPONSE.WRITE str_SQL_Cenario

set rs=conn_db.execute(str_SQL_Cenario)
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
window.location.href = "altera_escopo.asp?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"&Situacao1="+document.frm1.sit01.checked+"&Situacao2="+document.frm1.sit02.checked
}

function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location.href='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location.href='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location.href='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location.href='"+args[i+1]+"?txtOpc=3&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0'");
}
function MM_goToURL5() { //v3.0
  var i, args=MM_goToURL5.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location.href='"+args[i+1]+"?txtOpc=3&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}
function MM_goToURL6() { //v3.0
  var i, args=MM_goToURL6.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location.href='"+args[i+1]+"?txtOpc=3&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}

function Confirma2() 
{ 
	document.frm1.submit();
}

function Confirma() 
{
	document.frm1.submit();
}

function Limpa()
{
document.frm1.reset();
}

function muda_text()
{
document.frm1.atual.value=document.frm1.selEscopo.value;
}

function MM_swapImgRestore() 
{ //v3.0
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

<script>
</script>

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
<form name="frm1" method="post" action="valida_altera_escopo.asp">
  <input type="hidden" name="INC" size="20" value="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><input type="hidden" name="Atual" size="9" value="<%=valor_atual%>"></font>
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif" width="30" height="30"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif" width="30" height="30"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0" width="19" height="20"></a>&nbsp;</div>
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
          <td width="50"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Enviar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="751" border="0" cellpadding="2" cellspacing="0" name="tblSubProcesso" height="243">
    <tr>
      <td width="1" height="21"></td>
      <td width="1" height="21"></td>
      <td width="179" height="21"></td>
      <td width="546" height="21" colspan="7"> 
      </td>
    </tr>
    <tr>
      <td width="1" height="83"></td>
      <td width="1" height="83"></td>
      <td width="179" height="83"></td>
      <td width="546" height="83" colspan="7"> 
        <font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="3">Validação
        de Escopo de Cenário&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        </font><img border="0" src="../../imagens/preloader.gif" name="loader" align="absmiddle" width="190" height="50">
      </td>
    </tr>
    <tr>
      <td width="1" height="21" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="1" height="21" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="179" height="21" bgcolor="#FFFFE1">&nbsp;</td>
      <td width="546" height="21" bgcolor="#FFFFE1" colspan="7"> 
        <input type="hidden" name="txtOpc" value="1">
        <input type="hidden" name="SQL" size="78" value="<%=str_SQL_Cenario%>">
        &nbsp;
      </td>
    </tr>
    <tr> 
      <td width="1" height="25" bgcolor="#FFFFFF"> 
        &nbsp;
      </td>
      <td width="1" height="25" bgcolor="#FFFFFF"> 
        &nbsp;
      </td>
      <td width="179" height="25" bgcolor="#FFFFE1"> 
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Mega-Processo</font></b></font></div>
      </td>
      <td width="546" height="25" bgcolor="#FFFFE1" colspan="7"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','altera_escopo.asp');return document.MM_returnValue">
          <% 
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
    </tr>
    <tr> 
      <td width="1" height="25" bgcolor="#FFFFFF"> 
        &nbsp;
      </td>
      <td width="1" height="25" bgcolor="#FFFFFF"> 
        &nbsp;
      </td>
      <td width="179" height="25" bgcolor="#FFFFE1"> 
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Processo</font></b></font></div>
      </td>
      <td width="546" height="25" bgcolor="#FFFFE1" colspan="7"> 
        <select name="selProcesso" onChange="MM_goToURL2('self','altera_escopo.asp',this);return document.MM_returnValue">
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
        </select>
      </td>
    </tr>
    <tr> 
      <td width="1" height="25" bgcolor="#FFFFFF"> 
        &nbsp;
      </td>
      <td width="1" height="25" bgcolor="#FFFFFF"> 
        &nbsp;
      </td>
      <td width="179" height="25" bgcolor="#FFFFE1"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Sub-Processo</font></b></div>
      </td>
      <td width="546" height="25" bgcolor="#FFFFE1" colspan="7"> 
        <select name="selSubProcesso" onChange="MM_goToURL3('self','altera_escopo.asp',this);return document.MM_returnValue">
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
        </select>
      </td>
    </tr>
    <tr> 
      <td height="15" bgcolor="#FFFFFF" width="1">&nbsp;</td>
      <td height="15" bgcolor="#FFFFFF" width="1">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="179">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="13">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="9">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="101" align="left">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="34" align="left">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="142" align="left">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="24" align="left">&nbsp;</td>
      <td height="15" bgcolor="#FFFFE1" width="187" align="left">&nbsp;</td>
    </tr>
    <tr> 
      <td height="21" bgcolor="#FFFFFF" width="1">&nbsp;</td>
      <td height="21" bgcolor="#FFFFFF" width="1">&nbsp;</td>
      <td height="21" bgcolor="#FFFFE1" width="179">
        <p align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Situação</font></b></td>
      <td height="21" bgcolor="#FFFFE1" width="13">
        <p align="center">&nbsp;</td>
      <td height="21" bgcolor="#FFFFE1" width="9">
       <p align="center">
       <input type="checkbox" name="sit01" value="1" OnClick="manda()" <%=checado01%>>
       </p>
      </td>
      <td height="21" bgcolor="#FFFFE1" width="101" align="left">
        <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1">Fora
        do Escopo</font></td>
      <td height="21" bgcolor="#FFFFE1" width="34" align="left">
        <p align="center">
       <input type="checkbox" name="sit02" value="2" OnClick="manda()" <%=checado02%>>
       </td>
      <td height="21" bgcolor="#FFFFE1" width="142" align="left">
        <p align="left"><font face="Verdana, Arial, Helvetica, sans-serif" color="#330099" size="1">Dentro
        do
        Escopo</font></td>
      <td height="21" bgcolor="#FFFFE1" width="24" align="left">
        <p align="right">
       &nbsp;
        </td>
      <td height="21" bgcolor="#FFFFE1" width="187" align="left">
        <p align="left">&nbsp;</td>
    </tr>
    <tr> 
      <td height="21" bgcolor="#FFFFFF" width="1">&nbsp;</td>
      <td height="21" bgcolor="#FFFFFF" width="1">&nbsp;</td>
      <td height="21" bgcolor="#FFFFE1" width="106">&nbsp;</td>
      <td colspan="3" height="21" bgcolor="#FFFFE1" width="222">&nbsp;</td>
      <td colspan="4" height="21" bgcolor="#FFFFE1" width="403">&nbsp;</td>
    </tr>
  </table>
  &nbsp;
  <table border="0" width="740">
    <%
    tem=0
    do until rs.eof=true
    
    valor1=""
    valor2=""
    valor3=""
        
    select case rs("CENA_TX_SITUACAO_VALIDACAO")
    case "0"
    	valor1="checked"
    case "1"
    	valor2="checked"
    case "2"
       valor3="checked" 
    end select
    %>
    <tr>
      <td width="98" bgcolor="#330099" height="23" align="center">
        <p align="center"><font size="1" face="Verdana" color="#FFFFFF"><b>Código do Cenário</b></font></p>
      </td>
      <td width="299" bgcolor="#330099" height="23"><font size="1" face="Verdana" color="#FFFFFF"><b>Descrição</b></font></td>
      <td width="59" align="center" bgcolor="#330099" height="23"><font size="1" face="Verdana" color="#FFFFFF"><b>Fora do Escopo</b></font></td>
      <td width="70" align="center" bgcolor="#330099" height="23">
        <p align="center"><font size="1" face="Verdana" color="#FFFFFF"><b>Dentro
        do Escopo</b></font></p>
      </td>
      <td width="70" align="center" bgcolor="#FFFFFF" height="23">&nbsp;</td>
    </tr>
    <tr>
      <td width="98" height="28" align="center"><font size="1" face="Verdana"><b><%=rs("CENA_CD_CENARIO")%></b></font></td>
      <td width="299" height="28"><font size="1" face="Verdana"><%=rs("CENA_TX_TITULO_CENARIO")%></font></td>
      <%if valor3="" then%>
      <td width="59" align="center" height="28" bgcolor="#E6E6E6"><input type="radio" value="0" name="situacao_<%=rs("CENA_CD_CENARIO")%>" <%=valor1%>></td>
	   <%else%>
      <td width="82" align="center" height="28" bgcolor="#E6E6E6"></td>
      <%end if%>
      <td width="57" align="center" height="28" bgcolor="#E6E6E6"><input type="radio" value="1" name="situacao_<%=rs("CENA_CD_CENARIO")%>" <%=valor2%>></td>
      <%if valor1="" then%><%else%>
      <td width="69" align="center" height="28" bgcolor="#E6E6E6"></td>
      <%end if%>
      <td width="20" align="center" height="28" bgcolor="#FFFFFF" bordercolor="#FFFFFF"><img border="0" src="../../imagens/b04.gif" alt="Clique aqui para visualizar o histórico de Escopo" width="16" height="16"></td>
    </tr>
    <tr>
      <td width="98" valign="middle" align="center" bgcolor="#FFFFFF" height="36"><font size="1" face="Verdana" color="#330099"><b>Comentários</b></font></td>
      <td width="567" colspan="4" rowspan="2" valign="top" height="47">
        <p align="center"><textarea rows="2" name="coment_<%=rs("CENA_CD_CENARIO")%>" cols="72" OnChange="document.frm1.comentario_<%=tem%>.value=this.value"></textarea></p>
      </td>
    </tr>
    <tr>
      <td width="98" valign="top" height="7" align="center"></td>
    </tr>
    <tr>
      <td width="98" height="30" align="center"><b>&nbsp;</b></td>
      <td width="567" colspan="4" height="30">&nbsp;</td>
      
    </tr>
     <%
     tem=tem+1
	rs.movenext
    loop
    %>
  </table>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
  <%if tem=0 and inicial<>1 then %>
  <b><font color="#800000">Nenhum Registro Encontrado para a Seleção!</font></b>
  <%else
  if inicial=1 then
  %>
  &nbsp;
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <font color="#0000FF" face="Verdana" size="2">
  <b>			Efetue a Seleção Desejada nas opções acima</b>
  </font>
  <%
  end if
  end if%>
  </p>
  </form>
<p>&nbsp;</p>
</body>

<script>
MM_swapImage('loader','','../../Flash/branco.gif',1);
</script>

</html>
