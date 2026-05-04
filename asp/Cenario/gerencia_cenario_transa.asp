<%@LANGUAGE="VBSCRIPT"%> 
 
<%

Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso
Dim str_Cenario
Dim int_Total_Registros

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Cenario = 0
int_Total_Registros = 0

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

if (Request("selCenario") <> "") then 
    str_Cenario = Request("selCenario")
else
    str_Cenario = "0"
end if

if (Request("txtCenario") <> "") then 
    str_Cenario = UCase(Request("txtCenario"))
else
    str_Cenario = UCase(Request("selCenario"))
end if

ordena=request("order")
select case ordena
	case 1
		valor="" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
	case 2
		valor="" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
	case else
		valor="" & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
end select

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQl = " SELECT "
str_SQl = str_SQl & " CENA_CD_CENARIO, MEPR_CD_MEGA_PROCESSO, "
str_SQl = str_SQl & " PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO, CENA_TX_TITULO_CENARIO "
str_SQl = str_SQl & " FROM " & Session("PREFIXO") & "CENARIO "
str_SQl = str_SQl & " WHERE (CENA_CD_CENARIO = '" & str_Cenario & "')"
Set rdsCenario2= Conn_db.Execute(str_SQl)
if not rdsCenario2.EOF then
   str_MegaProcesso = rdsCenario2("MEPR_CD_MEGA_PROCESSO")
   str_Processo = rdsCenario2("PROC_CD_PROCESSO")
   str_SubProcesso = rdsCenario2("SUPR_CD_SUB_PROCESSO")
   str_TituloCenario = rdsCenario2("CENA_TX_TITULO_CENARIO")
else
   response.redirect "msg.asp?pOpt=0&txtCenario=" & str_Cenario   
end if
rdsCenario2.Close
set rdsCenario2 = Nothing

str_SQL_Cenario = ""
str_SQL_Cenario = str_SQL_Cenario & " SELECT "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.MEPR_CD_MEGA_PROCESSO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.PROC_CD_PROCESSO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.SUPR_CD_SUB_PROCESSO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_CD_CENARIO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.CETR_NR_SEQUENCIA," 
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.CENA_TX_TITULO_CENARIO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.CENA_TX_DESC_CENARIO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.CENA_TX_SITUACAO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.DESE_CD_DESENVOLVIMENTO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CLASSE_CENARIO.CLCE_TX_DESC_CLASSE_CENARIO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.BPPP_CD_BPP, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "OPERACOES_ESPEC.OPES_TX_DESC_OPERACAO_ESP, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_NR_SEQUENCIA_TRANS, "
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.CETR_TX_DESC_TRANSACAO,"
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_ABREVIA,"
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.CETR_TX_TIPO_RELACAO,"
str_SQL_Cenario = str_SQL_Cenario & " " & Session("PREFIXO") & "CENARIO.CENA_TX_SITUACAO_VALIDACAO"

str_SQL_Cenario = str_SQL_Cenario & " FROM " & Session("PREFIXO") & "CENARIO "
str_SQL_Cenario = str_SQL_Cenario & " INNER JOIN " & Session("PREFIXO") & "CENARIO_TRANSACAO ON " & Session("PREFIXO") & "CENARIO.CENA_CD_CENARIO = " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_CD_CENARIO"
str_SQL_Cenario = str_SQL_Cenario & " INNER JOIN " & Session("PREFIXO") & "CLASSE_CENARIO ON " & Session("PREFIXO") & "CENARIO.CLCE_CD_NR_CLASSE_CENARIO = " & Session("PREFIXO") & "CLASSE_CENARIO.CLCE_CD_NR_CLASSE_CENARIO"
str_SQL_Cenario = str_SQL_Cenario & " LEFT OUTER JOIN " & Session("PREFIXO") & "MEGA_PROCESSO ON " & Session("PREFIXO") & "CENARIO_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Cenario = str_SQL_Cenario & " LEFT OUTER JOIN " & Session("PREFIXO") & "OPERACOES_ESPEC ON " & Session("PREFIXO") & "CENARIO_TRANSACAO.OPES_CD_OPERACAO_ESP = " & Session("PREFIXO") & "OPERACOES_ESPEC.OPES_CD_OPERACAO_ESP"
str_SQL_Cenario = str_SQL_Cenario & " where " & Session("PREFIXO") & "CENARIO.CENA_CD_CENARIO = '" & str_Cenario & "'"
str_SQL_Cenario = str_SQL_Cenario & " order by " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_NR_SEQUENCIA_TRANS"
'response.Write(str_SQL_Cenario)
Set rdsCenario= Conn_db.Execute(str_SQL_Cenario)

if not rdsCenario.EOF then
   str_MegaProcesso = rdsCenario("MEPR_CD_MEGA_PROCESSO")
   str_Processo = rdsCenario("PROC_CD_PROCESSO")
   str_SubProcesso = rdsCenario("SUPR_CD_SUB_PROCESSO")
   str_TituloCenario = rdsCenario("CENA_TX_TITULO_CENARIO")
   str_Status1 = rdsCenario("CENA_TX_SITUACAO")
   str_Scopo1 = rdsCenario("CENA_TX_SITUACAO_VALIDACAO")
end if


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
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
Set rdsSubProcesso= Conn_db.Execute(str_SQL_Sub_Proc)
if not rdsSubProcesso.EOF then
   ls_Desc_MegaProcesso = rdsSubProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
   ls_Desc_Processo = rdsSubProcesso("PROC_TX_DESC_PROCESSO")
   ls_Desc_SubProcesso = rdsSubProcesso("SUPR_TX_DESC_SUB_PROCESSO")   
else
   ls_Desc_MegaProcesso = "Não Encontrado"
   ls_Desc_Processo = "Não Encontrado"
   ls_Desc_SubProcesso = "Não Encontrado"
end if
rdsSubProcesso.Close
set rdsSubProcesso = Nothing
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--

function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"'");
}
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?option=1&inc=2&id="+document.frm1.txtCenario.value+"'");
}
function MM_goToURL5() { //v3.0
  var i, args=MM_goToURL5.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=3) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"&p_CenarioChSequencia="+args[3]+"'");
}
function MM_goToURL4_2() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?option=1&inc=2&id="+document.frm1.txtCenario.value+"'");
}
function MM_goToURL6() { //v3.0
  var i, args=MM_goToURL6.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?option=1&inc=3&id="+document.frm1.txtCenario.value+"'");
}
function MM_goToURL9() 
{ //v3.0
     var i, args=MM_goToURL9.arguments; document.MM_returnValue = false;

     //if ((document.frm1.txtStatus1.value == "EE")||(document.frm1.txtStatus1.value == "DF"))
       if (document.frm1.txtStatus1.value == "DF")
        { 
	    if (document.frm1.txtScopo1.value == "1")
		   {
           for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"'");
           }	 
		 else
		   {
		   alert("Para mudar a situação é necessário que este Cenário esteja no ESCOPO !");
           return;		    
		   }
		 }  
	 else
        {
        for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"'");
	    }  
}

function Confirma2() 
{ 
	  document.frm1.submit();
}
function Confirma() 
{ 
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
if (document.frm1.selProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Processo é obrigatório!");
     document.frm1.selProcesso.focus();
     return;
     }
if (document.frm1.txtNovoSubProc1.value == "")
     { 
	 alert("Preencha um novo Sub Processo.");
     document.frm1.txtNovoSubProc1.focus();
     return;
     }	 
if (document.frm1.list2.options.length == 0)
     { 
	 alert("A seleção de uma Empresa/Unidade é obrigatória !");
     document.frm1.list2.focus();
     return;
     }	 
	 else
     {
 	  carrega_txt(document.frm1.list2);	  
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
<script>
function ver_desenv(trans)
{
var a=trans;
window.open("ver_desenv.asp?transacao=" + a + "","_blank","width=700,height=300,history=0,scrollbars=1,titlebar=0,resizable=0")
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../imagens/novo_registro_02.gif')">
<form name="frm1" method="post" action="">
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
    <td colspan="3" height="20">&nbsp; </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="24%">&nbsp;</td>
      <td width="50%">
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Edi&ccedil;&atilde;o</font><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
          de Cen&aacute;rio</font></div>
      </td>
    <td width="26%"><%'=str_SQL_Cenario%></td>
  </tr>
</table>
  <table width="740" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td width="14" height="15">&nbsp;</td>
      <td width="126" height="15"> 
        <div align="right"><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Mega-Proceso 
          :&nbsp; </font></font></div>
      </td>
      <td width="400" height="15"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><b><%=str_MegaProcesso%>-<%=ls_Desc_MegaProcesso%> 
        <input type="hidden" name="txtMegaProcesso" value="<%=str_MegaProcesso%>">
        </b> </font></td>
      <td width="200" height="15"><a href="selec_Mega_Proc_Sub_Cenario.asp"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a> 
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Novo 
        Mega-Processo</font></td>
    </tr>
    <tr> 
      <td width="14" height="15">&nbsp;</td>
      <td width="126" height="15"> 
        <div align="right"><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Proceso 
          :&nbsp; </font></font></div>
      </td>
      <td width="400" height="15"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><b><%=str_Processo%>-<%=ls_Desc_Processo%> 
        <input type="hidden" name="txtProcesso" value="<%=str_Processo%>">
        </b> </font></td>
      <td width="200" height="15"><a href="selec_Mega_Proc_Sub_Cenario.asp?selMegaProcesso=<%=str_MegaProcesso%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">&nbsp;Novo 
        Processo</font></td>
    </tr>
    <tr> 
      <td width="14" height="15">&nbsp;</td>
      <td width="126" height="15"> 
        <div align="right"><font size="1"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Sub-Proceso 
          :&nbsp; </font></font></div>
      </td>
      <td width="400" height="15"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><b><%=str_SubProcesso%>-<%=ls_Desc_SubProcesso%> 
        <input type="hidden" name="txtSubProcesso" value="<%=str_SubProcesso%>">
        </b> </font></td>
      <td width="200" height="15"><a href="selec_Mega_Proc_Sub_Cenario.asp?selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">&nbsp;Novo 
        Sub-Processo</font></td>
    </tr>
    <tr> 
      <td width="14" height="15">&nbsp;</td>
      <td width="126" height="15"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Cen&aacute;rio 
          :&nbsp; </font></div>
      </td>
      <td width="400" height="15"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">&nbsp; 
        </font> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="32%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
              <b><%=str_Cenario%> 
              <input type="hidden" name="txtCenario" value="<%=str_Cenario%>">
              </b> </font></td>
            <td width="60%"> 
              <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><b> 
                <input type="hidden" name="txtStatus1" value="<%=str_Status1%>">
                </b></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
                <%'=rdsCenario("CENA_TX_SITUACAO")%>
                Status :</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
                <font size="3"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><font size="3"><font size="2"><b> 
                <%If not rdsCenario.EOF then
			       If rdsCenario("CENA_TX_SITUACAO") = "EE" Then
				      ls_Situacao_Cenario = "EM ELABORAÇÃO"
				   elseIf rdsCenario("CENA_TX_SITUACAO") = "DF" Then
				      ls_Situacao_Cenario = "DEFINIDO"
				   elseIf rdsCenario("CENA_TX_SITUACAO") = "DS" Then
				      ls_Situacao_Cenario = "DESENHADO"
				   elseIf rdsCenario("CENA_TX_SITUACAO") = "PT" Then
				      ls_Situacao_Cenario = "PRONTO PARA TESTE"
				   elseIf rdsCenario("CENA_TX_SITUACAO") = "TD" Then
				      ls_Situacao_Cenario = "TESTADO NO PED"
				   elseIf rdsCenario("CENA_TX_SITUACAO") = "TQ" Then
				      ls_Situacao_Cenario = "TESTADO NO PEQ"
				   end if				  
				   'ls_Situacao_Cenario = "EM ELABORAÇÃO"
			  %>
                </b></font></font></font><font size="1"><b><%=ls_Situacao_Cenario%></b></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><font size="3"><font size="2"><b> 				 
                <% end if %>
                &nbsp;&nbsp; </b></font></font></font></font></font></div>
            </td>
            <td width="8%"><a href="javascript:MM_goToURL9('self','alt_situacao.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1331','','../../imagens/novo_registro_02.gif',1)"><img name="Image1331" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Altera Status"></a></td>
          </tr>
        </table>
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">&nbsp; 
          </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Escopo</font> 
          : <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">
		  <%
		    if str_Scopo1 = "0" then
			   str_Ds_Scopo1 = "Fora do escopo"
            elseif str_Scopo1 = "1" then
			   str_Ds_Scopo1 = "No escopo"
            elseif str_Scopo1 = "2" then
			   str_Ds_Scopo1 = "Cancelado"
			end if   %> 
		  <b><%=str_Ds_Scopo1%>
          <input type="hidden" name="txtScopo1" value="<%=str_Scopo1%>">
          </b></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">&nbsp; 
          </font></div></td>
      <td width="200" height="15"><a href="selec_Mega_Proc_Sub_Cenario.asp?selMegaProcesso=<%=str_MegaProcesso%>&selProcesso=<%=str_Processo%>&selSubProcesso=<%=str_SubProcesso%>"><img src="../../imagens/selecao_F02.gif" width="22" height="20" alt="Seleciona novo Mega Processo" border="0"></a><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">&nbsp;Novo 
        Cen&aacute;rio</font></td>
    </tr>
    <tr> 
      <td width="14" height="15">&nbsp;</td>
      <td width="126" height="15">&nbsp;</td>
      <td width="400" height="15">&nbsp;</td>
      <td width="200" height="15">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14" height="15">&nbsp;</td>
      <td width="126" height="15">&nbsp;</td>
      <td width="400" height="15"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <font size="3"> <font size="2"> <b> <%=str_TituloCenario%> 
        <%'=str_SQL_Cenario%>
        <%If rdsCenario.EOF then%>
        <font color="#FF0000">Este cenário não possui transações relacionadas.</font> 
        <% end if %>
        </b> </font> </font> </font></td>
      <td width="200" height="15">&nbsp;</td>
    </tr>
    <tr> 
      <td width="14" bgcolor="#0066CC"></td>
      <td width="126" bgcolor="#0066CC"></td>
      <td width="400" bgcolor="#0066CC"></td>
      <td width="200" height="3" bgcolor="#0066CC"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="14"></td>
      <td width="126" height="7"></td>
      <td width="400"></td>
      <td width="200" height="3"></td>
    </tr>
  </table>
  <%'If not rdsCenario.EOF then %>
  <table width="779" border="0" cellspacing="2" cellpadding="0" align="center">
    <tr> 
      <td width="201">&nbsp;</td>
      <td width="26">&nbsp;</td>
      <td width="216">&nbsp;</td>
      <td width="27">&nbsp;</td>
      <td width="216">&nbsp;</td>
      <td width="27">&nbsp;</td>
      <td width="1">&nbsp;</td>
      <td width="23">&nbsp;</td>
    </tr>
    <tr> 
      <td width="201"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Inclus&atilde;o 
          Nova Transa&ccedil;&atilde;o&nbsp;</font></b></div>
      </td>
      <td width="26"><a href="javascript:MM_goToURL4('self','cad_cenario_transacao.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image131','','../../imagens/novo_registro_02.gif',1)"><img name="Image131" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui nova Transa&ccedil;&atilde;o"></a> 
      </td>
      <td width="216"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Inclus&atilde;o 
          Exit/Interface&nbsp;</font></b></div>
      </td>
      <td width="27"><a href="javascript:MM_goToURL3('self','inc_oper_espe_exit_interface.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1332','','../../imagens/novo_registro_02.gif',1)"><img name="Image1332" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui nova Exit/interface"></a> 
      </td>
      <td width="201"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Inclus&atilde;o 
          Ponto de Controle&nbsp;</font></b></div>
      </td>
      <td width="26"><a href="javascript:MM_goToURL6('self','cad_cenario_transacao.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1311','','../../imagens/novo_registro_02.gif',1)"><img name="Image1311" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui nova Transa&ccedil;&atilde;o"></a> 
      </td>
      <td width="1">&nbsp;</td>
      <td width="23">&nbsp;</td>
    </tr>
    <tr> 
      <td width="216"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Inclus&atilde;o 
          Opera&ccedil;&atilde;o Manual&nbsp;</font></b></div>
      </td>
      <td width="27"><a href="javascript:MM_goToURL3('self','inc_oper_espe_manual.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image133','','../../imagens/novo_registro_02.gif',1)"><img name="Image133" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui nova Opera&ccedil;&atilde;o Manual"></a> 
      </td>
      <td width="216"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Inclus&atilde;o 
          Chamada Cen&aacute;rio&nbsp;</font></b></div>
      </td>
      <td width="27"><a href="javascript:MM_goToURL3('self','inc_chamada_cenario.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1321','','../../imagens/novo_registro_02.gif',1)"><img name="Image1321" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui nova Chamada Cen&aacute;rio"></a> 
      </td>
      <td width="216"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Novo</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
          Hist&oacute;rico&nbsp;</font></b></div>
      </td>
      <td width="27"><a href="javascript:MM_goToURL3('self','inc_historico.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image13211','','../../imagens/novo_registro_02.gif',1)"><img name="Image13211" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui novo Hist&oacute;rico"></a> 
      </td>
      <td width="1">&nbsp;</td>
      <td width="23">&nbsp;</td>
    </tr>
  </table>
  <%'end if %>
  <table width="758" align="center" cellspacing="2" cellpadding="0" border="0">
    <tr> 
      <td width="59" bgcolor="#0066CC" style="color: #FFFFFF" height="16"> 
        <div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">A&ccedil;&atilde;o</font></b></div>
      </td>
      <td width="28" bgcolor="#0066CC" style="color: #FFFFFF" height="16"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Seq</font></b></td>
      <td width="30" bgcolor="#0066CC" style="color: #FFFFFF" height="16"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Meg</font></b></td>
      <td width="347" bgcolor="#0066CC" style="color: #FFFFFF" height="16"><b><font size="2"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">Desc 
        Transa&ccedil;&atilde;o </font></font></b></td>
      <td width="92" bgcolor="#0066CC" style="color: #FFFFFF" height="16"> 
        <div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">C&oacute;d 
          Transa&ccedil;&atilde;o</font></b></div>
      </td>
      <td width="52" bgcolor="#0066CC" height="16"> 
        <div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Oper 
          Esp </font></b></div>
      </td>
      <td width="79" bgcolor="#0066CC" height="16"> 
        <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Cen&aacute;rio 
          ou &nbsp; Desenvolvim.</font></b></div>
      </td>
      <td width="60" bgcolor="#0066CC" height="16"> 
        <div align="center"><b><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">BPP</font></b></div>
      </td>
    </tr>
    <%do while not rdsCenario.EOF 
  int_Total_Registros = int_Total_Registros + 1
  If rdsCenario("CETR_TX_TIPO_RELACAO") = "0" then 
     ls_Cor_Linha = "#F7F7F7"
  elseIf rdsCenario("CETR_TX_TIPO_RELACAO") = "1" then 
     ls_Cor_Linha = "#FFFFC6"
  elseIf rdsCenario("CETR_TX_TIPO_RELACAO") = "2" then 
     ls_Cor_Linha = "#C6FFC6"
  elseIf rdsCenario("CETR_TX_TIPO_RELACAO") = "3" then 
     ls_Cor_Linha = "#CCFFFF"
  elseIf rdsCenario("CETR_TX_TIPO_RELACAO") = "5" then 
     ls_Cor_Linha = "#F7F7F7"
  End if
  %>
    <tr bgcolor="<%=ls_Cor_Linha%>"> 
      <td width="59" height="22"> 
        <table width="59" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr> 
            <td width="26"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
              <% If rdsCenario("CETR_TX_TIPO_RELACAO") = "0" or rdsCenario("CETR_TX_TIPO_RELACAO") = "5"  then %>
              <a href="javascript:MM_goToURL5('self','alt_cenario_transacao.asp',this,'<%=rdsCenario("CETR_NR_SEQUENCIA")%>')">Alt</a> 
              <% elseif rdsCenario("CETR_TX_TIPO_RELACAO") = "1" then %>
              <a href="javascript:MM_goToURL5('self','alt_oper_espe_cena_tran.asp',this,'<%=rdsCenario("CETR_NR_SEQUENCIA")%>')">Alt</a> 
              <% elseif rdsCenario("CETR_TX_TIPO_RELACAO") = "2" then %>
              <a href="javascript:MM_goToURL5('self','alt_chamada_cenario.asp',this,'<%=rdsCenario("CETR_NR_SEQUENCIA")%>')">Alt</a> 
              <% elseif rdsCenario("CETR_TX_TIPO_RELACAO") = "3" then %>
              <a href="javascript:MM_goToURL5('self','alt_oper_espe_exit_interface.asp',this,'<%=rdsCenario("CETR_NR_SEQUENCIA")%>')">Alt</a> 
              <% end if %>
              </font></b></td>
            <td width="31"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:MM_goToURL5('self','exc_cenario_transacao.asp',this,'<%=rdsCenario("CETR_NR_SEQUENCIA")%>')">Exc</a></font></b></td>
          </tr>
        </table>
      </td>
      <%IF Not IsNull(rdsCenario("CENA_NR_SEQUENCIA_TRANS")) Then 
           ls_str_Seq = rdsCenario("CENA_NR_SEQUENCIA_TRANS")
	    else
	       ls_str_Seq = "&nbsp;"
        End If 
	    IF Not IsNull(rdsCenario("OPES_TX_DESC_OPERACAO_ESP")) Then 
           ls_str_Desc_Ope_Esp = rdsCenario("OPES_TX_DESC_OPERACAO_ESP")
	    else
	       ls_str_Desc_Ope_Esp = "&nbsp;"
        End If 
	    IF Not IsNull(rdsCenario("CENA_CD_CENARIO_SEGUINTE")) Then 
           ls_str_Cod_Cena_Seq = rdsCenario("CENA_CD_CENARIO_SEGUINTE")
     	else
 	       ls_str_Cod_Cena_Seq = "&nbsp;"
        End If 
	    IF Not IsNull(rdsCenario("BPPP_CD_BPP")) Then 
           ls_str_Cod_BPP = rdsCenario("BPPP_CD_BPP")
	    else
	       ls_str_Cod_BPP = ""
        End If 
	    IF Not IsNull(rdsCenario("DESE_CD_DESENVOLVIMENTO")) Then 
           ls_str_Cod_Desenv = rdsCenario("DESE_CD_DESENVOLVIMENTO")
	    else
	       ls_str_Cod_Desenv = "nao"
        End If 
	    if ls_str_Cod_Desenv = "nao" then
	       ls_identifica = ls_str_Cod_Cena_Seq
     	else
	       ls_identifica = ls_str_Cod_Desenv
	    end if
   %>
      <td width="28" height="22"> 
        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ls_str_Seq%></font></div>
      </td>
      <td width="30" height="22"> 
        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsCenario("MEPR_TX_ABREVIA")%></font></div>
      </td>
      <td width="347" height="22"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsCenario("CETR_TX_DESC_TRANSACAO")%></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">&nbsp; 
        </font></td>
      <td width="92" height="22"> 
        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rdsCenario("TRAN_CD_TRANSACAO")%> </font></div>
      </td>
      <td width="52" height="22"> 
        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ls_str_Desc_Ope_Esp%></font></div>
      </td>
      <td width="79" height="22"> 
	     <%
	     set rsdesenv=Conn_db.Execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO_DESENV WHERE TRAN_CD_TRANSACAO='" & rdsCenario("TRAN_CD_TRANSACAO") & "'")
	     IF rsdesenv.eof=false then
	     %>
	     <div align="center">
          <p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ls_identifica%><a href="#" onclick="ver_desenv('<%=trim(rdsCenario("TRAN_CD_TRANSACAO"))%>')"><img border="0" src="../../imagens/b011.gif" width="16" height="16" alt="Clique para Visualizar os Desenvolvimentos relacionados..."></a></font></p>
        </div>
        <%else%>
   	     <div align="center">
          <p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ls_identifica%></font></p>
        </div>
			<%end if%>
	  </td>
      <td width="60" height="22"> 
        <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><% if rdsCenario("CETR_TX_TIPO_RELACAO") <> "5" then %> <a href="javascript:MM_goToURL5('self','inc_bpp_transa_cenario.asp',this,'<%=rdsCenario("CETR_NR_SEQUENCIA")%>')"><% If ls_str_Cod_BPP <> "" then %> <%=ls_str_Cod_BPP%> <% else %><img src="../../imagens/b04.gif" width="16" height="16" border="0"><% end if %></a><% else %>
          <font color="#FFCC66"> <font color="#FF9933">N&atilde;o aparecer&aacute; 
          no fluxo</font> 
          <% end if %></font></font></div>
      </td>
    </tr>
    <% rdsCenario.movenext
  Loop
  rdsCenario.Close
  set rdsCenario = Nothing
  %>
  </table>
  <table width="779" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="218">&nbsp;</td>
      <td width="78">&nbsp;</td>
      <td width="483">&nbsp;</td>
    </tr>
    <tr> 
      <td width="218"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Total 
          de Transa&ccedil;&otilde;es : </font></b></div>
      </td>
      <td width="78"> 
        <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=int_Total_Registros%></font></div>
      </td>
      <td width="483">&nbsp;</td>
    </tr>
    <tr> 
      <td width="218">&nbsp;</td>
      <td width="78">&nbsp;</td>
      <td width="483">&nbsp;</td>
    </tr>
  </table>
  </form>
</body>
</html>
<% conn_db.close %>
