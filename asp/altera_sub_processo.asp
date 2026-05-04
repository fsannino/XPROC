<%@ Language=VBScript %> 
<!--#include file="../asp/protege/protege.asp" -->
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"

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

if str_MegaProcesso <> "0" then
   Session("MegaProcesso") = str_MegaProcesso
else
    if Session("MegaProcesso") <> "" then
       str_MegaProcesso = Session("MegaProcesso") 
	end if   
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

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
str_SQL_Proc = str_SQL_Proc & " WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 

str_SQL_Sub_Proc = ""
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " SELECT "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO, "
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO,"
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_NR_SEQUENCIA,"
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_IMPACTO"
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

str_SQL_Empresa_Unid = ""
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " SELECT "
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA "
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " ," & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_TX_NOME_EMPRESA "
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE "

str_SQL_Rel_Sub_Emp_Falta = ""
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " SELECT " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_TX_NOME_EMPRESA, " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA"
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE"
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " WHERE " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA not in ("
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " SELECT " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA"
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " FROM " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE INNER JOIN"
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " " & Session("PREFIXO") & "EMPRESA_UNIDADE ON " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA = " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA"
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " WHERE "
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso  
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " AND " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.PROC_CD_PROCESSO = " & str_Processo 
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " AND " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & ")"
str_SQL_Rel_Sub_Emp_Falta = str_SQL_Rel_Sub_Emp_Falta & " order by EMPR_TX_NOME_EMPRESA  "

str_SQL_Rel_Sub_Emp = ""
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " SELECT " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA,"
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_TX_NOME_EMPRESA"
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " FROM " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE INNER JOIN"
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " " & Session("PREFIXO") & "EMPRESA_UNIDADE ON " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA = " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA"
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " WHERE "
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso  
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " AND " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.PROC_CD_PROCESSO = " & str_Processo 
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " AND " & Session("PREFIXO") & "SUB_PROCESSO_EMPRESA_UNIDADE.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Rel_Sub_Emp = str_SQL_Rel_Sub_Emp & " ORDER by " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_TX_NOME_EMPRESA  "

%>

<html>
<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<meta name="VI60_defaultClientScript" content="VBScript">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--
function carrega_txt(fbox) {
document.frm1.txtEmpSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtEmpSelecionada.value = document.frm1.txtEmpSelecionada.value + "," + fbox.options[i].value;
   }
}
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
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
if (document.frm1.txtDescSubProc.value == "")
     { 
	 alert("Preencha nova descrição do Sub Processo.");
     document.frm1.txtDescSubProc.focus();
     return;
     }
if(document.frm1.valImpacto.value==0 )
     { 
	 alert("Você deve selecionar o IMPACTO do sub-Processo.");
     return;
     }	 
//if (document.frm1.list2.options.length == 0)
//     { 
//	 alert("A seleção de uma Empresa/Unidade é obrigatória !");
//     document.frm1.list2.focus();
//     return;
//     }	 
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
<script language="javascript" src="js/troca_lista.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../imagens/continua_F02.gif','../imagens/continua2_F02.gif')">
<form name="frm1" method="post" action="grava_altera_sub_processo.asp">
              <input type="hidden" name="txtEmpSelecionada"><input type="hidden" name="txtOpc" value="<%=str_Opc%>">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="19%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="37%" valign="top" height="65"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20"> 
        <table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
            <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
            <td width="26">&nbsp;</td>
            <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
            <td width="27"><a href="javascript:Limpa()"><img src="../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></td>
            <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">limpa</font></b></font></td>
            <td width="28">&nbsp;</td>
            <td width="26">&nbsp;</td>
            <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <p align="center"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Altera 
        Sub-Processo</font></p>
  <table border="0" width="971" height="175">
    <tr>
      <td width="286" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Mega-Processo</b></font></td>
      <td width="1031" colspan="8" height="17"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','altera_sub_processo.asp?txtOpc=2');return document.MM_returnValue">
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
      <td width="286" height="21"></td>
      <td width="1031" colspan="8" height="13"></td>
    </tr>
    <tr>
      <td width="286" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Processos</b></font></td>
      <td width="1031" colspan="8" height="23"> 
              <select name="selProcesso" onChange="MM_goToURL2('self','altera_sub_processo.asp?txtOpc=3',this);return document.MM_returnValue">
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
      <td width="286" height="21"></td>
      <td width="1031" colspan="8" height="19"></td>
    </tr>
    <tr>
      <td width="286" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Sub-Processo</b></font></td>
      <td width="1031" colspan="8" height="23">
        <select name="selSubProcesso" onChange="MM_goToURL3('self','altera_sub_processo.asp?',this);return document.MM_returnValue">> 
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
		  <%ls_Desc_SubProcesso=(rdsSubProcesso.Fields.Item("SUPR_TX_DESC_SUB_PROCESSO").Value)
		  ls_Seq_SubProcesso=(rdsSubProcesso.Fields.Item("SUPR_NR_SEQUENCIA").Value)
		  %>
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
%>
        </select>
      </td>
    </tr>
    <tr>
      <td width="286" height="21"></td>
      <td width="1031" colspan="8" height="13"></td>
    </tr>
    <tr>
      <td width="286" height="34"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Nova 
        Desc Sub-Processo</b></font></td>
      <td width="1031" colspan="8" height="32"> 
              <input type="text" name="txtDescSubProc" size="50" maxlength="150" value="<%=ls_Desc_SubProcesso%>"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <input type="text" name="txtSeq1" size="7" value="<%=ls_Seq_SubProcesso%>">
                </font></td>
    </tr>
    <tr>
      <td width="286" height="1"></td>
      <td width="1031" colspan="8" height="1"></td>
    </tr>
    <%
    set temp=conn_db.execute("SELECT SUPR_TX_IMPACTO FROM SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " AND PROC_CD_PROCESSO=" & str_Processo & " AND SUPR_CD_SUB_PROCESSO=" & str_SubProcesso )    
    
    On Error Resume next
    impacto = TEMP("SUPR_TX_IMPACTO")
    
    select case impacto
    case 1
		v1="checked"
      	valorImp=1
    case 2
		v2="checked"
      	valorImp=2
    case 3
		v3="checked"
      	valorImp=3
    case else
      	valorImp=0
    end select
    
    %>
    <tr>
      <td width="286" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Impacto</b></font></td>
      <td width="26" height="21">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b><input type="radio" value="1" name="selImpacto" onClick="document.frm1.valImpacto.value=this.value" <%=v1%>></b></font> </td>
      <td width="68" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Alto</b></font> </td>
      <td width="32" height="21">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b><input type="radio" value="2" name="selImpacto" onClick="document.frm1.valImpacto.value=this.value" <%=v2%>></b></font> </td>
      <td width="93" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Médio</b></font> </td>
      <td width="21" height="21">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b><input type="radio" value="3" name="selImpacto" onClick="document.frm1.valImpacto.value=this.value" <%=v3%>></b></font> </td>
      <td width="280" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Baixo</b></font> </td>
      <td width="259" height="21"> <input type="hidden" name="valImpacto" size="20" value=<%=valorImp%>> </td>
      <td width="191" height="19"></td>
    </tr>
    
    <tr>
      <td width="1194" height="19" colspan="9"></td>
    </tr>
  </table>
  <table border="0" width="0" height="1">
    <tr>
      <td width="132" valign="top" height="1">
      </td>
      <td width="478" height="1"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Selecione
        Empresa/Unidade</b></font></td>
      <td width="24" height="1"></td>
      <td width="486" height="1"></td>
    </tr>
    <tr>
      <td width="132" valign="top" rowspan="5" height="154">
      </td>
      <td width="478" rowspan="5" height="154"><b><select name="list1" size="8" multiple>
                  <%Set rdsEmp_Unid = Conn_db.Execute(str_SQL_Rel_Sub_Emp_Falta)
While (NOT rdsEmp_Unid.EOF)
%>
                  <option value="<%=(rdsEmp_Unid.Fields.Item("EMPR_CD_NR_EMPRESA").Value)%>" ><%=(rdsEmp_Unid.Fields.Item("EMPR_TX_NOME_EMPRESA").Value)%></option>
                  <%
  rdsEmp_Unid.MoveNext()
Wend
If (rdsEmp_Unid.CursorType > 0) Then
  rdsEmp_Unid.MoveFirst
Else
  rdsEmp_Unid.Requery
End If
rdsEmp_Unid.close
set rdsEmp_Unid = Nothing
%>
                </select>
                </b></td>
      <td width="24" height="24"></td>
      <td width="486" rowspan="5" height="102"><font color="#000080"> 
                <select name="list2" size="8" multiple>
                  <%Set rdsEmp_Unid = Conn_db.Execute(str_SQL_Rel_Sub_Emp)
While (NOT rdsEmp_Unid.EOF)
%>
                  <option value="<%=(rdsEmp_Unid.Fields.Item("EMPR_CD_NR_EMPRESA").Value)%>" ><b><%=(rdsEmp_Unid.Fields.Item("EMPR_TX_NOME_EMPRESA").Value)%></b></option>
                  <%
  rdsEmp_Unid.MoveNext()
Wend
If (rdsEmp_Unid.CursorType > 0) Then
  rdsEmp_Unid.MoveFirst
Else
  rdsEmp_Unid.Requery
End If
rdsEmp_Unid.close
set rdsEmp_Unid = Nothing
%>
                </select>
                </font></td>
    </tr>
    <tr>
      <td width="24" height="29">
        <p align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2)"><img name="Image16" border="0" src="../imagens/continua_F01.gif" width="24" height="24" align="left"></a></td>
    </tr>
    <tr>
      <td width="24" height="23"></td>
    </tr>
    <tr>
      <td width="24" height="29">
        <p align="center"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1)"><img name="img01511" border="0" src="../imagens/continua2_F01.gif" width="24" height="24" align="left"></a></td>
    </tr>
    <tr>
      <td width="24" height="1"></td>
    </tr>
    <tr>
      <td width="132" height="1"></td>
      <td width="994" colspan="3" height="1">
        <p align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
                a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></td>
    </tr>
  </table>
              <p>&nbsp;</p>
  </form>
</body>
</html>


