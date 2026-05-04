<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"

if (Request("txtOpc") <> "") then 
   str_Opc = Request("txtOpc")
   if str_Opc = 2 then
      if (Request("selMegaProcesso") <> "") then 
         str_MegaProcesso = Request("selMegaProcesso")
	  end if
   end if
   if str_Opc = 3 then  	  
      if (Request("selProcesso") <> "") then 
         str_Trata = Request("selProcesso")
	     int_Tamanho = Len(Trim(str_Trata))
		 if int_Tamanho > 2 then
		    for i=1 to int_Tamanho
		        if Mid(str_Trata,i,1) = "/"  then
		           str_MegaProcesso = Trim(Mid(str_Trata,1,i-1))
			       str_Processo = Trim(Mid(str_Trata,i+1,int_Tamanho))
                   exit for
		        end if
		    next
         else
		    str_MegaProcesso = 0
		    str_Processo = 0			
	     end if
       else
		  str_MegaProcesso = 0
		  str_Processo = 0				   
	   end if
	end if
	if str_Opc = 4 then  	  
       if (Request("selSubProcesso") <> "") then 
          str_Trata = Request("selSubProcesso")
		  int_Tamanho = Len(Trim(str_Trata))
		  if int_Tamanho > 2 then
		     for i=1 to int_Tamanho
		        if Mid(str_Trata,i,1) = "/"  then
		           str_MegaProcesso = Trim(Mid(str_Trata,1,i-1))
			       str_Trata = Trim(Mid(str_Trata,i+1,int_Tamanho))
                   exit for
		        end if
		     next
		     int_Tamanho = Len(Trim(str_Trata))
		     for i=1 to int_Tamanho
		        if Mid(str_Trata,i,1) = "/"  then
		           str_Processo = Mid(str_Trata,1,i-1)
			       str_SubProcesso = Mid(str_Trata,i+1,int_Tamanho)
                   exit for
		        end if
		     next
          else
		     str_MegaProcesso = 0
		     str_Processo = 0
			 str_SubProcesso = 0		
		  end if
	   end if	
	end if	 
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
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

str_SQL_Atividade = ""
str_SQL_Atividade = str_SQL_Atividade & " SELECT "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.ATIV_CD_ATIVIDADE, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.ATIV_TX_DESC_ATIVIDADE"
str_SQL_Atividade = str_SQL_Atividade & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN"
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "PROCESSO ON "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " INNER JOIN " & Session("PREFIXO") & "ATIVIDADE ON "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " WHERE " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
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
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frmAtividade.selMegaProcesso.value+"'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frmAtividade.selProcesso.value+"'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frmAtividade.selSubProcesso.value+"'");
}
function Confirma() 
{ 
	  document.frmAtividade.submit();
}
function Confirma2() 
{ 
if (document.frmAtividade.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frmAtividade.selMegaProcesso.focus();
     return;
     }
if ((document.frmAtividade.selProcesso.selectedIndex == 0)&&
	(document.frmAtividade.txtNovoProcesso.value == ""))
     { 
	 alert("Selecione um Proceso ou cadastre um novo.");
     document.frmAtividade.selProcesso.focus();
     return;
     }	 
if ((document.frmAtividade.selSubProcesso.selectedIndex == 0)&&
	(document.frmAtividade.txtNovoSubProcesso.value == ""))
     { 
	 alert("Selecione um Sub Proceso ou cadastre um novo.");
     document.frmAtividade.selSubProcesso.focus();
     return;
     }	 
if ((document.frmAtividade.selAtividade.selectedIndex == 0)&&
	(document.frmAtividade.txtNovaAtividade.value == ""))
     { 
	 alert("Selecione uma Atividade ou cadastre uma nova.");
     document.frmAtividade.selAtividade.focus();
     return;
     }	 
	 else
     {
	  document.frmAtividade.submit();
	 }
 }

function Limpa(){
	document.frmAtividade.reset();
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


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
<form name="frmAtividade" method="post" action="grava_inc_sub_processo.asp">
  <table width="94%" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso" height="111">
    <tr> 
      <td width="23%">
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Mega-Processo</font></b></font></div>
      </td>
      <td width="56%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','form_inc_sub_processo.asp?txtOpc=2&amp;selMegaProcesso=');return document.MM_returnValue">
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
      <td width="21%"> 
        <%'=str_SQL_MegaProc%>
      </td>
    </tr>
    <tr> 
      <td width="23%" height="25">
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Processo</font></b></font></div>
      </td>
      <td width="56%" height="25"> 
        <select name="selProcesso" onChange="MM_goToURL2('self','form_inc_sub_processo.asp?txtOpc=3&amp;selProcesso=',this);return document.MM_returnValue">
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
          <option value="<%=(rdsProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value & "/" & rdsProcesso.Fields.Item("PROC_CD_PROCESSO").Value)%>"><%=(rdsProcesso.Fields.Item("PROC_TX_DESC_PROCESSO").Value)%></option>
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
      <td width="21%" height="25"> 
        <%'=str_SQL_Proc%>
      </td>
    </tr>
    <tr> 
      <td width="23%" height="25">
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Novo 
          Processo</font></b></font></div>
      </td>
      <td width="56%" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="txtNovoProcesso" maxlength="150" size="50">
        </font></td>
      <td width="21%" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <%'=str_SQL_Sub_Proc%>
        </font></td>
    </tr>
    <tr> 
      <td width="23%" height="21">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Sub-Processo</font></b></div>
      </td>
      <td width="56%"> 
        <select name="selSubProcesso" onChange="MM_goToURL3('self','form_inc_sub_processo.asp?txtOpc=4&amp;selSubProcesso=',this);return document.MM_returnValue">
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
          <option value="<%=rdsSubProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value & "/" & rdsSubProcesso.Fields.Item("PROC_CD_PROCESSO").Value & "/" & rdsSubProcesso.Fields.Item("SUPR_CD_SUB_PROCESSO").Value%>" ><%=(rdsSubProcesso.Fields.Item("SUPR_TX_DESC_SUB_PROCESSO").Value)%></option>
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
      <td width="21%" height="21"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <%'=str_SQL_Atividade%>
        </font></td>
    </tr>
    <tr> 
      <td width="23%" height="21">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Novo 
          Sub-Processo</font></b></div>
      </td>
      <td width="56%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="txtNovoSubProcesso" maxlength="150" size="50">
        </font></td>
      <td width="21%" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="23%">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Atividade</font></b></div>
      </td>
      <td width="56%"> 
        <select name="selAtividade">
          <option value="0" selected>Selecione uma Atividade</option>
          <%Set rdsAtividade = Conn_db.Execute(str_SQL_Atividade)
While (NOT rdsAtividade.EOF)%>
          <option value="<%=(rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)%>" ><%=(rdsAtividade.Fields.Item("ATIV_TX_DESC_ATIVIDADE").Value)%></option>
          <%
  rdsAtividade.MoveNext()
Wend
If (rdsAtividade.CursorType > 0) Then
  rdsAtividade.MoveFirst
Else
  rdsAtividade.Requery
End If
rdsAtividade.close
set rdsAtividade = Nothing
%>
        </select>
      </td>
      <td width="21%">&nbsp; </td>
    </tr>
    <tr> 
      <td width="23%">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Nova 
          Atividade</font></b></div>
      </td>
      <td width="56%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="txtNovaAtividade" maxlength="150" size="50">
        </font></td>
      <td width="21%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="23%">&nbsp;</td>
      <td width="56%">&nbsp;</td>
      <td width="21%">&nbsp;</td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"

if (Request("txtOpc") <> "") then 
   str_Opc = Request("txtOpc")
   if str_Opc = 2 then
      if (Request("selMegaProcesso") <> "") then 
         str_MegaProcesso = Request("selMegaProcesso")
	  end if
   end if
   if str_Opc = 3 then  	  
      if (Request("selProcesso") <> "") then 
         str_Trata = Request("selProcesso")
	     int_Tamanho = Len(Trim(str_Trata))
		 if int_Tamanho > 2 then
		    for i=1 to int_Tamanho
		        if Mid(str_Trata,i,1) = "/"  then
		           str_MegaProcesso = Trim(Mid(str_Trata,1,i-1))
			       str_Processo = Trim(Mid(str_Trata,i+1,int_Tamanho))
                   exit for
		        end if
		    next
         else
		    str_MegaProcesso = 0
		    str_Processo = 0			
	     end if
       else
		  str_MegaProcesso = 0
		  str_Processo = 0				   
	   end if
	end if
	if str_Opc = 4 then  	  
       if (Request("selSubProcesso") <> "") then 
          str_Trata = Request("selSubProcesso")
		  int_Tamanho = Len(Trim(str_Trata))
		  if int_Tamanho > 2 then
		     for i=1 to int_Tamanho
		        if Mid(str_Trata,i,1) = "/"  then
		           str_MegaProcesso = Trim(Mid(str_Trata,1,i-1))
			       str_Trata = Trim(Mid(str_Trata,i+1,int_Tamanho))
                   exit for
		        end if
		     next
		     int_Tamanho = Len(Trim(str_Trata))
		     for i=1 to int_Tamanho
		        if Mid(str_Trata,i,1) = "/"  then
		           str_Processo = Mid(str_Trata,1,i-1)
			       str_SubProcesso = Mid(str_Trata,i+1,int_Tamanho)
                   exit for
		        end if
		     next
          else
		     str_MegaProcesso = 0
		     str_Processo = 0
			 str_SubProcesso = 0		
		  end if
	   end if	
	end if	 
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
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

str_SQL_Atividade = ""
str_SQL_Atividade = str_SQL_Atividade & " SELECT "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.ATIV_CD_ATIVIDADE, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO, "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "ATIVIDADE.ATIV_TX_DESC_ATIVIDADE"
str_SQL_Atividade = str_SQL_Atividade & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO INNER JOIN"
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "PROCESSO ON "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " INNER JOIN " & Session("PREFIXO") & "ATIVIDADE ON "
str_SQL_Atividade = str_SQL_Atividade & " " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO"
str_SQL_Atividade = str_SQL_Atividade & " WHERE " & Session("PREFIXO") & "ATIVIDADE.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "ATIVIDADE.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "ATIVIDADE.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
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
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frmAtividade.selMegaProcesso.value+"'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frmAtividade.selProcesso.value+"'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frmAtividade.selSubProcesso.value+"'");
}
function Confirma() 
{ 
	  document.frmAtividade.submit();
}
function Confirma2() 
{ 
if (document.frmAtividade.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frmAtividade.selMegaProcesso.focus();
     return;
     }
if ((document.frmAtividade.selProcesso.selectedIndex == 0)&&
	(document.frmAtividade.txtNovoProcesso.value == ""))
     { 
	 alert("Selecione um Proceso ou cadastre um novo.");
     document.frmAtividade.selProcesso.focus();
     return;
     }	 
if ((document.frmAtividade.selSubProcesso.selectedIndex == 0)&&
	(document.frmAtividade.txtNovoSubProcesso.value == ""))
     { 
	 alert("Selecione um Sub Proceso ou cadastre um novo.");
     document.frmAtividade.selSubProcesso.focus();
     return;
     }	 
if ((document.frmAtividade.selAtividade.selectedIndex == 0)&&
	(document.frmAtividade.txtNovaAtividade.value == ""))
     { 
	 alert("Selecione uma Atividade ou cadastre uma nova.");
     document.frmAtividade.selAtividade.focus();
     return;
     }	 
	 else
     {
	  document.frmAtividade.submit();
	 }
 }

function Limpa(){
	document.frmAtividade.reset();
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


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
<form name="frmAtividade" method="post" action="grava_inc_sub_processo.asp">
  <table width="94%" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso" height="111">
    <tr> 
      <td width="23%">
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Mega-Processo</font></b></font></div>
      </td>
      <td width="56%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','form_inc_sub_processo.asp?txtOpc=2&amp;selMegaProcesso=');return document.MM_returnValue">
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
      <td width="21%"> 
        <%'=str_SQL_MegaProc%>
      </td>
    </tr>
    <tr> 
      <td width="23%" height="25">
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Processo</font></b></font></div>
      </td>
      <td width="56%" height="25"> 
        <select name="selProcesso" onChange="MM_goToURL2('self','form_inc_sub_processo.asp?txtOpc=3&amp;selProcesso=',this);return document.MM_returnValue">
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
          <option value="<%=(rdsProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value & "/" & rdsProcesso.Fields.Item("PROC_CD_PROCESSO").Value)%>"><%=(rdsProcesso.Fields.Item("PROC_TX_DESC_PROCESSO").Value)%></option>
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
      <td width="21%" height="25"> 
        <%'=str_SQL_Proc%>
      </td>
    </tr>
    <tr> 
      <td width="23%" height="25">
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Novo 
          Processo</font></b></font></div>
      </td>
      <td width="56%" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="txtNovoProcesso" maxlength="150" size="50">
        </font></td>
      <td width="21%" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <%'=str_SQL_Sub_Proc%>
        </font></td>
    </tr>
    <tr> 
      <td width="23%" height="21">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Sub-Processo</font></b></div>
      </td>
      <td width="56%"> 
        <select name="selSubProcesso" onChange="MM_goToURL3('self','form_inc_sub_processo.asp?txtOpc=4&amp;selSubProcesso=',this);return document.MM_returnValue">
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
          <option value="<%=rdsSubProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value & "/" & rdsSubProcesso.Fields.Item("PROC_CD_PROCESSO").Value & "/" & rdsSubProcesso.Fields.Item("SUPR_CD_SUB_PROCESSO").Value%>" ><%=(rdsSubProcesso.Fields.Item("SUPR_TX_DESC_SUB_PROCESSO").Value)%></option>
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
      <td width="21%" height="21"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <%'=str_SQL_Atividade%>
        </font></td>
    </tr>
    <tr> 
      <td width="23%" height="21">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Novo 
          Sub-Processo</font></b></div>
      </td>
      <td width="56%"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="txtNovoSubProcesso" maxlength="150" size="50">
        </font></td>
      <td width="21%" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="23%">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Atividade</font></b></div>
      </td>
      <td width="56%"> 
        <select name="selAtividade">
          <option value="0" selected>Selecione uma Atividade</option>
          <%Set rdsAtividade = Conn_db.Execute(str_SQL_Atividade)
While (NOT rdsAtividade.EOF)%>
          <option value="<%=(rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)%>" ><%=(rdsAtividade.Fields.Item("ATIV_TX_DESC_ATIVIDADE").Value)%></option>
          <%
  rdsAtividade.MoveNext()
Wend
If (rdsAtividade.CursorType > 0) Then
  rdsAtividade.MoveFirst
Else
  rdsAtividade.Requery
End If
rdsAtividade.close
set rdsAtividade = Nothing
%>
        </select>
      </td>
      <td width="21%">&nbsp; </td>
    </tr>
    <tr> 
      <td width="23%">
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Nova 
          Atividade</font></b></div>
      </td>
      <td width="56%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="txtNovaAtividade" maxlength="150" size="50">
        </font></td>
      <td width="21%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="23%">&nbsp;</td>
      <td width="56%">&nbsp;</td>
      <td width="21%">&nbsp;</td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
