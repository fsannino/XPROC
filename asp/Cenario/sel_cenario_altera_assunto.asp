<%@LANGUAGE="VBSCRIPT"%> 
<%
if (Request("txtOPT") <> "") then
    str_OPT = Request("txtOPT")
else
    str_OPT = "0"
end if

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if
if str_MegaProcesso <> "0" then
   Session("MegaProcesso") = str_MegaProcesso
else
    if Session("MegaProcesso") <> "" then
       str_MegaProcesso = Session("MegaProcesso") 
	end if   
end if
if request("selAssunto") then
   str_Assunto=request("selAssunto")
else
   str_Assunto = "0"
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

if (Request("selOnda") <> "") then 
    str_Onda = Request("selOnda")
else
    str_Onda = "0"
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
if str_Processo <> "0" then
   str_SQL_Cenario = str_SQL_Cenario & " AND " & Session("PREFIXO") & "CENARIO.PROC_CD_PROCESSO = " & str_Processo 
end if
if str_SubProcesso <> "0" then
   str_SQL_Cenario = str_SQL_Cenario & " AND " & Session("PREFIXO") & "CENARIO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso 
end if
if str_Onda <> "0" then
   str_SQL_Cenario = str_SQL_Cenario & " AND " & Session("PREFIXO") & "CENARIO.ONDA_CD_ONDA = " & str_Onda 
end if

if str_Assunto<>0 then
	str_SQL_Cenario = str_SQL_Cenario & " AND " & Session("PREFIXO") & "CENARIO.SUMO_NR_CD_SEQUENCIA = " & str_Assunto
end if

str_SQL_Cenario = str_SQL_Cenario & " order by " & Session("PREFIXO") & "CENARIO.CENA_TX_TITULO_CENARIO "

'response.Write(str_SQL_Cenario)

set rs_onda=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_TX_DESC_ONDA")

SQL_Assunto=""
SQL_Assunto = SQL_Assunto & " SELECT SUMO_NR_CD_SEQUENCIA"
SQL_Assunto = SQL_Assunto & " ,SUMO_TX_DESC_SUB_MODULO"
SQL_Assunto = SQL_Assunto & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
SQL_Assunto = SQL_Assunto & " FROM " & Session("PREFIXO") & "SUB_MODULO"
if str_MegaProcesso <> 0 then
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
else
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS = '9999'"
end if
SQL_Assunto=SQL_Assunto + " ORDER BY SUMO_TX_DESC_SUB_MODULO"

set rs_assunto=conn_db.execute(SQL_Assunto)
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
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOPT="+document.frm1.txtOPT.value+"&selAssunto="+document.frm1.selAssunto.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOPT="+document.frm1.txtOPT.value+"&selAssunto="+document.frm1.selAssunto.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOPT="+document.frm1.txtOPT.value+"&selAssunto="+document.frm1.selAssunto.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOPT="+document.frm1.txtOPT.value+"&selAssunto="+document.frm1.selAssunto.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"&selOnda="+document.frm1.selOnda.value+"'");
}

function manda_assunto()
{
window.location.href="sel_cenario_altera_assunto.asp?txtOPT="+document.frm1.txtOPT.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"&selAssunto="+document.frm1.selAssunto.value
}

function Confirma() 
{ 
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("Você deve selecionar um MEGA-PROCESSO");
     document.frm1.selMegaProcesso.focus();
     return;
     }
	 else
     {
      if(document.frm1.txtOPT.value == 1)
        {
        document.frm1.action="altera_assunto_em_massa.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 2)
        {
        document.frm1.action="relacao_cenario_sem_assunto.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }
      if(document.frm1.txtOPT.value == 3)
        {
        document.frm1.action="relacao_cenario_sem_empresa.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }
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
<form name="frm1" method="post" action="">
  <input type="hidden" name="INC" size="20" value="1"> 
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="39" valign="middle" align="center">
              <div align="center">
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
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
  <table width="96%" border="0" cellpadding="0" cellspacing="5" name="tblSubProcesso" height="251">
    <tr>
      <td height="39">&nbsp;</td>
      <td height="39"><div align="center">
          <input type="hidden" name="txtOPT" value="<%=str_OPT%>">
		  <% if str_OPT = "1" then 
    	        str_Texto = "Altera Assunto"
		     elseif str_OPT = "2" then 
			    str_Texto = "Relação de cenário sem Assunto"		
		     elseif str_OPT = "3" then 
			    str_Texto = "Relação de cenário sem Empresa"						
			 end if	
		  %>
          <font face="Verdana" color="#330099" size="3"><%=str_Texto%></font></div></td>
      <td height="39">&nbsp;</td>
      <td height="39">&nbsp;</td>
    </tr>
    <tr> 
      <td width="19%" height="39">&nbsp;</td>
      <td width="59%" height="39">&nbsp; </td>
      <td width="5%" height="39"> <%'=str_Opc%> </td>
      <td width="17%" height="39"> <%'=str_MegaProcesso%> <%'=str_Processo%> </td>
    </tr>
    <tr> 
      <td width="19%" height="25"> <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Mega-Processo&nbsp;&nbsp;</font></b></font></div></td>
      <td width="59%" height="25"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','sel_cenario_altera_assunto.asp');return document.MM_returnValue">
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
      <td width="5%" height="25">&nbsp; </td>
      <td width="17%" height="25"> <%'=str_SQL_MegaProc%> </td>
    </tr>
    <tr> 
      <td width="19%" height="22"> <p align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Assunto 
          &nbsp;</font></b> </td>
      <td width="59%" colspan="2" height="22"> <select size="1" name="selAssunto" onChange="manda_assunto()">
          <option value="0">Selecione um Assunto</option>
          <%
          do until rs_assunto.eof=true
          if trim(str_Assunto) = trim(rs_assunto("SUMO_NR_CD_SEQUENCIA")) then
          %>
          <option selected value="<%=rs_assunto("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_assunto("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%else%>
          <option value="<%=rs_assunto("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_assunto("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
          end if
          rs_assunto.movenext
          loop
          %>
        </select> </td>
      <td width="5%" height="25"> </td>
      <td width="17%" height="25"> </td>
    </tr>
    <tr> 
      <td width="19%" height="25"> <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Processo&nbsp;&nbsp;</font></b></font></div></td>
      <td width="59%" height="25"> <select name="selProcesso" onChange="MM_goToURL2('self','sel_cenario_altera_assunto.asp',this);return document.MM_returnValue">
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
      <td width="5%" height="25">&nbsp; </td>
      <td width="17%" height="25"> <%'=str_SQL_Proc%> </td>
    </tr>
    <tr> 
      <td width="19%" height="25"> <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Sub-Processo&nbsp;&nbsp;</font></b></div></td>
      <td width="59%" height="25"> <select name="selSubProcesso" onChange="MM_goToURL3('self','sel_cenario_altera_assunto.asp',this);return document.MM_returnValue">
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
      <td width="5%" height="25"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp; 
        </font></td>
      <td width="17%" height="25"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp; 
        </font></td>
    </tr>
    <tr> 
      <td width="21%"> <p align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Onda</font></b> 
      </td>
      <td width="52%"> <select size="1" name="selOnda" onChange="MM_goToURL4('self','sel_cenario_altera_assunto.asp',this);return document.MM_returnValue">
          <option value="0">Selecione a Onda</option>
          <%DO UNTIL RS_ONDA.EOF=TRUE
      IF TRIM(str_onda)=trim(rs_onda("ONDA_CD_ONDA")) then
      %>
          <option selected value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%else%>
          <option value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%
		END IF
		RS_ONDA.MOVENEXT
		LOOP
		%>
        </select> </td>
      <td width="5%" height="1">&nbsp;</td>
      <td width="17%" height="1">&nbsp;</td>
    </tr>
    <tr> 
      <td width="19%" height="1"> <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Cen&aacute;rio&nbsp;&nbsp;</font></b></div></td>
      <td width="59%" height="1"> <select name="selCenario" size="1">
          <% 
		  if str_Opc <> "1" then %>
          <option value="0" selected>Selecione um Cenário</option>
          <% else %>
          <option value="0" >Selecione um Cenário</option>
          <% end if %>
          <%Set rdsCenario = Conn_db.Execute(str_SQL_Cenario)
While (NOT rdsCenario.EOF)
           if (Trim(str_Cenario) = Trim(rdsCenario.Fields.Item("CENA_CD_CENARIO").Value)) then %>
          <option value="<%=rdsCenario.Fields.Item("CENA_CD_CENARIO").Value%>" selected ><%=rdsCenario.Fields.Item("CENA_CD_CENARIO").Value%> 
          - <%=(rdsCenario.Fields.Item("CENA_TX_TITULO_CENARIO").Value)%></option>
          <% else %>
          <option value="<%=rdsCenario.Fields.Item("CENA_CD_CENARIO").Value%>" ><%=rdsCenario.Fields.Item("CENA_CD_CENARIO").Value%> 
          - <%=(rdsCenario.Fields.Item("CENA_TX_TITULO_CENARIO").Value)%></option>
          <% end if %>
          <%
  rdsCenario.MoveNext()
Wend
If (rdsCenario.CursorType > 0) Then
  rdsCenario.MoveFirst
Else
  rdsCenario.Requery
End If
rdsCenario.close
set rdsCenario = Nothing
%>
        </select> </td>
      <td width="5%" height="1">&nbsp; </td>
      <td width="17%" height="1">&nbsp;</td>
    </tr>
    <tr> 
      <td width="19%" height="21"></td>
      <%
      if request("SEM")=1 THEN
      ORD="Cenário não encontrado!"
      else
      ORD=""
      end if
      %>
      <td width="59%" bgcolor="#FFFFFF" height="21"><font color="#800000" size="2" face="Verdana"><b><%=ord%></b></font> </td>
      <td width="5%" height="21"></td>
      <td width="17%" height="21"></td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>
