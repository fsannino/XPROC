<%@LANGUAGE="VBSCRIPT"%> 
 
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Funcao

str_MegaProcesso = "0"
str_Funcao = 0

str_Opc = Request("txtOpc")

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if

if (Request("selFuncao") <> "") then 
    str_Funcao = Request("selFuncao")
else
    str_Funcao = "0"
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

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " MEPR_CD_MEGA_PROCESSO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "
'response.write str_Sub_Modulo
set rs_SubModulo=Conn_db.execute(str_Sub_Modulo)

str_SQL_Funcao = ""
str_SQL_Funcao = str_SQL_Funcao & " SELECT "
str_SQL_Funcao = str_SQL_Funcao & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
str_SQL_Funcao = str_SQL_Funcao & " ," & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
str_SQL_Funcao = str_SQL_Funcao & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO"
str_SQL_Funcao = str_SQL_Funcao & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Funcao = str_SQL_Funcao & " order by " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "

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
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selFuncao="+document.frm1.selFuncao.value+"'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc=3&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0&selSubProcesso=0'");
}
function MM_goToURL5() { //v3.0
  var i, args=MM_goToURL5.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc=3&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
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
if (document.frm1.selFuncao.selectedIndex == 0)
     { 
	 alert("Selecione uma Função.");
     document.frm1.selFuncao.focus();
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


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../imagens/novo_registro_02.gif','../../imagens/atualiza_02.gif')">
<form name="frm1" method="post" action="cad_funcao_transacao2.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20"><%'=str_SQL_Funcao%></td>
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
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
          <td width="26"><a href="javascript:Confirma()"><img src="../../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"><a href="javascript:Limpa()"><img src="../../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">limpa</font></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table width="94%" border="0" cellpadding="2" cellspacing="7" name="tblSubProcesso" height="111">
    <tr> 
      <td width="142" height="5"> 
        <input type="hidden" name="txtOpc" value="1">
      </td>
      <td width="505" height="5"> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
          Fun&ccedil;&atilde;o R/3 x Transa&ccedil;&atilde;o</font></div>
      </td>
      <td width="17" height="5"> 
        <%'=str_Opc%>
      </td>
      <td width="17" height="5"> 
        <%'=str_MegaProcesso%>
        <%'=str_Funcao%>
      </td>
    </tr>
    <tr> 
      <td width="142">&nbsp;</td>
      <td width="505">&nbsp;</td>
      <td width="17">&nbsp;</td>
      <td width="17">&nbsp;</td>
    </tr>
    <tr> 
      <td width="142"> 
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Mega-Processo 
          : </font></b></font></div>
      </td>
      <td width="505"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','selec_Mega_Funcao.asp');return document.MM_returnValue">
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
      <td width="17">&nbsp; </td>
      <td width="17"> 
        <%'=str_SQL_MegaProc%>
      </td>
    </tr>
    <% If str_MegaProcesso = 11 then	 
	'if rs_mega("MEPR_CD_MEGA_PROCESSO") = 11 then
	%>
    <tr> 
      <td width="142"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b>Sub-Modulo 
          : </b></font></div>
      </td>
      <td width="505"> 
        <select size="1" name="selSubModulo">
          <option value="0">== Selecione o Sub Módulo ==</option>
          <%do until rs_SubModulo.eof=true%>
          <option value="<%=rs_SubModulo("SUMO_NR_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
					rs_SubModulo.movenext
					loop
					%>
        </select>
      </td>
      <td width="17">&nbsp;</td>
      <td width="17">&nbsp;</td>
    </tr>
    <% end if %>
    <tr> 
      <td width="142" height="25"> 
        <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Fun&ccedil;&atilde;o R/3 : </font></b></font></div>
      </td>
      <td width="505" height="25"> 
        <select name="selFuncao">
          <% 
		  if str_Opc <> "1" then %>
          <option value="0" selected>Selecione uma Função</option>
          <% else %>
          <option value="0" >Selecione uma Função</option>
          <% end if %>
          <%Set rdsFuncao = Conn_db.Execute(str_SQL_Funcao)
While (NOT rdsFuncao.EOF)  
    str_SQL_AREA = ""
	str_SQL_AREA = str_SQL_AREA & " SELECT " & Session("PREFIXO") & "ORGAO_AGLUTINADOR.AGLU_SG_AGLUTINADO "
    str_SQL_AREA = str_SQL_AREA & " FROM " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU "
	str_SQL_AREA = str_SQL_AREA & " , " & Session("PREFIXO") & "ORGAO_AGLUTINADOR "
	str_SQL_AREA = str_SQL_AREA & " WHERE " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU.AGLU_CD_AGLUTINADO = " & Session("PREFIXO") & "ORGAO_AGLUTINADOR.AGLU_CD_AGLUTINADO "
	str_SQL_AREA = str_SQL_AREA & " AND " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU.FUNE_CD_FUNCAO_NEGOCIO = '" & rdsFuncao("FUNE_CD_FUNCAO_NEGOCIO") & " '"
    Set rdsAREA = Conn_db.Execute(str_SQL_AREA)
	area = ""
	Do While not rdsAREA.eof
	   area = area & rdsAREA("AGLU_SG_AGLUTINADO") & " / "
	   rdsAREA.movenext
	Loop
	rdsAREA.close	
           if (Trim(str_Funcao) = Trim(rdsFuncao.Fields.Item("FUNE_CD_FUNCAO_NEGOCIO").Value)) then %>
          <option value="<%=(rdsFuncao.Fields.Item("FUNE_CD_FUNCAO_NEGOCIO").Value)%>" selected ><%= area & " = " & (rdsFuncao.Fields.Item("FUNE_CD_FUNCAO_NEGOCIO").Value)%> - <%=(rdsFuncao.Fields.Item("FUNE_TX_TITULO_FUNCAO_NEGOCIO").Value)%></option>
          <% else %>
          <option value="<%=(rdsFuncao.Fields.Item("FUNE_CD_FUNCAO_NEGOCIO").Value)%>"><%=area & " >> " & (rdsFuncao.Fields.Item("FUNE_CD_FUNCAO_NEGOCIO").Value)%> - <%=(rdsFuncao.Fields.Item("FUNE_TX_TITULO_FUNCAO_NEGOCIO").Value)%></option>
          <% end if %>
          <%
  rdsFuncao.MoveNext()
Wend
If (rdsFuncao.CursorType > 0) Then
  rdsFuncao.MoveFirst
Else
  rdsFuncao.Requery
End If

rdsFuncao.Close
set rdsFuncao = Nothing
set rdsAREA = Nothing
%>
        </select>
      </td>
      <td width="17" height="25">&nbsp;</td>
      <td width="17" height="25"> 
        <%'=str_SQL_Funcao%>
      </td>
    </tr>
    <tr> 
      <td width="142" height="25">&nbsp;</td>
      <td width="505" height="25">&nbsp; </td>
      <td width="17" height="25">&nbsp; </td>
      <td width="17" height="25">&nbsp; </td>
    </tr>
    <tr>
      <td width="142" height="25">&nbsp;</td>
      <td width="505" height="25">&nbsp;</td>
      <td width="17" height="25">&nbsp;</td>
      <td width="17" height="25">&nbsp;</td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>
