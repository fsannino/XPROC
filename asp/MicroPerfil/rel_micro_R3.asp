<%@LANGUAGE="VBSCRIPT"%> 
 
<%

Dim str_Opc
Dim str_MegaProcesso

str_MegaProcesso = "0"

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if

'if str_MegaProcesso <> "0" then
'   Session("MegaProcesso") = str_MegaProcesso
'else
'    if Session("MegaProcesso") <> "" then
'       str_MegaProcesso = Session("MegaProcesso") 
'	end if   
'end if

str_Desc_Macro = ""

if request("selFuncPrinc") <> "" then
   str_funcao = UCase(Trim(request("selFuncPrinc")))
else
   str_funcao = "0"
end if
if request("selMacroPerfil") <> "" then
   str_MacroPerfil = UCase(Trim(request("selMacroPerfil")))
else
   str_MacroPerfil = 0
end if

if request("pOPT2") <> "" then
   str_Opt2 = request("pOPT2")
else
   str_Opt2 = 0
end if
'response.Write("<p>" & str_Opt2)

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

Set rdsMegaProcesso = Conn_db.Execute(str_SQL_MegaProc)

'===========================================================================================================================
str_SQL_Fun_Neg = ""
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " SELECT DISTINCT "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
'str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO INNER JOIN "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO ON "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " INNER JOIN"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "MACRO_PERFIL ON "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = " & Session("PREFIXO") & "MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso  
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " and " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_INDICA_REFERENCIADA = '0'"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ORDER BY " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
'response.Write("<p>" & str_SQL_Fun_Neg)
set rs_funcao=conn_db.execute(str_SQL_Fun_Neg)

IF str_Opt2 = 2 Then
   str_SQL_Fun_Neg = ""
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " SELECT DISTINCT "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "MACRO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ," & Session("PREFIXO") & "MACRO_PERFIL.MCPE_TX_DESC_MACRO_PERFIL"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO INNER JOIN "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO ON "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " INNER JOIN"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "MACRO_PERFIL ON "
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = " & Session("PREFIXO") & "MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO"
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_funcao & "'"   
   set rs_funcao2=conn_db.execute(str_SQL_Fun_Neg)
   str_MacroPerfil = rs_funcao2("MCPR_NR_SEQ_MACRO_PERFIL")
   str_Desc_Macro = rs_funcao2("MCPE_TX_DESC_MACRO_PERFIL")
end if
'set rs_funcao=conn_db.execute(str_SQL_Fun_Neg)
'===========================================================================================================================
str_SQL_Mac_Per = ""
str_SQL_Mac_Per = str_SQL_Mac_Per & " SELECT DISTINCT "
str_SQL_Mac_Per = str_SQL_Mac_Per & " MCPE_TX_NOME_TECNICO, "
str_SQL_Mac_Per = str_SQL_Mac_Per & " FUNE_CD_FUNCAO_NEGOCIO, MCPR_NR_SEQ_MACRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " FROM " & Session("PREFIXO") & "MACRO_PERFIL"
str_SQL_Mac_Per = str_SQL_Mac_Per & " WHERE "
str_SQL_Mac_Per = str_SQL_Mac_Per & " MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 
'response.Write("<p>" & str_SQL_Mac_Per)
set rs_MacPer=conn_db.execute(str_SQL_Mac_Per)

IF str_Opt2 = 3 Then
   str_SQL_Mac_Per = ""
   str_SQL_Mac_Per = str_SQL_Mac_Per & " SELECT "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " MCPE_TX_NOME_TECNICO "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPE_TX_DESC_MACRO_PERFIL "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,FUNE_CD_FUNCAO_NEGOCIO"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPR_NR_SEQ_MACRO_PERFIL"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPE_TX_DESC_DETA_MACRO_PERFIL"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " ,MCPE_TX_ESPECIFICACAO"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " FROM " & Session("PREFIXO") & "MACRO_PERFIL"
   str_SQL_Mac_Per = str_SQL_Mac_Per & " WHERE "
   str_SQL_Mac_Per = str_SQL_Mac_Per & " MCPR_NR_SEQ_MACRO_PERFIL =" & str_MacroPerfil
   set rs_MacPer2=conn_db.execute(str_SQL_Mac_Per)
   str_funcao = rs_MacPer2("FUNE_CD_FUNCAO_NEGOCIO")
   str_Desc_Macro = rs_MacPer2("MCPE_TX_DESC_MACRO_PERFIL")
   str_desc_detal = rs_MacPer2("MCPE_TX_DESC_DETA_MACRO_PERFIL")
   str_espec = rs_MacPer2("MCPE_TX_ESPECIFICACAO")   
end if
set rs_MacPer=conn_db.execute(str_SQL_Mac_Per)
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
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"&selOnda="+document.frm1.selOnda.value+"'");
}
function MM_goToURL5() { //v3.0
  var i, args=MM_goToURL5.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc=3&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso=0'");
}

function manda1()
{
window.location.href='rel_micro_R3.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncPrinc='+document.frm1.selFuncPrinc.value+'&selMacroPerfil='+document.frm1.selMacroPerfil.value+'&pOPT2='+document.frm1.txtOPT2.value
}

function Confirma3()
{
if (document.frm1.ID2.value == "")
     { 
	 alert("Você deve especificar um CENÁRIO");
     document.frm1.ID2.focus();
     return;
     }	 

	 else
     {
	  document.frm1.submit();
	 }
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
<form name="frm1" method="post" action="gera_rel_micro_R3.asp">
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
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Cenario/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center">
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Cenario/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Cenario/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../Cenario/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Cenario/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../Cenario/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../Cenario/confirma_f02.gif" width="24" height="24" border="0"></a></td>
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
  <table width="97%" border="0" cellpadding="0" cellspacing="5" name="tblSubProcesso" height="147">
    <tr> 
      <td width="33%" height="1"></td>
      <td width="57%" height="1"> </td>
      <td width="5%" height="1"> </td>
      <td width="17%" height="1"> </td>
    </tr>
    <tr> 
      <td width="33%" height="1"><input type="hidden" name="txtOPT2" value="<%=str_OPT%>"></td>
      <td width="57%" height="1"> <input type="hidden" name="txtOpc" value="1"> 
        <p align="left"><font color="#330099" face="Verdana" size="3">Relatório 
          de Micro-Perfil criados no R/3</font></td>
      <td width="5%" height="1"> <%'=str_Opc%> </td>
      <td width="17%" height="1"> <%'=str_MegaProcesso%> <%'=str_Processo%> </td>
    </tr>
    <tr> 
      <td width="33%" height="17"> </td>
      <td width="57%" height="17"> </td>
      <td width="5%" height="17"> </td>
      <td width="17%" height="17"> </td>
    </tr>
    <tr> 
      <td width="33%" height="25"> <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Mega-Processo 
          :</font></b></font></div></td>
      <td width="57%" height="25"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso"  onChange="javascript:document.frm1.txtOPT2.value = 1;manda1()">
          <% 
		  if str_Opc <> "1" then %>
          <option value="0" selected>== TODOS ==</option>
          <% else %>
          <option value="0" >== TODOS ==</option>
          <% end if %>
          <%
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
      <td height="21"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          R/3 :</b></font></div></td>
      <td bgcolor="#FFFFFF" height="21"><b>
        <select size="1" name="selFuncPrinc" onChange="javascript:document.frm1.txtOPT2.value = 2;manda1()">
          <option value="0"><b>== Selecione uma Funcao R/3 ==</b></option>
          <%do until rs_funcao.eof=true
       if trim(str_Funcao)=trim(rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")) then
       %>
          <option selected value=<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>><b><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs_funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></b></option>
          <%ELSE%>
          <option value=<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>><b><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs_funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></b></option>
          <%
		end if
		rs_funcao.movenext
		loop
		%>
        </select>
        </b></td>
      <td height="21"></td>
      <td height="21"></td>
    </tr>
    <tr> 
      <td height="21"><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>ou 
          Macro Perfil :</b></font></div></td>
      <td bgcolor="#FFFFFF" height="21"><b>
        <select size="1" name="selMacroPerfil" onChange="javascript:document.frm1.txtOPT2.value = 3;manda1()">
          <option value="0"><b>== Selecione um Macro Perfil ==</b></option>
          <%do until rs_MacPer.eof=true
           if trim(str_MacroPerfil)=trim(rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")) then
          %>
          <option value="<%=rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")%>" selected><b><%=rs_MacPer("MCPE_TX_NOME_TECNICO")%></b></option>
          <%ELSE%>
          <option value="<%=rs_MacPer("MCPR_NR_SEQ_MACRO_PERFIL")%>"><b><%=rs_MacPer("MCPE_TX_NOME_TECNICO")%></b></option>
          <% end if
        rs_MacPer.movenext
        loop
        %>
        </select>
        </b></td>
      <td height="21"></td>
      <td height="21"></td>
    </tr>
    <tr> 
      <td width="33%" height="21"></td>
      <%
      if request("SEM")=1 THEN
      ORD="Cenário não encontrado!"
      else
      ORD=""
      end if
      %>
      <td width="57%" bgcolor="#FFFFFF" height="21"><font color="#800000" size="2" face="Verdana"><b><%=ord%></b></font> </td>
      <td width="5%" height="21"></td>
      <td width="17%" height="21"></td>
    </tr>
  </table>
  </form>
<p>&nbsp;</p>
</body>
</html>
