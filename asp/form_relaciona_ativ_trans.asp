<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
Dim str_Opc
dim str_MegaProcesso
dim str_Processo
dim str_SubProcesso
dim str_Atividade

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Atividade = "0"

str_Opc = Request("txtOpc")
str_MegaProcesso= Request("txtMegaProcesso")
str_Processo = Request("txtProcesso")
str_SubProcesso = Request("txtSubProcesso")
str_Atividade = Request("txtAtividade")

if str_MegaProcesso = "" or str_Processo = "" or str_SubProcesso = "" or str_Atividade = "" then
	response.redirect("http://S6000WS10.corp.petrobras.biz/xproc/" & "erro/erro_param_relac_trans.htm")
end if

if str_MegaProcesso = "0" or str_Processo = "0" or str_SubProcesso = "0" or str_Atividade = "0" then
	'response.redirect(application(ga_str_URL) & "/erro/erro_param_relac_trans.htm" 
	response.redirect("http://S6000WS10.corp.petrobras.biz/xproc/" & "erro/erro_param_relac_trans.htm")
	
end if

'int_MegaProcesso= Request.Form("selMegaProcesso")
'int_Processo = Request.Form("SelProcesso")
'str_NovoProcesso = UCase(Request.Form("txtNovoProcesso"))
'int_SubProcesso = Request.Form("SelSubProcesso")
'str_NovoSubProcesso = UCase(Request.Form("txtNovoSubProcesso"))
'int_Atividade = Request.Form("selAtividade")
'str_NovaAtividade = UCase(Request.Form("txtNovaAtividade"))

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

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
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "ATIVIDADE.ATIV_CD_ATIVIDADE = " & str_Atividade

Set rdsAtividade = Conn_db.Execute(str_SQL_Atividade)

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "TRANSACAO "
IF str_Opc <> "1" then
	str_SQL_Modulo = str_SQL_Modulo & " WHERE " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & str_Modulo
end if

str_SQL_Modulo = ""
str_SQL_Modulo = str_SQL_Modulo & " SELECT "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO, "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3.MODU_TX_DESC_MODULO "
str_SQL_Modulo = str_SQL_Modulo & " FROM " & Session("PREFIXO") & "MODULO_R3 "
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Projeto Sinergia</title>
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function Atualiza_txtTransacao(valor) 
{
if (document.frmCadastraPR.selTransacao.selectedIndex ==  -1)
     { alert("A seleção de uma Transação é obrigatória !");
       document.frmCadastraPR.selTransacao.focus();
     }
else
    {
document.frmCadastraPR.txtTranSelecionada.value =  document.frmCadastraPR.txtTranSelecionada.value  + '/' + document.frmCadastraPR.selTransacao.options[document.frmCadastraPR.selTransacao.selectedIndex].text;
	 }
}

function Confirma() 
{ 
if (document.frmCadastraPR.selCurso.selectedIndex == 0)
     { 
	 alert("A seleção de um Curso é obrigatório !");
     document.frmCadastraPR.selCurso.focus();
     return;
     }
  valor=document.frmCadastraPR.txtPergunta.value;
  tamanho = valor.length;
  if (tamanho == 0)
     { 
	 alert("O preenchimento do campo Pergunta é obrigatório !");
     document.frmCadastraPR.txtPergunta.focus();
     return;
     }
  valor=document.frmCadastraPR.txtResposta.value
  tamanho = valor.length   
  if (tamanho == 0)
     { 
	 alert("O preenchimento do campo Resposta é obrigatório !");
     document.frmCadastraPR.txtResposta.focus();
     return;
     }
  valor=document.frmCadastraPR.txtPalChaveTot.value
  tamanho = valor.length   
  if (tamanho == 0)
     { 
	 alert("O preenchimento do campo Palavra Chave é obrigatório !");
     document.frmCadastraPR.txtPalChaveTot.focus();
     return;
	 }
	 else
     {
	  document.frmCadastraPR.submit();
	 }
 }

function Limpa(){
	document.frmCadastraPR.reset();
}

function MM_changePropOO(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  var obj2 = MM_findObj(theValue);
  //alert("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../imagens/continua_F02.gif','../imagens/continua2_F02.gif')">
<form name="frmCadastraPR" method="post" action="file:///M|/aspscript/gra_Perg_Resp.asp">
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
  <table width="104%" cellspacing="0" cellpadding="0" border="0">
    <tr> 
      <td width="100%" bgcolor="#FFFFFF"> 
        <div align="left"></div>
        <table border=0 cellpadding=0 cellspacing=0 width="714" align="center">
          <tr> 
            <td valign=top width="714"> 
              <table width="75%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="20%">&nbsp;</td>
                  <td width="6%">&nbsp;</td>
                  <td width="74%">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega-Processo:&nbsp; 
                      </font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("MEPR_CD_MEGA_PROCESSO")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("MEPR_TX_DESC_MEGA_PROCESSO")%></font></font></td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Processo: 
                      &nbsp;</font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("PROC_CD_PROCESSO")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("PROC_TX_DESC_PROCESSO")%></font></font></td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Sub-Processo: 
                      &nbsp;</font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("SUPR_CD_SUB_PROCESSO")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("SUPR_TX_DESC_SUB_PROCESSO")%></font></font></td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade: 
                      &nbsp;</font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("ATIV_CD_ATIVIDADE")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("ATIV_TX_DESC_ATIVIDADE")%></font></font></td>
                </tr>
              </table>
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"> 
              <table width="443" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="233"> 
                    <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#003399">Transa&ccedil;&otilde;es 
                      existentes </font></font></div>
                  </td>
                  <td width="210"> 
                    <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#003399">Selecionada</font></font></div>
                  </td>
                </tr>
                <tr bgcolor="#0099CC"> 
                  <td width="233" height="7"></td>
                  <td width="210" height="7"></td>
                </tr>
                <tr> 
                  <td width="233"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2">M&oacute;dulo</font></font></td>
                  <td width="210">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="233"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
                    <select name="selModulo" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=2&amp;selModulo=',this);return document.MM_returnValue">
                      <% 
		  if str_Opc <> "1" then %>
                      <option value="0" selected><font color="#003300">Selecione 
                      um Módulo</font></option>
                      <% else %>
                      <option value="0" ><font color="#003300">Selecione um Módulo</font></option>
                      <% end if %>
                      <%Set rdsModulo = Conn_db.Execute(str_SQL_Modulo)
While (NOT rdsModulo.EOF)
  
           if (Trim(str_Modulo) = Trim(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)) then %>
                      <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>" selected ><font color="#003300"><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></font></option>
                      <% else %>
                      <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>"><font color="#003300"><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></font></option>
                      <% end if %>
                      <%
  rdsModulo.MoveNext()
Wend
If (rdsModulo.CursorType > 0) Then
  rdsModulo.MoveFirst
Else
  rdsModulo.Requery
End If
rdsModulo.Close
set rdsModulo = Nothing
%>
                    </select>
                    </b></font></font></td>
                  <td width="210">&nbsp;</td>
                </tr>
              </table>
              <table width="400" border="0">
                <tr> 
                  <td width="277"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
                    </b></font></font></td>
                  <td width="36">&nbsp;</td>
                  <td width="210">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="277">&nbsp;</td>
                  <td width="36" align="center">&nbsp;</td>
                  <td width="210">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="277"> <b> 
                    <select name="selTransacao" size="8">
                      <%Set rdsTransacao = Conn_db.Execute(str_SQL_Transacao)
While (NOT rdsTransacao.EOF)
%>
                      <option value="<%=(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" ><%=(rdsTransacao.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></option>
                      <%
  rdsTransacao.MoveNext()
Wend
If (rdsTransacao.CursorType > 0) Then
  rdsTransacao.MoveFirst
Else
  rdsTransacao.Requery
End If
%>
                    </select>
                    </b> 
                    <div align="center"></div>
                  </td>
                  <td width="36" align="center"> 
                    <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
                      <tr> 
                        <td><a href="javascript:;" onClick="Atualiza_txtTransacao(this.value)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','../imagens/continua_F02.gif',1)"><img name="Image8" border="0" src="../imagens/continua_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img015','','../imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC1','TEXTAREA')"><img name="img015" border="0" src="../imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td width="210"> 
                    <div align="center"> <font color="#000080">
                      <textarea name="txtTranSelecionada" rows="8" wrap="PHYSICAL" cols="20"></textarea>
                      </font></div>
                  </td>
                </tr>
                <tr>
                  <td width="277"><font color="#000080"> </font></td>
                  <td width="36" align="center">&nbsp;</td>
                  <td width="210">&nbsp;</td>
                </tr>
              </table>
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"> 
              <div align="center"> </div>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="100%">&nbsp; </td>
    </tr>
  </table>
    <p>&nbsp;</p>
  </form>
</body>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
Dim str_Opc
dim str_MegaProcesso
dim str_Processo
dim str_SubProcesso
dim str_Atividade

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Atividade = "0"

str_Opc = Request("txtOpc")
str_MegaProcesso= Request("txtMegaProcesso")
str_Processo = Request("txtProcesso")
str_SubProcesso = Request("txtSubProcesso")
str_Atividade = Request("txtAtividade")

if str_MegaProcesso = "" or str_Processo = "" or str_SubProcesso = "" or str_Atividade = "" then
	response.redirect("http://S6000WS10.corp.petrobras.biz/xproc/" & "erro/erro_param_relac_trans.htm")
end if

if str_MegaProcesso = "0" or str_Processo = "0" or str_SubProcesso = "0" or str_Atividade = "0" then
	'response.redirect(application(ga_str_URL) & "/erro/erro_param_relac_trans.htm" 
	response.redirect("http://S6000WS10.corp.petrobras.biz/xproc/" & "erro/erro_param_relac_trans.htm")
	
end if

'int_MegaProcesso= Request.Form("selMegaProcesso")
'int_Processo = Request.Form("SelProcesso")
'str_NovoProcesso = UCase(Request.Form("txtNovoProcesso"))
'int_SubProcesso = Request.Form("SelSubProcesso")
'str_NovoSubProcesso = UCase(Request.Form("txtNovoSubProcesso"))
'int_Atividade = Request.Form("selAtividade")
'str_NovaAtividade = UCase(Request.Form("txtNovaAtividade"))

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

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
str_SQL_Atividade = str_SQL_Atividade & " AND " & Session("PREFIXO") & "ATIVIDADE.ATIV_CD_ATIVIDADE = " & str_Atividade

Set rdsAtividade = Conn_db.Execute(str_SQL_Atividade)

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "TRANSACAO "
IF str_Opc <> "1" then
	str_SQL_Modulo = str_SQL_Modulo & " WHERE " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & str_Modulo
end if

str_SQL_Modulo = ""
str_SQL_Modulo = str_SQL_Modulo & " SELECT "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO, "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3.MODU_TX_DESC_MODULO "
str_SQL_Modulo = str_SQL_Modulo & " FROM " & Session("PREFIXO") & "MODULO_R3 "
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Projeto Sinergia</title>
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function Atualiza_txtTransacao(valor) 
{
if (document.frmCadastraPR.selTransacao.selectedIndex ==  -1)
     { alert("A seleção de uma Transação é obrigatória !");
       document.frmCadastraPR.selTransacao.focus();
     }
else
    {
document.frmCadastraPR.txtTranSelecionada.value =  document.frmCadastraPR.txtTranSelecionada.value  + '/' + document.frmCadastraPR.selTransacao.options[document.frmCadastraPR.selTransacao.selectedIndex].text;
	 }
}

function Confirma() 
{ 
if (document.frmCadastraPR.selCurso.selectedIndex == 0)
     { 
	 alert("A seleção de um Curso é obrigatório !");
     document.frmCadastraPR.selCurso.focus();
     return;
     }
  valor=document.frmCadastraPR.txtPergunta.value;
  tamanho = valor.length;
  if (tamanho == 0)
     { 
	 alert("O preenchimento do campo Pergunta é obrigatório !");
     document.frmCadastraPR.txtPergunta.focus();
     return;
     }
  valor=document.frmCadastraPR.txtResposta.value
  tamanho = valor.length   
  if (tamanho == 0)
     { 
	 alert("O preenchimento do campo Resposta é obrigatório !");
     document.frmCadastraPR.txtResposta.focus();
     return;
     }
  valor=document.frmCadastraPR.txtPalChaveTot.value
  tamanho = valor.length   
  if (tamanho == 0)
     { 
	 alert("O preenchimento do campo Palavra Chave é obrigatório !");
     document.frmCadastraPR.txtPalChaveTot.focus();
     return;
	 }
	 else
     {
	  document.frmCadastraPR.submit();
	 }
 }

function Limpa(){
	document.frmCadastraPR.reset();
}

function MM_changePropOO(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  var obj2 = MM_findObj(theValue);
  //alert("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../imagens/continua_F02.gif','../imagens/continua2_F02.gif')">
<form name="frmCadastraPR" method="post" action="file:///M|/aspscript/gra_Perg_Resp.asp">
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
  <table width="104%" cellspacing="0" cellpadding="0" border="0">
    <tr> 
      <td width="100%" bgcolor="#FFFFFF"> 
        <div align="left"></div>
        <table border=0 cellpadding=0 cellspacing=0 width="714" align="center">
          <tr> 
            <td valign=top width="714"> 
              <table width="75%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="20%">&nbsp;</td>
                  <td width="6%">&nbsp;</td>
                  <td width="74%">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega-Processo:&nbsp; 
                      </font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("MEPR_CD_MEGA_PROCESSO")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("MEPR_TX_DESC_MEGA_PROCESSO")%></font></font></td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Processo: 
                      &nbsp;</font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("PROC_CD_PROCESSO")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("PROC_TX_DESC_PROCESSO")%></font></font></td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Sub-Processo: 
                      &nbsp;</font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("SUPR_CD_SUB_PROCESSO")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("SUPR_TX_DESC_SUB_PROCESSO")%></font></font></td>
                </tr>
                <tr> 
                  <td width="20%"> 
                    <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade: 
                      &nbsp;</font></font></div>
                  </td>
                  <td width="6%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("ATIV_CD_ATIVIDADE")%></font> -</font></td>
                  <td width="74%"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2"><%=rdsAtividade("ATIV_TX_DESC_ATIVIDADE")%></font></font></td>
                </tr>
              </table>
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"> 
              <table width="443" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="233"> 
                    <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#003399">Transa&ccedil;&otilde;es 
                      existentes </font></font></div>
                  </td>
                  <td width="210"> 
                    <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#003399">Selecionada</font></font></div>
                  </td>
                </tr>
                <tr bgcolor="#0099CC"> 
                  <td width="233" height="7"></td>
                  <td width="210" height="7"></td>
                </tr>
                <tr> 
                  <td width="233"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2">M&oacute;dulo</font></font></td>
                  <td width="210">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="233"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
                    <select name="selModulo" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=2&amp;selModulo=',this);return document.MM_returnValue">
                      <% 
		  if str_Opc <> "1" then %>
                      <option value="0" selected><font color="#003300">Selecione 
                      um Módulo</font></option>
                      <% else %>
                      <option value="0" ><font color="#003300">Selecione um Módulo</font></option>
                      <% end if %>
                      <%Set rdsModulo = Conn_db.Execute(str_SQL_Modulo)
While (NOT rdsModulo.EOF)
  
           if (Trim(str_Modulo) = Trim(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)) then %>
                      <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>" selected ><font color="#003300"><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></font></option>
                      <% else %>
                      <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>"><font color="#003300"><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></font></option>
                      <% end if %>
                      <%
  rdsModulo.MoveNext()
Wend
If (rdsModulo.CursorType > 0) Then
  rdsModulo.MoveFirst
Else
  rdsModulo.Requery
End If
rdsModulo.Close
set rdsModulo = Nothing
%>
                    </select>
                    </b></font></font></td>
                  <td width="210">&nbsp;</td>
                </tr>
              </table>
              <table width="400" border="0">
                <tr> 
                  <td width="277"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
                    </b></font></font></td>
                  <td width="36">&nbsp;</td>
                  <td width="210">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="277">&nbsp;</td>
                  <td width="36" align="center">&nbsp;</td>
                  <td width="210">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="277"> <b> 
                    <select name="selTransacao" size="8">
                      <%Set rdsTransacao = Conn_db.Execute(str_SQL_Transacao)
While (NOT rdsTransacao.EOF)
%>
                      <option value="<%=(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" ><%=(rdsTransacao.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></option>
                      <%
  rdsTransacao.MoveNext()
Wend
If (rdsTransacao.CursorType > 0) Then
  rdsTransacao.MoveFirst
Else
  rdsTransacao.Requery
End If
%>
                    </select>
                    </b> 
                    <div align="center"></div>
                  </td>
                  <td width="36" align="center"> 
                    <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
                      <tr> 
                        <td><a href="javascript:;" onClick="Atualiza_txtTransacao(this.value)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','../imagens/continua_F02.gif',1)"><img name="Image8" border="0" src="../imagens/continua_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img015','','../imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC1','TEXTAREA')"><img name="img015" border="0" src="../imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td width="210"> 
                    <div align="center"> <font color="#000080">
                      <textarea name="txtTranSelecionada" rows="8" wrap="PHYSICAL" cols="20"></textarea>
                      </font></div>
                  </td>
                </tr>
                <tr>
                  <td width="277"><font color="#000080"> </font></td>
                  <td width="36" align="center">&nbsp;</td>
                  <td width="210">&nbsp;</td>
                </tr>
              </table>
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"> 
              <div align="center"> </div>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="100%">&nbsp; </td>
    </tr>
  </table>
    <p>&nbsp;</p>
  </form>
</body>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
</html>