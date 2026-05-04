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

if (Request("txtOpc") <> "") then 
   str_Opc = Request("txtOpc")
   if (Request("selAtividade") <> "") then 
       str_Atividade = Request("selAtividade")
	end if   
   if (Request("p_MegaProc") <> "") then 
       str_MegaProcesso = Request("p_MegaProc")
	end if   
   if (Request("p_Proc") <> "") then 
       str_Processo = Request("p_Proc")
	end if   
   if (Request("p_SubProc") <> "") then 
       str_SubProcesso = Request("p_SubProc")
	end if   

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

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "TRANSACAO "

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

function Atualiza_txtPalChaveTot(valor) 
{
if (document.frmCadastraPR.selPalChaveTot.selectedIndex ==  -1)
     { alert("A seleção de uma Palavra Chave é obrigatória !");
       document.frmCadastraPR.selPalChaveTot.focus();
     }
else
    {
document.frmCadastraPR.txtPalChaveTot.value =  document.frmCadastraPR.txtPalChaveTot.value  + ' ' + document.frmCadastraPR.selPalChaveTot.options[document.frmCadastraPR.selPalChaveTot.selectedIndex].text;
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
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('file:///M|/imagens/continua_F02.gif')" bgcolor="#FFFFFF">
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
<form name="frmCadastraPR" method="post" action="file:///M|/aspscript/gra_Perg_Resp.asp">
    
  <table width="104%" cellspacing="0" cellpadding="0" border="0">
    <tr> 
      <td width="100%" bgcolor="#FFFFFF"> 
        <div align="left"></div>
        <table border=0 cellpadding=0 cellspacing=0 width="714" align="center">
          <tr> 
            <td valign=top width="714"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Mega-Processo: 
              <%=rdsAtividade("MEPR_CD_MEGA_PROCESSO")%></b></font></font> -<font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("MEPR_TX_DESC_MEGA_PROCESSO")%></b></font></font> 
          <tr> 
            <td valign=top width="714">- <font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b>Processo: 
              </b></font><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("PROC_CD_PROCESSO")%></b></font></font> -<font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("PROC_TX_DESC_PROCESSO")%></b></font></font> <font face="Arial, Helvetica, sans-serif" size="2"><b> 
              </b></font></font> 
          <tr> 
            <td valign=top width="714"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Sub-Processo: 
              </b></font><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("SUPR_CD_SUB_PROCESSO")%></b></font></font> -<font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("SUPR_TX_DESC_SUB_PROCESSO")%></b></font></font> <font face="Arial, Helvetica, sans-serif" size="2"> 
              </font></font> 
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Atividade 
              <select name="selAtividade" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=2&amp;selAtividade=',this);return document.MM_returnValue">
                <% 
		  if str_Opc <> "1" then %>
                <option value="0" selected><font color="#003300">Selecione um 
                Processo</font></option>
                <% else %>
                <option value="0" ><font color="#003300">Selecione um Processo</font></option>
                <% end if %>
                <%Set rdsAtividade = Conn_db.Execute(str_SQL_Atividade)
While (NOT rdsAtividade.EOF)
  
           if (Trim(str_Atividade) = Trim(rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)) then %>
                <option value="<%=(rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)%>" selected ><font color="#003300"><%=(rdsAtividade.Fields.Item("ATIV_TX_DESC_ATIVIDADE").Value)%></font></option>
                <% else %>
                <option value="<%=rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)%>"><font color="#003300"><%=(rdsAtividade.Fields.Item("ATIV_TX_DESC_ATIVIDADE").Value)%></font></option>
                <% end if %>
                <%
  rdsAtividade.MoveNext()
Wend
If (rdsAtividade.CursorType > 0) Then
  rdsAtividade.MoveFirst
Else
  rdsAtividade.Requery
End If

rdsAtividade.Close
set rdsAtividade = Nothing
%>
              </select>
              </b></font></font> 
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"> 
              <table width="704" border="0">
                <tr> 
                  <td width="269"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Transa&ccedil;&otilde;es 
                    existentes </b></font></font></td>
                  <td width="52">&nbsp;</td>
                  <td width="121"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Selecionada</b></font></font></td>
                  <td width="41">&nbsp;</td>
                  <td width="199">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="269"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>M&oacute;dulo 
                    <select name="selTransacao" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=3&amp;selTransacao=',this);return document.MM_returnValue">
                    </select>
                                  <select name="selAtividade" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=2&amp;selAtividade=',this);return document.MM_returnValue">
                <% 
		  if str_Opc <> "1" then %>
                <option value="0" selected><font color="#003300">Selecione um 
                Processo</font></option>
                <% else %>
                <option value="0" ><font color="#003300">Selecione um Processo</font></option>
                <% end if %>
                <%Set rdsTransacao = Conn_db.Execute(str_SQL_Transacao)
While (NOT rdsTransacao.EOF)
  
           if (Trim(str_Transacao) = Trim(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)) then %>
                <option value="<%=(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" selected ><font color="#003300"><%=(rdsTransacao.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></font></option>
                <% else %>
                <option value="<%=rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)%>"><font color="#003300"><%=(rdsTransacao.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></font></option>
                <% end if %>
                <%
  rdsTransacao.MoveNext()
Wend
If (rdsTransacao.CursorType > 0) Then
  rdsTransacao.MoveFirst
Else
  rdsTransacao.Requery
End If
rdsTransacao.Close
set rdsTransacao = Nothing
%>
              </select>
</b></font></font></td>
                  <td width="52">&nbsp;</td>
                  <td width="121">&nbsp;</td>
                  <td width="41">&nbsp;</td>
                  <td width="199">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="269"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
                    <select name="selPalChaveTot" size="8">
                      <%
While (NOT rsdPalavraChave.EOF)
%>
                      <option value="<%=(rsdPalavraChave.Fields.Item("PALA_TX_PALAVRA_CHAVE").Value)%>" ><%=(rsdPalavraChave.Fields.Item("PALA_TX_PALAVRA_CHAVE").Value)%></option>
                      <%
  rsdPalavraChave.MoveNext()
Wend
If (rsdPalavraChave.CursorType > 0) Then
  rsdPalavraChave.MoveFirst
Else
  rsdPalavraChave.Requery
End If
%>
                    </select>
                    </b></font></td>
                  <td width="52" align="center"> 
                    <table width="53%" cellpadding="0" cellspacing="0" border="0">
                      <tr> 
                        <td><a href="javascript:;" onClick="Atualiza_txtPalChaveTot(this.value)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','file:///M|/imagens/continua_F02.gif',1)"><img name="Image8" border="0" src="file:///M|/imagens/continua_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img015','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC1','TEXTAREA')"><img name="img015" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td width="121"><font color="#000080"> 
                    <textarea name="txtPalChaveTot" rows="8" wrap="PHYSICAL" cols="20"></textarea>
                    </font></td>
                  <td width="41"> 
                    <table width="100%" border="0">
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC1','TEXTAREA')"><img name="img01" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img011','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC2','TEXTAREA')"><img name="img011" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img012','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC3','TEXTAREA')"><img name="img012" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img013','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC4','TEXTAREA')"><img name="img013" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img014','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC5','TEXTAREA')"><img name="img014" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td width="199">&nbsp; </td>
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
<%
rdsCursos.Close()
%>
<%
rsdPalavraChave.Close()
%>
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

if (Request("txtOpc") <> "") then 
   str_Opc = Request("txtOpc")
   if (Request("selAtividade") <> "") then 
       str_Atividade = Request("selAtividade")
	end if   
   if (Request("p_MegaProc") <> "") then 
       str_MegaProcesso = Request("p_MegaProc")
	end if   
   if (Request("p_Proc") <> "") then 
       str_Processo = Request("p_Proc")
	end if   
   if (Request("p_SubProc") <> "") then 
       str_SubProcesso = Request("p_SubProc")
	end if   

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

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "TRANSACAO "

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

function Atualiza_txtPalChaveTot(valor) 
{
if (document.frmCadastraPR.selPalChaveTot.selectedIndex ==  -1)
     { alert("A seleção de uma Palavra Chave é obrigatória !");
       document.frmCadastraPR.selPalChaveTot.focus();
     }
else
    {
document.frmCadastraPR.txtPalChaveTot.value =  document.frmCadastraPR.txtPalChaveTot.value  + ' ' + document.frmCadastraPR.selPalChaveTot.options[document.frmCadastraPR.selPalChaveTot.selectedIndex].text;
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
//-->
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('file:///M|/imagens/continua_F02.gif')" bgcolor="#FFFFFF">
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
<form name="frmCadastraPR" method="post" action="file:///M|/aspscript/gra_Perg_Resp.asp">
    
  <table width="104%" cellspacing="0" cellpadding="0" border="0">
    <tr> 
      <td width="100%" bgcolor="#FFFFFF"> 
        <div align="left"></div>
        <table border=0 cellpadding=0 cellspacing=0 width="714" align="center">
          <tr> 
            <td valign=top width="714"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Mega-Processo: 
              <%=rdsAtividade("MEPR_CD_MEGA_PROCESSO")%></b></font></font> -<font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("MEPR_TX_DESC_MEGA_PROCESSO")%></b></font></font> 
          <tr> 
            <td valign=top width="714">- <font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b>Processo: 
              </b></font><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("PROC_CD_PROCESSO")%></b></font></font> -<font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("PROC_TX_DESC_PROCESSO")%></b></font></font> <font face="Arial, Helvetica, sans-serif" size="2"><b> 
              </b></font></font> 
          <tr> 
            <td valign=top width="714"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Sub-Processo: 
              </b></font><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("SUPR_CD_SUB_PROCESSO")%></b></font></font> -<font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2"><b><%=rdsAtividade("SUPR_TX_DESC_SUB_PROCESSO")%></b></font></font> <font face="Arial, Helvetica, sans-serif" size="2"> 
              </font></font> 
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Atividade 
              <select name="selAtividade" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=2&amp;selAtividade=',this);return document.MM_returnValue">
                <% 
		  if str_Opc <> "1" then %>
                <option value="0" selected><font color="#003300">Selecione um 
                Processo</font></option>
                <% else %>
                <option value="0" ><font color="#003300">Selecione um Processo</font></option>
                <% end if %>
                <%Set rdsAtividade = Conn_db.Execute(str_SQL_Atividade)
While (NOT rdsAtividade.EOF)
  
           if (Trim(str_Atividade) = Trim(rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)) then %>
                <option value="<%=(rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)%>" selected ><font color="#003300"><%=(rdsAtividade.Fields.Item("ATIV_TX_DESC_ATIVIDADE").Value)%></font></option>
                <% else %>
                <option value="<%=rdsAtividade.Fields.Item("ATIV_CD_ATIVIDADE").Value)%>"><font color="#003300"><%=(rdsAtividade.Fields.Item("ATIV_TX_DESC_ATIVIDADE").Value)%></font></option>
                <% end if %>
                <%
  rdsAtividade.MoveNext()
Wend
If (rdsAtividade.CursorType > 0) Then
  rdsAtividade.MoveFirst
Else
  rdsAtividade.Requery
End If

rdsAtividade.Close
set rdsAtividade = Nothing
%>
              </select>
              </b></font></font> 
          <tr> 
            <td valign=top width="714">&nbsp; 
          <tr> 
            <td valign=top width="714"> 
              <table width="704" border="0">
                <tr> 
                  <td width="269"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Transa&ccedil;&otilde;es 
                    existentes </b></font></font></td>
                  <td width="52">&nbsp;</td>
                  <td width="121"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>Selecionada</b></font></font></td>
                  <td width="41">&nbsp;</td>
                  <td width="199">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="269"><font color="#003300">- <font face="Arial, Helvetica, sans-serif" size="2"><b>M&oacute;dulo 
                    <select name="selTransacao" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=3&amp;selTransacao=',this);return document.MM_returnValue">
                    </select>
                                  <select name="selAtividade" onChange="MM_goToURL('self','form_relaciona_ativ_trans.asp?txtOpc=2&amp;selAtividade=',this);return document.MM_returnValue">
                <% 
		  if str_Opc <> "1" then %>
                <option value="0" selected><font color="#003300">Selecione um 
                Processo</font></option>
                <% else %>
                <option value="0" ><font color="#003300">Selecione um Processo</font></option>
                <% end if %>
                <%Set rdsTransacao = Conn_db.Execute(str_SQL_Transacao)
While (NOT rdsTransacao.EOF)
  
           if (Trim(str_Transacao) = Trim(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)) then %>
                <option value="<%=(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" selected ><font color="#003300"><%=(rdsTransacao.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></font></option>
                <% else %>
                <option value="<%=rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)%>"><font color="#003300"><%=(rdsTransacao.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></font></option>
                <% end if %>
                <%
  rdsTransacao.MoveNext()
Wend
If (rdsTransacao.CursorType > 0) Then
  rdsTransacao.MoveFirst
Else
  rdsTransacao.Requery
End If
rdsTransacao.Close
set rdsTransacao = Nothing
%>
              </select>
</b></font></font></td>
                  <td width="52">&nbsp;</td>
                  <td width="121">&nbsp;</td>
                  <td width="41">&nbsp;</td>
                  <td width="199">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="269"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
                    <select name="selPalChaveTot" size="8">
                      <%
While (NOT rsdPalavraChave.EOF)
%>
                      <option value="<%=(rsdPalavraChave.Fields.Item("PALA_TX_PALAVRA_CHAVE").Value)%>" ><%=(rsdPalavraChave.Fields.Item("PALA_TX_PALAVRA_CHAVE").Value)%></option>
                      <%
  rsdPalavraChave.MoveNext()
Wend
If (rsdPalavraChave.CursorType > 0) Then
  rsdPalavraChave.MoveFirst
Else
  rsdPalavraChave.Requery
End If
%>
                    </select>
                    </b></font></td>
                  <td width="52" align="center"> 
                    <table width="53%" cellpadding="0" cellspacing="0" border="0">
                      <tr> 
                        <td><a href="javascript:;" onClick="Atualiza_txtPalChaveTot(this.value)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image8','','file:///M|/imagens/continua_F02.gif',1)"><img name="Image8" border="0" src="file:///M|/imagens/continua_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img015','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC1','TEXTAREA')"><img name="img015" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td width="121"><font color="#000080"> 
                    <textarea name="txtPalChaveTot" rows="8" wrap="PHYSICAL" cols="20"></textarea>
                    </font></td>
                  <td width="41"> 
                    <table width="100%" border="0">
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC1','TEXTAREA')"><img name="img01" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img011','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC2','TEXTAREA')"><img name="img011" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img012','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC3','TEXTAREA')"><img name="img012" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img013','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC4','TEXTAREA')"><img name="img013" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                      <tr> 
                        <td><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img014','','file:///M|/imagens/continua2_F02.gif',1)" onClick="MM_changePropOO('txtPalChaveTot','','value','txtNovaPC5','TEXTAREA')"><img name="img014" border="0" src="file:///M|/imagens/continua2_F01.gif" width="24" height="24"></a></td>
                      </tr>
                    </table>
                  </td>
                  <td width="199">&nbsp; </td>
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
<%
rdsCursos.Close()
%>
<%
rsdPalavraChave.Close()
%>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
