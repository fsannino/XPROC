<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
str_Opc = Request("txtOpc")

if Request("selMegaProcesso") = "" then
   str_MegaProcesso = "0"
else
	str_MegaProcesso = Request("selMegaProcesso")
end if

if Request("selProcesso") = "" then
   str_Processo = "0"
else
   str_Processo = Request("selProcesso")
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

str_SQL_Empresa_Unid = ""
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " SELECT "
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " " & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_CD_NR_EMPRESA "
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " ," & Session("PREFIXO") & "EMPRESA_UNIDADE.EMPR_TX_NOME_EMPRESA "
str_SQL_Empresa_Unid = str_SQL_Empresa_Unid & " FROM " & Session("PREFIXO") & "EMPRESA_UNIDADE "

str_SQL_Max_Seq_Sub_Proc = ""
str_SQL_Max_Seq_Sub_Proc = str_SQL_Max_Seq_Sub_Proc & " SELECT "
str_SQL_Max_Seq_Sub_Proc = str_SQL_Max_Seq_Sub_Proc & " MAX(SUPR_NR_SEQUENCIA) AS MaxSeq"
str_SQL_Max_Seq_Sub_Proc = str_SQL_Max_Seq_Sub_Proc & " FROM " & Session("PREFIXO") & "SUB_PROCESSO"
str_SQL_Max_Seq_Sub_Proc = str_SQL_Max_Seq_Sub_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO, "
str_SQL_Max_Seq_Sub_Proc = str_SQL_Max_Seq_Sub_Proc & " PROC_CD_PROCESSO"
str_SQL_Max_Seq_Sub_Proc = str_SQL_Max_Seq_Sub_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Max_Seq_Sub_Proc = str_SQL_Max_Seq_Sub_Proc & " AND PROC_CD_PROCESSO = " & str_Processo

Set rdsMaxSeqSubProcesso= Conn_db.Execute(str_SQL_Max_Seq_Sub_Proc)
if rdsMaxSeqSubProcesso.EOF then
   ls_int_MaxSubProcesso = 0
else
   IF not IsNull(rdsMaxSeqSubProcesso("MaxSeq")) then
      ls_int_MaxSubProcesso = rdsMaxSeqSubProcesso("MaxSeq")   
   else
      ls_int_MaxSubProcesso = 0
   end if	  
end if
rdsMaxSeqSubProcesso.close
set rdsMaxSeqSubProcesso = nothing

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
function carrega_txt(fbox) {
document.frm1.txtEmpSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtEmpSelecionada.value = document.frm1.txtEmpSelecionada.value + "," + fbox.options[i].value;
   }
}

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
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc=3&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso=0'");
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
if(document.frm1.valImpacto.value==0 )
     { 
	 alert("Você deve selecionar o IMPACTO do sub-Processo.");
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
function Recarrega_tela() { //v3.0

if (document.frm1.txtOpc.value != "1")
     { 
    document.history.go();
     }	 
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
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../imagens/novo_registro_02.gif','../imagens/atualiza_02.gif','../imagens/continua_F02.gif','../imagens/continua2_F02.gif')">
<form name="frm1" method="post" action="grava_sub_processo.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
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
  <table width="91%" border="0" cellpadding="0" cellspacing="0" height="515">
    <tr> 
      <td width="6%" height="21">&nbsp;</td>
      <td width="26%" height="21"><%'=str_Opc%><%'=str_MegaProcesso%></td>
      <td width="68%" colspan="7" height="21"><%'=str_Processo%></td>
    </tr>
    <tr> 
      <td width="6%" height="21">&nbsp;</td>
      <td width="26%" height="21"><%'=Session("MegaProcesso")%></td>
      <td width="68%" colspan="7" height="21"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Novo 
        Sub-Processo</font></td>
    </tr>
    <tr> 
      <td width="6%" height="21">&nbsp;</td>
      <td width="26%" height="21"><%'=ls_int_MaxSubProcesso%></td>
      <td width="68%" colspan="7" height="21"> 
        <div align="right"> 
          <input type="hidden" name="txtOpc" value="<%=str_Opc%>">
        </div>
      </td>
    </tr>
    <tr> 
      <td width="6%" height="25">&nbsp;</td>
      <td width="26%" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Mega-Processo</b></font></td>
      <td width="68%" colspan="7" height="25"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','cadas_sub_processo.asp');return document.MM_returnValue">
          <% 
		  if str_Opc <> "1" then %>
          <option value="0" selected>Selecione um Mega Processo</option>
          <% else %>
          <option value="0" >Selecione um Mega Processo</option>
          <% end if %>
          <%Set rdsMegaProcesso= Conn_db.Execute(str_SQL_MegaProc)
While (NOT rdsMegaProcesso.EOF)
  
           if (Trim(str_MegaProcesso) = Trim(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)) then %>
          <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>" selected ><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
          <% else %>
          <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>"><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
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
      </td>
    </tr>
    <tr> 
      <td width="6%" height="21">&nbsp;</td>
      <td width="26%" height="21"><%'=str_SQL_Max_Seq_Sub_Proc%></td>
      <td width="68%" colspan="7" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="6%" height="28">&nbsp;</td>
      <td width="26%" height="28"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Processos</b></font></td>
      <td width="68%" colspan="7" height="28"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="59%"> 
              <select name="selProcesso" onChange="MM_goToURL2('self','cadas_sub_processo.asp',this);return document.MM_returnValue">
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
            <td width="30"> 
              <table width="60" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="30"> 
                    <div align="center"><a href="javascript:MM_goToURL3('self','cadas_Processo.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image11','','../imagens/novo_registro_02.gif',1)"><img name="Image11" border="0" src="../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui novo Processo"></a></div>
                  </td>
                  <td width="30"> 
                    <div align="center"><a href="JavaScript:history.go()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image12','','../imagens/atualiza_02.gif',1)"><img name="Image12" border="0" src="../imagens/atualiza_02_off.gif" width="22" height="22" alt="Recarrega novo Processo"></a></div>
                  </td>
                </tr>
              </table>
            </td>
            <td width="0%">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="6%" height="21">&nbsp;</td>
      <td width="26%" height="21">&nbsp;</td>
      <td width="68%" colspan="7" height="21">&nbsp; </td>
    </tr>
    <tr> 
      <td width="6%" height="27">&nbsp;</td>
      <td width="26%" height="27"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Novo 
        Sub-Processo</b></font></td>
      <td width="68%" colspan="7" height="27"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoSubProc1" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq1" size="7" value="<%=ls_int_MaxSubProcesso+10%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="6%" height="22"></td>
      <td width="26%" height="22"></td>
      <td width="68%" colspan="7" height="22"> </td>
    </tr>
    <tr> 
      <td width="6%" height="21"></td>
      <td width="26%" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Impacto</b></font></td>
      <td width="4%" height="21">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b><input type="radio" value="1" name="selImpacto" onClick="document.frm1.valImpacto.value=this.value"></b></font> </td>
      <td width="6%" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Alto</b></font> </td>
      <td width="3%" height="21">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b><input type="radio" value="2" name="selImpacto" onClick="document.frm1.valImpacto.value=this.value"></b></font> </td>
      <td width="8%" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Médio</b></font> </td>
      <td width="3%" height="21">
        <p align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b><input type="radio" value="3" name="selImpacto" onClick="document.frm1.valImpacto.value=this.value"></b></font> </td>
      <td width="23%" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Baixo</b></font> </td>
      <td width="22%" height="21"><input type="hidden" name="valImpacto" size="20" value="0"> </td>
    </tr>
    <tr> 
      <td width="6%" height="21">&nbsp;</td>
      <td width="26%" height="21">&nbsp;</td>
      <td width="68%" colspan="7" height="21">&nbsp; </td>
    </tr>
    <tr> 
      <td width="6%" height="247">&nbsp;</td>
      <td width="26%" valign="top" height="247"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Selecione 
        Empresa/Unidade</b></font></td>
      <td width="68%" colspan="7" height="247"> 
        <table width="95%" border="0" align="left" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="52%"> 
              <div align="center"> <b> 
                <select name="list1" size="8" multiple>
                  <%Set rdsEmp_Unid = Conn_db.Execute(str_SQL_Empresa_Unid)
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
                </b></div>
            </td>
            <td width="5%" align="center"> 
              <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
                <tr> 
                  <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2)"><img name="Image16" border="0" src="../imagens/continua_F01.gif" width="24" height="24"></a></td>
                </tr>
                <tr> 
                  <td height="25">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="25"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1)"><img name="img01511" border="0" src="../imagens/continua2_F01.gif" width="24" height="24"></a></td>
                </tr>
              </table>
            </td>
            <td width="28%"> 
              <div align="center"><font color="#000080"> 
                <select name="list2" size="8" multiple>
                </select>
                </font></div>
            </td>
          </tr>
          <tr> 
            <td colspan="3">&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3">
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                Use a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></div>
            </td>
          </tr>
          <tr> 
            <td colspan="3"> 
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td width="52%"> 
              <%'=str_SQL_Sub_Proc%>
              <input type="hidden" name="txtEmpSelecionada">
            </td>
            <td width="5%" align="center">&nbsp;</td>
            <td width="28%"> 
              <%'=str_SQL_Proc%>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="6%" height="19">&nbsp;</td>
      <td width="26%" height="19">&nbsp;</td>
      <td width="68%" colspan="7" height="19">&nbsp; </td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
