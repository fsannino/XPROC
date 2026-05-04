 
<%
str_Opc = Request("txtOpc")

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if

if (Request("selMegaProcesso2") <> "") then 
    str_MegaProcesso2 = Request("selMegaProcesso2")
else
    str_MegaProcesso2 = str_MegaProcesso
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

if (Request("selAtividadeCarga") <> "") then 
    str_AtividadeCarga = Request("selAtividadeCarga")
else
    str_AtividadeCarga = "0"
end if

if (Request("selFuncao") <> "") then 
    str_Funcao = Request("selFuncao")
else
    str_Funcao = "0"
end if


set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso

Set rdsMegaProcesso = Conn_db.Execute(str_SQL_MegaProc)

If Not rdsMegaProcesso.EOF then
   ls_str_DescMegaProcesso = rdsMegaProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
else
   ls_str_DescMegaProcesso = "Não achou Mega"
end if
rdsMegaProcesso.close
set rdsMegaProcesso = Nothing

str_SQL_Funcao = ""
str_SQL_Funcao = str_SQL_Funcao & " SELECT "
str_SQL_Funcao = str_SQL_Funcao & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
str_SQL_Funcao = str_SQL_Funcao & " ," & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
str_SQL_Funcao = str_SQL_Funcao & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO"
str_SQL_Funcao = str_SQL_Funcao & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'"

Set rdsFuncao = Conn_db.Execute(str_SQL_Funcao)

If Not rdsFuncao.EOF then
   ls_str_TituloFuncao = rdsFuncao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
else
   ls_str_TituloFuncao = "Não achou Mega"
end if
rdsFuncao.close
set rdsFuncao = Nothing

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

str_SQL_Proc = ""
str_SQL_Proc = str_SQL_Proc & " SELECT DISTINCT "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " ," & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " ," & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO "
str_SQL_Proc = str_SQL_Proc & " FROM " & Session("PREFIXO") & "PROCESSO INNER JOIN "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "MEGA_PROCESSO ON "
str_SQL_Proc = str_SQL_Proc & " " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Proc = str_SQL_Proc & " WHERE " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2 
str_SQL_Proc = str_SQL_Proc & " order by  " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO "

str_SQL_Sub_Proc = ""
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " SELECT  DISTINCT "
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
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " order by  " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO "

str_SQL_Atividade_Carga = ""
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " SELECT  DISTINCT "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " ," & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA"
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT  DISTINCT  "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO not in ("
str_SQL_Transacao = str_SQL_Transacao & " SELECT TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2
str_SQL_Transacao = str_SQL_Transacao & " AND PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao = str_SQL_Transacao & " AND SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao = str_SQL_Transacao & " AND ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "')"

str_SQL_Transacao_Cad = ""
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " SELECT  DISTINCT  "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " WHERE " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso2
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'"

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
   var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value+"&selMegaProcesso2="+document.frm1.selMegaProcesso2.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"'");
}
function MM_goToURL2() { //v3.0
   var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value+"'");
}

function carrega_txt(fbox){
   document.frm1.txtTranSelecionada.value = "";
   for(var i=0; i<fbox.options.length; i++) 
     {
     document.frm1.txtTranSelecionada.value = document.frm1.txtTranSelecionada.value + "," + fbox.options[i].value;
     }
}
function Confirma2(){ 
	  document.frm1.submit();
}
function Confirma(){ 
   if (document.frm1.selMegaProcesso2.selectedIndex == 0) { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
   if (document.frm1.selProcesso.selectedIndex == 0) { 
	 alert("Selecione um Proceso.");
     document.frm1.selProcesso.focus();
     return;
     }	 
   if (document.frm1.selSubProcesso.selectedIndex == 0) { 
	 alert("Selecione um Sub Proceso.");
     document.frm1.selSubProcesso.focus();
     return;
     }	 
   if (document.frm1.selAtividadeCarga.selectedIndex == 0) { 
	 alert("Selecione uma Atividasde de Carga.");
     document.frm1.selAtividadeCarga.focus();
     return;
     }	 
	 else
     {
	 document.frm1.txtDescMegaProcesso2.value = document.frm1.selMegaProcesso2.options[document.frm1.selMegaProcesso2.selectedIndex].text;	 
	 document.frm1.txtDescProcesso.value = document.frm1.selProcesso.options[document.frm1.selProcesso.selectedIndex].text;	 
	 document.frm1.txtDescSubProcesso.value = document.frm1.selSubProcesso.options[document.frm1.selSubProcesso.selectedIndex].text;	 
	 document.frm1.txtDescAtividadeCarga.value = document.frm1.selAtividadeCarga.options[document.frm1.selAtividadeCarga.selectedIndex].text;	 	 
 	 carrega_txt(document.frm1.list2);
	 document.frm1.submit();
	 }
}
function Limpa(){
	document.frm1.reset();
}
function exibe_transacao(){
	window.open("exibe_funcao_trans.asp?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value,"_blank","width=620,height=240,history=0,scrollbars=1,titlebar=0,resizable=0")
}
function MM_changePropOO(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  var obj2 = MM_findObj(theValue);
  //alert("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
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
<script language="javascript" src="../Planilhas/js/troca_lista_sem_ordem.js"></script>
<body topmargin="0" leftmargin="0" onLoad="MM_preloadImages('../imagens/continua_F02.gif','../imagens/continua2_F02.gif','../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif')" bgcolor="#FFFFFF">
<form method="POST" action="grava_funcao_m_p_s_a_trans.asp" name="frm1">
  <table width="98%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
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
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
            <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
            <td width="26">&nbsp;</td>
            <td width="195"></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28"></td>
            <td width="26">&nbsp;</td>
            <td width="159"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table border="0" width="736" height="94" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td width="10" height="21">&nbsp;</td>
      <td width="160" height="21">&nbsp;</td>
      <td height="21" width="499"><font face="Verdana" color="#330099" size="3">Relação 
        Fun&ccedil;&atilde;o R/3 x Transação</font></td>
      <td height="21" width="67"> 
        <div align="center"><font face="Verdana" size="1" color="#330099">Transações 
          inseridas</font></div>
      </td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="160" height="14"> 
        <div align="right"><font size="2"><font face="Verdana" color="#330099">Mega-Processo: 
          </font></font></div>
      </td>
      <td width="499" height="14"><b><font face="Verdana" size="2" color="#330099"><%=str_MegaProcesso%>-<%=ls_str_DescMegaProcesso%> 
        <input type="hidden" name="txtMegaProcesso" size="46" value="<%=str_MegaProcesso%>">
        <input type="hidden" name="txtDescMegaProcesso" size="46" value="<%=ls_str_DescMegaProcesso%>">
        </font></b></td>
      <td height="14" width="67"> 
        <div align="center"><b><font face="Verdana" size="2" color="#330099"><a href="javascript:exibe_transacao()"><img border="0" src="../../imagens/icon_empresa.gif" align="absmiddle"></a></font></b></div>
      </td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="160" height="14"> 
        <div align="right"><font face="Verdana" size="2" color="#330099">Fun&ccedil;&atilde;o R/3: </font></div>
      </td>
      <td height="14" width="499"><b><font face="Verdana" size="2" color="#330099"><%=str_Funcao%> 
        <input type="hidden" name="txtFuncao" size="46" value="<%=str_Funcao%>">
        </font></b></td>
      <td height="14" width="67"> 
        <div align="center"><b><font face="Verdana" size="2" color="#330099"></font></b> 
        </div>
      </td>
    </tr>
    <tr> 
      <td width="10" height="23">&nbsp;</td>
      <td width="160" height="23"> 
        <input type="hidden" name="txtTranSelecionada" size="20">
        <b> 
        <%'=str_MegaProcesso%>
        <%'=str_MegaProcesso2%>
        <%'=str_Processo%>
        <%'=str_SubProcesso%>
        <%'=str_AtividadeCraga%>
        </b> </td>
      <td height="23" width="499"><b><font face="Verdana" size="2" color="#330099"><%=ls_str_TituloFuncao%></font><font face="Verdana" size="1" color="#330099"> 
        <input type="hidden" name="txtDescFuncao" size="46" value="<%=ls_str_TituloFuncao%>">
        </font></b></td>
      <td height="23" width="67">&nbsp;</td>
    </tr>
    <tr bgcolor="#0099CC"> 
      <td width="10" height="4"></td>
      <td width="160" height="4"></td>
      <td height="4" width="499"></td>
      <td height="4" width="67"></td>
    </tr>
  </table>
  <table border="0" width="779" height="6" align="center">
    <tr> 
      <td width="155" height="1"> </td>
      <td height="1" colspan="4"> </td>
    </tr>
    <tr> 
      <td width="155" height="1"> </td>
      <td height="1" colspan="4"> </td>
    </tr>
    <tr> 
      <td width="155" height="4"> 
        <p align="right"><font face="Verdana" size="2" color="#330099">Mega-Processo 
          :</font> 
      </td>
      <td height="4" colspan="4"> <font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso2" onChange="MM_goToURL1('self','cad_funcao_transacao_outro.asp');return document.MM_returnValue">
          <% 
		  if str_Opc <> "1" then %>
          <option value="0" selected>Selecione um Mega Processo</option>
          <% else %>
          <option value="0" >Selecione um Mega Processo</option>
          <% end if %>
          <%Set rdsMegaProcesso = Conn_db.Execute(str_SQL_MegaProc)
While (NOT rdsMegaProcesso.EOF)
         if (Trim(str_MegaProcesso2) = Trim(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)) then %>
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
        <b><font face="Verdana" size="1" color="#330099"> 
        <input type="hidden" name="txtDescMegaProcesso2" size="46" value="0">
        </font></b></font></td>
    </tr>
    <tr> 
      <td width="155" height="4"> 
        <div align="right"><font face="Verdana" size="2" color="#330099">Processo 
          : </font></div>
      </td>
      <td height="4" colspan="4"> 
        <select name="selProcesso" onChange="MM_goToURL1('self','cad_funcao_transacao_outro.asp',this);return document.MM_returnValue">
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
        <b><font face="Verdana" size="1" color="#330099"> 
        <input type="hidden" name="txtDescProcesso" size="46" value="0">
        </font></b> </td>
    </tr>
    <tr> 
      <td width="155" height="15"> 
        <p align="right"><font face="Verdana" size="2" color="#330099">Sub-Processo 
          : </font> 
      </td>
      <td height="15" colspan="4"> 
        <select name="selSubProcesso" onChange="MM_goToURL1('self','cad_funcao_transacao_outro.asp',this);return document.MM_returnValue">
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
        <b><font face="Verdana" size="1" color="#330099"> 
        <input type="hidden" name="txtDescSubProcesso" size="46" value="0">
        </font></b> </td>
    </tr>
  </table>
  <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="242">
    <tr> 
      <td width="392" height="4" bgcolor="#0099CC"></td>
      <td width="349" height="4" bgcolor="#0099CC"></td>
    </tr>
    <tr> 
      <td colspan="2" height="7"></td>
    </tr>
    <tr> 
      <td colspan="2" height="31"> 
        <div align="center"> 
          <table width="82%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="22%"> 
                <div align="center"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade 
                  </font></font><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b> 
                  </b></font></div>
              </td>
              <td width="78%"> <font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b> 
                <select name="selAtividadeCarga" onChange="MM_goToURL1('self','cad_funcao_transacao_outro.asp');return document.MM_returnValue">
                  <% 
		  if str_Opc <> "1" then %>
                  <option value="0" selected>Selecione uma Atividade</option>
                  <% else %>
                  <option value="0" >Selecione uma Atividade</option>
                  <% end if %>
                  <%Set rdsAtividadeCarga = Conn_db.Execute(str_SQL_Atividade_Carga)
While (NOT rdsAtividadeCarga.EOF)
         if (Trim(str_AtividadeCarga) = Trim(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)) then %>
                  <option value="<%=(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)%>" selected ><%=(rdsAtividadeCarga.Fields.Item("ATCA_TX_DESC_ATIVIDADE").Value)%></option>
                  <% else %>
                  <option value="<%=(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)%>" ><%=(rdsAtividadeCarga.Fields.Item("ATCA_TX_DESC_ATIVIDADE").Value)%></option>
                  <% end if %>
                  <%
  rdsAtividadeCarga.MoveNext()
Wend
If (rdsAtividadeCarga.CursorType > 0) Then
  rdsAtividadeCarga.MoveFirst
Else
  rdsAtividadeCarga.Requery
End If
rdsAtividadeCarga.Close
set rdsAtividadeCarga = Nothing
%>
                </select>
                <font face="Verdana" size="1" color="#330099"> 
                <input type="hidden" name="txtDescAtividadeCarga" size="46" value="0">
                </font></b></font></td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
    <tr> 
      <td height="7" width="392"></td>
      <td height="7" width="349"></td>
    </tr>
    <tr> 
      <td height="7" bgcolor="#0099CC" width="392"> 
        <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Transa&ccedil;&otilde;es 
          existentes</font></font></div>
      </td>
      <td height="7" bgcolor="#0099CC" width="349"> 
        <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Selecionada</font></font></div>
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="10"> 
        <%'=str_AtividadeCarga%>
        <%'=str_Modulo%>
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="10"> 
        <table width="616" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="266"> 
              <div align="center"> <b> 
                <select name="list1" size="8" multiple>
                  <%Set rdsTransacao = Conn_db.Execute(str_SQL_Transacao)
While (NOT rdsTransacao.EOF)
%>
                  <option value="<%=(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" ><%=(rdsTransacao.Fields.Item("TRAN_CD_TRANSACAO").Value) & "-" & (rdsTransacao.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></option>
                  <%
  rdsTransacao.MoveNext()
Wend
If (rdsTransacao.CursorType > 0) Then
  rdsTransacao.MoveFirst
Else
  rdsTransacao.Requery
End If
rdsTransacao.close
set rdsTransacao = Nothing
%>
                </select>
                </b></div>
            </td>
            <td width="24" align="center"> 
              <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
                <tr> 
                  <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2,0)"><img name="Image16" border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a></td>
                </tr>
                <tr> 
                  <td height="25">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="25"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24"></a></td>
                </tr>
              </table>
            </td>
            <td width="290"> 
              <div align="center"> <font color="#000080"> 
                <select name="list2" size="8" multiple>
                  <%Set rdsTransacao_cad = Conn_db.Execute(str_SQL_Transacao_Cad)
While (NOT rdsTransacao_cad.EOF)
%>
                  <option value="<%=(rdsTransacao_cad.Fields.Item("TRAN_CD_TRANSACAO").Value)%>"><%=(rdsTransacao_cad.Fields.Item("TRAN_CD_TRANSACAO").Value) & "-" & (rdsTransacao_cad.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></option>
                  <%
  rdsTransacao_cad.MoveNext()
Wend
If (rdsTransacao_cad.CursorType > 0) Then
  rdsTransacao_cad.MoveFirst
Else
  rdsTransacao_cad.Requery
End If
rdsTransacao_cad.close
set rdsTransacao_cad = Nothing
%>
                </select>
                </font></div>
            </td>
            <td width="1">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3">&nbsp;</td>
            <td width="1">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
                a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></div>
            </td>
            <td width="1">&nbsp;</td>
          </tr>
          <tr> 
            <td width="266"><font color="#000080"> 
              <%'=str_SQL_Transacao_Cad%>
              </font></td>
            <td width="24" align="center">&nbsp;</td>
            <td width="290"> 
              <input type="hidden" name="txtTranSelecionada2">
            </td>
            <td width="1">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  </form>
</body>
</html>
