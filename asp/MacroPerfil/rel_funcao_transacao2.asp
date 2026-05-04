 
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

if (Request("txtMacroPerfil") <> "") then 
    str_MacroPerfil = Request("txtMacroPerfil")
else
    str_MacroPerfil = "0"
end if

if (Request("txtNomeTecnico") <> "") then 
    str_NomeTecnico = Request("txtNomeTecnico")
else
    str_NomeTecnico = "0"
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

str_SQL_Proc_Sub_Proc = ""
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " SELECT DISTINCT "
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, "
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO, "
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, "
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " " & Session("PREFIXO") & "PROCESSO.PROC_TX_DESC_PROCESSO, "
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO,"
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, "
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE" 
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN " & Session("PREFIXO") & "SUB_PROCESSO ON " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO"
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA"                                                  
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " INNER JOIN " & Session("PREFIXO") & "PROCESSO ON " & Session("PREFIXO") & "SUB_PROCESSO.PROC_CD_PROCESSO = " & Session("PREFIXO") & "PROCESSO.PROC_CD_PROCESSO AND " & Session("PREFIXO") & "SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
if str_Processo <> "0" then
	str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & str_Processo
    if str_SubProcesso <> "0" then
	   str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
       if str_AtividadeCarga <> "0" then
	      str_SQL_Proc_Sub_Proc = str_SQL_Proc_Sub_Proc & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
       end if 
    end if  
end if

'response.write str_SQL_Proc_Sub_Proc
Set rdsProc_Sub_Proc = Conn_db.Execute(str_SQL_Proc_Sub_Proc)

str_SQL_Ativ_Trans = ""
str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " SELECT DISTINCT "
'str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, "
'str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE," 
str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO, "
str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO, "
str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " " & Session("PREFIXO") & "RELACAO_FINAL.MODU_CD_MODULO"
str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " FROM " & Session("PREFIXO") & "RELACAO_FINAL INNER JOIN " & Session("PREFIXO") & "ATIVIDADE_CARGA ON " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA"
str_SQL_Ativ_Trans = str_SQL_Ativ_Trans & " INNER JOIN " & Session("PREFIXO") & "TRANSACAO ON " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"

ls_Seq = 0
int_Conta_Transacao = 0
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
function MM_goToURL5() { //v3.0
  var i,x,args=MM_goToURL5.arguments; document.MM_returnValue = false;
  //for (i=0; i<(args.length-1); i+=4) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"&p_CenarioChSequencia="+args[3]+"'");
  //alert(document.frm1.imgMarca1.src);
  x=MM_findObj(args[4])
  // NÃO CONSIGO TESTAR EM DESENV OU PRODUÇÃO
  if(x.src == "http://S6000WS10.corp.petrobras.biz/xproc/imagens/func_tran_nao_marcada.gif") {
	 window.open("inc_Funcao_Trans.asp?"+args[3],"_blank","width=100,height=100,history=0,scrollbars=0,titlebar=0,resizable=0")
     MM_swapImage(x.name,'','../../imagens/func_tran_marcada.gif',1);
    // window.open("exibe_funcao_trans.asp?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value,"_blank","width=620,height=240,history=0,scrollbars=1,titlebar=0,resizable=0")

	}
	else 
	{
  //  if(document.frm1.imgMarca1.src == "http://S6000WS10.corp.petrobras.biz/xproc/imagens/b03.gif") 
	 window.open("exc_Funcao_Trans.asp?"+args[3],"_blank","width=100,height=100,history=0,scrollbars=0,titlebar=0,resizable=0")	
    MM_swapImage(x.name,'','../../imagens/func_tran_nao_marcada.gif',1);

    }
  //for (i=0; i<(args.length-1); i+=3) eval(args[i]+".location='"+args[i+1]+"?"+args[3]+"'");
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
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>
<script language="javascript" src="js/troca_lista_sem_ordem.js"></script>
</head>
<body topmargin="0" leftmargin="0" onLoad="MM_preloadImages('../imagens/continua_F02.gif','../imagens/continua2_F02.gif','../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif')" bgcolor="#FFFFFF" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" action="../Funcao/grava_funcao_m_p_s_a_trans.asp" name="frm1">
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
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
  <table border="0" width="736" height="94" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td width="10" height="21">&nbsp;</td>
      <td width="160" height="21">&nbsp;</td>
      <td height="21" width="499"><font face="Verdana" color="#330099" size="3">Fun&ccedil;&atilde;o R/3 x Transação</font></td>
      <td height="21" width="67"> </td>
    </tr>
    <tr> 
      <td width="10" height="14"></td>
      <td width="160" height="14"></td>
      <td width="499" height="14"></td>
      <td height="14" width="67"></td>
    </tr>
    <tr> 
      <td width="10" height="14"></td>
      <td width="160" height="14"> 
        <div align="right"><font size="2"><font face="Verdana" color="#330099">Id 
          Macro Perfil:</font></font></div>
      </td>
      <td width="499" height="14"><font face="Verdana" size="2" color="#330099"><%=str_MacroPerfil%></font></td>
      <td height="14" width="67"></td>
    </tr>
    <tr> 
      <td width="10" height="14"></td>
      <td width="160" height="14"> 
        <div align="right"><font size="2"><font face="Verdana" color="#330099">Nome 
          T&eacute;cnico:</font></font></div>
      </td>
      <td width="499" height="14"><b><font face="Verdana" size="2" color="#330099"><%=str_NomeTecnico%></font></b></td>
      <td height="14" width="67"></td>
    </tr>
    <tr> 
      <td width="10" height="14"></td>
      <td width="160" height="14"> </td>
      <td width="499" height="14"></td>
      <td height="14" width="67"> </td>
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
      <td height="14" width="67"> </td>
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
        <select name="selMegaProcesso2" onChange="MM_goToURL1('self','rel_funcao_transacao.asp');return document.MM_returnValue">
          <%Set rdsMegaProcesso = Conn_db.Execute(str_SQL_MegaProc)
While (NOT rdsMegaProcesso.EOF)
         if (Trim(str_MegaProcesso2) = Trim(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)) then %>
          <option value="<%=(rdsMegaProcesso.Fields.Item("MEPR_CD_MEGA_PROCESSO").Value)%>" selected ><%=(rdsMegaProcesso.Fields.Item("MEPR_TX_DESC_MEGA_PROCESSO").Value)%></option>
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
        <select name="selProcesso" onChange="MM_goToURL1('self','rel_funcao_transacao.asp',this);return document.MM_returnValue">
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
        <select name="selSubProcesso" onChange="MM_goToURL1('self','rel_funcao_transacao.asp',this);return document.MM_returnValue">
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
  <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="180">
    <tr> 
      <td width="352" height="4" bgcolor="#0099CC"></td>
      <td width="314" height="4" bgcolor="#0099CC"></td>
    </tr>
    <tr> 
      <td colspan="2" height="7"></td>
    </tr>
    <tr> 
      <td colspan="2" height="31"> 
        <table width="82%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr> 
            <td width="22%"> 
              <div align="center"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade 
                </font></font><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b> 
                </b></font></div>
            </td>
            <td width="78%"> <font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b> 
              <select name="selAtividadeCarga" onChange="MM_goToURL1('self','rel_funcao_transacao.asp');return document.MM_returnValue">
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
      </td>
    </tr>
    <tr> 
      <td height="2" bgcolor="#0099CC" width="352"> 
        <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Transa&ccedil;&otilde;es 
          existentes</font></font></div>
      </td>
      <td height="2" bgcolor="#0099CC" width="314"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="10"> 
        <% If rdsProc_Sub_Proc.EOF then
		        ls_msg = "Não encontrado transaçãoes para esta seleção"
			 else
				ls_msg = ""
			 end if			 	
		  %>
        <%=ls_msg%> 
        <% do while Not rdsProc_Sub_Proc.EOF
		  	'rdsProc_Sub_Proc("PROC_CD_PROCESSO") rdsProc_Sub_Proc("SUPR_CD_SUB_PROCESSO") 
		   %>
        <table width="666" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="572">&nbsp;</td>
            <td width="54">&nbsp;</td>
            <td width="40">&nbsp;</td>
          </tr>
          <tr> 
            <td width="572"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Processo 
              :<b> <%=rdsProc_Sub_Proc("PROC_TX_DESC_PROCESSO")%></b></font></td>
            <td width="54">&nbsp;</td>
            <td width="40">&nbsp;</td>
          </tr>
          <tr> 
            <td width="572"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Sub 
              Processo :<b> <%=rdsProc_Sub_Proc("SUPR_TX_DESC_SUB_PROCESSO")%></b></font></td>
            <td width="54">&nbsp;</td>
            <td width="40">&nbsp;</td>
          </tr>
          <tr> 
            <td width="572"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Atividade 
              :<b> <%=rdsProc_Sub_Proc("ATCA_TX_DESC_ATIVIDADE")%></b></font></td>
            <td width="54">&nbsp;</td>
            <td width="40">&nbsp;</td>
          </tr>
        </table>
        <%
		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans
		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " WHERE " & Session("PREFIXO") & "RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO = " & rdsProc_Sub_Proc("MEPR_CD_MEGA_PROCESSO")
		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND " & Session("PREFIXO") & "RELACAO_FINAL.PROC_CD_PROCESSO = " & rdsProc_Sub_Proc("PROC_CD_PROCESSO")
		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND " & Session("PREFIXO") & "RELACAO_FINAL.SUPR_CD_SUB_PROCESSO = " & rdsProc_Sub_Proc("SUPR_CD_SUB_PROCESSO")
		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND " & Session("PREFIXO") & "RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = " & rdsProc_Sub_Proc("ATCA_CD_ATIVIDADE_CARGA")		
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND " & Session("PREFIXO") & "RELACAO_FINAL.TRAN_CD_TRANSACAO not in ("
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " SELECT TRAN_CD_TRANSACAO"
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO"
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " WHERE MEPR_CD_MEGA_PROCESSO = " & rdsProc_Sub_Proc("MEPR_CD_MEGA_PROCESSO")
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND PROC_CD_PROCESSO = " & rdsProc_Sub_Proc("PROC_CD_PROCESSO")
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND SUPR_CD_SUB_PROCESSO = " & rdsProc_Sub_Proc("SUPR_CD_SUB_PROCESSO")
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND ATCA_CD_ATIVIDADE_CARGA = " & rdsProc_Sub_Proc("ATCA_CD_ATIVIDADE_CARGA")
'		str_SQL_Ativ_Trans2 = str_SQL_Ativ_Trans2 & " AND FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "')"
		
		Set rdsAtiv_Trans = Conn_db.Execute(str_SQL_Ativ_Trans2)
		
		%>
        <table width="666" border="0" cellspacing="3" cellpadding="0">
          <tr> 
            <td width="60"> 
              <div align="center"></div>
            </td>
            <td width="530"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Transa&ccedil;&atilde;o</font></div>
            </td>
            <td width="64">&nbsp;</td>
          </tr>
          <% Do While Not rdsAtiv_Trans.EOF 
		  if ls_Cor_Linha = "#F7F7F7" then
             ls_Cor_Linha = "#FFFFFF"
		  else		  
		     ls_Cor_Linha = "#F7F7F7"
		  end if
		  %>
          <tr bgcolor="<%=ls_Cor_Linha%>"> 
            <td width="60"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%'=rdsAtiv_Trans("ATCA_TX_DESC_ATIVIDADE")%>
              </font></td>
            <td width="530"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><a href="#" onclick="javascript:alert('Em desenvolvimento...')"><%=rdsAtiv_Trans("TRAN_CD_TRANSACAO")%> </b>- <%=rdsAtiv_Trans("TRAN_TX_DESC_TRANSACAO")%></a></font></td>
            <td width="64" bgcolor="<%=ls_Cor_Linha%>"> 
              <div align="center"> 
                <% int_Conta_Transacao = int_Conta_Transacao + 1 %>
                <% ls_Seq = ls_Seq + 1 %>
                <input type="hidden" name="txtOpt" value="<%=ls_Seq%>">
                <% ls_SQL = ""
				   ls_SQL = ls_SQL & " SELECT "
				   ls_SQL = ls_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO, "
                   ls_SQL = ls_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO, "
				   ls_SQL = ls_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO, "
				   ls_SQL = ls_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO, "
				   ls_SQL = ls_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA, "
    			   ls_SQL = ls_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO, "
    			   ls_SQL = ls_SQL & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MODU_CD_MODULO"
				   ls_SQL = ls_SQL & " FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO"
				   ls_SQL = ls_SQL & " WHERE " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'" 
				   ls_SQL = ls_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " &  rdsProc_Sub_Proc("MEPR_CD_MEGA_PROCESSO")
				   ls_SQL = ls_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = " & rdsProc_Sub_Proc("PROC_CD_PROCESSO")
				   ls_SQL = ls_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = " &  rdsProc_Sub_Proc("SUPR_CD_SUB_PROCESSO")
				   ls_SQL = ls_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = " & rdsProc_Sub_Proc("ATCA_CD_ATIVIDADE_CARGA")
				   ls_SQL = ls_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = '" & rdsAtiv_Trans("TRAN_CD_TRANSACAO") & "'" 
				   ls_SQL = ls_SQL & " AND " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MODU_CD_MODULO = " & rdsAtiv_Trans("MODU_CD_MODULO")		
				   Set rdsExiste = Conn_db.Execute(ls_SQL)
				   IF rdsExiste.EOF then
				      ls_Imagem = "func_tran_nao_marcada.gif"
                   else
				      ls_Imagem = "func_tran_marcada.gif"
				   end if
				   rdsExiste.close
				   set rdsExiste = Nothing	  				   
				%>
                <a href="edita_objetos_campo_macro_perfil.asp?selMacro_Perfil=<%=str_MacroPerfil%>&selDescMacro_Perfil=<%=str_NomeTecnico%>&selTransacao=<%=rdsAtiv_Trans("TRAN_CD_TRANSACAO")%>&selDescTransacao=<%=rdsAtiv_Trans("TRAN_TX_DESC_TRANSACAO")%>"><img src="../../imagens/b04.gif" width="16" height="16" border="0"></a> 
              </div>
            </td>
          </tr>
          <% 
		  	rdsAtiv_Trans.Movenext
		  Loop 
		  rdsAtiv_Trans.close
		  set rdsAtiv_Trans = Nothing
		  %>
        </table>
        <% 
		  	rdsProc_Sub_Proc.Movenext
		  Loop 
		  rdsProc_Sub_Proc.close
		  set rdsProc_Sub_Proc = Nothing		  
		  %>
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="10">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2" height="10">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2" height="10"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total 
        de transa&ccedil;&otilde;es listadas :<b> <%=int_Conta_Transacao%></b> </font> </td>
    </tr>
  </table>
  <p>&nbsp;</p>
</form>
</body>
</html>
