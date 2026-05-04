<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
Dim str_Opc
dim str_MegaProcesso
dim str_Processo
dim str_SubProcesso
dim str_Modulo

str_Opc = Request("txtOpc")
if str_Opc <> "1" then
   str_AtividadeCarga = Request("selAtividadeCarga")
   str_MegaProcesso= Request("selMegaProcesso")
   str_Processo = Request("selProcesso")
   str_SubProcesso = Request("selSubProcesso")   
   str_Modulo = Request("selModulo")
else
   str_AtividadeCarga = "0"
   str_MegaProcesso = "0"
   str_Processo = "0"
   str_SubProcesso = "0"
   str_Modulo = "0"
end if
if str_Opc = "6" then
   str_Processo = "0"
   str_SubProcesso = "0"
end if
if str_Opc = "7" then
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
   str_SubProcesso = "0"
end if


set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

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
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
Set rdsSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
if Not rdsSubProcesso.EOF then
   str_DescMegaProcesso = rdsSubProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
   str_DescProcesso = rdsSubProcesso("PROC_TX_DESC_PROCESSO")
   str_DescSubProcesso = rdsSubProcesso("SUPR_TX_DESC_SUB_PROCESSO")
else
   str_DescMegaProcesso = ""
   str_DescProcesso = ""
   str_DescSubProcesso = ""
end if
rdsSubProcesso.close
set rdsSubProcesso = Nothing

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

str_SQL_Atividade_Carga = ""
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " SELECT "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " ATCA_CD_ATIVIDADE_CARGA, "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " ATCA_TX_DESC_ATIVIDADE "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA"

str_SQL_Modulo = ""
str_SQL_Modulo = str_SQL_Modulo & " SELECT distinct "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO, "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3.MODU_TX_DESC_MODULO"
str_SQL_Modulo = str_SQL_Modulo & " FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA INNER JOIN"
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA ON "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA"
str_SQL_Modulo = str_SQL_Modulo & " INNER JOIN"
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3 ON "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO"
str_SQL_Modulo = str_SQL_Modulo & " WHERE " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA = "  & str_AtividadeCarga

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA INNER JOIN"
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO = '" & str_Modulo & "'"

'str_SQL_Transacao = ""
'str_SQL_Transacao = str_SQL_Transacao & " SELECT "
'str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO, "
'str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
'    lss_SQL = lss_SQL & " From PRODUTO as b "
'    lss_SQL = lss_SQL & " Where Convert(VarChar(5), b.GRPR_NR_CD_GRUPO_PRODUTO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),b.SUGR_NR_CD_SUB_GRUPO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),b.PROD_NR_CD_PRODUTO) not In "
'    lss_SQL = lss_SQL & " (Select Convert(varchar(5),h.GRPR_NR_CD_GRUPO_PRODUTO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),h.SUGR_NR_CD_SUB_GRUPO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),h.PROD_NR_CD_PRODUTO) "
'    lss_SQL = lss_SQL & " From PRODUTO_EMPRESA as h "
'    lss_SQL = lss_SQL & " Where h.EMPR_NR_CD_EMPRESA = " & txtPRFCod.Text
'    lss_SQL = lss_SQL & " ) "
'    If Len(Trim(txtPRDCodGrupoProduto)) <> 0 Then
'       lss_SQL = lss_SQL & " and b.GRPR_NR_CD_GRUPO_PRODUTO = " & txtPRDCodGrupoProduto.Text
'    End If
'    lss_SQL = lss_SQL & " and b.PROD_TX_SITUACAO_PRODUTO = 'C'"
'    lss_SQL = lss_SQL & " order by b.PROD_TX_NM_PRODUTO "

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA INNER JOIN"
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO = '" & str_Modulo & "'"
str_SQL_Transacao = str_SQL_Transacao & " and Convert(VarChar(5), " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO) "
str_SQL_Transacao = str_SQL_Transacao & "  Not In ("
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO INNER JOIN"
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO = '" & str_Modulo & "')"

str_SQL_Transacao_Cad = ""
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " SELECT "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO INNER JOIN"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " WHERE " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO = '" & str_Modulo & "'"

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
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selModulo="+document.frm1.selModulo.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selModulo="+document.frm1.selModulo.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selModulo="+document.frm1.selModulo.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"'");
}
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
function carrega_txt(fbox) {
document.frm1.txtTranSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtTranSelecionada.value = document.frm1.txtTranSelecionada.value + "[" + fbox.options[i].value;
   }
}

function carrega_txt2(fbox) {
document.frm1.txtTranNaoSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtTranNaoSelecionada.value = document.frm1.txtTranNaoSelecionada.value + "[" + fbox.options[i].value;
   }
}

function Confirma() 
{ 
if (document.frm1.selAtividadeCarga.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatória !");
     document.frm1.selAtividadeCarga.focus();
     return;
     }
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
if (document.frm1.selProcesso.selectedIndex == 0)
     { 
	 alert("Selecione um Proceso ou cadastre um novo.");
     document.frm1.selProcesso.focus();
     return;
     }	 
if (document.frm1.selSubProcesso.selectedIndex == 0)
     { 
	 alert("Selecione um Sub Proceso ou cadastre um novo.");
     document.frm1.selSubProcesso.focus();
     return;
     }	 
if (document.frm1.selModulo.selectedIndex == 0)
     { 
	 alert("A seleção de um Módulo é obrigatória !");
     document.frm1.selModulo.focus();
     return;
     }
if (document.frm1.list2.options.length == 0)
     { 
	 alert("A seleção de uma Transação é obrigatória !");
     document.frm1.list2.focus();
     return;
     }
	 else
     {
	  carrega_txt(document.frm1.list2);
  	  carrega_txt2(document.frm1.list1);
	  alert(document.frm1.txtTranSelecionada.value);
	  alert(document.frm1.txtTranNaoSelecionada.value);
	  document.frm1.txtDsMP.value = document.frm1.selMegaProcesso.options[document.frm1.selMegaProcesso.selectedIndex].text
	  document.frm1.txtDsP.value = document.frm1.selProcesso.options[document.frm1.selProcesso.selectedIndex].text
	  document.frm1.txtDsSP.value = document.frm1.selSubProcesso.options[document.frm1.selSubProcesso.selectedIndex].text
	  document.frm1.txtDsA.value = document.frm1.selAtividadeCarga.options[document.frm1.selAtividadeCarga.selectedIndex].text
	  document.frm1.txtDsM.value = document.frm1.selModulo.options[document.frm1.selModulo.selectedIndex].text
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script language="javascript" src="js/troca_lista.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../imagens/continua2_F02.gif','../imagens/continua_F02.gif','../imagens/novo_registro_02.gif','../imagens/atualiza_02.gif')">
<form name="frm1" method="post" action="grava_relaciona_ativ_trans2.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
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
  <table border=0 cellpadding=0 cellspacing=0 width="771" align="center">
    <tr> 
      <td valign=top width="786" height="11"> 
        <table width="83%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td width="20%"><%'=str_Opc%>
              <%=str_MegaProcesso%>
              <%'=str_Processo%>
              <%'=str_SubProcesso%></td>
            <td width="6%">&nbsp;</td>
            <td width="62%"><%'=str_AtividadeCarga%>
              <%'=str_Modulo%></td>
            <td width="12%">&nbsp;</td>
            <td width="12%">&nbsp;</td>
          </tr>
        </table>
    <tr> 
      <td valign=top width="786"> 
        <table width="81%" border="0" cellspacing="5" cellpadding="0" align="center">
          <tr> 
            <td width="30%"><div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade 
                de Carga: &nbsp;</font></font></div>
            </td>
            <td width="53%"><font color="#003366"> 
              <select name="selAtividadeCarga" onChange="MM_goToURL1('self','form_relaciona_ativ_trans4.asp?txtOpc=2');return document.MM_returnValue">
                <% 
		  if str_Opc <> "1" then %>
                <option value="0" selected><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Selecione 
                uma Atividade de Carga</font></option>
                <% else %>
                <option value="0" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Selecione 
                uma Atividade de Carga</font></option>
                <% end if %>
                <%Set rdsAtividadeCarga = Conn_db.Execute(str_SQL_Atividade_Carga)
While (NOT rdsAtividadeCarga.EOF)
         if (Trim(str_AtividadeCarga) = Trim(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)) then %>
                <option value="<%=(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)%>" selected ><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rdsAtividadeCarga.Fields.Item("ATCA_TX_DESC_ATIVIDADE").Value)%></font></option>
                <% else %>
                <option value="<%=(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)%>" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rdsAtividadeCarga.Fields.Item("ATCA_TX_DESC_ATIVIDADE").Value)%></font></option>
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
              </font></td>
            <td width="17%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="30%"> 
              <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega-Processo:&nbsp; 
                </font></font></div>
            </td>
            <td width="53%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              <select name="selMegaProcesso" onChange="MM_goToURL3('self','form_relaciona_ativ_trans4.asp?txtOpc=6');return document.MM_returnValue">
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
            <td width="17%"> 
              <input type="hidden" name="txtDsMP" value="<%=str_DescMegaProcesso%>">
              <input type="hidden" name="txtDsP" value="<%=str_DescProcesso%>">
              <input type="hidden" name="txtDsSP" value="<%=str_DescSubProcesso%>">
              <input type="hidden" name="txtDsA" value="<%=str_DescSubProcesso%>">
              <input type="hidden" name="txtDsM" value="<%=str_DescSubProcesso%>">
            </td>
          </tr>
          <tr> 
            <td width="30%"> 
              <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Processo: 
                &nbsp;</font></font></div>
            </td>
            <td width="53%"> 
              <select name="selProcesso" onChange="MM_goToURL4('self','form_relaciona_ativ_trans4.asp?txtOpc=7',this);return document.MM_returnValue">
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
            <td width="17%"> 
              <table width="60" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="30"> 
                    <div align="center"><a href="javascript:MM_goToURL1('self','cadas_Processo.asp?txtOpc=3',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','../imagens/novo_registro_02.gif',1)"><img name="Image15" border="0" src="../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui novo Processo"></a></div>
                  </td>
                  <td width="30"> 
                    <div align="center"><a href="JavaScript:history.go()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','../imagens/atualiza_02.gif',1)"><img name="Image17" border="0" src="../imagens/atualiza_02_off.gif" width="22" height="22" alt="Recarrega novo Processo"></a></div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td width="30%"> 
              <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Sub-Processo: 
                &nbsp;</font></font></div>
            </td>
            <td width="53%"> 
              <select name="selSubProcesso">
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
            </td>
            <td width="17%"> 
              <table width="60" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="30"> 
                    <div align="center"><a href="javascript:MM_goToURL2('self','cadas_sub_processo.asp?txtOpc=3',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image18','','../imagens/novo_registro_02.gif',1)"><img name="Image18" border="0" src="../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui novo Sub Processo"></a></div>
                  </td>
                  <td width="30"> 
                    <div align="center"><a href="JavaScript:history.go()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image19','','../imagens/atualiza_02.gif',1)"><img name="Image19" border="0" src="../imagens/atualiza_02_off.gif" width="22" height="22" alt="Recarrega novo Sub Processo"></a></div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td width="30%">&nbsp;</td>
            <td width="53%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
          </tr>
        </table>
        <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="242">
          <tr> 
            <td width="392" height="7" bgcolor="#0099CC"></td>
            <td width="349" height="7" bgcolor="#0099CC"></td>
          </tr>
          <tr> 
            <td colspan="2" height="7"></td>
          </tr>
          <tr> 
            <td colspan="2" height="21"> 
              <div align="center"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="19%"><font face="Arial, Helvetica, sans-serif" size="2" color="#003300">Agrupamento
                      (Master 
                      List R3)</font></td>
                    <td width="81%"><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b> 
                      <select name="selModulo" onChange="MM_goToURL2('self','form_relaciona_ativ_trans4.asp?txtOpc=3');return document.MM_returnValue">
                        <% 
		  if str_Opc <> "1" then %>
                        <option value="0" selected><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b>Selecione 
                        um Módulo</b></font></option>
                        <% else %>
                        <option value="0" ><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b>Selecione 
                        um Módulo</b></font></option>
                        <% end if %>
                        <%Set rdsModulo = Conn_db.Execute(str_SQL_Modulo)
While (NOT rdsModulo.EOF)
  
           if (Trim(str_Modulo) = Trim(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)) then %>
                        <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>" selected ><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></b></font></option>
                        <% else %>
                        <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>"><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></b></font></option>
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
                      </b></font></td>
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
              <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>N&atilde;o 
                selecionadas/N&atilde;o cadastradas</b></font></font></div>
            </td>
            <td height="7" bgcolor="#0099CC" width="349"> 
              <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Selecionadas/J&aacute; 
                cadastradas</b></font></font></div>
            </td>
          </tr>
          <tr> 
            <td colspan="2" height="10"><%'=str_AtividadeCarga%>
              <%'=str_Modulo%></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"> 
              <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="54%"> 
                    <div align="center"> <b> 
                      <select name="list1" size="8" multiple>
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
rdsTransacao.close
set rdsTransacao = Nothing
%>
                      </select>
                      </b></div>
                  </td>
                  <td width="8%" align="center"> 
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
                  <td width="38%"> 
                    <div align="center"><font color="#000080"> 
                      <select name="list2" size="8" multiple>
                        <%Set rdsTransacao_cad = Conn_db.Execute(str_SQL_Transacao_Cad)
While (NOT rdsTransacao_cad.EOF)
%>
                        <option value="<%=(rdsTransacao_cad.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" ><b><%=(rdsTransacao_cad.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></b></option>
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
                </tr>
                <tr> 
                  <td width="54%">
                    <%'=str_SQL_Sub_Proc%>
                    <input type="hidden" name="txtTranNaoSelecionada"> </td>
                  <td width="8%" align="center">&nbsp;</td>
                  <td width="38%"> 
                    <input type="hidden" name="txtTranSelecionada">
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
  </table>
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
dim str_Modulo

str_Opc = Request("txtOpc")
if str_Opc <> "1" then
   str_AtividadeCarga = Request("selAtividadeCarga")
   str_MegaProcesso= Request("selMegaProcesso")
   str_Processo = Request("selProcesso")
   str_SubProcesso = Request("selSubProcesso")   
   str_Modulo = Request("selModulo")
else
   str_AtividadeCarga = "0"
   str_MegaProcesso = "0"
   str_Processo = "0"
   str_SubProcesso = "0"
   str_Modulo = "0"
end if
if str_Opc = "6" then
   str_Processo = "0"
   str_SubProcesso = "0"
end if
if str_Opc = "7" then
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
   str_SubProcesso = "0"
end if


set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

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
str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND " & Session("PREFIXO") & "SUB_PROCESSO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
Set rdsSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
if Not rdsSubProcesso.EOF then
   str_DescMegaProcesso = rdsSubProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
   str_DescProcesso = rdsSubProcesso("PROC_TX_DESC_PROCESSO")
   str_DescSubProcesso = rdsSubProcesso("SUPR_TX_DESC_SUB_PROCESSO")
else
   str_DescMegaProcesso = ""
   str_DescProcesso = ""
   str_DescSubProcesso = ""
end if
rdsSubProcesso.close
set rdsSubProcesso = Nothing

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

str_SQL_Atividade_Carga = ""
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " SELECT "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " ATCA_CD_ATIVIDADE_CARGA, "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " ATCA_TX_DESC_ATIVIDADE "
str_SQL_Atividade_Carga = str_SQL_Atividade_Carga & " FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA"

str_SQL_Modulo = ""
str_SQL_Modulo = str_SQL_Modulo & " SELECT distinct "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO, "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3.MODU_TX_DESC_MODULO"
str_SQL_Modulo = str_SQL_Modulo & " FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA INNER JOIN"
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA ON "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA = " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA"
str_SQL_Modulo = str_SQL_Modulo & " INNER JOIN"
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODULO_R3 ON "
str_SQL_Modulo = str_SQL_Modulo & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO"
str_SQL_Modulo = str_SQL_Modulo & " WHERE " & Session("PREFIXO") & "ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA = "  & str_AtividadeCarga

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA INNER JOIN"
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO = '" & str_Modulo & "'"

'str_SQL_Transacao = ""
'str_SQL_Transacao = str_SQL_Transacao & " SELECT "
'str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO, "
'str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
'    lss_SQL = lss_SQL & " From PRODUTO as b "
'    lss_SQL = lss_SQL & " Where Convert(VarChar(5), b.GRPR_NR_CD_GRUPO_PRODUTO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),b.SUGR_NR_CD_SUB_GRUPO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),b.PROD_NR_CD_PRODUTO) not In "
'    lss_SQL = lss_SQL & " (Select Convert(varchar(5),h.GRPR_NR_CD_GRUPO_PRODUTO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),h.SUGR_NR_CD_SUB_GRUPO) "
'    lss_SQL = lss_SQL & " + Convert(varchar(5),h.PROD_NR_CD_PRODUTO) "
'    lss_SQL = lss_SQL & " From PRODUTO_EMPRESA as h "
'    lss_SQL = lss_SQL & " Where h.EMPR_NR_CD_EMPRESA = " & txtPRFCod.Text
'    lss_SQL = lss_SQL & " ) "
'    If Len(Trim(txtPRDCodGrupoProduto)) <> 0 Then
'       lss_SQL = lss_SQL & " and b.GRPR_NR_CD_GRUPO_PRODUTO = " & txtPRDCodGrupoProduto.Text
'    End If
'    lss_SQL = lss_SQL & " and b.PROD_TX_SITUACAO_PRODUTO = 'C'"
'    lss_SQL = lss_SQL & " order by b.PROD_TX_NM_PRODUTO "

str_SQL_Transacao = ""
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO, "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA INNER JOIN"
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.MODU_CD_MODULO = '" & str_Modulo & "'"
str_SQL_Transacao = str_SQL_Transacao & " and Convert(VarChar(5), " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO) "
str_SQL_Transacao = str_SQL_Transacao & "  Not In ("
str_SQL_Transacao = str_SQL_Transacao & " SELECT "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO "
str_SQL_Transacao = str_SQL_Transacao & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO INNER JOIN"
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao = str_SQL_Transacao & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao = str_SQL_Transacao & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO"
str_SQL_Transacao = str_SQL_Transacao & " WHERE " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Transacao = str_SQL_Transacao & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO = '" & str_Modulo & "')"

str_SQL_Transacao_Cad = ""
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " SELECT "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " FROM " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO INNER JOIN"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " INNER JOIN " & Session("PREFIXO") & "MODULO_R3 ON " & Session("PREFIXO") & "TRANSACAO.MODU_CD_MODULO = " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO"
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " WHERE " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.PROC_CD_PROCESSO = " & str_Processo
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & str_SubProcesso
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "ATIVIDADE_TRANSACAO.ATIV_CD_ATIVIDADE = " & str_AtividadeCarga
str_SQL_Transacao_Cad = str_SQL_Transacao_Cad & " AND " & Session("PREFIXO") & "MODULO_R3.MODU_CD_MODULO = '" & str_Modulo & "'"

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
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL2() { //v3.0
  var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selModulo="+document.frm1.selModulo.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"'");
}
function MM_goToURL3() { //v3.0
  var i, args=MM_goToURL3.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selModulo="+document.frm1.selModulo.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selModulo="+document.frm1.selModulo.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selProcesso="+document.frm1.selProcesso.value+"'");
}
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
function carrega_txt(fbox) {
document.frm1.txtTranSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtTranSelecionada.value = document.frm1.txtTranSelecionada.value + "[" + fbox.options[i].value;
   }
}

function carrega_txt2(fbox) {
document.frm1.txtTranNaoSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtTranNaoSelecionada.value = document.frm1.txtTranNaoSelecionada.value + "[" + fbox.options[i].value;
   }
}

function Confirma() 
{ 
if (document.frm1.selAtividadeCarga.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatória !");
     document.frm1.selAtividadeCarga.focus();
     return;
     }
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
if (document.frm1.selProcesso.selectedIndex == 0)
     { 
	 alert("Selecione um Proceso ou cadastre um novo.");
     document.frm1.selProcesso.focus();
     return;
     }	 
if (document.frm1.selSubProcesso.selectedIndex == 0)
     { 
	 alert("Selecione um Sub Proceso ou cadastre um novo.");
     document.frm1.selSubProcesso.focus();
     return;
     }	 
if (document.frm1.selModulo.selectedIndex == 0)
     { 
	 alert("A seleção de um Módulo é obrigatória !");
     document.frm1.selModulo.focus();
     return;
     }
if (document.frm1.list2.options.length == 0)
     { 
	 alert("A seleção de uma Transação é obrigatória !");
     document.frm1.list2.focus();
     return;
     }
	 else
     {
	  carrega_txt(document.frm1.list2);
  	  carrega_txt2(document.frm1.list1);
	  alert(document.frm1.txtTranSelecionada.value);
	  alert(document.frm1.txtTranNaoSelecionada.value);
	  document.frm1.txtDsMP.value = document.frm1.selMegaProcesso.options[document.frm1.selMegaProcesso.selectedIndex].text
	  document.frm1.txtDsP.value = document.frm1.selProcesso.options[document.frm1.selProcesso.selectedIndex].text
	  document.frm1.txtDsSP.value = document.frm1.selSubProcesso.options[document.frm1.selSubProcesso.selectedIndex].text
	  document.frm1.txtDsA.value = document.frm1.selAtividadeCarga.options[document.frm1.selAtividadeCarga.selectedIndex].text
	  document.frm1.txtDsM.value = document.frm1.selModulo.options[document.frm1.selModulo.selectedIndex].text
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script language="javascript" src="js/troca_lista.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../imagens/continua2_F02.gif','../imagens/continua_F02.gif','../imagens/novo_registro_02.gif','../imagens/atualiza_02.gif')">
<form name="frm1" method="post" action="grava_relaciona_ativ_trans2.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
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
  <table border=0 cellpadding=0 cellspacing=0 width="771" align="center">
    <tr> 
      <td valign=top width="786" height="11"> 
        <table width="83%" border="0" cellpadding="0" cellspacing="0" align="center">
          <tr> 
            <td width="20%"><%'=str_Opc%>
              <%=str_MegaProcesso%>
              <%'=str_Processo%>
              <%'=str_SubProcesso%></td>
            <td width="6%">&nbsp;</td>
            <td width="62%"><%'=str_AtividadeCarga%>
              <%'=str_Modulo%></td>
            <td width="12%">&nbsp;</td>
            <td width="12%">&nbsp;</td>
          </tr>
        </table>
    <tr> 
      <td valign=top width="786"> 
        <table width="81%" border="0" cellspacing="5" cellpadding="0" align="center">
          <tr> 
            <td width="30%"><div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Atividade 
                de Carga: &nbsp;</font></font></div>
            </td>
            <td width="53%"><font color="#003366"> 
              <select name="selAtividadeCarga" onChange="MM_goToURL1('self','form_relaciona_ativ_trans4.asp?txtOpc=2');return document.MM_returnValue">
                <% 
		  if str_Opc <> "1" then %>
                <option value="0" selected><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Selecione 
                uma Atividade de Carga</font></option>
                <% else %>
                <option value="0" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Selecione 
                uma Atividade de Carga</font></option>
                <% end if %>
                <%Set rdsAtividadeCarga = Conn_db.Execute(str_SQL_Atividade_Carga)
While (NOT rdsAtividadeCarga.EOF)
         if (Trim(str_AtividadeCarga) = Trim(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)) then %>
                <option value="<%=(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)%>" selected ><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rdsAtividadeCarga.Fields.Item("ATCA_TX_DESC_ATIVIDADE").Value)%></font></option>
                <% else %>
                <option value="<%=(rdsAtividadeCarga.Fields.Item("ATCA_CD_ATIVIDADE_CARGA").Value)%>" ><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><%=(rdsAtividadeCarga.Fields.Item("ATCA_TX_DESC_ATIVIDADE").Value)%></font></option>
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
              </font></td>
            <td width="17%">&nbsp;</td>
          </tr>
          <tr> 
            <td width="30%"> 
              <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Mega-Processo:&nbsp; 
                </font></font></div>
            </td>
            <td width="53%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
              <select name="selMegaProcesso" onChange="MM_goToURL3('self','form_relaciona_ativ_trans4.asp?txtOpc=6');return document.MM_returnValue">
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
            <td width="17%"> 
              <input type="hidden" name="txtDsMP" value="<%=str_DescMegaProcesso%>">
              <input type="hidden" name="txtDsP" value="<%=str_DescProcesso%>">
              <input type="hidden" name="txtDsSP" value="<%=str_DescSubProcesso%>">
              <input type="hidden" name="txtDsA" value="<%=str_DescSubProcesso%>">
              <input type="hidden" name="txtDsM" value="<%=str_DescSubProcesso%>">
            </td>
          </tr>
          <tr> 
            <td width="30%"> 
              <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Processo: 
                &nbsp;</font></font></div>
            </td>
            <td width="53%"> 
              <select name="selProcesso" onChange="MM_goToURL4('self','form_relaciona_ativ_trans4.asp?txtOpc=7',this);return document.MM_returnValue">
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
            <td width="17%"> 
              <table width="60" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="30"> 
                    <div align="center"><a href="javascript:MM_goToURL1('self','cadas_Processo.asp?txtOpc=3',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','../imagens/novo_registro_02.gif',1)"><img name="Image15" border="0" src="../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui novo Processo"></a></div>
                  </td>
                  <td width="30"> 
                    <div align="center"><a href="JavaScript:history.go()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image17','','../imagens/atualiza_02.gif',1)"><img name="Image17" border="0" src="../imagens/atualiza_02_off.gif" width="22" height="22" alt="Recarrega novo Processo"></a></div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td width="30%"> 
              <div align="right"><font color="#003366"><font face="Arial, Helvetica, sans-serif" size="2">Sub-Processo: 
                &nbsp;</font></font></div>
            </td>
            <td width="53%"> 
              <select name="selSubProcesso">
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
            </td>
            <td width="17%"> 
              <table width="60" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="30"> 
                    <div align="center"><a href="javascript:MM_goToURL2('self','cadas_sub_processo.asp?txtOpc=3',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image18','','../imagens/novo_registro_02.gif',1)"><img name="Image18" border="0" src="../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui novo Sub Processo"></a></div>
                  </td>
                  <td width="30"> 
                    <div align="center"><a href="JavaScript:history.go()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image19','','../imagens/atualiza_02.gif',1)"><img name="Image19" border="0" src="../imagens/atualiza_02_off.gif" width="22" height="22" alt="Recarrega novo Sub Processo"></a></div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>
          <tr> 
            <td width="30%">&nbsp;</td>
            <td width="53%">&nbsp;</td>
            <td width="17%">&nbsp;</td>
          </tr>
        </table>
        <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="242">
          <tr> 
            <td width="392" height="7" bgcolor="#0099CC"></td>
            <td width="349" height="7" bgcolor="#0099CC"></td>
          </tr>
          <tr> 
            <td colspan="2" height="7"></td>
          </tr>
          <tr> 
            <td colspan="2" height="21"> 
              <div align="center"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="19%"><font face="Arial, Helvetica, sans-serif" size="2" color="#003300">Agrupamento
                      (Master 
                      List R3)</font></td>
                    <td width="81%"><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b> 
                      <select name="selModulo" onChange="MM_goToURL2('self','form_relaciona_ativ_trans4.asp?txtOpc=3');return document.MM_returnValue">
                        <% 
		  if str_Opc <> "1" then %>
                        <option value="0" selected><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b>Selecione 
                        um Módulo</b></font></option>
                        <% else %>
                        <option value="0" ><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b>Selecione 
                        um Módulo</b></font></option>
                        <% end if %>
                        <%Set rdsModulo = Conn_db.Execute(str_SQL_Modulo)
While (NOT rdsModulo.EOF)
  
           if (Trim(str_Modulo) = Trim(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)) then %>
                        <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>" selected ><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></b></font></option>
                        <% else %>
                        <option value="<%=(rdsModulo.Fields.Item("MODU_CD_MODULO").Value)%>"><font face="Arial, Helvetica, sans-serif" size="2" color="#003300"><b><%=(rdsModulo.Fields.Item("MODU_TX_DESC_MODULO").Value)%></b></font></option>
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
                      </b></font></td>
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
              <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>N&atilde;o 
                selecionadas/N&atilde;o cadastradas</b></font></font></div>
            </td>
            <td height="7" bgcolor="#0099CC" width="349"> 
              <div align="center"><font color="#003300"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Selecionadas/J&aacute; 
                cadastradas</b></font></font></div>
            </td>
          </tr>
          <tr> 
            <td colspan="2" height="10"><%'=str_AtividadeCarga%>
              <%'=str_Modulo%></td>
          </tr>
          <tr> 
            <td colspan="2" height="10"> 
              <table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="54%"> 
                    <div align="center"> <b> 
                      <select name="list1" size="8" multiple>
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
rdsTransacao.close
set rdsTransacao = Nothing
%>
                      </select>
                      </b></div>
                  </td>
                  <td width="8%" align="center"> 
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
                  <td width="38%"> 
                    <div align="center"><font color="#000080"> 
                      <select name="list2" size="8" multiple>
                        <%Set rdsTransacao_cad = Conn_db.Execute(str_SQL_Transacao_Cad)
While (NOT rdsTransacao_cad.EOF)
%>
                        <option value="<%=(rdsTransacao_cad.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" ><b><%=(rdsTransacao_cad.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></b></option>
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
                </tr>
                <tr> 
                  <td width="54%">
                    <%'=str_SQL_Sub_Proc%>
                    <input type="hidden" name="txtTranNaoSelecionada"> </td>
                  <td width="8%" align="center">&nbsp;</td>
                  <td width="38%"> 
                    <input type="hidden" name="txtTranSelecionada">
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
  </table>
</form>
</body>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
</html>