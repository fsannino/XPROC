 
<%
str_Opt = Request("txtOPT")
str_Acao = Request("txtAcao")
'response.write Request("selMacroPerfil")
' ************ usado para edição de objetos ************************
if (Request("selMacroPerfil") <> "") then 
    str_MacroPerfil2 = Request("selMacroPerfil")
else
    str_MacroPerfil2 = ""
end if
'************** usado para edição de objetos ************************
if (Request("selFuncao") <> "") then 
    str_Funcao = Request("selFuncao")
else
    str_Funcao = "0"
end if

if (Request("txtMacroPerfil") <> "") then 
    str_MacroPerfil = Request("txtMacroPerfil")
else
    if str_MacroPerfil2 <> "" then
       str_MacroPerfil = str_MacroPerfil2
	else
       str_MacroPerfil = "0"
	end if   
end if

if (Request("txtNomeTecnico") <> "") then 
    str_NomeTecnico = Request("txtNomeTecnico")
else
    str_NomeTecnico = "0"
end if

'response.write str_MacroPerfil2
'response.write "   -  "
'response.write str_Funcao
'response.write "   -  "
'response.write str_MacroPerfil
'response.write "   -  "
'response.write str_NomeTecnico
'response.write "   -  "
'response.write str_Opt
'response.write "   -  "

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")


'response.Write(str_MacroPerfil)


'if str_Opt = 3 OR  str_Opt = 4 then
   str_SQL_Macro = ""
   str_SQL_Macro = str_SQL_Macro & " SELECT MCPR_NR_SEQ_MACRO_PERFIL, "
   str_SQL_Macro = str_SQL_Macro & " MCPE_TX_NOME_TECNICO, MCPE_TX_SITUACAO, "
   str_SQL_Macro = str_SQL_Macro & " FUNE_CD_FUNCAO_NEGOCIO"
   str_SQL_Macro = str_SQL_Macro & " FROM " & Session("PREFIXO") & "MACRO_PERFIL"
   str_SQL_Macro = str_SQL_Macro & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
   Set rdsMacro = Conn_db.Execute(str_SQL_Macro)
   str_Funcao = rdsMacro("FUNE_CD_FUNCAO_NEGOCIO")
   str_NomeTecnico = rdsMacro("MCPE_TX_NOME_TECNICO")
   str_Situacao = rdsMacro("MCPE_TX_SITUACAO")
'end if

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
   ls_str_TituloFuncao = "Não achou funcao"
end if
rdsFuncao.close
set rdsFuncao = Nothing

str_SQL_Trans = ""
str_SQL_Trans = str_SQL_Trans  & " SELECT " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.TRAN_CD_TRANSACAO, "
str_SQL_Trans = str_SQL_Trans  & " " & Session("PREFIXO") & "TRANSACAO.TRAN_TX_DESC_TRANSACAO, MEPR_CD_MEGA_PROCESSO,  "
str_SQL_Trans = str_SQL_Trans  & " " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.MCPT_NR_SITUACAO_ALTERACAO, " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.MCPT_NR_SITUACAO_ALTERACAO1, MCPT_NR_SITUACAO_ALTERACAO_FUNC "
str_SQL_Trans = str_SQL_Trans  & " FROM " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO INNER JOIN "
str_SQL_Trans = str_SQL_Trans  & " " & Session("PREFIXO") & "TRANSACAO ON "
str_SQL_Trans = str_SQL_Trans  & " " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.TRAN_CD_TRANSACAO = " & Session("PREFIXO") & "TRANSACAO.TRAN_CD_TRANSACAO"
str_SQL_Trans = str_SQL_Trans  & " WHERE " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.MCPR_NR_SEQ_MACRO_PERFIL = " & str_MacroPerfil
str_SQL_Trans = str_SQL_Trans  & " order by " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.TRAN_CD_TRANSACAO " 
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
  <table border="0" width="736" height="137" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td width="10" height="21">&nbsp;</td>
      <td width="160" height="21">&nbsp;</td>
      <td height="21" width="499"><font face="Verdana" color="#330099" size="3">Macro 
        Perfil x Transação</font></td>
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
      <td width="160" height="14"> <div align="right"><font size="2"></font></div></td>
      <td width="499" height="14">&nbsp;</td>
      <td height="14" width="67"></td>
    </tr>
    <tr> 
      <td width="10" height="14"></td>
      <td width="160" height="14"> <div align="right"><font size="2"><font face="Verdana" color="#330099">Nome 
          T&eacute;cnico:</font></font></div></td>
      <td width="499" height="14"><b><font face="Verdana" size="2" color="#330099"><%=str_NomeTecnico%></font></b></td>
      <td height="14" width="67"></td>
    </tr>
    <tr> 
      <td width="10" height="14"></td>
      <td width="160" height="14"> </td>
      <td width="499" height="14"><b><font face="Verdana" size="2" color="#330099"> 
        <input type="hidden" name="txtDescMegaProcesso" size="46" value="<%=ls_str_DescMegaProcesso%>">
        <input type="hidden" name="txtMegaProcesso" size="46" value="<%=str_MegaProcesso%>">
        </font></b></td>
      <td height="14" width="67"> </td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="160" height="14"> <div align="right"><font face="Verdana" size="2" color="#330099">Fun&ccedil;&atilde;o 
          R/3: </font></div></td>
      <td height="14" width="499"><b><font face="Verdana" size="2" color="#330099"><%=str_Funcao%> 
        <input type="hidden" name="txtFuncao" size="46" value="<%=str_Funcao%>">
        - <%=ls_str_TituloFuncao%></font><font face="Verdana" size="1" color="#330099"> 
        <input type="hidden" name="txtDescFuncao" size="46" value="<%=ls_str_TituloFuncao%>">
        </font></b></td>
      <td height="14" width="67"> <div align="center"><b><font face="Verdana" size="2" color="#330099"></font></b> 
        </div></td>
    </tr>
    <tr>
      <td height="23">&nbsp;</td>
      <td height="23">&nbsp;</td>
      <td height="23">&nbsp;</td>
      <td height="23">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10" height="23">&nbsp;</td>
      <td width="160" height="23"> <input type="hidden" name="txtTranSelecionada" size="20"> 
      </td>
      <% If str_Situacao = "EE" then
			str_Situacao = "Em elaboração"
		 elseIf str_Situacao = "AT" then
			str_Situacao = "Alterado transação"
		 elseIf str_Situacao = "EA" then
			str_Situacao = "Em aprovação"			  
		 elseIf str_Situacao = "NA" then
			str_Situacao = "Não aprovado"			  
		 elseIf str_Situacao = "EC" then
			str_Situacao = "Em criação no R/3"			  
		 elseIf str_Situacao = "RE" then
			str_Situacao = "Recusado no R/3"			  
		 elseIf str_Situacao = "EX" then
			str_Situacao = "Excluída a função"			  
		 elseIf str_Situacao = "MR" then
			str_Situacao = "Mudado para referência"			  
		 elseIf str_Situacao = "EL" then
			str_Situacao = "Excluído"			  
		 elseIf str_Situacao = "CR" then
			str_Situacao = "Criado no R3"			  
		 elseIf str_Situacao = "AR" then
			str_Situacao = "Em alteração no R/3"			  
		 elseIf str_Situacao = "ER" then
			str_Situacao = "Em exclusão no R/3"			  
		 elseIf str_Situacao = "AP" then
			str_Situacao = "Alterado no R/3"			  
		 elseIf str_Situacao = "EP" then
			str_Situacao = "Excluído no R/3"			  
         end if
	  %>
      <td height="23" width="499"><b><font face="Verdana" size="1" color="#330099"> 
        <%'=str_Opt%>
        </font></b><font face="Verdana" size="1" color="#330099">&nbsp; </font><font face="Verdana" size="2" color="#330099">Status:<b> 
        <%=str_Situacao%> </b></font></td>
      <td height="23" width="67">&nbsp;</td>
    </tr>
  </table>
  <table width="800" border="0" cellpadding="0" cellspacing="0" align="center" height="180">
    <tr> 
      <td colspan="2" height="7"></td>
    </tr>
    <tr> 
      <td height="2" bgcolor="#0099CC" width="179"> <strong><font color="#003300" size="2"></font></strong></td>
      <td height="2" bgcolor="#0099CC" width="754"> <table width="815" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="188">&nbsp;</td>
            <td width="286"> <div align="right"><strong><font size="2"><img src="../../imagens/aprova_02.gif" width="14" height="12"> 
                <font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Objetos 
                alterados pelo Elaborador</font></font> </strong></div></td>
            <td width="334"> <div align="right"><strong><font size="2"><img src="../../imagens/aprova_03.gif" width="14" height="12"> 
                <font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Objetos 
                alterados pelo Validador</font></font> </strong></div></td>
            <td width="7">&nbsp;</td>
          </tr>
          <tr> 
            <td><strong><font color="#003300" size="2"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Transa&ccedil;&otilde;es 
              existentes</font></font></strong></td>
            <td><div align="right"><strong><font size="2"><img src="../../imagens/novo_registro_02_off.gif" width="22" height="22"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
                Transa&ccedil;&atilde;o inclu&iacute;da na fun&ccedil;&atilde;o</font></font> 
                </strong></div></td>
            <td><div align="right"><strong><font size="2"><img src="../../imagens/desiste_F01.gif" width="24" height="24"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Transa&ccedil;&atilde;o 
                exclu&iacute;da na fun&ccedil;&atilde;o</font></font></strong></div></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><div align="right"><strong><font size="2"><img src="../../imagens/b04.gif" width="16" height="16"> 
                <font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Sem 
                objetos cr&iacute;ticos</font></font></strong></div></td>
            <td><div align="right"><strong><font size="2"><img src="../../imagens/b03.gif" width="16" height="16"> 
                <font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Com</font></font><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif"> 
                objetos cr&iacute;ticos</font></strong></div></td>
            <td>&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="10"> <table width="926" border="0" cellspacing="3" cellpadding="0">
          <tr> 
            <td width="35"> <div align="center"></div></td>
            <td width="666"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Transa&ccedil;&atilde;o</font></div></td>
            <td width="32">&nbsp;</td>
            <td width="32">&nbsp;</td>
            <td width="32">&nbsp;</td>
            <td width="108">&nbsp;</td>
          </tr>
          <% 
          Set rdsTransacao = Conn_db.Execute(str_SQL_Trans)
		  int_MegaProcesso = 0
		  Do While Not rdsTransacao.EOF 
		     if ls_Cor_Linha = "#F7F7F7" then
                ls_Cor_Linha = "#FFFFFF"
             else		  
		        ls_Cor_Linha = "#F7F7F7"
		     end if		  
		  %>
          <% 'response.write int_MegaProcesso
		     'response.write rdsTransacao("MEPR_CD_MEGA_PROCESSO")
		  If Trim(int_MegaProcesso) <> Trim(rdsTransacao("MEPR_CD_MEGA_PROCESSO")) then 
             'int_MegaProcesso = rdsTransacao("MEPR_CD_MEGA_PROCESSO")
		     str_SQL = ""
		     str_SQl = str_SQL & " SELECT MEPR_CD_MEGA_PROCESSO, MEPR_TX_DESC_MEGA_PROCESSO "
             str_SQl = str_SQL & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
             str_SQl = str_SQL & " WHERE MEPR_CD_MEGA_PROCESSO = " & rdsTransacao("MEPR_CD_MEGA_PROCESSO")
			 Set rdsMegaProc = Conn_db.Execute(str_SQL)
		  %>
          <tr bgcolor="<%=ls_Cor_Linha%>"> 
            <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
              <%'=rdsMegaProc("MEPR_CD_MEGA_PROCESSO")%>
              <%'=rdsMegaProc("MEPR_TX_DESC_MEGA_PROCESSO")%>
              </font></b></font></td>
            <td width="32" bgcolor="<%=ls_Cor_Linha%>">&nbsp;</td>
            <td width="32" bgcolor="<%=ls_Cor_Linha%>">&nbsp;</td>
            <td width="32" bgcolor="<%=ls_Cor_Linha%>">&nbsp;</td>
            <td width="108" bgcolor="<%=ls_Cor_Linha%>">&nbsp;</td>
          </tr>
          <% end if %>
          <tr bgcolor="<%=ls_Cor_Linha%>"> 
            <td width="35"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
              <%'=rdsMegaProc("MEPR_CD_MEGA_PROCESSO")%>
              </font></b></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
              <%'=int_MegaProcesso%>
              </font></b></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            <%   str_SQL = ""
				   str_SQL = str_SQL & " SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
				   str_SQL = str_SQL & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
                   str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "TRANSACAO_MEGA INNER JOIN"
                   str_SQL = str_SQL & " " & Session("PREFIXO") & "MEGA_PROCESSO ON "
                   str_SQL = str_SQL & " " & Session("PREFIXO") & "TRANSACAO_MEGA.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
                   str_SQL = str_SQL & " WHERE " & Session("PREFIXO") & "TRANSACAO_MEGA.TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'" 
				   str_SQL = str_SQL & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "				   
				   Set rdsExiste2 = Conn_db.Execute(str_SQL)				   
				   loo_Existe = False
				   str_Mega = "         - Dono : "
				   a = ""
				   IF not rdsExiste2.EOF then
				      Do While not rdsExiste2.EOF
				         'if Trim(rdsExiste2("MEPR_CD_MEGA_PROCESSO")) = Trim(str_MegaProcesso2) then
						 'InStr([start, ]string1, string2[, compare])
						 'response.write "        ---      "
						 'response.write "," & Trim(rdsExiste2("MEPR_CD_MEGA_PROCESSO")) & ","
						 'response.write "        ---      "
						 'response.write "," & Session("AcessoUsuario") & ","
				         'if InStr("," & Session("AcessoUsuario") & ",","," &  Trim(rdsExiste2("MEPR_CD_MEGA_PROCESSO")) & ",") <> 0 then						 
				         '   loo_Existe = True
						'	 str_Mega = ""							 
                         '   exit do
	                     'end if
					     str_Mega = str_Mega & rdsExiste2("MEPR_TX_DESC_MEGA_PROCESSO") & " / "
					     rdsExiste2.Movenext
				      Loop
				   else
				   	  str_Mega = str_Mega & "   -  em processo de definição de dono "
				   end if
				   rdsExiste2.close
				   set rdsExiste2 = Nothing
				 %>
            <td width="666"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><%=rdsTransacao("TRAN_CD_TRANSACAO")%> </b>- <%=rdsTransacao("TRAN_TX_DESC_TRANSACAO")%><i><font color="#999999"><%=str_Mega%></font></i></font></td>
            <% 
			str_SQL = ""
			str_SQL = str_SQL & " Select * "
			str_SQL = str_SQL & " from TRANSACAO_OBJETO "
			str_SQL = str_SQL & " WHERE TROB_TX_CRITICO = 'X' " 
			str_SQL = str_SQL & " and TRAN_CD_TRANSACAO = '" & rdsTransacao("TRAN_CD_TRANSACAO") & "'" 
			Set rdsCritico = Conn_db.Execute(str_SQL)
			if rdsCritico.EOF then
			   str_Critico = "N"
			else
			   str_Critico = "S"
			end if
			rdsCritico.CLOSE
			if rdsTransacao("MCPT_NR_SITUACAO_ALTERACAO_FUNC") = 0 then
			      str_Var1 = "func_tran_nao_marcada.gif"
				  str_Texto1 = ""				  			
			   elseif rdsTransacao("MCPT_NR_SITUACAO_ALTERACAO_FUNC") = 1 then
                  str_Var1 = "novo_registro_02_off.gif"
				  str_Texto1 = "Esta transação foi incluída"				  
			   elseif rdsTransacao("MCPT_NR_SITUACAO_ALTERACAO_FUNC") = 2 then
                  str_Var1 = "desiste_F01.gif"
				  str_Texto1 = "Esta transação foi excluída"				  
			   end if
			%>
            <td width="32" bgcolor="<%=ls_Cor_Linha%>"> 
              <div align="center"><img src="../../imagens/<%=str_Var1%>" width="15" height="15" alt="<%=str_Texto1%>"></div></td>
            <% if rdsTransacao("MCPT_NR_SITUACAO_ALTERACAO1") = 0 then
			      str_Var2 = "func_tran_nao_marcada.gif"
				  str_Texto2 = ""			
			   elseif rdsTransacao("MCPT_NR_SITUACAO_ALTERACAO1") = 1 then
                  str_Var2 = "aprova_03.gif"
				  str_Texto2 = "Indica que foi alterado objeto desta Transação pelo Validador"
			   end if
			%>
            <td width="32" bgcolor="<%=ls_Cor_Linha%>"> 
              <div align="center"><img src="../../imagens/<%=str_Var2%>" width="15" height="15" alt="<%=str_Texto2%>"></div></td>
            <% if rdsTransacao("MCPT_NR_SITUACAO_ALTERACAO") = 0 then
			      str_Var3 = "func_tran_nao_marcada.gif"
				  str_Texto3 = ""			
			   elseif rdsTransacao("MCPT_NR_SITUACAO_ALTERACAO") = 1 then
                  str_Var3 = "aprova_02.gif"
				  str_Texto3 = "Indica que foi alterado objeto desta Transação pelo Elaborador"
			   end if
			%>
            <td width="32" bgcolor="<%=ls_Cor_Linha%>"> 
              <div align="center"><img src="../../imagens/<%=str_Var3%>" width="15" height="15" alt="<%=str_Texto3%>"></div></td>
            <td width="108" bgcolor="<%=ls_Cor_Linha%>"> 
              <div align="center"> 
                <% int_Conta_Transacao = int_Conta_Transacao + 1 
                   ls_Seq = ls_Seq + 1 
				   IF str_Critico = "N" THEN
				      str_Imag = "imagens/b04.gif"
				   else 	  
				      str_Imag = "imagens/b03.gif"
				   end if
				   if str_Opt <> 1 and str_Opt <> 3 then 
                      if loo_Existe = false then %>
                         <a href="edita_objetos_campo_macro_perfil.asp?selMacro_Perfil=<%=str_MacroPerfil%>&selDescMacro_Perfil=<%=str_NomeTecnico%>&selTransacao=<%=rdsTransacao("TRAN_CD_TRANSACAO")%>&selDescTransacao=<%=rdsTransacao("TRAN_TX_DESC_TRANSACAO")%>&txtFuncao=<%=str_Funcao%>&txtOpt=<%=str_Opt%>"><img src="../../<%=str_Imag%>" width="16" height="16" border="0" alt="Click para editar os objetos"></a> 
                <%    end if %>
                <% else %>
                      <a href="edita_objetos_campo_macro_perfil.asp?selMacro_Perfil=<%=str_MacroPerfil%>&selDescMacro_Perfil=<%=str_NomeTecnico%>&selTransacao=<%=rdsTransacao("TRAN_CD_TRANSACAO")%>&selDescTransacao=<%=rdsTransacao("TRAN_TX_DESC_TRANSACAO")%>&txtFuncao=<%=str_Funcao%>&txtOpt=<%=str_Opt%>"><img src="../../<%=str_Imag%>" width="16" height="16" border="0" alt="Click para editar os objetos"></a> 
                <% end if %>
              </div></td>
          </tr>
          <% 
		  	rdsTransacao.Movenext
			if not rdsTransacao.EOF then
			   int_MegaProcesso = rdsTransacao("MEPR_CD_MEGA_PROCESSO")
			end if   
		  Loop 
		  rdsTransacao.close
		  set rdsTransacao = Nothing
		  %>
        </table></td>
    </tr>
    <tr> 
      <td colspan="2" height="10"><%'=Session("AcessoUsuario")%></td>
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
