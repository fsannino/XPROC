 
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso
Dim str_Cenario
dim str_CenarioTrSequencia
Dim str_CenarioChSequencia
dim str_MegaProcesso2

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Cenario = "0"
str_CenarioTrSequencia = "0"
str_MegaProcesso2 = "0"

str_Opc = Request("txtOpc")

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if

if (Request("selMegaProcesso2") <> "") then 
    str_MegaProcesso2 = Request("selMegaProcesso2")
else
    str_MegaProcesso2 = "0"
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

if (Request("selSubProcesso") <> "") then 
    str_SubProcesso = Request("selSubProcesso")
else
    str_SubProcesso = "0"
end if

if (Request("selCenario") <> "") then 
    str_Cenario = Request("selCenario")
else
    str_Cenario = "0"
end if

if (Request("txtCenarioTrSequencia") <> "") then 
    str_CenarioTrSequencia = Request("txtCenarioTrSequencia")
else
    str_CenarioTrSequencia = ""
end if

if (Request("p_CenarioChSequencia") <> "") then 
    str_CenarioChSequencia = Request("p_CenarioChSequencia")
else
    str_CenarioChSequencia = "0"
end if

if (Request("txtDescTransacao") <> "") then 
    str_DescTransacao = Request("txtDescTransacao")
else
    str_DescTransacao = ""
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
Set rdsSubProcesso= Conn_db.Execute(str_SQL_Sub_Proc)
if not rdsSubProcesso.EOF then
   ls_Desc_MegaProcesso = rdsSubProcesso("MEPR_TX_DESC_MEGA_PROCESSO")
   ls_Desc_Processo = rdsSubProcesso("PROC_TX_DESC_PROCESSO")
   ls_Desc_SubProcesso = rdsSubProcesso("SUPR_TX_DESC_SUB_PROCESSO")   
else
   ls_Desc_MegaProcesso = "Não Encontrado"
   ls_Desc_Processo = "Não Encontrado"
   ls_Desc_SubProcesso = "Não Encontrado"
end if
rdsSubProcesso.Close
set rdsSubProcesso = Nothing

str_SQL_Cen_Tran = ""
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " SELECT "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_CD_CENARIO "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " , " & Session("PREFIXO") & "CENARIO_TRANSACAO.CETR_NR_SEQUENCIA "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " , " & Session("PREFIXO") & "CENARIO_TRANSACAO.BPPP_CD_BPP "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " , " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_CD_CENARIO_SEGUINTE " 
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " , " & Session("PREFIXO") & "CENARIO_TRANSACAO.OPES_CD_OPERACAO_ESP "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " , " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_NR_SEQUENCIA_TRANS "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " , " & Session("PREFIXO") & "CENARIO_TRANSACAO.CETR_TX_DESC_TRANSACAO "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " , " & Session("PREFIXO") & "CENARIO_TRANSACAO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO "
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " WHERE " & Session("PREFIXO") & "CENARIO_TRANSACAO.CENA_CD_CENARIO = '" & str_Cenario & "'" 
str_SQL_Cen_Tran = str_SQL_Cen_Tran & " AND " & Session("PREFIXO") & "CENARIO_TRANSACAO.CETR_NR_SEQUENCIA = " &  str_CenarioChSequencia
Set rdsCen_Tran= Conn_db.Execute(str_SQL_Cen_Tran)
if not rdsCen_Tran.EOF then
   str_CenarioSeguinte = rdsCen_Tran("CENA_CD_CENARIO_SEGUINTE")
   str_MegaProcesso2 = rdsCen_Tran("MEPR_CD_MEGA_PROCESSO")
   str_CenarioTrSequencia = rdsCen_Tran("CENA_NR_SEQUENCIA_TRANS")
   str_DescTransacao = rdsCen_Tran("CETR_TX_DESC_TRANSACAO")
else
   str_CenarioSeguinte = "0"
   str_MegaProcesso2 = "0"
   str_CenarioTrSequencia = "0"
   str_DescTransacao = "0"
end if
rdsCen_Tran.Close
set rdsCen_Tran = Nothing

str_SQL_OpEs = ""
str_SQL_OpEs = str_SQL_OpEs & " SELECT "
str_SQL_OpEs = str_SQL_OpEs & " OPES_CD_OPERACAO_ESP, "
str_SQL_OpEs = str_SQL_OpEs & " OPES_TX_DESC_OPERACAO_ESP"
str_SQL_OpEs = str_SQL_OpEs & " FROM " & Session("PREFIXO") & "OPERACOES_ESPEC"
str_SQL_OpEs = str_SQL_OpEs & " order by OPES_TX_DESC_OPERACAO_ESP "

str_SQL_Cenario = ""
str_SQL_Cenario = str_SQL_Cenario & " SELECT "
str_SQL_Cenario = str_SQL_Cenario & " CENA_CD_CENARIO"
str_SQL_Cenario = str_SQL_Cenario & " , CENA_TX_TITULO_CENARIO"
str_SQL_Cenario = str_SQL_Cenario & " FROM " & Session("PREFIXO") & "CENARIO"
str_SQL_Cenario = str_SQL_Cenario & " where CENA_CD_CENARIO <> '" &  str_Cenario & "'"
str_SQL_Cenario = str_SQL_Cenario & " and MEPR_CD_MEGA_PROCESSO = " &  str_MegaProcesso2 
str_SQL_Cenario = str_SQL_Cenario & " order by CENA_TX_TITULO_CENARIO"

str_SQL_DescCenario = ""
str_SQL_DescCenario = str_SQL_DescCenario & " SELECT "
str_SQL_DescCenario = str_SQL_DescCenario & " CENA_CD_CENARIO"
str_SQL_DescCenario = str_SQL_DescCenario & " , CENA_TX_TITULO_CENARIO"
str_SQL_DescCenario = str_SQL_DescCenario & " FROM " & Session("PREFIXO") & "CENARIO"
str_SQL_DescCenario = str_SQL_DescCenario & " where CENA_CD_CENARIO = '" &  str_Cenario & "'"
Set rdsDescCenario= Conn_db.Execute(str_SQL_DescCenario)

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selMegaProcesso2="+document.frm1.selMegaProcesso2.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"&txtCenarioTrSequencia="+document.frm1.txtCenarioTrSequencia.value+"'");
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
if (document.frm1.txtCenarioTrSequencia.value == "")
     { 
	 alert("O preenchimento do campo Sequência é obrigatório.");
     document.frm1.txtCenarioTrSequencia.focus();
     return;
     }	 
  ano=document.frm1.txtCenarioTrSequencia.value 
  ok = 1
  for(i=0; i<4; ++i) {
      digito=(ano.substr(i,1));
      caracter=digito.charCodeAt(0);
      if ((caracter<48||caracter>57)) 
         ok = 0;
  }   
  if (ok == 0) {
     alert("O preenchimento do campo Sequência é obrigatório com número !");
     document.frm1.txtCenarioTrSequencia.focus();
     return;
     }	 
if (document.frm1.selCenario.value == "0")
     { 
	 alert("A seleção de um Cenário é obrigatória!.");
     document.frm1.selCenario.focus();
     return;
     }	 
	 else
if (document.frm1.txtDescTransacao.value == "")
     { 
	 alert("O preenchimento do campo Descrição da Transação é obrigatório.");
     document.frm1.txtDescTransacao.focus();
     return;
     }	 
	 else	 
	 {
	 //alert(document.frm1.selCenario.options[document.frm1.selCenario.selectedIndex].text);
	 //document.frm1.txtDescTransacao.value = document.frm1.selCenario.options[document.frm1.selCenario.selectedIndex].text;	 
	 document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
}

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form name="frm1" method="POST" action="grava_operacoes_especiais.asp">
  <table width="773" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td height="66" colspan="2">&nbsp;</td>
      <td height="66" colspan="2">&nbsp;</td>
      <td valign="top" colspan="2" height="66"> 
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="111">&nbsp; </td>
      <td height="20" width="30"><a href="javascript:Confirma()"><img src="../../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
      <td height="20" width="213"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Enviar</b></font></td>
      <td colspan="2" height="20">
        <div align="right"><a href="javascript:Limpa()"><img src="../../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></div>
      </td>
      <td height="20" width="334"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Limpa</b></font></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%" height="21">
<%'=str_SQL_DescCenario%></td>
      <td width="62%" height="21">&nbsp;</td>
      <td width="14%" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alterar</font><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        chamada para um Cen&aacute;rio</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="54%">&nbsp;</td>
      <td width="27%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%"> 
        <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Mega-Proceso 
          :&nbsp; </font></div>
      </td>
      <td width="54%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=str_MegaProcesso%>-<%=ls_Desc_MegaProcesso%> 
        <input type="hidden" name="txtMegaProcesso" value="<%=str_MegaProcesso%>">
        </font></b></td>
      <td width="27%"><%=str_MegaProcesso%></td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%"> 
        <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Proceso 
          :&nbsp; </font></div>
      </td>
      <td width="54%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=str_Processo%>-<%=ls_Desc_Processo%> 
        <input type="hidden" name="txtProcesso" value="<%=str_Processo%>">
        </font></b></td>
      <td width="27%"><%=str_Processo%></td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%"> 
        <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Sub-Proceso 
          :&nbsp; </font></div>
      </td>
      <td width="54%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=str_SubProcesso%>-<%=ls_Desc_SubProcesso%> 
        <input type="hidden" name="txtSubProcesso" value="<%=str_SubProcesso%>">
        </font></b></td>
      <td width="27%"><%=str_SubProcesso%></td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Cen&aacute;rio 
          :&nbsp; </font></div>
      </td>
      <td width="54%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <%=str_Cenario%> 
        <input type="hidden" name="txtCenario" value="<%=str_Cenario%>">
        </font></b></td>
      <td width="27%"><%=str_Cenario%></td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="54%">&nbsp;</td>
      <td width="27%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="2%">&nbsp;</td>
      <td width="17%">&nbsp;</td>
      <td width="54%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <%If not rdsDescCenario.EOF then%>
        <%=rdsDescCenario("CENA_TX_TITULO_CENARIO")%> 
        <% end if 
		%>
        </font> 
        <%If rdsDescCenario.EOF then%>
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Este 
        cenário não possuir transações relacionadas.</font> 
        <% end if 
     	rdsDescCenario.close
		set rdsDescCenario = Nothing
		%>
        </b></td>
      <td width="27%">&nbsp;</td>
    </tr>
    <tr bgcolor="#0066CC"> 
      <td width="2%"></td>
      <td width="17%"></td>
      <td width="54%"></td>
      <td width="27%" height="3"></td>
    </tr>
  </table>
  <table width="779" border="0" cellspacing="2" cellpadding="0">
    <tr> 
      <td width="49">&nbsp;</td>
      <td width="139"> 
        <input type="hidden" name="txtAcao" value="ACC">
      </td>
      <td width="583"> 
        <input type="hidden" name="txtCenarioChSequencia" value="<%=str_CenarioChSequencia%>">
      </td>
    </tr>
    <tr> 
      <td width="49"> 
        <%'=str_MegaProcesso2%>
      </td>
      <td width="139"> 
        <div align="right"><b><font color="#000099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Sequ&ecirc;ncia 
          </font></b></div>
      </td>
      <td width="583"> 
        <input type="text" name="txtCenarioTrSequencia" size="4" maxlength="4" value="<%=str_CenarioTrSequencia%>">
      </td>
    </tr>
    <tr> 
      <td width="49">&nbsp;</td>
      <td width="139"> 
        <div align="right"><b><font color="#000099" size="2" face="Verdana, Arial, Helvetica, sans-serif">Mega-Processo 
          </font></b></div>
      </td>
      <td width="583"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"> 
        <select name="selMegaProcesso2" onChange="MM_goToURL1('self','inc_chamada_cenario.asp');return document.MM_returnValue">
          <option value="0" >Selecione um Mega Processo</option>
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
        </font></td>
    </tr>
    <tr> 
      <td width="49">&nbsp;</td>
      <td width="139"> 
        <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
          Chama Cen&aacute;rio </font></b></div>
      </td>
      <td width="583"> 
        <select name="selCenario">
          <option value="0" selected>Sem Cenário</option>
          <%Set rdsCenario= Conn_db.Execute(str_SQL_Cenario)
While (NOT rdsCenario.EOF)
         if (Trim(str_CenarioSeguinte) = Trim(rdsCenario.Fields.Item("CENA_CD_CENARIO").Value)) then %>
          <option value="<%=(rdsCenario.Fields.Item("CENA_CD_CENARIO").Value)%>" selected ><%=(rdsCenario.Fields.Item("CENA_CD_CENARIO").Value)%>-<%=(rdsCenario.Fields.Item("CENA_TX_TITULO_CENARIO").Value)%></option>
          <% else %>
          <option value="<%=(rdsCenario.Fields.Item("CENA_CD_CENARIO").Value)%>"><%=(rdsCenario.Fields.Item("CENA_CD_CENARIO").Value)%>-<%=(rdsCenario.Fields.Item("CENA_TX_TITULO_CENARIO").Value)%></option>
          <% end if %>
          <%
  rdsCenario.MoveNext()
Wend
If (rdsCenario.CursorType > 0) Then
  rdsCenario.MoveFirst
Else
  rdsCenario.Requery
End If

rdsCenario.Close
set rdsCenario = Nothing
%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="49">&nbsp;</td>
      <td width="139" height="25"> 
        <div align="right"><b><font size="2" color="#000099"><font face="Verdana, Arial, Helvetica, sans-serif">Desc</font><font face="Verdana, Arial, Helvetica, sans-serif"> 
          Opera&ccedil;&atilde;o especial</font></font></b></div>
      </td>
      <td width="598" height="25"> 
        <input type="text" name="txtDescTransacao" size="60" maxlength="150" value="<%=str_DescTransacao%>">
      </td>
    </tr>
    <tr> 
      <td width="49">&nbsp;</td>
      <td width="139">&nbsp;</td>
      <td width="583"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="4%" height="20"><img src="../../imagens/b021.gif" width="16" height="16"></td>
            <td width="96%" height="20"><font size="2" face="Arial, Helvetica, sans-serif" color="#CC6600">As 
              transa&ccedil;&otilde;es (R/3, manual, interface, exit, chamada) 
              dever&atilde;o ser cadastradas no infinitivo.</font></td>
          </tr>
          <tr> 
            <td width="4%" height="20">&nbsp;</td>
            <td width="96%" height="20"><font size="2" face="Arial, Helvetica, sans-serif" color="#CC6600">Exemplo: 
              &quot;Criar pedido ...&quot;</font></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="49">&nbsp;</td>
      <td width="139">&nbsp;</td>
      <td width="583"> 
        <input type="hidden" name="txtDescTransacao2" value="0">
      </td>
    </tr>
    <tr> 
      <td width="49">&nbsp;</td>
      <td width="139">&nbsp;</td>
      <td width="583">&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp; </p>
</form>
</body>
</html>
