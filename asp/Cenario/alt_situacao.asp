 
<%
Dim str_Opc
Dim str_MegaProcesso
Dim str_Processo
Dim str_SubProcesso
Dim str_Cenario

str_MegaProcesso = "0"
str_Processo = "0"
str_SubProcesso = "0"
str_Cenario = 0

str_Opc = Request("txtOpc")

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
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

str_SQL_DescCenario = ""
str_SQL_DescCenario = str_SQL_DescCenario & " SELECT "
str_SQL_DescCenario = str_SQL_DescCenario & " CENA_CD_CENARIO"
str_SQL_DescCenario = str_SQL_DescCenario & " , CENA_TX_TITULO_CENARIO"
str_SQL_DescCenario = str_SQL_DescCenario & " , CENA_TX_SITUACAO"
str_SQL_DescCenario = str_SQL_DescCenario & " FROM " & Session("PREFIXO") & "CENARIO"
str_SQL_DescCenario = str_SQL_DescCenario & " where CENA_CD_CENARIO = '" &  str_Cenario & "'"
Set rdsDescCenario= Conn_db.Execute(str_SQL_DescCenario)


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
function exibe_historico()
{
window.open("exibe_historico.asp?txtCenario="+document.frm1.txtCenario.value,"_blank","width=650,height=240,history=0,scrollbars=1,titlebar=0,resizable=0,left=100,top=100")
}

function Confirma2() 
{ 
	  document.frm1.submit();
}
function Confirma()
{
if (document.frm1.txtStNovo.value=="")
{
alert("Não existe mais nenhum status disponível para alteração");
return;
}
  if (document.frm1.txtStNovo.value=="DS")
     {
     window.location.href = "altera_status.asp?id="+document.frm1.txtCenario.value;
     }
	 else
	 {
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
      <td width="24%"><%'=str_SQL_DescCenario%></td>
      <td width="62%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Altera 
        status de um Cen&aacute;rio</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
      <td width="71%" colspan="2">&nbsp;</td>
      <td width="12%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="16%"> 
        <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Mega-Proceso 
          :&nbsp; </font></div>
      </td>
      <td width="40%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=str_MegaProcesso%>-<%=ls_Desc_MegaProcesso%></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <input type="hidden" name="selMegaProcesso" value="<%=str_MegaProcesso%>">
        </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        </font></b></td>
      <td width="31%"></td>
      <td width="12%"> 
        <div align="center"></div>
      </td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="16%"> 
        <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Proceso 
          :&nbsp; </font></div>
      </td>
      <td width="40%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=str_Processo%>-<%=ls_Desc_Processo%></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <input type="hidden" name="selProcesso" value="<%=str_Cenario%>">
        </font></b></td>
      <td width="31%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2"><a href="valida_altera_status.asp?option=1&txtCenario=<%=str_cenario%>"><img border="0" src="../../imagens/b011.gif" width="21" height="21" align="absmiddle"></a>
        Zerar Status</font></td>
      <td width="12%"> 
        <div align="center"><b><font face="Verdana" size="2" color="#330099"></font></b></div>
      </td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="16%"> 
        <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Sub-Proceso 
          :&nbsp; </font></div>
      </td>
      <td width="40%"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"><%=str_SubProcesso%>-<%=ls_Desc_SubProcesso%></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <input type="hidden" name="txtCenario4" value="<%=str_Cenario%>">
        </font></b></td>
      <td width="31%"><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="1"><b>(O
        Status retornará para EM ELABORAÇÃO)</b></font></td>
      <td width="12%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="16%"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Cen&aacute;rio 
          :&nbsp; </font></div>
      </td>
      <td width="40%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <%=str_Cenario%> 
        <input type="hidden" name="txtCenario" value="<%=str_Cenario%>">
        </font></b></td>
      <td width="31%"></td>
      <td width="12%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
      <td width="71%" colspan="2"><b></b></td>
      <td width="12%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="1%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
      <td width="71%" colspan="2"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        <%If not rdsDescCenario.EOF then%>
          <%=rdsDescCenario("CENA_TX_TITULO_CENARIO")%> 
		  <%ls_Situacao_Anterior = rdsDescCenario("CENA_TX_SITUACAO")
		    ls_Cod_Cenario = rdsDescCenario("CENA_CD_CENARIO")
         end if 
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
      <td width="12%">&nbsp;</td>
    </tr>
    <tr bgcolor="#0066CC"> 
      <td width="1%"></td>
      <td width="16%"></td>
      <td width="71%" colspan="2"></td>
      <td width="12%" height="3"></td>
    </tr>
  </table>
  <table width="779" border="0" cellspacing="2" cellpadding="0">
    <tr> 
      <td width="15">&nbsp;</td>
      <td width="152"> 
        <input type="hidden" name="txtAcao" value="AS">
      </td>
      <td width="119">&nbsp;</td>
      <td width="483">&nbsp; </td>
    </tr>
    <tr> 
      <td width="15">&nbsp;</td>
      <td width="152">&nbsp;</td>
      <td width="119">&nbsp;</td>
      <td width="483">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15" height="12">&nbsp;</td>
      <td width="152" valign="top" height="12"> 
        <div align="right"><b></b></div>
      </td>
      <td width="119" valign="top" height="12">&nbsp;</td>
      <td width="483" valign="top" height="12">&nbsp; </td>
    </tr>
    <tr> 
      <%
	ls_SQL = ""
	ls_SQL = ls_SQL + " Select SITU_TX_REFERENTE, SITU_NR_SEQUENCIA_STATUS "
	ls_SQL = ls_SQL + " , SITU_TX_DESC_SITUACAO, SITU_TX_CD_STATUS "
	ls_SQL = ls_SQL + " from " & Session("PREFIXO") & "SITUACAO_GERAL "
	ls_SQL = ls_SQL + " where SITU_TX_CD_STATUS = '" & ls_Situacao_Anterior & "'"
    Set rdsStatus= Conn_db.Execute(ls_SQL)
    if not rdsStatus.EOF then
	   ls_Desc_Situacao_Anterior = rdsStatus("SITU_TX_DESC_SITUACAO")
       ls_Seq_Situacao_Anterior = rdsStatus("SITU_NR_SEQUENCIA_STATUS")
	   ls_Referente = rdsStatus("SITU_TX_REFERENTE")
	   ls_Cod_Situacao_Anterior = rdsStatus("SITU_TX_CD_STATUS")
    else
       ls_Desc_Situacao_Anterior = ""
       ls_Seq_Situacao_Anterior = 0
	   ls_Referente = "0"
	   ls_Cod_Situacao_Anterior = 0
    end if
    rdsStatus.close
	set rdsStatus = Nothing

	ls_SQL = ""
	ls_SQL = ls_SQL + " Select Max(SITU_NR_SEQUENCIA_STATUS) as Maximo from " & Session("PREFIXO") & "SITUACAO_GERAL "
	ls_SQL = ls_SQL + " where SITU_TX_REFERENTE = '" & ls_Referente & "'"
	Set rdsMaxSeq= Conn_db.Execute(ls_SQL)
 	ls_Maximo = rdsMaxSeq("Maximo")
	rdsMaxSeq.close
	set rdsMaxSeq = nothing
    
	if ls_Seq_Situacao_Anterior < ls_Maximo  then
	
       ls_SQL = ""
	   ls_SQL = ls_SQL + " Select SITU_NR_SEQUENCIA_STATUS, SITU_TX_CD_STATUS, SITU_TX_DESC_SITUACAO "
	   ls_SQL = ls_SQL + " from " & Session("PREFIXO") & "SITUACAO_GERAL "
	   ls_SQL = ls_SQL + " where SITU_TX_REFERENTE = '" & ls_Referente & "'"
	   ls_SQL = ls_SQL + " and SITU_NR_SEQUENCIA_STATUS = " & ls_Seq_Situacao_Anterior + 1
	   Set rdsStatusSeq= Conn_db.Execute(ls_SQL)
       if not rdsStatusSeq.EOF then
          ls_Seq_Situacao_Nova = rdsStatusSeq("SITU_NR_SEQUENCIA_STATUS")
		  ls_Desc_Situacao_Nova = rdsStatusSeq("SITU_TX_DESC_SITUACAO")
          ls_Situacao_Nova = rdsStatusSeq("SITU_TX_CD_STATUS")
       else
	      ls_Seq_Situacao_Nova = "0"
		  ls_Desc_Situacao_Nova = "0"
		  ls_Situacao_Nova = "0"
       end if
       rdsStatusSeq.close
	   set rdsStatusSeq = Nothing
	
	else
	   ls_MSG = "Este Cenário já está no último Status."    
    end if 
	%>
      <td width="15">&nbsp;</td>
      <td width="152"> 
        <div align="right"><b><font size="2" color="#000099" face="Verdana, Arial, Helvetica, sans-serif">Status 
          do Cen&aacute;rio : </font></b></div>
      </td>
      <td width="119"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Status 
        Atual</font> :<font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#000099">&nbsp; 
        </font></td>
      <td width="483"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#000099"><b><font color="#FF0000"><%=ls_Desc_Situacao_Anterior%> 
        <input type="hidden" name="txtStAtual" value="<%=Trim(ls_Cod_Situacao_Anterior)%>">
        </font></b> </font> </td>
    </tr>
    <tr> 
      <td width="15">&nbsp;</td>
      <td width="152">&nbsp;</td>
      <td width="119"><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Pr&oacute;ximo 
        Status </font> :</td>
      <td width="483"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#000099"> 
        <% if Trim(ls_Cod_Situacao_Anterior)="DS" then
		      ls_Situacao_Nova = "DS" 
			  ls_Desc_Situacao_Nova = "DESENHADO"
		   end if
		%>
		<%
		if trim(ls_Desc_Situacao_Nova)="" then
		   ls_Desc_Situacao_Nova="NENHUM STATUS À SEGUIR"
		   nenhum=1
		end if
		%>
        <%=ls_Desc_Situacao_Nova%> 
        <input type="hidden" name="txtStNovo" value="<%=Trim(ls_Situacao_Nova)%>">
        </font></td>
    </tr>
    <tr>
      <td width="15">&nbsp;</td>
      <td width="152">&nbsp;</td>
      <td width="119">&nbsp;</td>
      <td width="483">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15">&nbsp;</td>
      <td width="152">&nbsp;</td>
      <td width="119">&nbsp;</td>
      <%
	  if nenhum<>1 then
	  %>
	  <td width="483"><font face="Verdana, Arial, Helvetica, sans-serif" size="3" color="#000099"><font size="2">(Clicar 
        <b>Enviar</b> para confirmar Pr&oacute;ximo Status)</font></font></td>
    </tr>
	<%end if%>
  </table>
  <p>&nbsp; </p>
</form>
</body>
</html>
