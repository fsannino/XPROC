 
<%
server.scripttimeout=99999999
on error resume next

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

if (Request("selMegaProcesso") <> "") then 
    str_MegaProcesso = Request("selMegaProcesso")
else
    str_MegaProcesso = "0"
end if

if (Request("txtDescMegaProcesso") <> "") then 
    str_DescMegaProcesso = Request("txtDescMegaProcesso")
else
    str_DescMegaProcesso = "Não passou"
end if

response.write str_DescMegaProcesso

if (Request("selFuncao") <> "") then 
    str_Funcao = Request("selFuncao")
else
    str_Funcao = "0"
end if

if (Request("txtDescFuncao") <> "") then 
    str_DescFuncao = Request("txtDescFuncao")
else
    str_DescFuncao = "Não passado"
end if

response.write str_DescFuncao

ls_str_SQL_Func_Tran = ""
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & ""
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " SELECT "
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO,"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " FUN_NEG_TRANSACAO.PROC_CD_PROCESSO,"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO,"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA,"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " MEGA_PROCESSO.MEPR_TX_ABREVIA, "
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " PROCESSO.PROC_TX_DESC_PROCESSO, "
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " SUB_PROCESSO.SUPR_TX_DESC_SUB_PROCESSO, "
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " ATIVIDADE_CARGA.ATCA_TX_DESC_ATIVIDADE, "
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO, "
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " TRANSACAO.TRAN_TX_DESC_TRANSACAO"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " FROM FUN_NEG_TRANSACAO "
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " INNER JOIN MEGA_PROCESSO ON FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " INNER JOIN TRANSACAO ON FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = TRANSACAO.TRAN_CD_TRANSACAO"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " INNER JOIN ATIVIDADE_CARGA ON FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA = ATIVIDADE_CARGA.ATCA_CD_ATIVIDADE_CARGA"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " INNER JOIN SUB_PROCESSO ON FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = SUB_PROCESSO.MEPR_CD_MEGA_PROCESSO AND FUN_NEG_TRANSACAO.SUPR_CD_SUB_PROCESSO = SUB_PROCESSO.SUPR_CD_SUB_PROCESSO AND FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = SUB_PROCESSO.PROC_CD_PROCESSO"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " INNER JOIN PROCESSO ON FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = PROCESSO.MEPR_CD_MEGA_PROCESSO AND FUN_NEG_TRANSACAO.PROC_CD_PROCESSO = PROCESSO.PROC_CD_PROCESSO"
ls_str_SQL_Func_Tran = ls_str_SQL_Func_Tran & " WHERE FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Funcao & "'"

'Set rdsFunc_Tran = Conn_db.Execute(ls_str_SQL_Func_Tran)

%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--
function MM_goToURL4() { //v3.0
  var i, args=MM_goToURL4.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value+"'");
}
function MM_goToURL5() { //v3.0
  var i, args=MM_goToURL5.arguments; document.MM_returnValue = false;
  //for (i=0; i<(args.length-1); i+=3) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"&p_CenarioChSequencia="+args[3]+"'");
  for (i=0; i<(args.length-1); i+=3) eval(args[i]+".location='"+args[i+1]+"?"+args[3]+"'");

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

<body topmargin="0" leftmargin="0" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../imagens/novo_registro_02.gif')">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
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
    <td colspan="3" height="9">&nbsp; </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="24%">&nbsp;</td>
    <td width="50%"> 
      <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Edi&ccedil;&atilde;o</font><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
        Fun&ccedil;&atilde;o R/3 x Transa&ccedil;&atilde;o</font> 
      </div>
    </td>
    <td width="26%"> 
      <%'=str_SQL_Cenario%>
    </td>
  </tr>
</table>
<table width="586" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="34" height="22">&nbsp;</td>
    <td width="179" height="22">&nbsp;</td>
    <td width="373" height="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="34" height="22">&nbsp;</td>
    <td width="179" height="22"> 
      <div align="right"> <font color="#330099" face="Verdana" size="2">Mega-Processo 
        : </font> </div>
    </td>
    <td width="373" height="22"><b><font color="#330099" face="Verdana" size="2"><%=str_MegaProcesso%> - 
      <%str_DescMegaProcesso%>
      <input type="hidden" name="txtMegaProcesso" value="<%=str_MegaProcesso%>">
      </font></b></td>
  </tr>
  <tr> 
    <td width="34">&nbsp;</td>
    <td width="179"> 
      <div align="right"> <font color="#330099" face="Verdana" size="2">Fun&ccedil;&atilde;o R/3 : </font> </div>
    </td>
    <td width="373"><b><font color="#330099" face="Verdana" size="2"><%=str_Funcao%> 
      <input type="hidden" name="txtFuncao" value="<%=str_Funcao%>">
      </font></b></td>
  </tr>
  <tr> 
    <td width="34">&nbsp;</td>
    <td width="179">&nbsp;</td>
    <td width="373">&nbsp;</td>
  </tr>
  <tr> 
    <td width="34">&nbsp;</td>
    <td width="179">&nbsp;</td>
    <td width="373"><b><font color="#330099" face="Verdana" size="2"><%=str_DescFuncao%></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099"> 
      </font> </b></td>
  </tr>
  <% If rdsFunc_Tran.EOF then %>
  <tr> 
    <td width="34">&nbsp;</td>
    <td width="179">&nbsp;</td>
    <td width="373"><b> 
      <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">Esta 
      fun&ccedil;&atilde;o não possuir transações associadas.</font>       
      </b></td>
  </tr>
  <% end if %>
</table>
<table width="779" border="0" cellspacing="2" cellpadding="0" align="center">
  <tr>
    <td width="201">&nbsp;</td>
    <td width="26">&nbsp;</td>
    <td width="216">&nbsp;</td>
    <td width="27">&nbsp;</td>
    <td width="216">&nbsp;</td>
    <td width="27">&nbsp;</td>
    <td width="1">&nbsp;</td>
    <td width="23">&nbsp;</td>
  </tr>
  <tr> 
    <td width="201"> 
      <div align="right"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Inclus&atilde;o 
        Nova Transa&ccedil;&atilde;o&nbsp;</font></b></div>
    </td>
    <td width="26"><a href="javascript:MM_goToURL4('self','cad_cenario_transacao.asp',this)" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image131','','../../imagens/novo_registro_02.gif',1)"><img name="Image131" border="0" src="../../imagens/novo_registro_02_off.gif" width="22" height="22" alt="Inclui nova Transa&ccedil;&atilde;o"></a> 
    </td>
    <td width="216">&nbsp;</td>
    <td width="27">&nbsp;</td>
    <td width="216">&nbsp;</td>
    <td width="27">&nbsp;</td>
    <td width="1">&nbsp;</td>
    <td width="23">&nbsp;</td>
  </tr>
  <tr> 
    <td width="216">&nbsp;</td>
    <td width="27">&nbsp;</td>
    <td width="216">&nbsp;</td>
    <td width="27">&nbsp;</td>
    <td width="216">&nbsp;</td>
    <td width="27">&nbsp;</td>
    <td width="1">&nbsp;</td>
    <td width="23">&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="1" cellpadding="2" width="870" bordercolor="#000000">
  <tr> 
    <td width="61" bgcolor="#330099"> 
      <div align="center"><font size="2"><b><font face="Verdana" color="#FFFFFF">A&ccedil;&atilde;o</font></b></font></div>
    </td>
    <td width="167" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Mega-Processo</font></b></td>
    <td width="132" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Processo</font></b></td>
    <td width="184" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Sub-Processo</font></b></td>
    <td width="171" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Atividade</font></b></td>
    <td width="124" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Transação</font></b></td>
  </tr>
  <% If Not rdsFunc_Tran.EOF then %>
  <%Do while not rdsFunc_Tran.EOF %>
  <tr> 
    <td width="61" height="24"> 
      <div align="center"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="javascript:MM_goToURL5('self','exc_funcao_trans.asp',this,'selFuncao=<%=str_Funcao%>&selMegaProcesso=<%=rdsFunc_Tran("MEPR_CD_MEGA_PROCESSO")%>&selProcesso=<%=rdsFunc_Tran("PROC_CD_PROCESSO")%>&selSubProcesso=<%=rdsFunc_Tran("SUPR_CD_SUB_PROCESSO")%>&selAtivCarga=<%=rdsFunc_Tran("ATCA_CD_ATIVIDADE_CARGA")%>&selTrans=<%=rdsFunc_Tran("TRAN_CD_TRANSACAO")%>')">Exc</a></font></b> 
      </div>
    </td>
    <td width="167" height="24"><%=rdsFunc_Tran("MEPR_TX_DESC_MEGA_PROCESSO")%></td>
    <td width="132" height="24"><font face="Verdana" size="1"><%=rdsFunc_Tran("PROC_TX_DESC_PROCESSO")%></font></td>
    <td width="184" height="24"><font face="Verdana" size="1"><%=rdsFunc_Tran("SUPR_TX_DESC_SUB_PROCESSO")%></font></td>
    <td width="171" height="24"><font face="Verdana" size="1"><%=rdsFunc_Tran("ATCA_TX_DESC_ATIVIDADE")%></font></td>
    <td width="124" height="24"><font face="Verdana" size="1"><%=rdsFunc_Tran("TRAN_CD_TRANSACAO")%></font></td>
  </tr>
  <%
  rdsFunc_Tran.MOVENEXT
  Loop
  rdsFunc_Tran.close
  set rdsFunc_Tran = Nothing
  %>
</table>
<% else %>
<b><font face="Verdana" size="2" color="#800000">
Nenhum Registro Encontrado</font></b>
<% end if %>
</body>
</html>