<%@LANGUAGE="VBSCRIPT"%> 
<%
str_Opc = Request("txtOpc")

if Request("selMegaProcesso") = "" then
   str_MegaProcesso = "0"
else
	str_MegaProcesso = Request("selMegaProcesso")
end if

if str_MegaProcesso <> "0" then
   Session("MegaProcesso") = str_MegaProcesso
else
    if Session("MegaProcesso") <> "" then
       str_MegaProcesso = Session("MegaProcesso") 
	end if   
end if

'RESPONSE.Write(Session("Conn_String_Cogest_Gravacao"))
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

str_SQL_Max_Seq_Proc = ""
str_SQL_Max_Seq_Proc = str_SQL_Max_Seq_Proc & " SELECT "
str_SQL_Max_Seq_Proc = str_SQL_Max_Seq_Proc & " MAX(" & Session("PREFIXO") & "PROCESSO.PROC_NR_SEQUENCIA) AS MaxSeq"
str_SQL_Max_Seq_Proc = str_SQL_Max_Seq_Proc & " FROM " & Session("PREFIXO") & "PROCESSO"
str_SQL_Max_Seq_Proc = str_SQL_Max_Seq_Proc & " GROUP BY " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_SQL_Max_Seq_Proc = str_SQL_Max_Seq_Proc & " HAVING " & Session("PREFIXO") & "PROCESSO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso

Set rdsMaxSeqProcesso= Conn_db.Execute(str_SQL_Max_Seq_Proc)
if rdsMaxSeqProcesso.EOF then
   ls_int_MaxProcesso = 0
   AA = "AQUI 1"
else
   IF not IsNull(rdsMaxSeqProcesso("MaxSeq")) then
      ls_int_MaxProcesso = rdsMaxSeqProcesso("MaxSeq")   
   else
      ls_int_MaxProcesso = 0
   end if	  
   AA = "AQUI 2"
end if
rdsMaxSeqProcesso.close
set rdsMaxSeqProcesso = Nothing
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
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?txtOpc="+document.frm1.txtOpc.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
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
if ((document.frm1.txtNovoProc1.value == "")&&
	(document.frm1.txtNovoProc2.value == "")&&
	(document.frm1.txtNovoProc3.value == "")&&
	(document.frm1.txtNovoProc4.value == "")&&
	(document.frm1.txtNovoProc5.value == "")&&
	(document.frm1.txtNovoProc6.value == "")&&
	(document.frm1.txtNovoProc7.value == "")&&
	(document.frm1.txtNovoProc8.value == "")&&
	(document.frm1.txtNovoProc9.value == "")&&
	(document.frm1.txtNovoProc10.value == ""))
     { 
	 alert("Preencha um novo Processo.");
     document.frm1.txtNovoProc1.focus();
     return;
     }	 
	 else
     {
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="grava_processo.asp">
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
  <table width="90%" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%"><%'=Session("MegaProcesso")%></td>
      <td width="70%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%"><%'=str_Opc%></td>
      <td width="70%"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Novo 
        Processo</font></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%"><%'=str_SQL_Max_Seq_Proc%></td>
      <td width="70%"> <%'=str_MegaProcesso%> 
        <%'=Session("Conn_String_Cogest_Gravacao")%>
        <%'=ls_int_MaxProcesso%>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Mega-Processo</b></font></td>
      <td width="70%"> 
        <select name="selMegaProcesso" onChange="MM_goToURL1('self','cadas_processo.asp');return document.MM_returnValue">
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
      <td width="10%">&nbsp;</td>
      <td width="20%">
        <input type="hidden" name="txtOpc" value="<%=str_OPC%>">
      </td>
      <td width="70%"> 
        <table width="89%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%">&nbsp;</td>
            <td width="39%"> 
              <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Sequ&ecirc;ncia</b></font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%" height="30">&nbsp;</td>
      <td width="20%" height="30"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Novos 
        Processos</b></font></td>
      <td width="70%" height="30"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc1" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq1" size="7" value="<%=ls_int_MaxProcesso+10%>">
                </font></div>
            </td>
          </tr>
        </table>
        
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc2" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq2" size="7" value="<%=ls_int_MaxProcesso+20%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc3" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq3" size="7" value="<%=ls_int_MaxProcesso+30%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc4" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq4" size="7" value="<%=ls_int_MaxProcesso+40%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc5" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq5" size="7" value="<%=ls_int_MaxProcesso+50%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc6" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq6" size="7" value="<%=ls_int_MaxProcesso+60%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc7" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq7" size="7" value="<%=ls_int_MaxProcesso+70%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc8" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq8" size="7" value="<%=ls_int_MaxProcesso+80%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc9" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq9" size="7" value="<%=ls_int_MaxProcesso+90%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"> 
        <table width="90%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="61%"> 
              <input type="text" name="txtNovoProc10" size="50" maxlength="150">
            </td>
            <td width="39%"> 
              <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"> 
                <input type="text" name="txtSeq10" size="7" value="<%=ls_int_MaxProcesso+100%>">
                </font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%">&nbsp; </td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%"><img src="../imagens/aprova_02.gif" width="56" height="17" onClick="MM_openBrWindow('teste.asp','','')"></td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
      <td width="70%">&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
