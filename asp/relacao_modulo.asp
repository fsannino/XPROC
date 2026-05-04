<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
server.scripttimeout=999999999

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Atividade=request("selAtiv")
str_Modulo=request("selModu")
de=request("range1")
ate=request("range2")

'response.write str_Atividade

set rs_destino=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA="& str_Atividade & " AND MODU_CD_MODULO=" & str_Modulo)

ssql = ""
ssql="SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE"

SET RS=CONN_DB.EXECUTE(SSQL)

ssql = ""
ssql="SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY MODU_TX_DESC_MODULO"

SET RS2=CONN_DB.EXECUTE(SSQL)

if len(de)=0 and len(ate)=0 then
	RS_TRANSACAO = "SELECT TRAN_CD_TRANSACAO,TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO"
else
	if len(ate)=0 then
		RS_TRANSACAO = "SELECT TRAN_CD_TRANSACAO,TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO >'" & de & "%' ORDER BY TRAN_CD_TRANSACAO"
	else
		RS_TRANSACAO = "SELECT TRAN_CD_TRANSACAO,TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO >'" & de & "%' AND TRAN_CD_TRANSACAO <='" & ate & "%' ORDER BY TRAN_CD_TRANSACAO"
	end if
end if

%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--

function Confirma2() 
{
if (document.frm1.selModulo.selectedIndex == 0)
     { 
     document.frm1.selModulo.focus();
     return;
     }
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	  document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href='relacao_modulo.asp?selAtiv='+document.frm1.selAtividade.value + '&selModu='+document.frm1.selModulo.value+ '&Range1='+document.frm1.range1.value+'&Range2='+document.frm1.range2.value
	 }
 }

function Confirma3() 
{
if (document.frm1.selModulo.selectedIndex == 0)
     { 
     document.frm1.selModulo.focus();
     return;
     }
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href='relacao_modulo.asp?selAtiv='+document.frm1.selAtividade.value + '&selModu='+document.frm1.selModulo.value+ '&Range1='+document.frm1.range1.value+'&Range2='+document.frm1.range2.value
	 }
 }


function carrega_txt(fbox) {
document.frm1.txtEmpSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtEmpSelecionada.value = document.frm1.txtEmpSelecionada.value + "," + fbox.options[i].value;
   }
}

function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
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
<form name="frm1" method="post" action="grava_relmodulo.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
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
          <td width="26"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:Confirma3()"></td>
          <td width="169"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Alterar
            Intervalo</b></font></td>
            <td width="27">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <div align="left">
  <table width="1056" border="0" cellpadding="0" cellspacing="0" align="left">
    <tr>
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
        <td width="1050" height="21"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Relação 
          Agrupamento ( Master List R/3) x Atividade x Transaçao</font></td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21"> 
        <input type="hidden" name="txtOpc" value="<%=str_Opc%>">
      </td>
    </tr>
    <tr> 
      <td width="4" height="25">&nbsp;</td>
        <td width="134" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Agrupamento 
          ( Master List R/3 )</b></font></td>
      <td width="1050" height="25"> 
        <select name="selModulo" size="1" onchange="javascript:Confirma3()">
          <option value="0" selected>Selecione um Agrupamento ( Master List R/3 )</option>
          <%DO UNTIL RS2.EOF=TRUE
          if trim(str_Modulo)=trim(RS2("MODU_CD_MODULO")) then
          %>
              <option selected value=<%=RS2("MODU_CD_MODULO")%>><%=RS2("MODU_TX_DESC_MODULO")%></option>
          <%else%>
              <option value=<%=RS2("MODU_CD_MODULO")%>><%=RS2("MODU_TX_DESC_MODULO")%></option>
          <%
			end if
			RS2.MOVENEXT
			LOOP
			%>
          </select>
      </td>
    </tr>
    <tr> 
      <td width="4" height="18"></td>
      <td width="134" height="18"></td>
      <td width="1050" height="18"></td>
    </tr>
    <tr> 
      <td width="4" height="25"></td>
      <td width="134" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Atividade</b></font></td>
      <td width="1050" height="25"><select name="selAtividade" size="1" onchange="javascript:Confirma2()">
          <option value="0" selected>Selecione uma Atividade</option>
          <%DO UNTIL RS.EOF=TRUE
          if trim(str_Atividade)=trim(RS("ATCA_CD_ATIVIDADE_CARGA")) then
          %>
              <option selected value=<%=RS("ATCA_CD_ATIVIDADE_CARGA")%>><%=RS("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%else%>
              <option value=<%=RS("ATCA_CD_ATIVIDADE_CARGA")%>><%=RS("ATCA_TX_DESC_ATIVIDADE")%></option>
			<%
			end if
			RS.MOVENEXT
			LOOP
			%>
          </select>
      </td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="4" height="21"></td>
      <td width="134" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Intervalo
        Transações</b></font></td>
      <td width="1050" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>de
        <input type="text" name="range1" size="11" value="<%=de%>">
         à 
         <input type="text" name="range2" size="10" value="<%=ate%>">
        </b></font> </td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21">&nbsp; </td>
    </tr>
    <tr> 
      <td width="4" height="226">&nbsp;</td>
      <td width="134" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Selecione
        Transações&nbsp;</b></font></td>
      <td width="1050" height="226"> 
        <table width="1050" border="0" align="left" cellpadding="0" cellspacing="0">
          <tr> 
            <td valign="middle" align="left"> 
              <div align="LEFT"> <b> 
                <select name="list1" size="8" multiple>
                 <%
                 Set RS1 = Conn_db.Execute(RS_TRANSACAO)
                 
                 DO UNTIL RS1.EOF=TRUE
                 JATEM=0
                 
                 ON ERROR RESUME NEXT
                 
                 RS_DESTINO.MOVEFIRST
                 
                 DO UNTIL RS_DESTINO.EOF=TRUE
                 		IF TRIM(RS1("TRAN_CD_TRANSACAO"))=TRIM(RS_DESTINO("TRAN_CD_TRANSACAO")) THEN
                 			JATEM=JATEM+1	
						END IF
						RS_DESTINO.MOVENEXT
                 LOOP
                 	IF JATEM=0 THEN
                 	%>
                  <option value="<%=RS1("TRAN_CD_TRANSACAO")%>" ><%=RS1("TRAN_CD_TRANSACAO")%>-<%=RS1("TRAN_TX_DESC_TRANSACAO")%></option>
                  <%
                 END IF
  					RS1.MoveNext
					LOOP
					%>
                </select>
                </b></div>
            </td>
            <td width="26" align="left" valign="middle"> 
              <table width="53%" cellpadding="0" cellspacing="0" border="0" align="LEFT">
                <tr> 
                  <td width="100%" valign="middle" align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2)"><img name="Image16" border="0" src="../imagens/continua_F01.gif" width="24" height="24"></a></td>
                </tr>
                <tr> 
                  <td height="25">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="25" valign="middle" align="center"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1)"><img name="img01511" border="0" src="../imagens/continua2_F01.gif" width="24" height="24"></a></td>
                </tr>
              </table>
            </td>
            <td valign="middle" align="left"> 
              <div align="LEFT"><font color="#000080"> 
                <select name="list2" size="8" multiple>
                  <%
                RS_DESTINO.MOVEFIRST
                
                DO UNTIL RS_DESTINO.EOF=TRUE
                SSQL1="SELECT TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & rs_destino("TRAN_CD_TRANSACAO") & "'"
                SET RS_ATUAL=CONN_DB.EXECUTE(SSQL1)
                ATUAL=RS_ATUAL("TRAN_TX_DESC_TRANSACAO")
                %>
                  <option value=<%=RS_DESTINO("TRAN_CD_TRANSACAO")%>><%=RS_DESTINO("TRAN_CD_TRANSACAO")%>-<%=ATUAL%></option>
                  <%
                RS_DESTINO.MOVENEXT
                LOOP
                %>
                </select>
                </font></div>
            </td>
          </tr>
          <tr>
            <td colspan="3" width="1048">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3" width="1048"> 
              <div align="LEFT"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
                a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></div>
            </td>
          </tr>
          <tr> 
            <td width="570"> 
              <%'=str_SQL_Sub_Proc%>
            </td>
            <td width="26" align="left" valign="middle">&nbsp;</td>
            <td width="448"> 
              <input type="hidden" name="txtEmpSelecionada">
            </td>
          </tr>
        </table>
      </td>
    </tr>
    
  </table>
  </div>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
server.scripttimeout=999999999

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Atividade=request("selAtiv")
str_Modulo=request("selModu")
de=request("range1")
ate=request("range2")

'response.write str_Atividade

set rs_destino=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA="& str_Atividade & " AND MODU_CD_MODULO=" & str_Modulo)

ssql = ""
ssql="SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE"

SET RS=CONN_DB.EXECUTE(SSQL)

ssql = ""
ssql="SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY MODU_TX_DESC_MODULO"

SET RS2=CONN_DB.EXECUTE(SSQL)

if len(de)=0 and len(ate)=0 then
	RS_TRANSACAO = "SELECT TRAN_CD_TRANSACAO,TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO"
else
	if len(ate)=0 then
		RS_TRANSACAO = "SELECT TRAN_CD_TRANSACAO,TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO >'" & de & "%' ORDER BY TRAN_CD_TRANSACAO"
	else
		RS_TRANSACAO = "SELECT TRAN_CD_TRANSACAO,TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO >'" & de & "%' AND TRAN_CD_TRANSACAO <='" & ate & "%' ORDER BY TRAN_CD_TRANSACAO"
	end if
end if

%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--

function Confirma2() 
{
if (document.frm1.selModulo.selectedIndex == 0)
     { 
     document.frm1.selModulo.focus();
     return;
     }
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	  document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href='relacao_modulo.asp?selAtiv='+document.frm1.selAtividade.value + '&selModu='+document.frm1.selModulo.value+ '&Range1='+document.frm1.range1.value+'&Range2='+document.frm1.range2.value
	 }
 }

function Confirma3() 
{
if (document.frm1.selModulo.selectedIndex == 0)
     { 
     document.frm1.selModulo.focus();
     return;
     }
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href='relacao_modulo.asp?selAtiv='+document.frm1.selAtividade.value + '&selModu='+document.frm1.selModulo.value+ '&Range1='+document.frm1.range1.value+'&Range2='+document.frm1.range2.value
	 }
 }


function carrega_txt(fbox) {
document.frm1.txtEmpSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtEmpSelecionada.value = document.frm1.txtEmpSelecionada.value + "," + fbox.options[i].value;
   }
}

function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"'");
}
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
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
<form name="frm1" method="post" action="grava_relmodulo.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top" height="65"> 
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
          <td width="26"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:Confirma3()"></td>
          <td width="169"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Alterar
            Intervalo</b></font></td>
            <td width="27">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <div align="left">
  <table width="1056" border="0" cellpadding="0" cellspacing="0" align="left">
    <tr>
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
        <td width="1050" height="21"><font color="#003366" size="3" face="Verdana, Arial, Helvetica, sans-serif">Relação 
          Agrupamento ( Master List R/3) x Atividade x Transaçao</font></td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21"> 
        <input type="hidden" name="txtOpc" value="<%=str_Opc%>">
      </td>
    </tr>
    <tr> 
      <td width="4" height="25">&nbsp;</td>
        <td width="134" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Agrupamento 
          ( Master List R/3 )</b></font></td>
      <td width="1050" height="25"> 
        <select name="selModulo" size="1" onchange="javascript:Confirma3()">
          <option value="0" selected>Selecione um Agrupamento ( Master List R/3 )</option>
          <%DO UNTIL RS2.EOF=TRUE
          if trim(str_Modulo)=trim(RS2("MODU_CD_MODULO")) then
          %>
              <option selected value=<%=RS2("MODU_CD_MODULO")%>><%=RS2("MODU_TX_DESC_MODULO")%></option>
          <%else%>
              <option value=<%=RS2("MODU_CD_MODULO")%>><%=RS2("MODU_TX_DESC_MODULO")%></option>
          <%
			end if
			RS2.MOVENEXT
			LOOP
			%>
          </select>
      </td>
    </tr>
    <tr> 
      <td width="4" height="18"></td>
      <td width="134" height="18"></td>
      <td width="1050" height="18"></td>
    </tr>
    <tr> 
      <td width="4" height="25"></td>
      <td width="134" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Atividade</b></font></td>
      <td width="1050" height="25"><select name="selAtividade" size="1" onchange="javascript:Confirma2()">
          <option value="0" selected>Selecione uma Atividade</option>
          <%DO UNTIL RS.EOF=TRUE
          if trim(str_Atividade)=trim(RS("ATCA_CD_ATIVIDADE_CARGA")) then
          %>
              <option selected value=<%=RS("ATCA_CD_ATIVIDADE_CARGA")%>><%=RS("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%else%>
              <option value=<%=RS("ATCA_CD_ATIVIDADE_CARGA")%>><%=RS("ATCA_TX_DESC_ATIVIDADE")%></option>
			<%
			end if
			RS.MOVENEXT
			LOOP
			%>
          </select>
      </td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="4" height="21"></td>
      <td width="134" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Intervalo
        Transações</b></font></td>
      <td width="1050" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>de
        <input type="text" name="range1" size="11" value="<%=de%>">
         à 
         <input type="text" name="range2" size="10" value="<%=ate%>">
        </b></font> </td>
    </tr>
    <tr> 
      <td width="4" height="21">&nbsp;</td>
      <td width="134" height="21">&nbsp;</td>
      <td width="1050" height="21">&nbsp; </td>
    </tr>
    <tr> 
      <td width="4" height="226">&nbsp;</td>
      <td width="134" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Selecione
        Transações&nbsp;</b></font></td>
      <td width="1050" height="226"> 
        <table width="1050" border="0" align="left" cellpadding="0" cellspacing="0">
          <tr> 
            <td valign="middle" align="left"> 
              <div align="LEFT"> <b> 
                <select name="list1" size="8" multiple>
                 <%
                 Set RS1 = Conn_db.Execute(RS_TRANSACAO)
                 
                 DO UNTIL RS1.EOF=TRUE
                 JATEM=0
                 
                 ON ERROR RESUME NEXT
                 
                 RS_DESTINO.MOVEFIRST
                 
                 DO UNTIL RS_DESTINO.EOF=TRUE
                 		IF TRIM(RS1("TRAN_CD_TRANSACAO"))=TRIM(RS_DESTINO("TRAN_CD_TRANSACAO")) THEN
                 			JATEM=JATEM+1	
						END IF
						RS_DESTINO.MOVENEXT
                 LOOP
                 	IF JATEM=0 THEN
                 	%>
                  <option value="<%=RS1("TRAN_CD_TRANSACAO")%>" ><%=RS1("TRAN_CD_TRANSACAO")%>-<%=RS1("TRAN_TX_DESC_TRANSACAO")%></option>
                  <%
                 END IF
  					RS1.MoveNext
					LOOP
					%>
                </select>
                </b></div>
            </td>
            <td width="26" align="left" valign="middle"> 
              <table width="53%" cellpadding="0" cellspacing="0" border="0" align="LEFT">
                <tr> 
                  <td width="100%" valign="middle" align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2)"><img name="Image16" border="0" src="../imagens/continua_F01.gif" width="24" height="24"></a></td>
                </tr>
                <tr> 
                  <td height="25">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="25" valign="middle" align="center"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1)"><img name="img01511" border="0" src="../imagens/continua2_F01.gif" width="24" height="24"></a></td>
                </tr>
              </table>
            </td>
            <td valign="middle" align="left"> 
              <div align="LEFT"><font color="#000080"> 
                <select name="list2" size="8" multiple>
                  <%
                RS_DESTINO.MOVEFIRST
                
                DO UNTIL RS_DESTINO.EOF=TRUE
                SSQL1="SELECT TRAN_TX_DESC_TRANSACAO FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & rs_destino("TRAN_CD_TRANSACAO") & "'"
                SET RS_ATUAL=CONN_DB.EXECUTE(SSQL1)
                ATUAL=RS_ATUAL("TRAN_TX_DESC_TRANSACAO")
                %>
                  <option value=<%=RS_DESTINO("TRAN_CD_TRANSACAO")%>><%=RS_DESTINO("TRAN_CD_TRANSACAO")%>-<%=ATUAL%></option>
                  <%
                RS_DESTINO.MOVENEXT
                LOOP
                %>
                </select>
                </font></div>
            </td>
          </tr>
          <tr>
            <td colspan="3" width="1048">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3" width="1048"> 
              <div align="LEFT"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
                a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></div>
            </td>
          </tr>
          <tr> 
            <td width="570"> 
              <%'=str_SQL_Sub_Proc%>
            </td>
            <td width="26" align="left" valign="middle">&nbsp;</td>
            <td width="448"> 
              <input type="hidden" name="txtEmpSelecionada">
            </td>
          </tr>
        </table>
      </td>
    </tr>
    
  </table>
  </div>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
