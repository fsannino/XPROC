 
<%
cenario=request("ID")
opt=request("OPTION")
inclu=request("INC")

server.scripttimeout=999999999

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql3="SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='" & cenario & "'"
set rsexiste=db.execute(ssql3)
ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& trim(cenario) & "'"
set rs_cenario=db.execute(ssql)

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& rs_cenario("MEPR_CD_MEGA_PROCESSO"))
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& rs_cenario("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO="& rs_cenario("PROC_CD_PROCESSO"))
set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& rs_cenario("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO="& rs_cenario("PROC_CD_PROCESSO") & " AND SUPR_CD_SUB_PROCESSO="& rs_cenario("SUPR_CD_SUB_PROCESSO"))

str_mega=0
str_sub=0
str_proc=0

if opt=1 then
	str_mega=rs_cenario("MEPR_CD_MEGA_PROCESSO")
else
	str_mega=request("selMegaProcesso")
	str_proc=request("selProcesso")
	str_sub=request("selSubProcesso")
end if

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO<>'"& trim(cenario) & "' AND MEPR_CD_MEGA_PROCESSO="& str_mega & " ORDER BY CENA_CD_CENARIO"
set rs_cenario2=db.execute(ssql)

'response.write str_sub

if session("MegaProcesso")<>0 and str_mega=0 then
	str_mega=session("MegaProcesso")
end if

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

COMPL1=""

if str_mega<>0 then
	compl1="MEPR_CD_MEGA_PROCESSO=" & str_mega
end if

if str_proc<>0 then
if len(compl1)=0 then
	compl1="PROC_CD_PROCESSO=" & str_proc
else
	compl1=compl1 & " AND PROC_CD_PROCESSO=" & str_proc
end if
end if

if str_sub <> 0 then
if len(compl1)=0 then
	compl1= " SUPR_CD_SUB_PROCESSO=" & str_sub
else
	compl1 = compl1 & " AND SUPR_CD_SUB_PROCESSO=" & str_sub
end if
end if

if len(compl1)>0 then
	compl1="WHERE " + COMPL1
end if

if str_mega<>0 and str_proc<>0 then
	ssql2="SELECT DISTINCT SUPR_CD_SUB_PROCESSO, SUPR_TX_DESC_SUB_PROCESSO, PROC_CD_PROCESSO, MEPR_CD_MEGA_PROCESSO FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega &" AND PROC_CD_PROCESSO=" & str_proc & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO"
	set rs_sub = db.execute(ssql2)
else
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega<>0 then
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
else
	set rs_PROC=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if


if str_mega<>0 then
	SSQL5="SELECT DISTINCT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "RELACAO_FINAL " & compl1 & " ORDER BY TRAN_CD_TRANSACAO"
	set rs_trans=db.execute(SSQL5)
else
	set rs_trans=db.execute("SELECT DISTINCT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY TRAN_CD_TRANSACAO")
end if

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">
<!--

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function carrega_txt(fbox) {
document.frm1.txtTranSelecionada.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtTranSelecionada.value = document.frm1.txtTranSelecionada.value + "," + fbox.options[i].value;
   }
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

<script language="javascript" src="../js/troca_lista_sem_ordem2.js"></script>
<script language="javascript" src="../js/troca_lista_sem_ordem3.js"></script>

<%if rsexiste.eof=true then%>

<script>
function Confirma()
{
if(document.frm1.selMegaProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MEGA-PROCESSO!");
document.frm1.selMegaProcesso.focus();
return;
}
if ((document.frm1.list2.options.length == 0) && (document.frm1.selDuplicaCenario.selectedIndex == 0))
{ 
alert("A seleção de pelo menos uma TRANSAÇÃO é obrigatória !");
document.frm1.list1.focus();
return;
}
else
{
carrega_txt(document.frm1.list2)
document.frm1.submit();
}
}

function redefine()
{
window.location.href='cad_cenario.asp'
}

function manda1()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value
}

function manda2()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value
}

function manda3()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value+'&selSubProcesso='+document.frm1.selSubProcesso.value
}

function manda4()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value+'&selSubProcesso='+document.frm1.selSubProcesso.value+'&selAtividade='+document.frm1.selAtividade.value
}

function exibe_transacao()
{
window.open("exibe_trans.asp?ID="+document.frm1.ID.value,"_blank","width=530,height=240,history=0,scrollbars=0,titlebar=0,resizable=0")
}

</script>

<%else%>

<script>
function Confirma()
{
if(document.frm1.selMegaProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MEGA-PROCESSO!");
document.frm1.selMegaProcesso.focus();
return;
}
if (document.frm1.list2.options.length == 0)
{ 
alert("A seleção de pelo menos uma TRANSAÇÃO é obrigatória !");
document.frm1.list1.focus();
return;
}
else
{
carrega_txt(document.frm1.list2)
document.frm1.submit();
}
}

function redefine()
{
window.location.href='cad_cenario.asp'
}

function manda1()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value
}

function manda2()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value
}

function manda3()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value+'&selSubProcesso='+document.frm1.selSubProcesso.value
}

function manda4()
{
window.location.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value+'&selSubProcesso='+document.frm1.selSubProcesso.value+'&selAtividade='+document.frm1.selAtividade.value
}

function exibe_transacao()
{
window.open("exibe_trans.asp?ID="+document.frm1.ID.value,"_blank","width=530,height=240,history=0,scrollbars=0,titlebar=0,resizable=0")
}

</script>

<%end if%>
<body topmargin="0" leftmargin="0">
<form method="POST" action="valida_cad_cenario_transacao.asp" name="frm1">
        <input type="hidden" name="txtTranSelecionada" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
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
          <td width="21"></td>
          <td width="217"></td>
            <td width="18"></td>  <td width="16"></td>
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
      <td width="135" height="21">&nbsp;</td>
      <td height="21" width="524"><font face="Verdana" color="#330099" size="3">Relação 
        Cenário x Transação</font></td>
      <td height="21" width="67"> 
        <div align="center"><font face="Verdana" size="1" color="#330099">Transações 
          inseridas</font></div>
      </td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="135" height="14"> 
        <div align="right"><font size="1"><font face="Verdana" color="#330099">Mega-Processo 
          :&nbsp;</font></font></div>
      </td>
      <td width="524" height="14"><b><font face="Verdana" size="1" color="#330099"><%=RS1("MEPR_TX_DESC_MEGA_PROCESSO")%></font></b></td>
      <td height="14" width="67"> 
        <div align="center"><b><font face="Verdana" size="2" color="#330099"><a href="#" onClick="javascript:exibe_transacao()" );"><img border="0" src="../../imagens/icon_empresa.gif" align="absmiddle"></a></font></b></div>
      </td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="135" height="14"> 
        <div align="right"><font size="1"><font face="Verdana" color="#330099">Processo 
          :&nbsp;</font></font></div>
      </td>
      <td height="14" width="524"><b><font face="Verdana" size="1" color="#330099"><%=RS2("PROC_TX_DESC_PROCESSO")%></font></b></td>
      <td height="14" width="67">&nbsp;</td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="135" height="14"> 
        <div align="right"><font size="1"><font face="Verdana" color="#330099">Sub-Processo 
          :&nbsp;</font></font></div>
      </td>
      <td height="14" width="524"><b><font face="Verdana" size="1" color="#330099"><%=RS3("SUPR_TX_DESC_SUB_PROCESSO")%></font></b></td>
      <td height="14" width="67"> 
        <div align="center"></div>
      </td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="135" height="14"> 
        <div align="right"><font face="Verdana" size="2" color="#330099">Cenário 
          :&nbsp;</font></div>
      </td>
      <td height="14" width="524"><b><font face="Verdana" size="2" color="#330099"><%=CENARIO%></font></b></td>
      <td height="14" width="67"> 
        <div align="center"><b><font face="Verdana" size="2" color="#330099"></font></b> 
        </div>
      </td>
    </tr>
    <tr> 
      <td width="10" height="14">&nbsp;</td>
      <td width="135" height="14">&nbsp;</td>
      <td height="14" width="524"><b><font face="Verdana" size="2" color="#330099"><%=rs_cenario("CENA_TX_TITULO_CENARIO")%> </font></b></td>
      <td height="14" width="67">&nbsp;</td>
    </tr>
    <tr bgcolor="#330099"> 
      <td width="10" height="4"></td>
      <td width="135" height="4"></td>
      <td height="4" width="524"></td>
      <td height="4" width="67"></td>
    </tr>
  </table>
  <table border="0" width="732" height="1" align="center">
    <tr> 
      <td width="226" height="1"> </td>
      <td height="1" colspan="4"> </td>
    </tr>
    <%IF rsexiste.EOF=TRUE THEN%>
    <tr> 
      <td height="7" colspan="5" align="center"> <font face="Verdana" size="2" color="#330099">Copiar 
        Transações do Cenário</font> </td>
    </tr>
    <tr> 
      <td height="7" colspan="5" align="center"> 
        <select size="1" name="selDuplicaCenario">
          <option value="0">== Selecione o Cenário ==</option>
          <%
        do until rs_cenario2.eof=true
        set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO_TRANSACAO WHERE CENA_CD_CENARIO='"& trim(rs_cenario2("CENA_CD_CENARIO"))& "'")
        if atual.eof=false then
        %>
          <option value="<%=rs_cenario2("CENA_CD_CENARIO")%>"><%=rs_cenario2("CENA_CD_CENARIO")%>-<%=LEFT(rs_cenario2("CENA_TX_TITULO_CENARIO"),75)%></option>
          <%
        end if
        rs_cenario2.movenext
        loop
        %>
        </select>
      </td>
    </tr>
    <tr> 
      <td height="4" bgcolor="#330099" colspan="5"></td>
      <%end if%>
    </tr>
    <tr> 
      <td width="226" height="4"> 
        <p align="right"><font face="Verdana" size="2" color="#330099"> Mega-Processo 
          :</font> 
      </td>
      <td height="4" colspan="4"> 
        <select size="1" name="selMegaProcesso" onchange="javascript:manda1()">
          <option value="0">== Selecione o Mega-Processo ==</option>
          <%do until rs_mega.eof=true
          if trim(str_mega)=trim(rs_mega("MEPR_CD_MEGA_PROCESSO")) then
          %>
          <option selected value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%ELSE%>
          <option value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
		end if
		rs_mega.movenext
		loop
		%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="226" height="4"> 
        <p align="right"><font face="Verdana" size="2" color="#330099">Processo 
          : </font><font face="Verdana" color="#330099" size="1">(Opcional )</font> 
      </td>
      <td height="4" colspan="4"> 
        <select size="1" name="selProcesso" onchange="javascript:manda2()">
          <option value="0">== Selecione o Processo ==</option>
          <%do until rs_proc.eof=true
          if trim(str_proc)=trim(rs_proc("PROC_CD_PROCESSO")) then
          %>
          <option selected value=<%=RS_proc("PROC_CD_PROCESSO")%>><%=RS_proc("proc_TX_DESC_PROCESSO")%></option>
          <%ELSE%>
          <option value=<%=RS_proc("PROC_CD_PROCESSO")%>><%=RS_proc("proc_TX_DESC_PROCESSO")%></option>
          <%
		end if
		rs_proc.movenext
		loop
		%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="226" height="4"> 
        <p align="right"><font face="Verdana" size="2" color="#330099">Sub-Processo 
          : (<font color="#330099" size="1" face="Verdana">Opcional)</font>&nbsp;</font> 
      </td>
      <td height="4" colspan="4"> 
        <select size="1" name="selSubProcesso" onchange="javascript:manda3()">
          <option value="0">== Selecione o Sub-Processo ==</option>
          <%do until rs_sub.eof=true
       	if trim(str_sub)=trim(rs_sub("SUPR_CD_SUB_PROCESSO")) then
          %>
          <option selected value=<%=RS_sub("SUPR_CD_SUB_PROCESSO")%>><%=RS_SUB("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%ELSE%>
          <option value=<%=RS_sub("SUPR_CD_SUB_PROCESSO")%>><%=RS_SUB("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%
			end if
			rs_sub.movenext
			loop
		%>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="226" height="7"> 
        <div align="right"><b><font face="Verdana" size="2" color="#330099">Transações</font></b> 
          <input type="hidden" name="INC" size="11" value="<%=INCLU%>">
        </div>
      </td>
      <td height="7" width="20">&nbsp; </td>
      <td width="69" height="7"> 
        <input type="hidden" name="ID" size="46" value="<%=cenario%>">
      </td>
      <td width="154" height="7">&nbsp; </td>
      <td width="241" height="7"> </td>
    </tr>
  </table>
  <table border="0" width="102%">
  <tr>
    <td width="8%" rowspan="5">
      <p style="margin-top: 0; margin-bottom: 0"></td>
    <td width="41%" rowspan="5"> 
        <p style="margin-top: 0; margin-bottom: 0"> 
        <select size="10" name="list1" multiple>
          <%do until rs_trans.eof=true
    	set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & rs_trans("TRAN_CD_TRANSACAO") & "'")
		on error resume next
		VALOR = LEFT(atual("TRAN_TX_DESC_TRANSACAO"),35)
		if err.number<>0 then
			VALOR = atual("TRAN_TX_DESC_TRANSACAO")
		end if
		err.clear
    	%>
          <option value="<%=rs_trans("TRAN_CD_TRANSACAO")%>"><%=rs_trans("TRAN_CD_TRANSACAO")%>-<%=VALOR%></option>
          <%
    	rs_trans.movenext
    	loop
    	%>
        </select>
        </p>
    </td>
    <td width="5%" align="center">
      <p style="margin-top: 0; margin-bottom: 0"></td>
    <td width="82%" rowspan="5"> 
        <p style="margin-top: 0; margin-bottom: 0"> 
        <select size="10" name="list2" multiple>
        </select>
        </p>
    </td>
  </tr>
  <tr>
    <td width="5%" align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2,0)"><img name="Image16" border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a> 
    </td>
  </tr>
  <tr>
    <td width="5%" align="center"></td>
  </tr>
  <tr>
    <td width="5%" align="center"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="alterar(document.frm1.list2,document.frm1.list1,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24"></a></td>
  </tr>
  <tr>
    <td width="5%" align="center"></td>
  </tr>
</table>
  </form>

</body>

</html>
