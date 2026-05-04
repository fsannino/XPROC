 
<!--#include file="../../asp/protege/protege.asp" -->
<%
cenario=request("ID")
opt=request("OPTION")
inclu=request("INC")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& trim(cenario) & "'"
set rs_cenario=db.execute(ssql)

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& rs_cenario("MEPR_CD_MEGA_PROCESSO"))
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& rs_cenario("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO="& rs_cenario("PROC_CD_PROCESSO"))
set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO="& rs_cenario("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO="& rs_cenario("PROC_CD_PROCESSO") & " AND SUPR_CD_SUB_PROCESSO="& rs_cenario("SUPR_CD_SUB_PROCESSO"))

str_mega=0
str_sub=0

if opt=1 then
	str_mega=rs_cenario("MEPR_CD_MEGA_PROCESSO")
else
	str_mega=request("selMegaProcesso")
	str_sub=request("selSubProcesso")
end if

if session("MegaProcesso")<>0 and str_mega=0 then
	str_mega=session("MegaProcesso")
end if

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

COMPL1=""

if str_sub<>0 then
	COMPL1=" AND SUPR_CD_SUB_PROCESSO=" & str_sub
end if

if str_mega<>0 then
	ssql2="SELECT DISTINCT SUPR_TX_DESC_SUB_PROCESSO, SUPR_CD_SUB_PROCESSO FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega &" ORDER BY SUPR_TX_DESC_SUB_PROCESSO"
	RESPONSE.WRITE SSQL2
	set rs_sub = db.execute(ssql2)
else
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO==0")
end if

if str_mega<>0 then
	SSQL5="SELECT DISTINCT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & + COMPL1 +" ORDER BY TRAN_CD_TRANSACAO"
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

<script language="javascript" src="../Planilhas/js/troca_lista_sem_ordem2.js"></script>
<script language="javascript" src="../Planilhas/js/troca_lista_sem_ordem3.js"></script>

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
window.location.href.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value
}

function manda2()
{
window.location.href.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value
}

function manda3()
{
window.location.href.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubProcesso='+document.frm1.selSubProcesso.value
}

function manda4()
{
window.location.href.href='cad_cenario_transacao.asp?option=2&INC='+document.frm1.INC.value+'&ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value+'&selSubProcesso='+document.frm1.selSubProcesso.value+'&selAtividade='+document.frm1.selAtividade.value
}

function exibe_transacao()
{
window.open("exibe_trans.asp?ID="+document.frm1.ID.value,"_blank","width=530,height=240,history=0,scrollbars=0,titlebar=0,resizable=0")
}

</script>

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
          <td width="26"><img border="0" src="../../imagens/confirma_f02.gif" onclick="javascript:Confirma()"></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table border="0" width="687" height="94" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td width="15" height="21">&nbsp;</td>
      <td width="147" height="21">&nbsp;</td>
      <td height="21" width="415"><font face="Verdana" color="#330099" size="3">Relação 
        Cenário x Transação</font></td>
      <td height="21" width="92">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15" height="21">&nbsp;</td>
      <td width="147" height="21"> 
        <div align="right"><font face="Verdana" size="1" color="#330099">Mega-Processo 
          :&nbsp;</font></div>
      </td>
      <td width="415" height="21"><font face="Verdana" size="1" color="#330099"><%=RS1("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
      <td height="21" width="92">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15" height="21">&nbsp;</td>
      <td width="147" height="21"> 
        <div align="right"><font face="Verdana" size="1" color="#330099">Processo 
          :&nbsp;</font></div>
      </td>
      <td height="21" width="415"><font face="Verdana" size="1" color="#330099"><%=RS2("PROC_TX_DESC_PROCESSO")%></font></td>
      <td height="21" width="92">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15" height="21">&nbsp;</td>
      <td width="147" height="21"> 
        <div align="right"><font face="Verdana" size="1" color="#330099">Sub-Processo 
          :&nbsp;</font></div>
      </td>
      <td height="21" width="415"><font face="Verdana" size="1" color="#330099"><%=RS3("SUPR_TX_DESC_SUB_PROCESSO")%></font></td>
      <td height="21" width="92"><b><font face="Verdana" size="2" color="#330099">Transações 
        existente</font></b></td>
    </tr>
    <tr> 
      <td width="15" height="20">&nbsp;</td>
      <td width="147" height="20"> 
        <div align="right"><font face="Verdana" size="2" color="#330099">Cenário 
          : </font></div>
      </td>
      <td height="20" width="415"><font face="Verdana" size="2" color="#330099"><%=CENARIO%> </font></td>
      <td height="20" width="92">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15" height="20">&nbsp;</td>
      <td width="147" height="20"> 
        <div align="right"><b><font face="Verdana" size="2" color="#330099">&nbsp;</font></b></div>
      </td>
      <td height="20" width="415"><font face="Verdana" size="2" color="#330099"><%=rs_cenario("CENA_TX_TITULO_CENARIO")%></font></td>
      <td height="20" width="92"> 
        <div align="center"><b><font face="Verdana" size="2" color="#330099"><a href="#" onClick="javascript:exibe_transacao()" );"><img border="0" src="../../imagens/icon_empresa.gif" align="absmiddle"></a></font></b> 
        </div>
      </td>
    </tr>
    <tr> 
      <td width="15" height="4" bgcolor="#330099"></td>
      <td width="147" height="4" bgcolor="#330099"></td>
      <td height="4" width="415" bgcolor="#330099"></td>
      <td height="4" width="92" bgcolor="#330099"></td>
    </tr>
  </table>
  <table border="0" width="779" height="6">
    <tr> 
      <td width="93" height="1"></td>
      <td width="163" height="1"> 
      </td>
      <td height="1" colspan="4"> 
      </td>
    </tr>
    <tr> 
      <td width="93" height="1"></td>
      <td width="163" height="1"> 
      </td>
      <td height="1" colspan="4"> 
      </td>
    </tr>
    <tr> 
      <td width="93" height="4"></td>
      <td width="163" height="4"> 
        <p align="right"><b><font face="Verdana" size="2" color="#330099"> Mega-Processo
          :</font></b>
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
      <td width="93" height="4"></td>
      <td width="163" height="4"> 
        <p align="right"><b><font face="Verdana" size="2" color="#330099">Sub-Processo
        :&nbsp;</font></b>
      </td>
      <td height="4" colspan="4"> 
        <select size="1" name="selSubProcesso" onchange="javascript:manda3()">
          <option value="0">== Selecione o Sub-Processo ==</option>
           <%do until rs_sub.eof=true
       if trim(str_sub)=trim(rs_sub("SUPR_CD_SUB_PROCESSO")) then
       %>
          <option selected value=<%=RS_sub("SUPR_CD_SUB_PROCESSO")%>><%=RS_SUB("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%ELSE%>
          <option value=<%=RS_sub("SUPR_TX_DESC_SUB_PROCESSO")%>><%=RS_SUB("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%
		end if
		rs_sub.movenext
		loop
		%>
        </select>
        (<font color="#330099" size="1" face="Verdana"><b>Opcional)</b></font>
      </td>
    </tr>
    <tr> 
      <td width="93" height="4">&nbsp;</td>
      <td width="163" height="4"> 
        <div align="right"></div>
      </td>
      <td height="4" colspan="4"> 
      </td>
    </tr>
    <tr> 
      <td width="93" height="7"></td>
      <td width="163" height="7"> 
        <div align="right">
          <input type="hidden" name="INC" size="11" value="<%=INCLU%>"><b><font face="Verdana" size="2" color="#330099">Transações : </font></b> 
        </div>
      </td>
      <td height="7">&nbsp; </td>
      <td width="22" height="7"> 
        <input type="hidden" name="ID" size="46" value="<%=cenario%>">
      </td>
      <td width="68" height="7"></td>
      <td width="122" height="7"></td>
    </tr>
  </table>
  <table border="0" width="102%">
  <tr>
    <td width="8%" rowspan="5"></td>
    <td width="41%" rowspan="5"> 
        <select size="6" name="list1" multiple>
          <%do until rs_trans.eof=true
    	set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & rs_trans("TRAN_CD_TRANSACAO") & "'")
    	VALOR = LEFT(atual("TRAN_TX_DESC_TRANSACAO"),35)
    	%>
          <option value="<%=rs_trans("TRAN_CD_TRANSACAO")%>"><%=rs_trans("TRAN_CD_TRANSACAO")%>-<%=VALOR%></option>
          <%
    	rs_trans.movenext
    	loop
    	%>
        </select>
    </td>
    <td width="5%" align="center"></td>
    <td width="82%" rowspan="5"> 
        <select size="6" name="list2" multiple>
        </select>
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
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>

</body>

</html>
