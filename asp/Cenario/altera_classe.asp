 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

cenario=request("ID")

set origem=db.execute("SELECT * FROM " & Session("PREFIXO") & "CLASSE_CENARIO WHERE CLCE_CD_NR_CLASSE_CENARIO='" & CENARIO & "'")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
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
document.frm1.txtMega.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtMega.value = document.frm1.txtMega.value + "," + fbox.options[i].value;
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

<script language="javascript" src="../Planilhas/js/troca_lista.js"></script>

<script>
function limpa()
{
document.frm1.txtDescClasse.focus();
}

function Confirma()
{
if(document.frm1.txtDescClasse.value == '')
{
alert("É obrigatória a descrição da CLASSE!");
document.frm1.txtDescClasse.focus();
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

</script>

<body topmargin="0" leftmargin="0" onload="javascript:limpa()">
<form method="POST" action="valida_altera_classe.asp" name="frm1">
<input type="hidden" name="ID" size="46" value="<%=cenario%>"><input type="hidden" name="INC" size="11" value="<%=INCLU%>"><input type="hidden" name="txtTranSelecionada" size="20">
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
  <table border="0" width="742" height="94" cellpadding="0" cellspacing="0" align="center">
    <tr> 
      <td width="15" height="21">&nbsp;</td>
      <td width="173" height="21">&nbsp;</td>
      <td height="21" width="453">
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Alteração
        de Classe</font></p>
      </td>
      <td height="21" width="83">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15" height="21">&nbsp;</td>
      <td width="173" height="21"> 
      </td>
      <td width="453" height="21"></td>
      <td height="21" width="83">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15" height="21">&nbsp;</td>
      <td width="173" height="21"> 
        <font color="#330099" face="Verdana" size="2"><b>Descrição da Classe :</b></font>
      </td>
      <td height="21" width="453"><input type="text" name="txtDescClasse" size="63" value="<%=ORIGEM("CLCE_TX_DESC_CLASSE_CENARIO")%>"></td>
      <td height="21" width="83">&nbsp;</td>
    </tr>
    <tr bgcolor="#330099"> 
      <td width="15" height="4" bgcolor="#FFFFFF"></td>
      <td width="173" height="4" bgcolor="#FFFFFF"></td>
      <td height="4" width="453" bgcolor="#FFFFFF"><input type="hidden" name="txtMega" size="20"></td>
      <td height="4" width="83" bgcolor="#FFFFFF"></td>
    </tr>
  </table>
  <table border="0" width="779" height="6">
    <tr> 
      <td width="93" height="1"></td>
      <td width="163" height="1"> 
      </td>
      <td height="1"> 
      <font color="#330099" face="Verdana" size="2"><b>Mega-Processo</b></font> 
      </td>
    </tr>
  </table>
  <table border="0" width="102%">
  <tr>
    <td width="8%" rowspan="5"></td>
    <td width="40%" rowspan="5"> 
        <p align="right"> 
        <select size="10" name="list1" multiple>
       <%do until rs_mega.eof=true
       set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CLASSE_CENARIO_MEGA_PROCESSO WHERE CLCE_CD_NR_CLASSE_CENARIO='" & cenario & "' AND MEPR_CD_MEGA_PROCESSO=" & rs_mega("MEPR_CD_MEGA_PROCESSO"))
    	if atual.eof=true then
    	%>
          <option value="<%=rs_mega("MEPR_CD_MEGA_PROCESSO")%>"><%=rs_mega("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
       <%
       end if
    	rs_mega.movenext
    	loop
    	rs_mega.movefirst
    	%>
        </select>
        </p>
    </td>
    <td width="5%" align="center"></td>
    <td width="83%" rowspan="5"> 
        <select size="10" name="list2" multiple>
       <%do until rs_mega.eof=true
       set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CLASSE_CENARIO_MEGA_PROCESSO WHERE CLCE_CD_NR_CLASSE_CENARIO='" & cenario & "' AND MEPR_CD_MEGA_PROCESSO=" & rs_mega("MEPR_CD_MEGA_PROCESSO"))
    	if atual.eof=false then
    	%>
          <option value="<%=rs_mega("MEPR_CD_MEGA_PROCESSO")%>"><%=rs_mega("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
       <%
       end if
    	rs_mega.movenext
    	loop
    	%>

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
    <td width="5%" align="center"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24"></a></td>
  </tr>
  <tr>
    <td width="5%" align="center"></td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>

</body>

</html>
