<%@LANGUAGE="VBSCRIPT"%>
<%
server.scripttimeout=99999999
response.buffer=false

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

operacao = request("opti")

if request("str01")<>"" then
	orgao_1=request("str01")
	orgao_1_= left(formatnumber(ORGAO_1), len(formatnumber(orgao_1))-3)
else
	orgao_1=0
end if

if request("str02")<>"" then
	orgao_2=request("str02")
	orgao_2=right((left(orgao_2,5)),3)	

if(left(orgao_2,1))=0 then
	orgao_2=right(orgao_2,(len(orgao_2))-1)
end if

else
	orgao_2="000"
end if

if request("str03")<>"" then
	orgao_3=request("str03")
else
	orgao_3=0
end if

if request("str04")<>"" then
	orgao_4=request("str04")
else
	orgao_4=0
end if

SSQL1=""
SSQL1="SELECT AGLU_SG_AGLUTINADO, AGLU_CD_AGLUTINADO FROM dbo.ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO"

set str1=db.execute(ssql1)

SSQL2=""
SSQL2="SELECT dbo.ORGAO_MAIOR.ORLO_CD_ORG_LOT, "
SSQL2=SSQL2+"dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT, dbo.ORGAO_MAIOR.ORLO_NR_ORDEM, dbo.ORGAO_MAIOR.ORLO_NM_ORG_LOT, dbo.ORGAO_MAIOR.ORLO_CD_STATUS"
SSQL2=SSQL2+" FROM dbo.ORGAO_MAIOR "
SSQL2=SSQL2+" WHERE (dbo.ORGAO_MAIOR.ORLO_CD_STATUS = 'A') AND (dbo.ORGAO_MAIOR.AGLU_CD_AGLUTINADO = '" & orgao_1_ & "')"
SSQL2=SSQL2+" ORDER BY dbo.ORGAO_MAIOR.ORLO_SG_ORG_LOT"
	
set str2=db.execute(ssql2)

ssql3=""
ssql3="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql3=ssql3+" WHERE (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORME_CD_STATUS = 'A')"
ssql3=ssql3+" AND SUBSTRING(ORME_CD_ORG_MENOR,3,3)='" & right("000"& ORGAO_2,3) & "' AND SUBSTRING(ORME_CD_ORG_MENOR,11,5)='00000' AND SUBSTRING(ORME_CD_ORG_MENOR,8,3)<>'000'"

set str3=db.execute(ssql3)

ssql4=""
ssql4="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql4=ssql4+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql4=ssql4+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,10)='" & ORGAO_3 & "' AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)='00' AND SUBSTRING(ORME_CD_ORG_MENOR,13,3) <> '000'" 

set str4=db.execute(ssql4)

ssql5=""
ssql5="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql5=ssql5+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql5=ssql5+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,13)='" & ORGAO_4 & "'  AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)<>'00'" 

set str5=db.execute(ssql5)

set rs_onda = db.execute("SELECT * FROM ONDA ORDER BY ONDA_TX_DESC_ONDA")

set rsmotivo = db.execute("SELECT * FROM MOTIVO ORDER BY MOTI_TX_DESC_MOTIVO")

onda = request("onda")

select case onda
case 1
	set rs_funcao = db.execute ("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_NM_ANTECIPADA = 1 ORDER BY FUNE_TX_DESC_FUNCAO_NEGOCIO")
case 2
	set rs_funcao = db.execute ("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_NM_ANTECIPADA = 1 ORDER BY FUNE_TX_DESC_FUNCAO_NEGOCIO")
case 5
	set rs_funcao = db.execute ("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_NM_ANTECIPADA = 1 ORDER BY FUNE_TX_DESC_FUNCAO_NEGOCIO")
case 7
	set rs_funcao = db.execute ("SELECT * FROM FUNCAO_NEGOCIO WHERE FUNE_NM_ANTECIPADA = 1 ORDER BY FUNE_TX_DESC_FUNCAO_NEGOCIO")
case else
	set rs_funcao = db.execute ("SELECT * FROM FUNCAO_NEGOCIO ORDER BY FUNE_TX_DESC_FUNCAO_NEGOCIO")
end select

set rs_usuario = db.execute("SELECT * FROM USUARIO_MAPEAMENTO ORDER BY USMA_TX_NOME_USUARIO")
%>
<html>
<head>

<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function carrega_txt1(fbox) {
document.frm1.txtfuncao.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtfuncao.value = document.frm1.txtfuncao.value + "'" + fbox.options[i].value + "',";
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

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
alert(message);
return false;
}
}
if (document.layers) {
if (e.which == 3) {
alert(message);
return false;
}
}
}
if (document.layers) {
document.captureEvents(Event.MOUSEDOWN);
}
document.onmousedown=click;

</script>

<script>
function manda01()
{
window.location="seleciona_lm.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&Onda="+document.frm1.selOnda.value
}

function manda02()
{
window.location="seleciona_lm.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&Onda="+document.frm1.selOnda.value
}

function manda03()
{
window.location="seleciona_lm.asp?str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&Onda="+document.frm1.selOnda.value
}

function Confirma() 
{
if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0)&&(document.frm1.list2.options.length == 0)&&(document.frm1.txtChave.value == ""))
{
alert("Você deve Selecionar pelo menos um dos parâmetros!");
document.frm1.Str01.focus();
return;
}
if((document.frm1.selMotivo.value==0))
{
alert("Você deve selecionar um MOTIVO!");
document.frm1.selMotivo.focus();
return;
}
else
{
carrega_txt1(document.frm1.list2);
document.frm1.submit();
}
}

function apaga_item()
{
var a=event.keyCode;
if (a==46)
{
var f=document.frm1.list1.selectedIndex;
if (f!=-1){
	document.frm1.list1.options[f] = null;
	document.frm1.list1.selectedIndex=f-1;
}
}
}

function apaga_item2()
{
var f = document.frm1.list2.options.length;
var items = '';
for(var i = 0; i < f; i++)
{
if (document.frm1.list2.options[i].selected)
{
	items = items + ';' + i
}
}
items=items + ';';
var t = document.frm1.list2.options.length;
var f = -1;
for(var d = 0; d < t + 1; d++)
{
var s = ';'+d+';';
if(items.search(s)!=-1)
{
if(f==-1)
{
document.frm1.list2.options[d] = null;
f=d;
}
else
{
document.frm1.list2.options[f] = null;
}
}
}
}


function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
alert("Tecla não permitida!");
return;
}
}

</SCRIPT>


<script language="javascript" src="troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF">
<form name="frm1" method="POST" action="verifica_lm.asp">

  <table width="972" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="border-collapse: collapse" bordercolor="#111111">
    <tr> 
      <td width="158" height="20" colspan="2">&nbsp;</td>
      <td width="493" height="60" colspan="3">&nbsp;</td>
      <td width="321" valign="top" colspan="2"> <table width="195" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC" style="border-collapse: collapse" bordercolor="#111111">
          <tr>
            <td bgcolor="#330099" width="96" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="66" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="82" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr>
            <td bgcolor="#330099" height="12" width="96" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="66" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="82" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="38">&nbsp; </td>
      <td height="20" width="120"> <p align="right"> <img border="0" src="../../imagens/confirma_f02.gif" onclick="Confirma()"> 
      </td>
      <td height="20" width="181"> <font size="2" face="Verdana" color="#000080"><b>&nbsp;Enviar</b></font> 
      </td>
      <td height="20" width="44">&nbsp;</td>
      <td height="20" width="268">&nbsp;</td>
      <td height="20" width="84">&nbsp; </td>
      <td height="20" width="237">&nbsp; </td>
    </tr>
  </table>
  <table border="0" width="96%" height="244">
    <tr>
      <td width="1%" height="33">&nbsp;</td>
      <td width="4%" height="33">&nbsp;</td>
      <td width="65%" height="33" colspan="3"><p align="center"><font face="Verdana" color="#000080">Treinamento - Liberação Manual de Usuários em Cursos (<b>LM</b>)</font></td>
    </tr>
    <tr>
      <td width="1%" height="23">&nbsp;</td>
      <td width="4%" height="23">&nbsp;</td>
      <td width="65%" height="23" colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="1%" height="1"></td>
      <td width="4%" height="1"></td>
      <td width="65%" height="1" colspan="3"></td>
    </tr>
    <tr>
      <td width="1%" height="18"></td>
      <td width="4%" height="18"><font color="#000080" face="System" size="2"><b>Órgão</b></font></td>
      <td width="65%" height="18" colspan="3"><select size="1" name="Str01" onChange="manda01()" style="font-family: Verdana; font-size: 8 pt">
        <OPTION VALUE="0">== Selecione Órgão Aglutinador ==</OPTION>
        <%do until str1.eof=true
        if trim(orgao_1)=trim(str1("AGLU_CD_AGLUTINADO")) then
        %>
        <option selected value="<%=str1("AGLU_CD_AGLUTINADO")%>"><%=str1("AGLU_SG_AGLUTINADO")%></option>
        <%else%>
        <option value="<%=str1("AGLU_CD_AGLUTINADO")%>"><%=str1("AGLU_SG_AGLUTINADO")%></option>
        <%
        end if
        str1.movenext
        looP
        %>
        </select></td>
    </tr>
    <tr>
      <td width="1%" height="24"></td>
      <td width="4%" height="24"><b><font face="System" size="2" color="#000080">Unidade</font></b></td>
      <td width="65%" height="24" colspan="3"><select size="1" name="Str02" onChange="manda02()" style="font-family: Verdana; font-size: 8 pt">
      <OPTION VALUE="000">== Selecione Órgão de Lotação ==</OPTION>
          <%do until str2.eof=true
        if trim(orgao_2)=trim(str2("ORLO_CD_ORG_LOT"))then
        %>
        <option selected value="<%=ORGAO_1 & right(("000" & str2("ORLO_CD_ORG_LOT")),3) & right(("000" & str2("ORLO_NR_ORDEM")),2)%>"><%=str2("ORLO_SG_ORG_LOT")%></option>
        <%else%>
        <option value="<%=ORGAO_1 & right(("000" & str2("ORLO_CD_ORG_LOT")),3) & right(("000" & str2("ORLO_NR_ORDEM")),2)%>"><%=str2("ORLO_SG_ORG_LOT")%></option>
        <%
        end if
        str2.movenext
        looP
        %>
        </select></td>
    </tr>
    <tr>
      <td width="1%" height="13"></td>
      <td width="4%" height="13"><font color="#000080" face="System" size="2"><b> Gerência</b></font></td>
      <td width="65%" height="13" colspan="3"><select size="1" name="Str03" style="font-family: Verdana; font-size: 8 pt">
      <OPTION VALUE="0">== Selecione Gerência ==</OPTION>
        <%
        do until str3.eof=true
        if trim(orgao_3)=trim(left((str3("ORME_CD_ORG_MENOR")),10)) then
        %>
        <option selected value="<%=trim(left((str3("ORME_CD_ORG_MENOR")),10))%>"><%=str3("ORME_SG_ORG_MENOR")%></option>
        <%else%>
        <option value="<%=trim(left((str3("ORME_CD_ORG_MENOR")),10))%>"><%=str3("ORME_SG_ORG_MENOR")%></option>
        <%

        end if
        str3.movenext
        looP
        %>
        </select></td>
    </tr>
    <tr>
      <td width="1%" height="23">&nbsp;</td>
      <td width="4%" height="23" valign="top">&nbsp;</td>
      <td width="32%" height="23" valign="bottom">&nbsp;</td>
      <td width="2%" height="13" valign="bottom"></td>
      <td width="33%" height="23" valign="bottom">&nbsp;</td>
    </tr>
    <tr>
      <td width="1%" height="23">&nbsp;</td>
      <td width="4%" height="23" valign="top"><b><font face="System" size="2" color="#000080">Onda</font></b></td>
      <td width="32%" height="23" valign="bottom"><select size="1" name="selOnda" style="font-family: Verdana; font-size: 8 pt" onChange="manda03()">
      <option value="0">== Selecione a Onda ==</option>
      <%
      do until rs_onda.eof=true
      if trim(rs_onda("ONDA_CD_ONDA")) = trim(request("onda")) then
      	chec = "selected"
      else
      	chec = ""
      end if
      %>
      <option <%=chec%> value="<%=rs_onda("ONDA_CD_ONDA")%>"><%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
      <%
      rs_onda.movenext
      loop
      %>
      </select></td>
      <td width="2%" height="13" valign="bottom"></td>
      <td width="33%" height="23" valign="bottom">&nbsp;</td>
    </tr>
    <tr>
      <td width="1%" height="23"></td>
      <td width="4%" height="23" valign="top"></td>
      <td width="32%" height="23" valign="bottom"><font color="#000080" face="Verdana" size="1">Funções Disponíveis</font></td>
      <td width="2%" height="13" valign="bottom"></td>
      <td width="33%" height="23" valign="bottom"><font color="#000080" face="Verdana" size="1">Funções Selecionadas</font></td>
    </tr>
    <tr>
      <td width="1%" height="27" rowspan="3">&nbsp;</td>
      <td width="4%" height="27" valign="top" rowspan="3"><b><font face="System" size="2" color="#000080">Função</font></b></td>
      <td width="32%" height="27" rowspan="3"><select size="5" name="selFuncao" style="font-family: Verdana; font-size: 7 pt" multiple>
      <%
      i=0
      reg = rs_funcao.RecordCount
      %>
      <option value="0">== Todas as Funções ==</option>
      <%
      do until i = reg
      %>
      <option value="<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=LEFT(rs_funcao("FUNE_TX_DESC_FUNCAO_NEGOCIO"),90)%></option>
      <%
      i = i + 1
      rs_funcao.movenext
      loop
      %>
      </select></td>
      <td width="2%" height="9"><img border="0" src="../../imagens/continua_F01.gif" onClick="move(document.frm1.selFuncao,document.frm1.list2,1)"></td>
      <td width="33%" height="27" rowspan="3"><select size="5" name="list2" style="font-family: Verdana; font-size: 7 pt" multiple></select></td>
    </tr>
    <tr>
      <td width="2%" height="9"></td>
    </tr>
    <tr>
      <td width="2%" height="9"><img border="0" src="../../imagens/continua2_F01.gif" onClick="apaga_item2()"></td>
    </tr>
    <tr>
      <td width="1%" height="16"></td>
      <td width="4%" height="16"></td>
      <td width="65%" height="16" colspan="3">
      </td>
    </tr>
    <tr>
      <td width="1%" height="20">&nbsp;</td>
      <td width="4%" height="20"><b><font face="System" size="2" color="#000080">Chave</font></b></td>
      <td width="65%" height="20" colspan="3">
      <input type="text" name="txtChave" size="14" style="font-family: Verdana; font-size: 8 pt" maxlength="4"></td>
    </tr>
    <tr>
      <td width="1%" height="22">&nbsp;</td>
      <td width="4%" height="22">&nbsp;</td>
      <td width="65%" height="22" colspan="3">
      &nbsp;</td>
    </tr>
    <tr>
      <td width="1%" height="22">&nbsp;</td>
      <td width="4%" height="22"><b><font size="2" color="#000080"><font face="System">Motivo</font></font></b></td>
      <td width="65%" height="22" colspan="3">
      <select size="1" name="selMotivo" style="font-family: Verdana; font-size: 8 pt">
      <option value="0">== Selecione o Motivo ==</option>
      	<%
    	do until rsmotivo.eof=true	  	
  	   	%>
    	  <option value="<%=RSMOTIVO("MOTI_NR_CD_MOTIVO")%>"><%=RSMOTIVO("MOTI_TX_DESC_MOTIVO")%></option>
    	  <%
    	  rsmotivo.movenext
    	  loop
    	  %>
      </select>
      <input type="hidden" name="txtfuncao" size="20">
      </td>
    </tr>
  </table>
  <p align="left">
</form>
<%
db.close
set db = nothing
%>
</body>
</html>