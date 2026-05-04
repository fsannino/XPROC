<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Apoio/conn_consulta.asp" -->
<%
server.ScriptTimeout=99999999

operacao = request("opti")

chave=request("chave")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set Rusuario=db.execute("SELECT * FROM " & Session("Prefixo") & "USUARIO_MAPEAMENTO WHERE USMA_CD_USUARIO='" & chave & "'")

usuario=Rusuario("USMA_TX_NOME_USUARIO")

sel1=0
sel2=0
sel3=0
sel4=0
sel5=0

if request("str01")<>"" then
	orgao_1 = request("str01")
	orgao_1_= left(formatnumber(ORGAO_1), len(formatnumber(orgao_1))-3)
	sel1=1
else
	orgao_1=0
end if

if orgao_1=0 then
	sel1=0
end if

if request("str02")<>"" then
	orgao_2=request("str02")
	orgao_2=right((left(orgao_2,5)),3)
	sel2=1
if(left(orgao_2,1))=0 then
	orgao_2=right(orgao_2,(len(orgao_2))-1)
end if
else
	orgao_2="000"
end if

if request("str03")<>"" then
	orgao_3=request("str03")
	sel3=1
else
	orgao_3=0
end if

if request("str04")<>"" then
	orgao_4=request("str04")
	sel4=1	
else
	orgao_4=0
end if

'response.write sel1 & "<br>"
'response.write sel2 & "<br>"
'response.write sel3 & "<br>"
'response.write sel4 & "<br>"

SSQL1=""
SSQL1="SELECT AGLU_SG_AGLUTINADO, AGLU_CD_AGLUTINADO FROM dbo.ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO"

SET str1=db.execute(ssql1)

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
document.frm1.txtorgao.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtorgao.value = document.frm1.txtorgao.value + "," + fbox.options[i].value;
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

<title>Base de Dados de Coordenadores Locais</title>
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
carrega_txt1(document.frm1.list1);
window.location="cad_orgao_cli.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
}

function manda02()
{
carrega_txt1(document.frm1.list1);
window.location="cad_orgao_cli.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
}

function manda03()
{
carrega_txt1(document.frm1.list1);
window.location="cad_orgao_cli.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
}

function manda04()
{
carrega_txt1(document.frm1.list1);
window.location="cad_orgao_cli.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&str04="+document.frm1.Str04.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
}


function Confirma() 
{
if(document.frm1.list1.options.length == 0)
{
alert("É obrigatória a seleção de pelo menos um ORGÃO!");
document.frm1.list1.focus();
return;
}
else
{
carrega_txt1(document.frm1.list1);
document.frm1.submit();
}
}

function apaga_item()
{
var f=document.frm1.list1.selectedIndex;
if (f!=-1){
	document.frm1.list1.options[f] = null;
	document.frm1.list1.selectedIndex=f-1;
}
}

function apaga_item2()
{
var f = document.frm1.list1.options.length;
var items = '';
for(var i = 0; i < f; i++)
{
if (document.frm1.list1.options[i].selected)
{
	items = items + ';' + i
}
}
items=items + ';';
var t = document.frm1.list1.options.length;
var f = -1;
for(var d = 0; d < t + 1; d++)
{
var s = ';'+d+';';
if(items.search(s)!=-1)
{
if(f==-1)
{
document.frm1.list1.options[d] = null;
f=d;
}
else
{
document.frm1.list1.options[f] = null;
}
}
}
}

function verifica_tecla(e)
{
if(window.event.keyCode==16)
{
//alert("Tecla não permitida!");
//return;
}
}

</SCRIPT>


<script language="javascript" src="../../Apoio/Clis/troca_lista2.js"></script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onKeyDown="verifica_tecla()">
<form name="frm1" method="POST" action="../../Apoio/Clis/valida_cad_orgao_cli.asp">

  <input type="hidden" name="chave" size="13" value="<%=REQUEST("CHAVE")%>">
  <input type="hidden" name="atribb" size="13" value="<%=REQUEST("ATRIB")%>">
  <input type="hidden" name="txtorgao" size="78" value="<%=REQUEST("LISTA")%>">

  <table width="798" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="158" height="20" colspan="2">&nbsp;</td>
      <td width="349" height="60" colspan="3">&nbsp;</td>
      <td width="285" valign="top" colspan="2"> <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> <div align="center"> 
                <p align="center"> </div></td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> <div align="center">&nbsp;</div></td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> <div align="center">&nbsp;</div></td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center">&nbsp;</div></td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center">&nbsp;</div></td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center">&nbsp;</div></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="37">&nbsp; </td>
      <td height="20" width="119"> <p align="right"> <img src="../../../imagens/confirma_f02.gif" width="24" height="24" border="0" onclick="Confirma()"> 
      </td>
      <td height="20" width="178"> <font size="2" face="Verdana" color="#000080"><b>Enviar</b></font> 
      </td>
      <td height="37">
        <p align="right"><a href="../../Apoio/Clis/menu_cli.asp?cli=<%=Session("cli")%>"><img src="../../../imagens/volta_f02.gif" width="24" height="24" border="0"></a> </td>
      <td height="37"> <font color="#000080" size="2" face="Verdana"><b>Menu Principal</b></font> </td>
      <td height="20" width="142">&nbsp; </td>
      <td height="20" width="141">&nbsp; </td>
    </tr>
  </table>
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#000080">Coordenadores Locais de Implanta&ccedil;&atilde;o </font></p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#000080" size="2">Associação
de Órgãos</font></b></p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" width="80%" cellspacing="3" cellpadding="2">
    <tr> 
      <td width="19%" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="18%" bgcolor="#000080"><b><font face="Verdana" size="2" color="#FFFFFF">Usuário&nbsp;&nbsp;</font></b></td>
      <td width="82%"><font color="#000080" face="Verdana" size="2"><%=CHAVE%> - <%=USUARIO%></font></td>
    </tr>
  </table>
  <table border="0" width="93%">
    <tr>
      <td width="53%"></td>
      <td width="47%"> <p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>  </td>
    </tr>
  </table>
  <table border="0" width="100%" height="329">
    <tr>
      <td width="22%" height="18"></td>
      <td width="23%" height="18"><font color="#000080" face="Verdana" size="2"><b>Órgão 
        Aglutinador</b></font></td>
      <td width="6%" align="center" height="18"></td>
	  <td width="47%" rowspan="11">

<p><font color="#000066" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>&Oacute;rg&atilde;os Selecionados</strong></font>
<br>
          <select name="list1" size="16" multiple>
  <%
if operacao=1 then

str_valor = request("lista")

if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if
tamanho = Len(str_valor)
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If
tamanho = Len(str_valor)
contador = 1
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)

    If str_temp = "," Then
    
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
        quantos2=len(str_atual)
        
        select case quantos2
        	
        	case 2
        		       	
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & str_atual &"'")        	
				valor_nome=temp_orgao("AGLU_SG_AGLUTINADO")

        	case 7
				
				org_aglu=left(str_atual,2)
				org_maior=right(left(str_atual,5),3)
				org_seq=right(str_atual,2)
				
				if left(org_seq,1)=0 then
					org_seq=right(org_seq,1)
				end if
												
				SSQL=""
				SSQL="SELECT * FROM " & Session("Prefixo") & "ORGAO_MAIOR "
				SSQL=SSQL+"WHERE AGLU_CD_AGLUTINADO=" & org_aglu
				SSQL=SSQL+"AND ORLO_CD_ORG_LOT=" & org_maior
				SSQL=SSQL+"AND ORLO_NR_ORDEM=" & org_seq
				
				set temp_orgao=db.execute(ssql)				
				
				valor_nome=temp_orgao("ORLO_SG_ORG_LOT")

        	case 10
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & str_atual & "00000'")        	
				valor_nome=temp_orgao("ORME_SG_ORG_MENOR")

        	case 13
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & str_atual & "00'")        	
				valor_nome=temp_orgao("ORME_SG_ORG_MENOR")

        	case 15
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & str_atual & "'")        	
				valor_nome=temp_orgao("ORME_SG_ORG_MENOR")

		end select
%>
            <option value="<%=str_atual%>"><%=valor_nome%></option>
            <%

    quantos = 0
    End If
    contador = contador + 1
Loop

else

		set orgaos=db.execute("SELECT * FROM " & Session("prefixo") & "CLI_ORGAO WHERE USMA_CD_USUARIO='" & REQUEST("CHAVE") & "'")

       do until orgaos.eof=true
       
       str_atual=orgaos("ORME_CD_ORG_MENOR")
       
       quantos2=len(str_atual)
        
        select case quantos2
        	
        	case 2
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & str_atual &"'")        	
				valor_nome=temp_orgao("AGLU_SG_AGLUTINADO")

        	case 7
				
				org_aglu=left(str_atual,2)
				org_maior=right(left(str_atual,5),3)
				org_seq=right(str_atual,2)
				
				if left(org_seq,1)=0 then
					org_seq=right(org_seq,1)
				end if
												
				SSQL=""
				SSQL="SELECT * FROM " & Session("Prefixo") & "ORGAO_MAIOR "
				SSQL=SSQL+"WHERE AGLU_CD_AGLUTINADO=" & org_aglu
				SSQL=SSQL+"AND ORLO_CD_ORG_LOT=" & org_maior
				SSQL=SSQL+"AND ORLO_NR_ORDEM=" & org_seq
				
				set temp_orgao=db.execute(ssql)				
				
				valor_nome=temp_orgao("ORLO_SG_ORG_LOT")

        	case 10
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & str_atual & "00000'")        	
				valor_nome=temp_orgao("ORME_SG_ORG_MENOR")

        	case 13
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & str_atual & "00'")        	
				valor_nome=temp_orgao("ORME_SG_ORG_MENOR")

        	case 15
				set temp_orgao=db.execute("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & str_atual & "'")        	
				valor_nome=temp_orgao("ORME_SG_ORG_MENOR")

		end select
%>
            <option value="<%=str_atual%>"><%=valor_nome%></option>
            <%
orgaos.movenext
loop

end if
%>
          </select>
          &nbsp; <a href="#" onclick="apaga_item2()"><img src="../../Apoio/excluir.gif" width="110" height="25" border="0" align="absmiddle"></a><BR>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
          a tecla Ctrl com o mouse para selecionar </font><BR>
          <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">mais 
      de uma op&ccedil;&atilde;o ou para desmarcar um item selecionado.</font>      </td>
	   <td width="2%" rowspan="11">
      </td>
    </tr>
    <tr>
      <td width="22%" height="28"></td>
      <td width="23%" height="28"><select size="1" name="Str01" onChange="manda01()">
        <OPTION VALUE="0">== Selecione um nível ==</OPTION>
        <%
		do until str1.eof=true
        if trim(orgao_1)=trim(str1("AGLU_CD_AGLUTINADO")) then
        %>
        <option selected value="<%=str1("AGLU_CD_AGLUTINADO")%>"><%=str1("AGLU_SG_AGLUTINADO")%></option>
        <%
		else%>
        <option value="<%=str1("AGLU_CD_AGLUTINADO")%>"><%=str1("AGLU_SG_AGLUTINADO")%></option>
        <%
        end if
        str1.movenext
        looP
        %>
        </select></td>
	  <%
	  if sel1=1 and sel2=0 and sel3=0 and sel4=0 and sel5=0 then
	  %>
	  <td width="6%" align="center" height="28"><a href="#" onClick="move(document.frm1.Str01,document.frm1.list1,0)"><img name="Image16" border="0" src="../../../imagens/continua_F01.gif" width="24" height="24"></a></td>
	  <%else%>		  
	  <td width="6%" align="center" height="28"></td>	  
	  <%end if%>
    </tr>
    <tr>
      <td width="22%" height="22"></td>
      <td width="23%" height="22"><font color="#000080" face="Verdana" size="2"><b>Órgão
        de Lotação</b></font></td>
      <td width="6%" align="center" height="22"></td>
    </tr>
    <tr>
      <td width="22%"></td>
      <td width="23%"><select size="1" name="Str02" onChange="manda02()">
      <OPTION VALUE="000">== Selecione o Nível ==</OPTION>
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
	  <%
	  if sel1=1 and sel2=1 and sel3=0 and sel4=0 and sel5=0 then
	  %>
	  <td width="6%" align="center"><a href="#" onClick="move(document.frm1.Str02,document.frm1.list1,0)"><img name="Image16" border="0" src="../../../imagens/continua_F01.gif" width="24" height="24"></a></td>
	  <%else%>
  	  <td width="6%" align="center"></td>
	  <%end if%>
    </tr>
    <tr>
      <td width="22%" height="22"></td>
      <td width="23%" height="22"><font color="#000080" face="Verdana" size="2"><b> Gerência</b></font></td>
      <td width="6%" align="center" height="22"></td>
    </tr>
    <tr>
      <td width="22%" height="28"></td>
      <td width="23%" height="28"><select size="1" name="Str03" onChange="manda03()">
      <OPTION VALUE="0">== Selecione o Nível ==</OPTION>
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
      <%
	  if sel1=1 and sel2=1 and sel3=1 and sel4=0 and sel5=0 then
	  %>
	  <td width="6%" align="center" height="28"><a href="#" onClick="move(document.frm1.Str03,document.frm1.list1,0)"><img name="Image16" border="0" src="../../../imagens/continua_F01.gif" width="24" height="24"></a></td>
	  <%else%>
	  <td width="6%" align="center" height="28"></td>	  
	  <%end if%>
    </tr>
    <tr>
      <td width="22%" height="21"></td>
      <td width="23%" height="21"><font color="#000080" face="Verdana" size="2"><b> Gerência Setorial</b></font></td>
      <td width="6%" align="center" height="21"></td>
    </tr>
    <tr>
      <td width="22%"></td>
      <td width="23%"><select size="1" name="Str04" onChange="manda04()">
      <OPTION VALUE="0">== Selecione o Nível ==</OPTION>
        <%
        do until str4.eof=true
        if trim(orgao_4)=trim(left((str4("ORME_CD_ORG_MENOR")),13)) then
        %>
        <option selected value="<%=LEFT((str4("ORME_CD_ORG_MENOR")),13)%>"><%=str4("ORME_SG_ORG_MENOR")%></option>
        <%else%>
        <option value="<%=LEFT((str4("ORME_CD_ORG_MENOR")),13)%>"><%=str4("ORME_SG_ORG_MENOR")%></option>
        <%

        end if
        str4.movenext
        looP
        %>
        </select></td>
      <%
	  if sel1=1 and sel2=1 and sel3=1 and sel4=1 and sel5=0 then
	  %>
      <td width="6%" align="center"><a href="#" onClick="move(document.frm1.Str04,document.frm1.list1,0)"><img name="Image16" border="0" src="../../../imagens/continua_F01.gif" width="24" height="24"></a></td>
	  <%
	  else
	  %>
      <td width="6%" align="center"></td>	  
	  <%end if%>		  
    </tr>
    <tr>
      <td width="22%" height="21"></td>
      <td width="23%" height="21">&nbsp;</td>
      <td width="6%" align="center" height="21"></td>
    </tr>
    <tr>
      <td width="22%" height="27"></td>
      <td width="23%" height="27">&nbsp;</td>
      <td width="6%" align="center" height="27">&nbsp;</td>
    </tr>
    <tr>
      <td width="22%" height="21"></td>
      <td width="23%" height="21"></td>
      <td width="6%" align="center" height="21"></td>
    </tr>
  </table>
  <p align="left">
</form>
</body>
</html>
