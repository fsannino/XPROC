<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Apoio/conn_consulta.asp" -->
<%
server.ScriptTimeout=99999999

operacao = request("opti")

chave=request("chave")
atrib=request("atrib")
recebe=request("recebe")

if atrib=1 then
	tipo="APOIADOR LOCAL"
else
	tipo="MULTIPLICADOR"
end if

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set Rusuario=db.execute("SELECT * FROM " & Session("Prefixo") & "USUARIO_MAPEAMENTO WHERE USMA_CD_USUARIO='" & chave & "'")

usuario=Rusuario("USMA_TX_NOME_USUARIO")

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
ssql3=ssql3+" WHERE (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (AGLU_CD_AGLUTINADO = " & ORGAO_1 & ") AND (ORME_CD_STATUS = 'A')"

'ssql3=ssql3+" AND SUBSTRING(ORME_CD_ORG_MENOR,8,3)<>'000'"
'ssql3=ssql3+" AND SUBSTRING(ORME_CD_ORG_MENOR,3,3)='" & right("000"& ORGAO_2,3) & "' AND SUBSTRING(ORME_CD_ORG_MENOR,11,5)='00000' AND SUBSTRING(ORME_CD_ORG_MENOR,8,3)<>'000'"

set str3=db.execute(ssql3)

ssql4=""
ssql4="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql4=ssql4+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql4=ssql4+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,10)='" & ORGAO_3 & "' AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)='00' AND SUBSTRING(ORME_CD_ORG_MENOR,13,3) <> '000'" 

'set str4=db.execute(ssql4)

ssql5=""
ssql5="SELECT  ORLO_CD_ORG_LOT, ORME_CD_ORG_MENOR, AGLU_CD_AGLUTINADO, ORME_SG_ORG_MENOR, ORME_CD_STATUS, ORME_NM_ORG_MENOR FROM dbo.ORGAO_MENOR "
ssql5=ssql5+" WHERE (AGLU_CD_AGLUTINADO = '" & ORGAO_1 & "') AND (ORLO_CD_ORG_LOT = " & ORGAO_2 & ") AND (ORME_CD_STATUS = 'A')"
ssql5=ssql5+" AND SUBSTRING(ORME_CD_ORG_MENOR,1,13)='" & ORGAO_4 & "'  AND SUBSTRING(ORME_CD_ORG_MENOR,14,2)<>'00'" 

'set str5=db.execute(ssql5)
%>
<html>
<head>

<script language="JavaScript">
<!--
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

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
//-->
</script>

<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function manda01()
{
carrega_txt1(document.frm1.list1);
window.location="cad_orgao.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
}

function manda02()
{
carrega_txt1(document.frm1.list1);
window.location="cad_orgao.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
}

function manda03()
{
carrega_txt1(document.frm1.list1);
window.location="cad_orgao.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
}

function manda04()
{
carrega_txt1(document.frm1.list1);
window.location="cad_orgao.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&str04="+document.frm1.Str04.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value
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
var a=event.keyCode;
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

function Seleciona_tudo()
{
var items = document.frm1.Str03.length;
for(var i=0;i<items;i++)
{
document.frm1.Str03.options[i].selected=true;
}
}

</SCRIPT>

<script language="javascript" src="../../Apoio/troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../imagens/continua_F02.gif')">
<form name="frm1" method="POST" action="../../Apoio/valida_cad_orgao.asp">

  <input type="hidden" name="chave" size="13" value="<%=REQUEST("CHAVE")%>">
  <input type="hidden" name="atribb" size="13" value="<%=REQUEST("ATRIB")%>">
  <input type="hidden" name="txtorgao" size="78" value="<%=REQUEST("LISTA")%>">
  <table width="798" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="158" height="20" colspan="2">&nbsp;</td>
      <td width="349" height="60" colspan="3">&nbsp;</td>
      <td width="285" valign="top" colspan="2"> <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td width="55" align="center" valign="middle" bgcolor="#330099"> <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../Apoio/voltar.gif"></a> 
              </div></td>
            <td width="45" align="center" valign="middle" bgcolor="#330099"> <div align="center"><a href="JavaScript:history.forward()"><img src="../../../imagens/avancar.gif" width="30" height="30" border="0"></a></div></td>
            <td width="54" align="center" valign="middle" bgcolor="#330099"> <div align="center"></div></td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../Apoio/imprimir.gif"></a></div></td>
            <td bgcolor="#330099" height="12" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../../imagens/atualizar.gif"></a></div></td>
            <td bgcolor="#330099" height="12" valign="middle" align="center"> 
              <div align="center"><a href="../../Apoio/menu.asp"><img src="../../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="25" width="37">&nbsp; </td>
      <td height="25" width="119"> 
        <p align="right"> <img border="0" src="../../../imagens/confirma_f02.gif" onclick="Confirma()"> 
      </td>
      <td height="25" width="178"> <font size="2" face="Verdana" color="#000080"><b>Enviar</b></font> 
      </td>
      <td height="25" width="44">&nbsp;</td>
      <td height="25" width="123">&nbsp;</td>
      <td height="25" width="142">&nbsp; </td>
      <td height="25" width="141">&nbsp; </td>
    </tr>
  </table>
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana" color="#000080">Apoiadores
  Locais e Multiplicadores</font></p>
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#000080" size="2">Associação
de Órgãos Apoiados</font></b></p>
<p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" width="80%" cellspacing="3" cellpadding="2">
    <tr> 
      <td width="21%" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="13%" bgcolor="#EEEEEE"><div align="right"><b><font face="Verdana" size="2" color="#333333">Usuário:</font></b></div></td>
      <td width="66%"><font color="#000080" face="Verdana" size="2"><%=CHAVE%> - <%=USUARIO%></font></td>
    </tr>
    <tr> 
      <td width="21%" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="13%" bgcolor="#EEEEEE"><div align="right"><b><font face="Verdana" size="2" color="#333333">Atribuição:</font></b></div></td>
      <td width="66%"><font color="#000080" face="Verdana" size="2"><%=ATRIB%> - <%=TIPO%></font></td>
    </tr>
  </table>
  <table border="0" width="79%">
    <tr>
      <td width="30%"></td>
      <td width="70%">
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  </td>
    </tr>
  </table>
  <table border="0" width="100%" height="317">
    <tr> 
      <td width="15%" height="22"></td>
      <td width="20%" height="22"><font color="#000080" face="Verdana" size="2"><b>Órgão 
        Aglutinador</b></font></td>
      <td width="4%" align="center" height="24"></td>
      <td width="15%" rowspan="10" valign="top"> <table width="100%" border="0">
          <tr> 
            <td height="22"> <div align="left"><font color="#000080" face="Verdana" size="2"><b>Selecionados</b></font></div></td>
          </tr>
          <tr> 
            <td valign="top"> 
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

		set orgaos=db.execute("SELECT * FROM " & Session("prefixo") & "APOIO_LOCAL_ORGAO WHERE USMA_CD_USUARIO='" & REQUEST("CHAVE") & "' AND APLO_NR_ATRIBUICAO=" & REQUEST("ATRIB"))

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
            </td>
          </tr>
        </table></td>
      <td width="46%" rowspan="10">&nbsp; </td>
    </tr>
    <tr> 
      <td width="15%" height="28"></td>
      <td width="20%" height="28"><select name="Str01" size="1" onChange="manda01()">
          <OPTION VALUE="0">== Selecione Orgão ==</OPTION>
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
      <td width="4%" align="center" height="28">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15%" height="22"></td>
      <td width="20%" height="22"><font color="#000080" face="Verdana" size="2"><b>Órgão 
        de Lotação</b></font></td>
      <td width="4%" align="center" height="22"></td>
    </tr>
    <tr> 
      <td width="15%" height="39"></td>
      <td width="20%"> <select name="Str02" size="1" onChange="manda02()">
          <option value="000">== Selecione a Lotação ==</option>
          <%
		  conta=0
		  do until str2.eof=true
			if trim(orgao_2)=trim(str2("ORLO_CD_ORG_LOT"))then
	        %>
          <option selected value="<%=ORGAO_1 & right(("000" & str2("ORLO_CD_ORG_LOT")),3) & right(("000" & str2("ORLO_NR_ORDEM")),2)%>"><%=str2("ORLO_SG_ORG_LOT")%></option>
          <%
			else
			%>
          <option value="<%=ORGAO_1 & right(("000" & str2("ORLO_CD_ORG_LOT")),3) & right(("000" & str2("ORLO_NR_ORDEM")),2)%>"><%=str2("ORLO_SG_ORG_LOT")%></option>
          <%
        end if
		conta_item = conta_item +1
        str2.movenext
        looP

		if recebe=1 then
			conta_item=0
		end if
		
		if conta_item=1 then
        %>
          <script>
		{
		document.frm1.Str02.options[1].selected=true
		carrega_txt1(document.frm1.list1);
		window.location="cad_orgao.asp?opti=1&lista="+document.frm1.txtorgao.value+"&str01="+document.frm1.Str01.value+"&str02="+document.frm1.Str02.value+"&Chave="+document.frm1.chave.value+"&Atrib="+document.frm1.atribb.value+"&recebe=1"
		}
		</script>
          <%
		end if
		conta_item=0		
		%>
        </select></td>
      <td width="4%" align="center">&nbsp;</td>
    </tr>
    <tr> 
      <td width="15%" height="22"></td>
      <td width="20%" height="22"><font color="#000080" face="Verdana" size="2"><b> 
        Gerência</b></font></td>
      <td width="4%" align="center" height="22"></td>
    </tr>
    <tr> 
      <td width="15%" height="28" rowspan="2"></td>
      <td width="20%" height="28" rowspan="2"> <table width="59%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="42%" rowspan="2"><select name="Str03" size="5" multiple>
                <%
        do until str3.eof=true
        if trim(orgao_3)=trim(left((str3("ORME_CD_ORG_MENOR")),15)) then
        %>
                <option selected value="<%=trim(left((str3("ORME_CD_ORG_MENOR")),15))%>"><%=str3("ORME_SG_ORG_MENOR")%></option>
                <%else%>
                <option value="<%=trim(left((str3("ORME_CD_ORG_MENOR")),15))%>"><%=str3("ORME_SG_ORG_MENOR")%></option>
                <%

        end if
        str3.movenext
        looP
        %>
              </select></td>
            <td width="58%"><div align="center"></div></td>
          </tr>
          <tr> 
            <td><div align="center"><a href="#"><img src="../../Apoio/selecionar.gif" width="58" height="40" border="0" onClick="Seleciona_tudo()"></a></div></td>
          </tr>
        </table></td>
      <td width="4%" height="42" align="center" valign="middle"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image161','','../../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.Str03,document.frm1.list1,0)"><img src="../../../imagens/continua_F01.gif" alt="Incluir &Iacute;tem" name="Image161" width="25" height="24" border="0" id="Image161"></a></td>
    </tr>
    <tr>
      <td height="30" align="center" valign="middle"><a href="#"><img src="../../../imagens/continua2_F01.gif" alt="Excluir &Iacute;tem Selecionado" width="24" height="24" border="0" onClick="apaga_item2()"></a></td>
    </tr>
    <tr> 
      <td width="15%" height="21"></td>
      <td width="20%" height="21"><div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
          a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
          ou para desmarcar um item selecionado.</font></div></td>
      <td width="4%" align="center" height="21"></td>
    </tr>
    <tr> 
      <td height="21" colspan="3"> <div align="center"></div></td>
    </tr>
    <tr> 
      <td width="15%" height="2"></td>
      <td width="20%" height="2"></td>
      <td width="4%" align="center" height="2"></td>
    </tr>
  </table>
  <p align="left">
</form>
</body>
</html>
