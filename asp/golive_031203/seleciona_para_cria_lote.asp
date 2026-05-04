<%@LANGUAGE="VBSCRIPT"%>
<%
response.Buffer=false

operacao = request("opti")

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation=3

if request("txtDescLote") <> "" then
   str_DescLote = request("txtDescLote")
else
   str_DescLote = ""
end if

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

if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = 0
end if

if request("selOnda") <> 0 then
   str_Onda = request("selOnda")
else
   str_Onda = 0
end if

if request("txtFuncSel") <> "" then
   str_FuncSel = request("txtFuncSel")
else
   str_FuncSel = ""
end if

if request("txtOrgSel") <> "" then
   str_OrgSel = request("txtOrgSel")
else
   str_OrgSel = ""
end if

if request("txtDescOrgao") <> "" then
   str_DescOrgao = request("txtDescOrgao")
else
   str_DescOrgao = ""
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

str_SQL = ""
str_SQL = str_SQL & " SELECT * "
str_SQL = str_SQL & " FROM FUNCAO_NEGOCIO "
str_SQL = str_SQL & " WHERE MEPR_CD_MEGA_PROCESSO > 0 "
if str_MegaProcesso <> 0 then
	str_SQL = str_SQL & " AND MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
end if	
if str_Onda = 1 or str_Onda = 2 or str_Onda = 5 then
	str_SQL = str_SQL & " AND FUNE_NM_ANTECIPADA  = 1 "
end if	
str_SQL = str_SQL &  " ORDER BY FUNE_TX_TITULO_FUNCAO_NEGOCIO"
'response.Write(str_SQL)
'response.End()
set rs_funcao = db.execute (str_SQL)

set rs_usuario = db.execute("SELECT * FROM USUARIO_MAPEAMENTO ORDER BY USMA_TX_NOME_USUARIO")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO > 0 "
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
set rs_mega=db.execute(str_SQL_MegaProc)

%>
<html>
<head>
<link href="../../css/objinterface.css" rel="stylesheet" type="text/css">
<script language="javascript" src="troca_lista.js"></script>
<script language="JavaScript">
<!--

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
<script language="JavaScript">

function manda()
	{
    carrega_txt1(document.frm1.selFuncaoSele);
	carrega_txt2(document.frm1.list1);
	carrega_txt3(document.frm1.list1);
	document.frm1.action="seleciona_para_cria_lote.asp"	
	document.frm1.submit();
	}

function manda01()
	{
	window.location="seleciona_para_cria_lote.asp?opti=1&str01="+document.frm1.Str01.value+"&pDescLote="+document.frm1.txtDescLote.value
	}

function manda02()
	{
	window.location="seleciona_para_cria_lote.asp?opti=1&str01="+document.frm1.Str01.value+"&pDescLote="+document.frm1.txtDescLote.value+"&str02="+document.frm1.Str02.value
	}

function manda03()
	{
	//alert("str03="+document.frm1.Str03.value)
	//alert("selMegaProcesso="+document.frm1.selMegaProcesso.value)
	window.location="seleciona_para_cria_lote.asp?opti=1&str01="+document.frm1.Str01.value+"&pDescLote="+document.frm1.txtDescLote.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value
}

function manda04()
{
	//alert("str03="+document.frm1.Str03.value)
	//alert("selMegaProcesso="+document.frm1.selMegaProcesso.value)
	window.location="seleciona_para_cria_lote.asp?opti=1&str01="+document.frm1.Str01.value+"&pDescLote="+document.frm1.txtDescLote.value+"&str02="+document.frm1.Str02.value+"&str03="+document.frm1.Str03.value+"&selMegaProcesso="+document.frm1.selMegaProcesso.value+"&selOnda="+document.frm1.selOnda.value
}

/*
function carrega_txt1(fbox) {
document.frm1.txtorgao.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtorgao.value = document.frm1.txtorgao.value + "," + fbox.options[i].value;
}
}
*/

function carrega_txt1(fbox) 
	{
	document.frm1.txtFuncSel.value = "";
	for(var i=0; i<fbox.options.length; i++) 
		{
		document.frm1.txtFuncSel.value = document.frm1.txtFuncSel.value + "," + fbox.options[i].value;
		}
	//alert(document.frm1.txtFuncSel.value)		
	}

function carrega_txt2(fbox) 
	{
	document.frm1.txtOrgSel.value = "";
	for(var i=0; i<fbox.options.length; i++) 
		{
		document.frm1.txtOrgSel.value = document.frm1.txtOrgSel.value + "," + fbox.options[i].value;
		}
	//alert(document.frm1.txtOrgSel.value)
	}

function carrega_txt3(fbox) 
	{
	//alert(fbox.options[0].text)
	document.frm1.txtDescOrgao.value = "";
	for(var i=0; i<fbox.options.length; i++) 
		{
		document.frm1.txtDescOrgao.value = document.frm1.txtDescOrgao.value + "," + fbox.options[i].text;
		}
	//alert(document.frm1.txtDescOrgao.value)
	}

function Confirma() 
	{

	if(document.frm1.txtDescLote.value=="")
		{
		alert("Você deve preencher o campo descrição do Lote!");
		document.frm1.txtDescLote.focus();
		return;		
		}
	if(document.frm1.list1.options.length == 0)
		{
		alert("É obrigatória a seleção de pelo menos um Órgão!");
		document.frm1.list1.focus();
		return;
		} 
	if(document.frm1.selFuncaoSele.options.length == 0)
		{
		alert("É obrigatória a seleção de pelo menos uma Função!");
		document.frm1.selFuncaoSele.focus();
		return;
		} 

/*	if(document.frm1.list1.options.length == 0)
		{
		alert("É obrigatória a seleção de pelo menos um ORGÃO!");
		document.frm1.list1.focus();
		return;
	}		
	if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0)&&(document.frm1.selMegaProcesso.value==0)&&(document.frm1.selFuncaoSele.options.length==0))
	{ 
	alert("Você deve Selecionar pelo menos um dos parâmetros!");
	document.frm1.selOnda.focus();
	return;
	}
*/	
    carrega_txt1(document.frm1.selFuncaoSele)
	carrega_txt2(document.frm1.list1);
	carrega_txt3(document.frm1.list1);
	document.frm1.action="cria_lote.asp"
	document.frm1.submit();
	}

function Confirma33() 
{
	if(document.frm1.txtDescLote.value=="")
		{
		alert("Você deve preencher o campo Descrição do Lote!");
		document.frm1.txtDescLote.focus();
		return;		
		}
	if((document.frm1.Str01.value==0)&&(document.frm1.Str02.value==0)&&(document.frm1.Str03.value==0)&&(document.frm1.selMegaProcesso.value==0)&&(document.frm1.selFuncaoSele.options.length==0))
	{ 
	alert("Você deve selecionar pelo menos um dos parâmetros!");
	document.frm1.selOnda.focus();
	return;
	}
	else
	{
    carrega_txt1(document.frm1.selFuncaoSele)
	document.frm1.submit();
	}
}

function apaga_item()
	{
	var a=event.keyCode;
	if (a==46)
		{
		var f=document.frm1.list1.selectedIndex;
		if (f!=-1)
			{
			document.frm1.list1.options[f] = null;
			document.frm1.list1.selectedIndex=f-1;
			}
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
	carrega_txt1(document.frm1.selFuncaoSele);
	carrega_txt2(document.frm1.list1);		
	}

</SCRIPT>

<script src="troca_lista_sem_retirar.js" language="javascript"></script>

<body bgcolor="#FFFFFF" text="#000000" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../imagens/continua_F02.gif')">
<form name="frm1" method="POST" action="">

  <table width="1015" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" style="border-collapse: collapse" bordercolor="#111111">
    <tr> 
      <td width="158" height="20" colspan="2">&nbsp;</td>
      <td width="493" height="60" colspan="3">&nbsp;</td>
      <td width="364" valign="top" colspan="2"> <table width="179" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr>
            <td bgcolor="#330099" width="1" valign="middle" align="right"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a> 
              </div>
            </td>
            <td bgcolor="#330099" width="1" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr>
            <td bgcolor="#330099" height="12" width="1" valign="middle" align="right"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="1" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="38">&nbsp; </td>
      <td height="20" width="120"> <p align="right"> <a href="#"><img border="0" src="../../imagens/confirma_f02.gif" onclick="Confirma()"></a> 
      </td>
      <td height="20" width="181"> <font size="2" face="Verdana" color="#000080"><b>&nbsp;Enviar</b></font> 
      </td>
      <td height="20" width="44">&nbsp;</td>
      <td height="20" width="268">&nbsp;</td>
      <td height="20" width="84">&nbsp; </td>
      <td height="20" width="280">&nbsp; </td>
    </tr>
  </table>
  <table width="96%" height="498" border="0" cellspacing="2">
    <tr>
      <td width="2%" height="33">&nbsp;</td>
      <td height="33" colspan="3"><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="14%">&nbsp;</td>
            <td width="62%"><font color="#000080" face="Verdana">Cria&ccedil;&atilde;o de Lote para exporta&ccedil;&atilde;o </font></td>
            <td width="24%"><img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></td>
          </tr>
        </table>
      </div></td>
    </tr>
    <tr>
      <td width="2%" height="1"></td>
      <td height="1" colspan="3"></td>
    </tr>
    <tr>
      <td height="18">&nbsp;</td>
      <td height="18" colspan="3"><font color="#000080" face="System" size="2"><b>Descri&ccedil;&atilde;o do Lote</b></font></td>
    </tr>
    <tr>
      <td height="18">&nbsp;</td>
      <td height="18" colspan="3"><input name="txtDescLote" type="text" id="txtDescLote" value="<%=str_DescLote%>" size="50" maxlength="100"> 
        <font color="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif">* este campo serve para voc&ecirc; identifica o lote.</font></td>
    </tr>
    <tr>
      <td height="18">&nbsp;</td>
      <td height="18" colspan="3"><font color="#000080" face="System" size="2"><b>Órgão Aglutinador </b></font></td>
    </tr>
    <tr>
      <td width="2%" height="18"><div align="right"><font color="#000080" face="System" size="2"></font></div></td>
      <td height="18">        <table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="55%"><select name="Str01" size="1" class="listbox02" onChange="manda()">
              <option value="0">== Todos ==</option>
              <% 'response.Flush()
		do until str1.eof=true
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
            <td width="45%"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16111','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.Str01,document.frm1.list1,0);carrega_txt1(document.frm1.selFuncaoSele);carrega_txt2(document.frm1.list1);"><img src="../../imagens/continua_F01.gif" alt="Incluir &Iacute;tem" name="Image16111" width="25" height="24" border="0" id="Image161"></a></td>
          </tr>
        </table></td>
      <td rowspan="5"><font color="#000080" face="System" size="2">&nbsp;
</font>
        <table width="78%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="68%"><font color="#000080" face="System" size="2">
              <select name="list1" size="7" multiple>
                <%
		str_valor = request("txtOrgSel")		
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
				%>
              </select>
            </font></td>
            <td width="32%"><table width="12%" height="50"  border="0" cellpadding="1" cellspacing="5">
              <tr>
                <td><a href="#"><img src="../../imagens/delete_98.gif" alt="Excluir &Iacute;tem Selecionado" width="24" height="24" border="0" onClick="apaga_item2()"></a></td>
              </tr>
            </table>              <font color="#000080" face="System" size="2">&nbsp;            </font></td>
          </tr>
        </table>
      <font color="#000080" face="System" size="2">&nbsp;      </font></td>
      <td height="18">&nbsp;</td>
    </tr>
    <tr>
      <td height="24">&nbsp;</td>
      <td height="24"><b><font face="System" size="2" color="#000080">Unidade - &Oacute;rg&atilde;o maior </font></b></td>
      <td height="24">&nbsp;</td>
    </tr>
    <tr>
      <td width="2%" height="24"><div align="right"><b></b></div></td>
      <td height="24">        <table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="55%"><select name="Str02" size="1" class="listbox02" onChange="manda()">
              <option value="000">== Todas ==</option>
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
            <td width="45%"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1611','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.Str02,document.frm1.list1,0);carrega_txt1(document.frm1.selFuncaoSele);carrega_txt2(document.frm1.list1);"><img src="../../imagens/continua_F01.gif" alt="Incluir &Iacute;tem" name="Image1611" width="25" height="24" border="0" id="Image161"></a></td>
          </tr>
        </table></td>
      <td height="24">&nbsp;</td>
    </tr>
    <tr>
      <td height="13">&nbsp;</td>
      <td height="13"><font color="#000080" face="System" size="2"><b>Gerência - &Oacute;rg&atilde;o menor </b></font></td>
      <td height="13">&nbsp;</td>
    </tr>
    <tr>
      <td width="2%" height="13"><div align="right"><font color="#000080" face="System" size="2"><b> </b></font></div></td>
      <td height="13">        <table width="100%"  border="0" cellspacing="0" cellpadding="1">
          <tr>
            <td width="55%"><select size="1" name="Str03" class="listbox02" >
              <option value="0">== Todas ==</option>
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
            <td width="45%"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1612','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.Str03,document.frm1.list1,0);carrega_txt1(document.frm1.selFuncaoSele);carrega_txt2(document.frm1.list1);"><img src="../../imagens/continua_F01.gif" alt="Incluir &Iacute;tem" name="Image1612" width="25" height="24" border="0" id="Image161"></a></td>
          </tr>
        </table></td>
      <td height="13">&nbsp;</td>
    </tr>
    <tr>
      <td height="23" valign="top">&nbsp;</td>
      <td height="23" valign="bottom"><font color="#000080" face="System" size="2"><b>Mega Processo</b></font></td>
      <td height="23" valign="bottom"></td>
      <td height="23" valign="bottom">&nbsp;</td>
    </tr>
    <tr>
      <td height="23" valign="top"><div align="right"><font color="#000080" face="System" size="2"></font></div></td>
      <td height="23" valign="bottom"><select size="1" name="selMegaProcesso" class="listbox02" onChange="manda()">
        <option value="0">== Todos ==</option>
        <%do until rs_mega.eof=true
       if trim(str_MegaProcesso)=trim(rs_mega("MEPR_CD_MEGA_PROCESSO")) then
       %>
        <option selected value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
        <%ELSE%>
        <option value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
        <%
		end if
		rs_mega.movenext
		loop
		%>
      </select></td>
      <td height="23" valign="bottom"></td>
      <td height="23" valign="bottom">&nbsp;</td>
    </tr>
    <tr>
      <td height="23" valign="top">&nbsp;</td>
      <td height="23" valign="bottom"><b><font face="System" size="2" color="#000080">Onda</font></b></td>
      <td height="23" valign="bottom"></td>
      <td height="23" valign="bottom">&nbsp;</td>
    </tr>
    <tr>
      <td height="23" valign="top"><div align="right"><b></b></div></td>
      <td height="23" valign="bottom"><select size="1" name="selOnda" class="listbox02" onChange="manda()">
        <option value="0">== Todas ==</option>
        <%
      do until rs_onda.eof=true
		if trim(str_Onda)=trim(rs_onda("ONDA_CD_ONDA")) then    	  
      %>
        <option selected value="<%=rs_onda("ONDA_CD_ONDA")%>"><%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
        <%ELSE%>
        <option value="<%=rs_onda("ONDA_CD_ONDA")%>"><%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
        <%
		end if
      rs_onda.movenext
      loop
      %>
      </select></td>
      <td height="23" valign="bottom"></td>
      <td height="23" valign="bottom">&nbsp;</td>
    </tr>
    <tr>
      <td height="23" valign="top"></td>
      <td height="23" valign="bottom"><b><font face="System" size="2" color="#000080">Fun&ccedil;&atilde;o</font></b></td>
      <td height="23" valign="bottom"></td>
      <td height="23" valign="bottom">&nbsp;</td>
    </tr>
    <tr>
      <td width="2%" height="23" valign="top"></td>
      <td width="38%" height="23" valign="bottom"><font color="#000080" face="Verdana" size="1">Funções Disponíveis</font></td>
      <td width="29%" height="23" valign="bottom"></td>
      <td width="31%" height="23" valign="bottom"><font color="#000080" face="Verdana" size="1">Funções Selecionadas</font></td>
    </tr>
    <tr>
      <td width="2%" height="16"></td>
      <td colspan="3"><table width="798" border="0">
        <tr>
          <td width="350">
            <select size="10" name="selFuncaoDisp" style="font-family: Verdana; font-size: 7 pt">
			<% IF NOT rs_funcao.eof THEN %>
			<option value="0">==== Todas as funções ====</option>
      <% 		END IF
      i=0
      reg = rs_funcao.RecordCount
      do until i = reg
      %>
      <option value="<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=LEFT(rs_funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO"),110)%></option>
      <%
      i = i + 1
      rs_funcao.movenext
      loop
      %>
      </select>
          </td>
          <td width="32">
            <table width="30" border="0">
              <tr>
                <td width="24"><img src="../../imagens/continua_F01.gif" alt="Seleciona Função" name="imgSetaDireita1" width="24" height="24" id="imgSetaDireita1" onClick="move(document.frm1.selFuncaoDisp,document.frm1.selFuncaoSele,1)"></td>
              </tr>
              <tr>
                <td><img src="../../imagens/delete_98.gif" alt="Excluir &Iacute;tem Selecionado" name="imgSetaDireita1" width="24" height="24" id="imgSetaDireita1" onClick="deleta(document.frm1.selFuncaoSele,1)"></td>
              </tr>
          </table></td>
          <td width="354">
            <select size="10" name="selFuncaoSele" style="font-family: Verdana; font-size: 7 pt">
            </select>
          </td>
          </tr>
      </table>      </td>
    </tr>
    <tr>
      <td width="2%" height="22"></td>
      <td height="22" colspan="3"><input name="txtFuncSel" type="hidden" value="<%=str_FuncSel%>">
      <input name="txtOrgSel" type="hidden" value="<%=str_OrgSel%>">
      <input type="hidden" name="txtorgao" size="78" value="<%=REQUEST("LISTA")%>">
      <input type="hidden" name="txtDescOrgao" size="78" value="<%=str_DescOrgao%>"></td>
    </tr>
  </table>
  <p align="left">
</form>
<%
db.close
set db = nothing
%>
</body>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>