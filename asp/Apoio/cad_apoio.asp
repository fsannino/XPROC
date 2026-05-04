<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="conn_consulta.asp" -->

<html>
<head>

<script language="JavaScript">
<!--
function carrega_txt1(fbox) {
document.frm1.txtmodulo.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtmodulo.value = document.frm1.txtmodulo.value + "," + fbox.options[i].value;
}
}

function carrega_txt2(fbox) {
document.frm1.txtonda.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtonda.value = document.frm1.txtonda.value + "," + fbox.options[i].value;
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

<title>Base de Dados de Apoiadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
opt=request("op")
Chave=Session("cdusuario")
inicio=0

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

'Esta linha está errada para testar o DEBUG do asp
'conn_db.Open Session("Conn_String_Cogest_Gravacao")

tipo=request("tipo")
valor=request("valor")

MOM_1=""
MOM_2=""

ATV_1="checked"
ATV_2=""

momento=0

if trim(valor)="" or valor="null" then
	set rs=db.execute("SELECT * FROM " & Session("Prefixo") & "USUARIO_MAPEAMENTO WHERE USMA_TX_MATRICULA=''")
	inicio=1
	tipo=0
	tem=0
else
	on error resume next
	set rs=db.execute("SELECT * FROM " & Session("Prefixo") & "USUARIO_MAPEAMENTO WHERE USMA_TX_MATRICULA=" & valor & "")
	if rs.eof=true or err.number<>0 then
		set rs=db.execute("SELECT * FROM " & Session("Prefixo") & "USUARIO_MAPEAMENTO WHERE USMA_CD_USUARIO='" & valor & "'")
	end if
end if

if rs.eof=false then
	tem=1
end if

set rs_onda=db.execute("SELECT * FROM " & Session("Prefixo") & "ONDA ORDER BY ONDA_TX_DESC_ONDA")

ssql=""
ssql="SELECT (SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA1, (SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 5),2)) AS FAIXA2 ,(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA3, dbo.MEGA_PROCESSO.MEPR_TX_ABREVIA AS FAIXA4 , dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA,  dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO FROM dbo.SUB_MODULO "
ssql=ssql+"LEFT JOIN dbo.MEGA_PROCESSO ON "
ssql=ssql+"(CONVERT(char(3),dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO)) IN ((REPLACE((REPLACE(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS,'0','')),'-',', '))) "
ssql=ssql+" ORDER BY FAIXA1, FAIXA2, FAIXA3, FAIXA4, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "

set rs_modulo=db.execute(ssql)

EMPRESA = rs("USMA_TX_EMPRESA")

if tem=1 then
	a = "SELECT * FROM " & Session("Prefixo") & "APOIO_LOCAL_MULT WHERE APLO_NR_ATRIBUICAO=" & TIPO & " AND USMA_CD_USUARIO='" & RS("USMA_CD_USUARIO")& "'"
	set fonte=db.execute(a)
end if

if fonte.eof=true and opt=2 then%>
<script>
alert('Registro não Encontrado');
history.go(-1);
</script>
<%
end if

if fonte.eof=false then

	edita=1
	tipo=fonte("APLO_NR_ATRIBUIÇÃO")
	obs=fonte("APLO_TX_OBS")
	
	IF FONTE("APLO_NR_RELACAO_EMPREGO")="E" THEN
		EMPRESA="PETROBRAS"
	ELSE
		EMPRESA="CONTRATADA"
	END IF
	
	select case fonte("APLO_NR_MOMENTO")
	
	case 12
		MOM_1="checked"
		MOM_2="checked"
	case 1
		MOM_1="checked"
		MOM_2=""
	case 2
		MOM_1=""
		MOM_2="checked"
	case else
		MOM_1=""
		MOM_2=""
	end SELECT
	
	momento=fonte("APLO_NR_MOMENTO")
	
	select case fonte("APLO_NR_SITUACAO")
	
	case 1
		ATV_1="checked"
		ATV_2=""
	case 2
		ATV_1=""
		ATV_2="checked"
	case else
		ATV_1="checked"
		ATV_2=""
	end SELECT

	set fonte_modulo=db.execute("SELECT * FROM " & Session("Prefixo") & "APOIO_LOCAL_MODULO WHERE APLO_NR_ATRIBUICAO=" & TIPO & " AND USMA_CD_USUARIO='" & RS("USMA_CD_USUARIO")& "'")
	SSQL=""
	SSQL="SELECT * FROM " & Session("Prefixo") & "APOIO_LOCAL_ONDA WHERE APLO_NR_ATRIBUICAO=" & TIPO & " AND USMA_CD_USUARIO='" & RS("USMA_CD_USUARIO")& "'"
	Set fonte_onda=db.execute(SSQL)
		
else
	edita=0
end if

%>

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
function Confirma() 
{
if(document.frm1.list2.options.length == 0)
{
alert("É obrigatória a seleção de pelo menos um MÓDULO!");
document.frm1.list1.focus();
return;
}
if(document.frm1.list4.options.length == 0)
{
alert("É obrigatória a seleção de pelo menos uma ONDA!");
document.frm1.list3.focus();
return;
}
<%IF TIPO=1 THEN%>
if((document.frm1.strmomento1.checked == false) && (document.frm1.strmomento2.checked == false))
{
alert("É obrigatória a seleção de pelo menos um MOMENTO!");
document.frm1.strmomento1.focus();
return;
}
<%END IF%>
else
{
carrega_txt1(document.frm1.list2);
carrega_txt2(document.frm1.list4);
document.frm1.submit();
}
}

function pega_func()
{
if(document.frm1.sel_apoio.value!=0)
{
	window.open("procura.asp?apoio=" + document.frm1.sel_apoio.value + "","_blank","width=240,height=150,history=0,scrollbars=0,titlebar=0,resizable=0,status=0")
}
}

function carrega_momento()
{
		if(document.frm1.strmomento1.checked==true && document.frm1.strmomento2.checked==true)
		{
		document.frm1.txtmomento.value ='12'
		}
		
		if(document.frm1.strmomento1.checked==true && document.frm1.strmomento2.checked==false)
		{
		document.frm1.txtmomento.value ='1'
		}
		
		if(document.frm1.strmomento1.checked==false && document.frm1.strmomento2.checked==true)
		{
		document.frm1.txtmomento.value ='2'
		}
		
		if(document.frm1.strmomento1.checked==false && document.frm1.strmomento2.checked==false)
		{
		document.frm1.txtmomento.value ='0'
		}
}

function deleta()
{
if (confirm("Confirma exclusão da associação do Usário Atual?"))
{
window.location="exclui_apoio.asp?chave="+document.frm1.txtchave.value+"&tipo="+document.frm1.txttipo.value+"&CdUser="+document.frm1.CDUser.value
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
<script language="javascript" src="../js/troca_lista.js"></script>

<body bgcolor="#FFFFFF" text="#000000" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif')">
<form name="frm1" method="POST" action="valida_cad_apoio.asp">
  <table width="984" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="97">
    <tr> 
      <td width="142" height="60" colspan="2">&nbsp;</td>
      <td width="384" height="60" colspan="4">
        <p align="center"><font size="1" face="Arial Narrow" color="#FFFFFF"></font></p>
      </td>
      <td width="452" valign="top" colspan="3" height="60"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"> 
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center">&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="37" width="26">&nbsp; </td>
      <td height="37" width="114">
        <p align="right">
        <%if tipo<>0 and tem<>0 then%>
        <img border="0" src="../../imagens/confirma_f02.gif" onclick="Confirma()">
        <%end if%>
        </td>
      <td height="37" width="163">
      <font color="#000080">
      <%if tipo<>0 and tem<>0 then%>
      <font size="2" face="Verdana" color="#000080"><b>Enviar</b></font>
      <%end if%>
      </font>
      </td>
      <td height="37" width="41">
      <p align="right"><a href="menu.asp"><img border="0" src="../../imagens/volta_f02.gif"></a>
      </td>
      <td height="37" width="150">
      <font color="#000080" size="2" face="Verdana"><b>Menu Principal</b></font>
      </td>
      <td height="37" width="28">&nbsp;
        
 </td>
      <td height="37" width="55">&nbsp; </td>
      <td height="37" width="159">
        <p align="center">&nbsp;</p>
 </td>
      <td height="37" width="236">&nbsp; </td>
    </tr>
  </table>
  <p align="center"><font size="3" face="Verdana" color="#000080">Cadastro de
  Apoiadores Locais / Multiplicadores</font></p>
  <table border="0" width="837">
    <tr>
      <%if tem<>0 then%>
      <td width="133" bgcolor="#000080" height="19">
        <p align="center"><font color="#FFFFFF" size="2" face="Verdana"><b>Atribuição</b></font></td>
      
      <td width="325" height="19">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><select size="1" name="sel_apoio" onChange="pega_func()">
          <%if tipo=1 then%>
          <option selected value="1">APOIADOR LOCAL</option>
          <%else%>
          <option selected value="2">MULTIPLICADOR</option>
          <%end if%>
          </select> 
        <%if fonte.eof=false then%>
        <input type="button" value="Excluir Registro" name="E1" onclick="deleta()"></p>
        <%end if%>
        <%end if%>
      </td>
      
      <td width="359" height="19">
      <%if tipo=1 then%>  
      <%end if%>
      </td>
    </tr>
  </table>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
  
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
 <%
  if tipo<>0 and rs.eof=false then%>
  <table border="0" width="28%" height="23">
    <tr>
      <td width="7%" bgcolor="#FFFFFF" height="17">
        <p align="center"><b><font color="#FFFFFF" face="Verdana" size="2"><%=valor1%><input type="radio" name="str_ativo" value="1" <%=ATV_1%>></font></b></td>
      <td width="42%" height="17" bgcolor="#FFFFFF"><font face="Verdana" size="2"><b>Ativo</b></font></td>
      <td width="13%" bgcolor="#FFFFFF" height="17">
        <p align="center"><b><font color="#FFFFFF" size="2" face="Verdana"><%=valor2%><input type="radio" name="str_ativo" value="2" <%=ATV_2%>></font></b></td>
      <td width="45%" height="17" bgcolor="#FFFFFF"><font face="Verdana" size="2"><b>Inativo</b></font></td>
    </tr>
  </table>
  <table border="0" width="28%" height="23">
    <% 
		IF trim(EMPRESA)="PETROBRAS" then
    		valor1="X"
    		valor2=""
    		EMP="E"
    	else
			valor1=""
			valor2="X"    	
			EMP="C"
    	end if
    	
    	IF EMPRESA="" THEN
    		valor1=""
    		valor2=""
    		EMP=""
    	END IF
    %>
    <tr>
      <td width="13%" bgcolor="#000080" height="17">
        <p align="center"><b><font color="#FFFFFF" face="Verdana" size="2"><%=valor1%></font></b></td>
      <td width="38%" height="17"><font face="Verdana" size="2"><b>Empregado</b></font></td>
      <td width="14%" bgcolor="#000080" height="17">
        <p align="center"><b><font color="#FFFFFF" size="2" face="Verdana"><%=valor2%></font></b></td>
      <td width="46%" height="17"><font face="Verdana" size="2"><b>Contratado</b></font></td>
    </tr>
  </table>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" width="96%" cellpadding="3" height="57">
    <tr> 
      <%
    IF RS("USMA_TX_MATRICULA")=0 THEN
    	MATRIC=""
    ELSE
    	MATRIC=RS("USMA_TX_MATRICULA")
    END IF
    %>
      <td width="12%" bgcolor="#000080" height="6"><font color="#FFFFFF" size="2" face="Verdana"><b>Matrícula</b></font></td>
      <td width="21%" height="6"><font face="Verdana" size="2"><%=MATRIC%></font></td>
      <td width="12%" bgcolor="#000080" height="6"><font color="#FFFFFF" size="2" face="Verdana"><b>Nome 
        </b></font></td>
      <td width="55%" height="6"><font face="Verdana" size="2"><%=RS("USMA_TX_NOME_USUARIO")%></font></td>
    </tr>
    <tr> 
      <td width="12%" bgcolor="#000080" height="25"><font color="#FFFFFF" size="2" face="Verdana"><b>Lotação</b>&nbsp;</font></td>
       <%
      SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR=" & RS("ORME_CD_ORG_MENOR"))
      %>
	  <td width="21%" height="25"><font face="Verdana" size="2"><%=temp("ORME_SG_ORG_MENOR")%></font></td>
      <td width="12%" bgcolor="#000080" valign="top" height="25"><font color="#FFFFFF" size="2" face="Verdana"><b>Ramal 
        </b></font></td>
      <td height="25"><font face="Verdana" size="2"><%=RS("USUA_TX_RAMAL")%></font></td>
    </tr>
    <tr> 
      <%
      SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR=" & RS("ORME_CD_ORG_MENOR"))
      %>
      <td width="12%" bgcolor="#000080" height="18"><font color="#FFFFFF" size="2" face="Verdana"><b>Chave</b></font></td>
      <td width="21%" height="18"><font face="Verdana" size="2"><%=RS("USMA_CD_USUARIO")%></font></td>
      <td width="12%" bgcolor="#FFFFFF" valign="top" height="18"></td>
      <td height="18">&nbsp;</td>
    </tr>
  </table>
  <table width="90%" height="132" border="0" cellpadding="2" cellspacing="3">
    <tr> 
      <td height="19" bgcolor="#000080" colspan="4"><font face="Verdana" size="2" color="#FFFFFF"><b>Assunto</b></font></td>
    </tr>
    <tr> 
      <td width="22%" height="100" rowspan="5" align="center"> <p align="left"> 
          <select size="4" name="list1" multiple>
            <%
        DO UNTIL RS_MODULO.EOF=TRUE
        IF RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>33 AND RS_MODULO("SUMO_NR_CD_SEQUENCIA")<>34 THEN
        SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("Prefixo") & "APOIO_LOCAL_MODULO WHERE APLO_NR_ATRIBUICAO=" & TIPO & " AND USMA_CD_USUARIO='" & RS("USMA_CD_USUARIO")& "' AND SUMO_NR_CD_SEQUENCIA=" & RS_MODULO("SUMO_NR_CD_SEQUENCIA"))        
        IF TEMP.EOF=TRUE THEN
		IF ISNULL(RS_MODULO("FAIXA4")) THEN
        IF RS_MODULO("FAIXA1") = RS_MODULO("FAIXA2") THEN
		%>
            <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("FAIXA1")%>-<%=RS_MODULO("FAIXA3")%> 
            - <%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
            <%
		ELSE
		%>
            <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("FAIXA1")%>-<%=RS_MODULO("FAIXA2")%>-<%=RS_MODULO("FAIXA3")%> 
            - <%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
            <%
		END IF
		ELSE
		%>
            <option value="<%=RS_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=RS_MODULO("FAIXA1")%> 
            - <%=RS_MODULO("SUMO_TX_DESC_SUB_MODULO")%></option>
            <%
		END IF
        END IF
        END IF
        RS_MODULO.MOVENEXT
        LOOP
        %>
          </select>
      </td>
      <td width="20%" height="7" align="center"></td>
      <td width="57%" height="100" rowspan="5" align="center"> <p align="left"> 
          <select size="4" name="list2" multiple>
            <%
        IF EDITA=1 THEN
        do until fonte_modulo.eof=true
		
		ssql=""
		ssql="SELECT (SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA1, (SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = LEFT(RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 5),2)) AS FAIXA2 ,(SELECT MEPR_TX_ABREVIA FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = RIGHT(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, 2)) AS FAIXA3, dbo.MEGA_PROCESSO.MEPR_TX_ABREVIA AS FAIXA4 , dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS, dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA,  dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO FROM dbo.SUB_MODULO "
		ssql=ssql+"LEFT JOIN dbo.MEGA_PROCESSO ON "
		ssql=ssql+"(CONVERT(char(3),dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO)) IN ((REPLACE((REPLACE(dbo.SUB_MODULO.MEPR_CD_MEGA_PROCESSO_TODOS,'0','')),'-',', '))) "
		ssql=ssql+"WHERE dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA=" & fonte_MODULO("SUMO_NR_CD_SEQUENCIA") & " "
		ssql=ssql+" ORDER BY FAIXA1, FAIXA2, FAIXA3, FAIXA4, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
		
		SET TEMP=DB.EXECUTE(SSQL)
        NOME=TEMP("SUMO_TX_DESC_SUB_MODULO")
		IF ISNULL(TEMP("FAIXA4")) THEN
		IF TEMP("FAIXA1") = TEMP("FAIXA2") THEN
        %>
            <option value="<%=fonte_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=TEMP("FAIXA1")%>-<%=TEMP("FAIXA3")%> 
            - <%=NOME%></option>
            <%ELSE%>
            <option value="<%=fonte_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=TEMP("FAIXA1")%>-<%=TEMP("FAIXA2")%>-<%=RS_MODULO("FAIXA3")%> 
            - <%=NOME%></option>
            <%
		END IF
		ELSE
		%>
            <option value="<%=fonte_MODULO("SUMO_NR_CD_SEQUENCIA")%>"><%=TEMP("FAIXA1")%> 
            - <%=NOME%></option>
            <%
		END IF
        fonte_modulo.movenext
        loop
        end if
        %>
          </select>
      </td>
      <td width="1%" height="100" rowspan="5"></td>
    </tr>
    <tr> 
      <td width="20%" height="35" align="center"><img src="../../imagens/continua_F01.gif" name="Image161" width="24" height="24" border="0" id="Image161" onClick="move(document.frm1.list1,document.frm1.list2,1)" onMouseOver="MM_swapImage('Image161','','../../imagens/continua_F02.gif',1)" onMouseOut="MM_swapImgRestore()"></td>
    </tr>
    <tr> 
      <td width="20%" height="10" align="center"></td>
    </tr>
    <tr> 
      <td width="20%" height="22" align="center"><img src="../../imagens/continua2_F01.gif" name="img015111" width="24" height="24" border="0" id="img015111" onClick="move(document.frm1.list2,document.frm1.list1,1)" onMouseOver="MM_swapImage('img015111','','../../imagens/continua2_F02.gif',1)" onMouseOut="MM_swapImgRestore()"></td>
    </tr>
    <tr> 
      <td width="20%" height="11" align="center"></td>
    </tr>
  </table>
  <table border="0" width="90%" height="144" cellspacing="3" cellpadding="2">
    <tr>
      <td width="100%" height="19" bgcolor="#000080" colspan="4"><font face="Verdana" size="2" color="#FFFFFF"><b>Onda</b></font></td>
    </tr>
    <tr>
      <td width="6%" height="100" rowspan="5" align="center">
        <p align="left">
          <select size="6" name="list3" multiple>
            <%
        DO UNTIL RS_ONDA.EOF=TRUE
        SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("Prefixo") & "APOIO_LOCAL_ONDA WHERE APLO_NR_ATRIBUICAO=" & TIPO & " AND USMA_CD_USUARIO='" & RS("USMA_CD_USUARIO")& "' AND ONDA_CD_ONDA=" & RS_ONDA("ONDA_CD_ONDA"))        
        IF TEMP.EOF=TRUE THEN
        %>
        <option value="<%=RS_ONDA("ONDA_CD_ONDA")%>"><%=RS_ONDA("ONDA_TX_DESC_ONDA")%></option>
        <%
        END IF
        RS_ONDA.MOVENEXT
        LOOP
        %>
        </select></td>
      <td width="4%" height="23" align="center"></td>
      <td width="46%" height="100" rowspan="5" align="center">
        <p align="left">
          <select size="6" name="list4" multiple>
            <%
        IF EDITA=1 THEN
        do until fonte_onda.eof=true
        SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("Prefixo") & "ONDA WHERE ONDA_CD_ONDA=" & fonte_ONDA("ONDA_CD_ONDA"))
        NOME2=TEMP("ONDA_TX_DESC_ONDA")
        %>
        <option value="<%=fonte_ONDA("ONDA_CD_ONDA")%>"><%=NOME2%></option>
        &nbsp;
        <%
        fonte_onda.movenext
        loop
        end if
        %>
        &nbsp;
        </select></td>
      <td width="44%" height="100" rowspan="5"></td>
    </tr>
    <tr>
      <td width="4%" height="23" align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list3,document.frm1.list4,1)"><img name="Image16" border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a></td>
    </tr>
    <tr>
      <td width="4%" height="22" align="center"></td>
    </tr>
    <tr>
      <td width="4%" height="22" align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list4,document.frm1.list3,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24"></a></td>
    </tr>
    <tr>
      <td width="4%" height="2" align="center"></td>
    </tr>
  </table>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; 
  
  <input type="hidden" name="txtonda" size="13">
  <input type="hidden" name="txtmodulo" size="13">
  <input type="hidden" name="txtchave" size="13" value="<%=RS("USMA_CD_USUARIO")%>">
  <input type="hidden" name="txtedita" size="13" value="<%=edita%>">
  <input type="hidden" name="txttipo" size="13" value="<%=tipo%>">
  <input type="hidden" name="txtmomento" size="13" value="<%=momento%>">
  <input type="hidden" name="txtvinculo" size="13" value="<%=EMP%>">
  <input type="hidden" name="txtorgao" size="13" value="<%=RS("ORME_CD_ORG_MENOR")%>">
  <input type="hidden" name="CDUser" size="13" value="<%=Session("CdUsuario")%>">
  
  <%if tipo=1 then%>
  <table border="0" width="90%" cellspacing="3">
    <tr>
      <td width="100%" bgcolor="#000080">
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#FFFFFF"><b>Momento da Chegada do Apoiador Local</b></font></p>
      </td>
    </tr>
    <tr>
      <td width="100%"><font face="Verdana" size="2"><input type="checkbox" name="strmomento1" value="1" onclick="carrega_momento()" <%=MOM_1%>>
        Momento
        1 - Completeza; Mapeamentos Treinamento e Perfil; Testes Integrados</font></td>
    </tr>
    <tr>
      <td width="100%"><font face="Verdana" size="2"><input type="checkbox" name="strmomento2" value="2" onclick="carrega_momento()" <%=MOM_2%>>
        Momento
        2 - Partida e Estabilização</font></td>
    </tr>
  </table>
  <%end if%>
  <table border="0" width="90%" cellspacing="3" cellpadding="2">
    <tr>
      <td width="100%" bgcolor="#000080"><font face="Verdana" size="2" color="#FFFFFF"><b>Observação</b></font></td>
    </tr>
    <tr>
      <td width="100%"><textarea rows="4" name="str_obs" cols="108"><%=fonte("APLO_TX_OBS")%></textarea></td>
    </tr>
  </table>
  <%
  tem=1
  end if
  if tem=0 and inicio<>1 then
  %>
	<center>  
  <font color="#800000"><b>Nenhum Registro Encontrado</b></font>
  </center>
  <%
  end if
  %>
</form>
</body>
</html>
