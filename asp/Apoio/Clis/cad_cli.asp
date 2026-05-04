<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../conn_consulta.asp" -->

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
//-->
</script>

<title>Base de Dados de Coordenadores Locais</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
opt=request("op")
Chave=Session("cdusuario")
inicio=0

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

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

if tem=1 then
	a = "SELECT * FROM " & Session("Prefixo") & "CLI WHERE USMA_CD_USUARIO='" & RS("USMA_CD_USUARIO")& "'"
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
else
	edita=0
end if

%>

<script language="JavaScript">

var message="SINERGIA - Conteúdo Protegido"; 

function click(e) {
if (document.all) {
if (event.button == 2) {
//alert(message);
//return false;
}
}
if (document.layers) {
if (e.which == 3) {
//alert(message);
//return false;
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
document.frm1.submit();
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
window.location="exclui_cli.asp?chave="+document.frm1.txtchave.value+"&tipo="+document.frm1.txttipo.value
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

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" vlink="#0000FF" alink="#0000FF" onKeyDown="verifica_tecla()">
<form name="frm1" method="POST" action="valida_cad_cli.asp">
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
        <%if tem<>0 then%>
          <img src="../../../imagens/confirma_f02.gif" width="24" height="24" border="0" onclick="Confirma()"> 
          <%end if%>
      </td>
      <td height="37" width="163">
      <font color="#000080">
      <%if tem=1 and opt=1 and edita=0 then%>
      <font size="2" face="Verdana" color="#000080"><b>Incluir</b></font>
	  <%else%>
	  <font size="2" face="Verdana" color="#000080"><b>Alterar</b></font>
      <%end if%>
      </font>
      </td>
      <td height="37" width="41">
      <p align="right"><a href="menu_cli.asp?cli=<%=Session("cli")%>"><img src="../../../imagens/volta_f02.gif" width="24" height="24" border="0"></a> 
      </td>
      <td height="37" width="150">
      <font color="#000080" size="2" face="Verdana"><b>Menu Principal</b></font>
      </td>
      <td height="37" width="28">&nbsp;
        
 </td>
      <td height="37" width="55">&nbsp; </td>
      <td height="37" width="159"><p align="center">&nbsp; </p>
 </td>
      <td height="37" width="236">&nbsp; </td>
    </tr>
  </table>
  <%if edita=0 then%>
  <p align="center" style="margin-bottom: 0"><font size="3" face="Verdana" color="#000080">Cadastro 
    de Coordenador</font></p>
  <%else%>
  <p align="center" style="margin-bottom: 0"><font size="3" face="Verdana" color="#000080">Alteração de Dados  
    de Coordenador</font></p>
	<%end if%>
  <p>
  
  <%
  if rs.eof=false then%>
  </p>
  <table border="0" width="62%" height="40">
    <% 
		IF rs("USMA_TX_MATRICULA")<>0 then
    	valor1="X"
    	valor2=""
        	else
		valor1=""
		valor2="X"    	
    	end if
    %>
    <tr> 
      <td width="17%" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="5%" bgcolor="#000080" height="17"> <p align="center"><b><font color="#FFFFFF" face="Verdana" size="2"><%=valor1%></font></b></td>
      <td width="18%" height="17"><font face="Verdana" size="2"><b>Empregado</b></font></td>
      <td width="5%" bgcolor="#000080" height="17"> <p align="center"><b><font color="#FFFFFF" size="2" face="Verdana"><%=valor2%></font></b></td>
      <td width="21%" height="17"><font face="Verdana" size="2"><b>Contratado</b></font></td>
      <td width="34%"><p align="left"> 
          <%IF edita=1 THEN%>
          <input type="button" value="Excluir Registro" name="E1" onClick="deleta()">
          <%END IF%>
        </p>
 </td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF" height="17">&nbsp;</td>
      <td height="17" bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#FFFFFF" height="17">&nbsp;</td>
      <td height="17" bgcolor="#FFFFFF">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <br>

  <table border="0" width="76%" cellpadding="3" height="142">
    <tr> 
      <td width="14%" bgcolor="#FFFFFF">&nbsp;</td>
      <%
    IF RS("USMA_TX_MATRICULA")=0 THEN
    	MATRIC=" - "
    ELSE
    	MATRIC=RS("USMA_TX_MATRICULA")
    END IF
    %>
      <td width="23%" bgcolor="#000080" height="6"><font color="#FFFFFF" size="2" face="Verdana"><b>Matrícula</b></font></td>
      <td width="63%" height="6"><font face="Verdana" size="2"><%=MATRIC%></font></td>
    </tr>
    <tr> 
      <td width="14%" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="23%" bgcolor="#000080" height="28"><font color="#FFFFFF" size="2" face="Verdana"><b>Nome</b></font></td>
      <td width="63%" height="28"><font face="Verdana" size="2"><%=RS("USMA_TX_NOME_USUARIO")%></font></td>
    </tr>
    <tr> 
      <td width="14%" bgcolor="#FFFFFF">&nbsp;</td>
      <%
      SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("Prefixo") & "ORGAO_MENOR WHERE ORME_CD_ORG_MENOR=" & RS("ORME_CD_ORG_MENOR"))
      %>
      <td width="23%" bgcolor="#000080" height="26"><font color="#FFFFFF" size="2" face="Verdana"><b>Lotação</b></font></td>
      <td width="63%" height="26"><font face="Verdana" size="2"><%=temp("ORME_SG_ORG_MENOR")%></font></td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF">&nbsp;</td>
      <td bgcolor="#000080" height="26"><font color="#FFFFFF" size="2" face="Verdana"><b>Chave</b></font></td>
      <td height="26"><font face="Verdana" size="2"><%=RS("USMA_CD_USUARIO")%></font></td>
    </tr>
    <tr> 
      <td width="14%" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="23%" bgcolor="#000080" height="26"><font color="#FFFFFF" size="2" face="Verdana"><b>Ramal 
        </b></font></td>
      <td width="63%" height="26"> <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2"><%=RS("USUA_TX_RAMAL")%></font></p></td>
    </tr>
  </table>
  <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> 
    <input type="hidden" name="txtchave" size="13" value="<%=RS("USMA_CD_USUARIO")%>">
  <input type="hidden" name="txtedita" size="13" value="<%=edita%>">
  <input type="hidden" name="txttipo" size="13" value="<%=tipo%>">
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
