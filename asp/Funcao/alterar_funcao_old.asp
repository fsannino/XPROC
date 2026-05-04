 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_funcao=request("selFuncao")

set rs_func=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TIPO_QUALIFICACAO ORDER BY TPQU_TX_DESC_TIPO_QUALIFICACAO")
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "TIPO_PUBLICO_PRINCIPAL ORDER BY TPPP_TX_DESC_PUB_PRINCIPAL")

set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")
set rs4=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_PUB_PRI WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

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

function carrega_txt1(fbox) {
document.frm1.txtQua.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtQua.value = document.frm1.txtQua.value + "," + fbox.options[i].value;
}
}

function carrega_txt2(fbox) {
document.frm1.txtpub.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtpub.value = document.frm1.txtpub.value + "," + fbox.options[i].value;
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
function Confirma()
{
if(document.frm1.selMegaProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MEGA-PROCESSO!");
document.frm1.selMegaProcesso.focus();
return;
}

if(document.frm1.txtfuncao.value == "")
{
alert("É obrigatória a definição da FUNÇÃO DE NEGÓCIO!");
document.frm1.txtfuncao.focus();
return;
}

if(document.frm1.txtdescfuncao.value == "")
{
alert("É obrigatória a descrição da FUNÇÃO DE NEGÓCIO!");
document.frm1.txtdescfuncao.focus();
return;
}

if (document.frm1.list2.options.length == 0)
{ 
alert("A seleção de pelo menos uma QUALIFICAÇÃO NÃO R/3 é obrigatória !");
document.frm1.list1.focus();
return;
}
else
{
carrega_txt1(document.frm1.list2)
carrega_txt2(document.frm1.list4)
document.frm1.submit();
}
}


</script>
<body topmargin="0" leftmargin="0">
<form method="POST" action="valida_altera_funcao.asp" name="frm1">
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
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
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
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Alteração
        de Fun&ccedil;&atilde;o R/3</font></p>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <table border="0" width="832" height="164">
          <tr>
            <td width="73" height="25"></td>
            <td width="252" height="25" valign="top"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo</b></font></td>
            <td width="494" height="25"><select size="1" name="selMegaProcesso">
                <option value="0">== Selecione o Mega-Processo ==</option>
                	<%do until rs.eof=true
                	if trim(rs_func("MEPR_CD_MEGA_PROCESSO"))=trim(rs("MEPR_CD_MEGA_PROCESSO")) then
                	%>
                <option selected value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
					<%else%>                
                <option value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
					<%
					end if
					rs.movenext
					loop
					%>
              </select></td>
          </tr>
          <tr>
            <td width="73" height="25"></td>
            <td width="252" height="25" valign="top"><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o R/3</b></font></td>
            <td width="494" height="25"><input type="text" name="txtfuncao" size="58" value="<%=rs_func("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%>" maxlength="50"></td>
          </tr>
          <tr>
            <td width="73" height="83"></td>
            <td width="252" height="83" valign="top"><font face="Verdana" size="2" color="#330099"><b>Descrição
              da Fun&ccedil;&atilde;o R/3</b></font>
              <p><input type="hidden" name="Funcao" size="20" value="<%=str_funcao%>"><input type="hidden" name="txtQua" size="20"><input type="hidden" name="txtpub" size="20"></p>
            </td>
            <td width="494" height="83" valign="top">
              <p align="center" style="margin-top: 0; margin-bottom: 0"><textarea rows="4" name="txtdescfuncao" cols="59"><%=rs_func("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></textarea></td>
          </tr>
        </table>
  <table border="0">
  <tr>
    <td width="152" rowspan="5" height="125">
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>Qualificação
              Não R/3</b></font></td>
    <td rowspan="5" height="125" width="366"> 
        <p style="margin-top: 0; margin-bottom: 0" align="right"> 
        <select size="6" name="list1" multiple>
        <%do until rs1.eof=true
        SET RSTEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "' AND TPQU_CD_TIPO_QUALIFICACAO=" & rs1("TPQU_CD_TIPO_QUALIFICACAO"))
        if rstemp.eof=true then
        %>
        <option value="<%=rs1("TPQU_CD_TIPO_QUALIFICACAO")%>"><%=rs1("TPQU_TX_DESC_TIPO_QUALIFICACAO")%></option>
        <%
        end if
        rs1.movenext
        loop
        %>
        </select>
        </p>
    </td>
    <td align="center" height="21" width="28">
      <p style="margin-top: 0; margin-bottom: 0"></td>
    <td width="38" rowspan="5" height="125"> 
        <p style="margin-top: 0; margin-bottom: 0" align="left"> 
        <select size="6" name="list2" multiple>
        <%do until rs3.eof=true
        SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TIPO_QUALIFICACAO WHERE TPQU_CD_TIPO_QUALIFICACAO=" & rs3("TPQU_CD_TIPO_QUALIFICACAO"))
        VALOR=TEMP("TPQU_TX_DESC_TIPO_QUALIFICACAO")
        %>
        <option value="<%=rs3("TPQU_CD_TIPO_QUALIFICACAO")%>"><%=VALOR%></option>
        <%
        rs3.movenext
        loop
        %>
        </select>
        </p>
    </td>
  </tr>
  <tr>
    <td align="left" height="26" width="28"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image16" border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a> 
    </td>
  </tr>
  <tr>
    <td align="center" height="21" width="28"></td>
  </tr>
  <tr>
    <td align="center" height="26" width="28"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24" align="left"></a></td>
  </tr>
  <tr>
    <td align="center" height="15" width="28"></td>
  </tr>
</table>
  <table border="0">
  <tr>
    <td width="183" rowspan="5" height="123">
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>Público
              Principal&nbsp;/ Depto&nbsp;</b></font>
      <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b> ( Opcional )
      <img border="0" src="../../imagens/novo_registro_02.gif" align="absmiddle"></b></font></td>
    <td rowspan="5" height="123" width="334"> 
        <p style="margin-top: 0; margin-bottom: 0" align="right"> 
        <select size="6" name="list3" multiple>
           <%do until rs2.eof=true
           SET RSTEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_PUB_PRI WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "' AND TPPP_CD_TIPO_PUB_PRINCIPAL=" & rs2("TPPP_CD_TIPO_PUB_PRINCIPAL"))
           if rstemp.eof=true then
           %>
           <option value="<%=rs2("TPPP_CD_TIPO_PUB_PRINCIPAL")%>"><%=rs2("TPPP_TX_DESC_PUB_PRINCIPAL")%></option>
           <%
           end if
           rs2.movenext
           loop
           %>
           </select>
        </p>
    </td>
    <td align="center" height="21" width="28">
      <p style="margin-top: 0; margin-bottom: 0"></td>
    <td width="39" rowspan="5" height="123"> 
        <p style="margin-top: 0; margin-bottom: 0" align="left"> 
        <select size="6" name="list4" multiple>
           <%do until rs4.eof=true
          SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TIPO_PUBLICO_PRINCIPAL WHERE TPPP_CD_TIPO_PUB_PRINCIPAL=" & rs4("TPPP_CD_TIPO_PUB_PRINCIPAL"))
	       VALOR=TEMP("TPPP_TX_DESC_PUB_PRINCIPAL")
           %>
           <option value="<%=rs4("TPPP_CD_TIPO_PUB_PRINCIPAL")%>"><%=VALOR%></option>
           <%
           rs4.movenext
           loop
           %>
        &nbsp;
        </select>
        </p>
    </td>
  </tr>
  <tr>
    <td align="center" height="26" width="28"> 
    <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list3,document.frm1.list4,1)"><img name="Image16" border="0" src="../../imagens/continua_F01.gif" width="24" height="24" align="left"></a> 
    </td>
  </tr>
  <tr>
    <td align="center" height="21" width="28"></td>
  </tr>
  <tr>
    <td align="center" height="26" width="28"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list4,document.frm1.list3,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24" align="left"></a></td>
  </tr>
  <tr>
    <td align="center" height="13" width="28"></td>
  </tr>
</table>
  </form>

</body>

</html>
