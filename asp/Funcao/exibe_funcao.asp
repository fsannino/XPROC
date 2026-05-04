 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_funcao=request("selFuncao")

set rs_func=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TIPO_QUALIFICACAO ORDER BY TPQU_TX_DESC_TIPO_QUALIFICACAO")
'set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "TIPO_PUBLICO_PRINCIPAL ORDER BY TPPP_TX_DESC_PUB_PRINCIPAL")

set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")
'set rs4=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_PUB_PRI WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

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
function mover()
{
window.moveTo(50,100)
}
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
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" onLoad="javascript:mover()">
<form method="POST" action="valida_altera_funcao.asp" name="frm1">
        
  <table width="529" border="0" cellspacing="0" cellpadding="0" height="37">
    <tr>
      <td width="527" height="19"></td>
    </tr>
    <tr>
      <td width="527" height="18">
        <p align="right"><b><font color="#800000" face="Verdana" size="2">&nbsp;</font><a href="javascript:window.close()"><font color="#800000" face="Verdana" size="2">Fechar
        Janela</font></a></b></p>
      </td>
    </tr>
  </table>
  <table border="0" width="532" height="136">
          <tr>
            
      <td width="19" height="25"></td>
            
      <td width="145" height="25" valign="middle"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo</b></font></td>
            
      <td width="356" height="25" valign="middle"> 
        <font face="Verdana" size="2" color="#330099"> 
                 	<%do until rs.eof=true
                	if trim(rs_func("MEPR_CD_MEGA_PROCESSO"))=trim(rs("MEPR_CD_MEGA_PROCESSO")) then
                	%>
              	<%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>					<%
					end if
					rs.movenext
					loop
					%>
              </tr>
          <tr>
            
      <td width="19" height="25"></td>
            
      <td width="145" height="25" valign="middle"><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o R/3</b></font></td>
            
      <td width="356" height="25" valign="middle"> 
        <font face="Verdana" size="2" color="#330099"><%=rs_func("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
          </tr>
          <tr>
            
      <td width="19" height="55"></td>
            
      <td width="145" height="55" valign="middle"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Descrição da</b></font></td>
          </tr>
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o R/3</b></font></td>
          </tr>
        </table>
        <p><input type="hidden" name="Funcao" size="20" value="<%=str_funcao%>"><input type="hidden" name="txtQua" size="20"><input type="hidden" name="txtpub" size="20"></p>
            </td>
            
      <td width="356" height="55" valign="middle"> 
        <p align="left" style="margin-top: 0; margin-bottom: 0">
          <font face="Verdana" size="2" color="#330099"><%=rs_func("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></font>
      </td>
          </tr>
        </table>
  <div align="left">
  <table width="533" border="0" cellpadding="0" cellspacing="0" align="left" height="144">
    <tr> 
      <td height="8" width="285"> 
        <div align="center"><font face="Verdana" size="2" color="#330099"><font color="#003366"><b>Qualificação 
          Não R/3</b></font></font></div>
      </td>
      <td height="8" width="325"> 
        <p align="center"><font face="Verdana" size="2" color="#330099"><b> 
            &nbsp;&nbsp;</b></font> 
      </td>
      <td height="8" width="3"> 
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="1" width="612"> 
        <div align="left">
        <table width="527" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="247" height="114"> 
              <div align="center"> 
                <p align="center"> <font color="#000080">
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
                </font></div>
            </td>
            <td width="276" height="114">
              <p align="center">&nbsp;<font color="#000080">
                </font></p>
            </td>
          </tr>
        </table>
        </div>
      </td>
      <td height="1" width="3"> 
      </td>
    </tr>
  </table>
  </div>
  </form>

</body>

</html>
