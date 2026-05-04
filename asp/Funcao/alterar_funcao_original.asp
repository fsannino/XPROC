 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_funcao=request("selFuncao")

set rs_func=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = rs_func("MEPR_CD_MEGA_PROCESSO")
end if

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TIPO_QUALIFICACAO ORDER BY TPQU_TX_DESC_TIPO_QUALIFICACAO")
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")

set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")
set rs4=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " " & Session("PREFIXO") & "SUB_MODULO.MEPR_CD_MEGA_PROCESSO,  "
str_Sub_Modulo = str_Sub_Modulo & " " & Session("PREFIXO") & "SUB_MODULO.SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " " & Session("PREFIXO") & "SUB_MODULO.SUMO_NR_SEQUENCIA, "
str_Sub_Modulo = str_Sub_Modulo & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO, " & Session("PREFIXO") & "MEGA_PROCESSO"
str_Sub_Modulo = str_Sub_Modulo & " WHERE " & Session("PREFIXO") & "SUB_MODULO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_Sub_Modulo = str_Sub_Modulo & " and " & Session("PREFIXO") & "SUB_MODULO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_Sub_Modulo = str_Sub_Modulo & " order by " & Session("PREFIXO") & "SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "
set rs_SubModulo=db.execute(str_Sub_Modulo)

'response.write " tp fun_neg "
'response.write rs_func("FUNE_TX_TP_FUN_NEG")

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
function carrega_txt2(fbox) 
{
document.frm1.txtImp.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtImp.value = document.frm1.txtImp.value + "," + fbox.options[i].value;
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

<script language="javascript" src="../js/troca_lista.js"></script>

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
else
{
carrega_txt1(document.frm1.list2)
carrega_txt2(document.frm1.list4)

document.frm1.submit();
}
}

function pega_tamanho()
{
valor=document.frm1.txtdescfuncao.value.length;
document.frm1.txttamanho.value=valor
if (valor > 500) {
	str1=document.frm1.txtdescfuncao.value;
	str2=str1.slice(0,500);
	document.frm1.txtdescfuncao.value=str2;
	valor=str2.length;
	document.frm1.txttamanho.value=valor;
}
}
</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif');pega_tamanho()">
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
        
  <table width="750" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="748">
        &nbsp;
        <div align="center"><font face="Verdana" color="#330099" size="3">Alteração 
          de Fun&ccedil;&atilde;o R/3 - <b><%=str_funcao%> 
          <input type="hidden" name="selFuncao" value="<%=str_funcao%>">
          </b></font></div>
      </td>
    </tr>
    <tr>
      <td width="748">&nbsp; </td>
    </tr>
  </table>
  <table border="0" width="749" height="164">
    <tr> 
      <td width="72" height="25"></td>
      <td width="150" height="25" valign="top">&nbsp;</td>
      <td width="367" height="25"> <font size="2"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
        <font color="#330099">
        <input type="hidden" name="selMegaProcesso" value="<%=str_MegaProcesso%>">
        <%'=rs_SubModulo("MEPR_TX_DESC_MEGA_PROCESSO")%> </font></font></font></td>
      <td width="146" height="25"> 
        <p align="right"><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o 
          Gen&eacute;rica: 
          <%if rs_func("FUNE_TX_TP_FUN_NEG")="G" then%>
          <input type="checkbox" name="selGenerica" value="1" checked>
          <%else%>
          <input type="checkbox" name="selGenerica" value="0">
          <%end if%>
          </b></font> 
      </td>
    </tr>
    <% If str_MegaProcesso = 11 or str_MegaProcesso = 10 then	 
	%>

    <tr>
      <td width="72" height="25"></td>
      <td width="150" height="25" valign="top">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Sub-Modulo</b></font></td>
          </tr>
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><font size="1">(Este 
              campo somente ser&aacute; utilizado para facilitar a consulta. &Eacute; 
              mais uma forma de filtro.) - FACULTATIVO</font></font></td>
          </tr>
        </table>
      </td>
      <td width="513" height="25" colspan="2"> 
        <select size="1" name="selSubModulo">
          <option value="0">== Selecione o Sub Módulo ==</option>
          <%do until rs_SubModulo.eof=true
		  if trim(rs_func("SUMO_NR_SEQUENCIA"))=trim(rs_SubModulo("SUMO_NR_SEQUENCIA")) then
		  %>
          <option selected value="<%=rs_SubModulo("SUMO_NR_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <% else %>
          <option value="<%=rs_SubModulo("SUMO_NR_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
		            end if
					rs_SubModulo.movenext
					loop
					%>
        </select>
      </td>
    </tr>
	<% end if %>
    <tr> 
      <td width="72" height="25"></td>
      <td width="150" height="25" valign="top"><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o R/3</b></font></td>
      <td width="513" height="25" colspan="2"> 
        <input type="text" name="txtfuncao" size="58" value="<%=rs_func("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%>" maxlength="50">
      </td>
    </tr>
    <tr> 
      <td width="72" height="83"></td>
      <td width="150" height="83" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Descrição da</b></font></td>
          </tr>
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o R/3</b></font></td>
          </tr>
        </table>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0"> 
          <input type="hidden" name="Funcao" size="20" value="<%=str_funcao%>">
          <input type="hidden" name="txtQua" size="20">
          <input type="hidden" name="txtImp" size="20">
        </p>
      </td>
      <td width="513" height="83" valign="top" colspan="2"> 
        <p align="left" style="margin-top: 0; margin-bottom: 0"> 
          <textarea rows="4" name="txtdescfuncao" cols="59" onkeydown="javascript:pega_tamanho()"><%=rs_func("FUNE_TX_DESC_FUNCAO_NEGOCIO")%></textarea>
        <p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp; 
        <p align="left" style="margin-top: 0; margin-bottom: 0"> <font face="Verdana" size="2" color="#330099"><b>Caracteres 
          digitados&nbsp; 
          <input type="text" name="txttamanho" size="5" value="0" maxlength="50">
          </b></font><font face="Verdana" color="#330099" size="1">(Máximo 500 
          caracteres)</font> 
      </td>
    </tr>
  </table>
  <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="180">
    <tr> 
      <td width="351" height="4" bgcolor="#0099CC"></td>
      <td width="315" height="4" bgcolor="#0099CC"></td>
    </tr>
    <tr> 
      <td height="7" width="351">&nbsp;</td>
      <td height="7" width="315">&nbsp;</td>
    </tr>
    <tr> 
      <td height="7" colspan="2"> 
        <div align="center"><font face="Verdana" size="2" color="#330099"><font color="#003366"><b>Qualificação 
          Não R/3</b></font></font></div>
      </td>
    </tr>
    <tr> 
      <td height="7" colspan="2"> 
        <div align="center"><font face="Verdana" size="2" color="#330099"><font color="#003366"></font></font></div>
      </td>
    </tr>
    <tr> 
      <td colspan="2" height="10"> 
        <table width="616" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="266"> 
              <div align="center"> <b>
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
                </b></div>
            </td>
            <td width="24" align="center"> 
              <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
                <tr> 
                  <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image161','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image161" border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a></td>
                </tr>
                <tr> 
                  <td height="25">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="25"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img015111','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img015111" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24" align="left"></a></td>
                </tr>
              </table>
            </td>
            <td width="290"> 
              <div align="center"> <font color="#000080">
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
            <td width="1">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3"></td>
            <td width="1"></td>
          </tr>
          <tr>
            <td width="636" colspan="4"> 
              <p style="margin-top: 0; margin-bottom: 0" align="center"><b><font color="#003366" face="Verdana" size="2">&Aacute;rea 
                de abrangência</font></b> 
            </td>
          </tr>
          <tr>
            <td width="364"> 
            </td>
            <td width="26" align="center"> 
            </td>
            <td width="242"> 
            </td>
            <td width="4"></td>
          </tr>
          <tr>
            <td width="364"> 
            </td>
            <td width="26" align="center"> 
            </td>
            <td width="242"> 
            </td>
            <td width="4"></td>
          </tr>
          <tr>
            <td width="364"> 
              <div align="center"> 
                <select size="9" name="list3" multiple>
                  <% A = "NÃO PASOU"
				  do until rs2.eof=true
                    SET RSTEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "' AND AGLU_CD_AGLUTINADO='" & rs2("AGLU_CD_AGLUTINADO") & "'")
					    A = "PASSOU"
        				if rstemp.eof=true then
                    %>
                    <option value="<%=rs2("AGLU_CD_AGLUTINADO")%>"><%=rs2("AGLU_SG_AGLUTINADO")%></option>
                    <%
                    END IF
           			rs2.movenext
           			loop
                    %>
                  </select></div>
            </td>
            <td width="26" align="center"> 
              <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
                <tr> 
                  <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list3,document.frm1.list4,1)"><img name="Image16" border="0" src="../../imagens/continua_F01.gif" width="24" height="24" align="left"></a></td>
                </tr>
                <tr> 
                  <td height="25">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="25"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list4,document.frm1.list3,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24" align="left"></a></td>
                </tr>
              </table>
            </td>
            <td width="242"> 
              <div align="center"> 
                <select size="9" name="list4" multiple>
                  <%do until rs4.eof=true
        				SET TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO ='" & rs4("AGLU_CD_AGLUTINADO") & "'")
			          VALOR=TEMP("AGLU_SG_AGLUTINADO")
        				%>
                  <option value="<%=rs4("AGLU_CD_AGLUTINADO")%>"><%=VALOR%></option>
                  <%
        			rs4.movenext
        			loop
        			%>
                  
                  </select></div>
            </td>
            <td width="4">&nbsp;</td>
          </tr>
          <tr>
            <td colspan="3" width="636">
              <p style="margin-top: 0; margin-bottom: 0" align="center">
            </td>
            <td width="4">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3"></td>
            <td width="1"></td>
          </tr>
          <tr> 
            <td colspan="3">&nbsp;
              <p>&nbsp;
            </td>
            <td width="1">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
                a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></div>
            </td>
            <td width="1">&nbsp;</td>
          </tr>
          <tr> 
            <td width="266"><font color="#000080">&nbsp; </font></td>
            <td width="24" align="center">&nbsp;</td>
            <td width="290">&nbsp; </td>
            <td width="1">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="42">
    <tr> 
      <td width="351" height="1" bgcolor="#0099CC"></td>
      <td width="315" height="1" bgcolor="#0099CC"></td>
    </tr>
    <tr> 
      <td height="7" width="351">&nbsp;</td>
      <td height="7" width="315">&nbsp;</td>
    </tr>
  </table>
  </form>

</body>

</html>