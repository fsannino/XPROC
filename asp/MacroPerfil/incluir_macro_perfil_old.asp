 
<%

if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = 0
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

set rs_mega=db.execute(str_SQL_MegaProc)

str_SQL_Fun_Neg = ""
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " SELECT DISTINCT " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO, " & Session("PREFIXO") & "FUN_NEG_TRANSACAO " 
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " WHERE " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso  
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ORDER BY " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
'RESPONSE.WRITE str_SQL_Fun_Neg
set rs1=db.execute(str_SQL_Fun_Neg)

'***********************************
set rs=db.execute("SELECT MEPR_TX_ABREVIA, MEPR_TX_DESC_MEGA_PROCESSO FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso )
if not rs.eof then
   str_PrefixoNomeTecnico = "Z:" & Trim(rs("MEPR_TX_ABREVIA")) & "_PB000_"
else
   str_PrefixoNomeTecnico = ""
end if

rs.CLOSE
SET rs = NOTHING

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">
<!--
function manda1()
{
window.location.href='incluir_macro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function carrega_txt1(fbox) 
{
document.frm1.txtFuncSelec.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtFuncSelec.value = document.frm1.txtFuncSelec.value + "," + fbox.options[i].value;
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

if(document.frm1.txtNomeTecnico.value == "")
{
alert("É obrigatória a especificação do NOME TÉCNICO!");
document.frm1.txtNomeTecnico.focus();
return;
}

if(document.frm1.selFuncPrinc.selectedIndex == 0)
{
alert("É obrigatória a seleção de uma Função!");
document.frm1.selFuncPrinc.focus();
return;
}
if(document.frm1.txtDescMacroPerfil.value == "")
{
alert("É obrigatória a especificação da DESCRIÇÃO DO MACROPERFIL!");
document.frm1.txtDescMacroPerfil.focus();
return;
}
//if (document.frm1.list2.options.length == 0)
//{ 
//alert("A seleção de pelo menos uma FUNÇÃO é obrigatória !");
//document.frm1.list2.focus();
//return;
//}
//if (document.frm1.list2.options.length > 1)
//{ 
//alert("Somente uma FUNÇÃO deve ser selecionada !");
//document.frm1.list2.focus();
//return;
//}
else
{
carrega_txt1(document.frm1.list2)
document.frm1.submit();
}
}

function pega_tamanho()
{
valor=document.frm1.txtDescMacroPerfil.value.length;
document.frm1.txttamanho.value=valor
if (valor > 61) {
	str1=document.frm1.txtDescMacroPerfil.value;
	str2=str1.slice(0,61);
	document.frm1.txtDescMacroPerfil.value=str2;
	valor=str2.length;
	document.frm1.txttamanho.value=valor;
}
}


</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif');pega_tamanho()">
<form method="POST" action="grava_macro_perfil.asp" name="frm1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20"> 
        <table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
            <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
            <td width="26">&nbsp;</td>
            <td width="195"></td>
            <td width="27"></td>
            <td width="50"></td>
            <td width="28"></td>
            <td width="26">&nbsp;</td>
            <td width="159"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <div align="center"><font face="Verdana" color="#330099" size="3">Inclus&atilde;o 
          de Macro Perfil</font></div>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="764" height="164">
    <tr> 
      <td width="66" height="25"></td>
      <td width="156" height="25" valign="top">&nbsp;</td>
      <td width="341" height="25">&nbsp; </td>
      <td width="287" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="66" height="24"></td>
      <td width="156" height="24" valign="top"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          : </b></font></div>
      </td>
      <td width="341" height="24"> 
        <select size="1" name="selMegaProcesso" onChange="javascript:manda1()">
          <option value="0">== Selecione o Mega-Processo ==</option>
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
        </select>
      </td>
      <td width="287" height="24"> 
        <p align="left">&nbsp; 
      </td>
    </tr>
    <% 'If str_MegaProcesso = 11 then
	if InStrRev("11/10", Right("00" & str_MegaProcesso, 2)) <> 0 then	 
	'if rs_mega("MEPR_CD_MEGA_PROCESSO") = 11 then
	%>
    <%
	end if	
	%>
    <tr> 
      <td width="66" height="23"></td>
      <td width="156" height="23"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Nome 
          T&eacute;cnico : </b></font><font face="Verdana" size="2" color="#330099"></font></div>
      </td>
      <td height="23" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099"><%=str_PrefixoNomeTecnico%></font> 
        <input type="hidden" name="txtPrefixoNomeTecnico" value="<%=str_PrefixoNomeTecnico%>">
        <input type="text" name="txtNomeTecnico" size="20" maxlength="19">
        <input type="hidden" name="txtAcao" value="C">
      </td>
    </tr>
    <tr> 
      <td width="66" height="25"></td>
      <td width="156" height="25" valign="top">
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          : </b></font></div>
      </td>
      <td height="25" valign="top" colspan="2"><b> 
        <select size="1" name="selFuncPrinc">
   		<option value="0">== Selecione uma Função ==</option>
          <%do until rs1.eof=true%>
          <option value="<%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs1("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
        rs1.movenext
        loop
        %>
        </select>
        </b></td>
    </tr>
    <tr> 
      <td width="66" height="83"></td>
      <td width="156" height="83" valign="top"> 
        <div align="right"><font face="Verdana" size="2" color="#330099"><b> </b></font> 
          <font face="Verdana" size="2" color="#330099"><b>Descrição : </b></font> 
          <input type="hidden" name="txtFuncSelec" size="20">
          <input type="hidden" name="txtImp" size="20">
        </div>
      </td>
      <td height="83" valign="top" colspan="2"> 
        <p align="left" style="margin-top: 0; margin-bottom: 0"> 
          <textarea rows="3" name="txtDescMacroPerfil" cols="49" onkeyup="javascript:pega_tamanho()"></textarea>
        <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="2" color="#330099"><b>Caracteres 
          digitados&nbsp; 
          <input type="text" name="txttamanho" size="5" value="0" maxlength="50">
          </b></font><font face="Verdana" color="#330099" size="1">(Máximo 61 
          caracteres)</font> 
      </td>
    </tr>
  </table>
  <table width="666" border="0" cellpadding="0" cellspacing="0" height="180">
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
        <div align="center"><font face="Verdana" size="2" color="#330099"><font color="#003366"><b>Fun&ccedil;&otilde;es 
          Similares</b></font></font></div>
      </td>
    </tr>
    <tr> 
      <td height="7" colspan="2"> 
        <div align="center"><font face="Verdana" size="2" color="#330099"><font color="#003366"></font></font></div>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="10"> 
        <table width="644" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="364"> 
              <div align="center"> <b> 
                <select size="6" name="list1" multiple>				  
                  <% rs1.movefirst
				  do until rs1.eof=true%>
                  <option value="<%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs1("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
                  <%
        rs1.movenext
        loop
        %>
                </select>
                </b></div>
            </td>
            <td width="26" align="center"> 
              <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
                <tr> 
                  <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1611','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="../Funcao/continua_F01.gif" width="24" height="24"></a></td>
                </tr>
                <tr> 
                  <td height="25">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="25"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img0151111','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="../Funcao/continua2_F01.gif" width="24" height="24"></a></td>
                </tr>
              </table>
            </td>
            <td width="242"> 
              <div align="center"> <font color="#000080"> 
                <select size="6" name="list2" multiple>
                </select>
                </font></div>
            </td>
            <td width="4">&nbsp;</td>
          </tr>
          <tr> 
            <td width="364"> </td>
            <td width="26" align="center"> </td>
            <td width="242"> </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td width="364"> </td>
            <td width="26" align="center"> </td>
            <td width="242"> </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td width="364"> </td>
            <td width="26" align="center"> </td>
            <td width="242"> </td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td colspan="3" width="636"> 
              <p style="margin-top: 0; margin-bottom: 0" align="center"> 
            </td>
            <td width="4">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3" width="636"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
                a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></div>
            </td>
            <td width="4">&nbsp;</td>
          </tr>
          <tr> 
            <td width="364"><font color="#000080"> </font></td>
            <td width="26" align="center">&nbsp;</td>
            <td width="242">&nbsp; </td>
            <td width="4">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="2">
    <tr> 
      <td width="351" height="1" bgcolor="#FFFFFF"></td>
      <td width="315" height="1" bgcolor="#FFFFFF"></td>
    </tr>
  </table>
</form>
</body>

</html>
