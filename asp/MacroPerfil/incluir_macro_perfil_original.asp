<%

if request("selMegaProcesso") <> "" then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = "0"
end if
if request("selSubModulo") <> "" then
   str_SubModulo = request("selSubModulo")
else
   str_SubModulo = "0"
end if

'response.Write(str_SubModulo)
'response.Write("  a1ui  ")
'response.Write(str_MegaProcesso)
if InStrRev("11/10", Right("00" & str_MegaProcesso, 2)) = 0 then
   str_SubModulo = 0
end if
'response.Write(str_SubModulo)
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
'str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO, " & Session("PREFIXO") & "FUN_NEG_TRANSACAO " 
'str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO INNER JOIN"
'str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUN_NEG_TRANSACAO ON "
'str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO"
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " FROM FUNCAO_NEGOCIO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " INNER JOIN FUN_NEG_TRANSACAO ON FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO "
'str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " LEFT OUTER JOIN MACRO_PERFIL ON FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " WHERE "
' & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " " & Session("PREFIXO") & "FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_INDICA_REFERENCIADA = 0 "
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND MACRO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO IS NULL "
if str_SubModulo <> "0" then
   str_SQL_Fun_Neg = str_SQL_Fun_Neg & " AND FUNCAO_NEGOCIO.SUMO_NR_CD_SEQUENCIA = " & str_SubModulo
end if
str_SQL_Fun_Neg = str_SQL_Fun_Neg & " ORDER BY " & Session("PREFIXO") & "FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
'response.Write(str_SQL_Fun_Neg)
'response.Write("           aaaaaaaaaaaa            ")
'RESPONSE.WRITE str_SQL_Fun_Neg
'set rs1=db.execute(str_SQL_Fun_Neg)

'***********************************
if str_MegaProcesso <> 15 then
   if str_SubModulo <> "0" then
      str_SQL = ""
      str_SQL = str_SQL & " Select SUMO_TX_ABREV "
      str_SQL = str_SQL & " from " & Session("PREFIXO") & "SUB_MODULO "
      str_SQL = str_SQL & " where SUMO_NR_CD_SEQUENCIA = " & str_SubModulo
      str_SQL = str_SQL & "  " 'and MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & str_MegaProcesso & "%'"
      'response.Write("<p>" & str_SQL)
      set rsSubMod = db.Execute(str_SQL)   
      ls_meio = "_" & Trim(rsSubMod("SUMO_TX_ABREV"))
      rsSubMod.close
      set rsSubMod = Nothing
   else
      ls_meio = ""
   end if
else
   ls_meio = ""   
END IF
if str_MegaProcesso <> 15 then
   set rs=db.execute("SELECT MEPR_TX_ABREVIA, MEPR_TX_DESC_MEGA_PROCESSO FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso )
   if not rs.eof then
      str_PrefixoNomeTecnico = "Z:" & Trim(rs("MEPR_TX_ABREVIA")) & ls_meio & "_PB000_"
   else
      str_PrefixoNomeTecnico = ""
   end if
   rs.CLOSE
   SET rs = NOTHING
else
   str_PrefixoNomeTecnico = "Z:BW_C_APL_PB_"
end if

'str_Sub_Modulo = ""
'str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
'str_Sub_Modulo = str_Sub_Modulo & " MEPR_CD_MEGA_PROCESSO_TODOS, "
'str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
'str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_CD_SEQUENCIA"
'str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
'str_Sub_Modulo = str_Sub_Modulo & " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '" & str_MegaProcesso & "%'"
'str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "
'response.write str_Sub_Modulo
'set rs_SubModulo=db.execute(str_Sub_Modulo)

SQL_Assunto=""
SQL_Assunto = SQL_Assunto & " SELECT SUMO_NR_CD_SEQUENCIA"
SQL_Assunto = SQL_Assunto & " ,SUMO_TX_DESC_SUB_MODULO"
SQL_Assunto = SQL_Assunto & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
SQL_Assunto = SQL_Assunto & " FROM " & Session("PREFIXO") & "SUB_MODULO"
if str_MegaProcesso <> 0 then
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
else
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS = '9999'"
end if
SQL_Assunto=SQL_Assunto + " ORDER BY SUMO_TX_DESC_SUB_MODULO"
'response.write "<p>" & SQL_Assunto
set rs_SubModulo=db.execute(SQL_Assunto)


%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">
<!--
function manda1()
{
//alert(document.frm1.selSubModulo.selectedIndex)
//alert(document.frm1.selSubModulo.options[document.frm1.selSubModulo.selectedIndex].value)
window.location.href='incluir_macro_perfil.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value
    
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
//carrega_txt1(document.frm1.list2)
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
          de Macro Perfil - 2</font></div>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="938" height="242">
    <tr> 
      <td width="27" height="25"></td>
      <td width="151" height="25" valign="top">&nbsp;</td>
      <td width="429" height="25">&nbsp; </td>
      <td width="313" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="27" height="24"></td>
      <td width="151" height="24" valign="top"> <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          : </b></font></div></td>
      <td width="429" height="24"> <select size="1" name="selMegaProcesso" onChange="javascript:manda1()">
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
        <font face="Verdana, Arial, Helvetica, sans-serif" size="2">&nbsp; </font> </td>
      <td width="313" height="24"> <p align="left">&nbsp; </td>
    </tr>
    <tr> 
      <td height="23"></td>
      <td height="23"><div align="right"><font face="Verdana" size="2" color="#330099"><b>Assunto 
          :</b></font></div></td>
      <td height="23" colspan="2"><select size="1" name="selSubModulo" onChange="javascript:manda1()">
          <option value="0">== Selecione o Sub Módulo ==</option>
		  <% 		  
            do until rs_SubModulo.eof=true		  
		  if trim(str_SubModulo)=trim(rs_SubModulo("SUMO_NR_CD_SEQUENCIA")) then
		  %>
          <option selected value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <% else %>
          <option value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
		     end if
					rs_SubModulo.movenext
					loop
					%>
        </select></td>
    </tr>
    <tr> 
      <td width="27" height="23"></td>
      <td width="151" height="23"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Nome 
          T&eacute;cnico : </b></font><font face="Verdana" size="2" color="#330099"></font></div></td>
      <td height="23" colspan="2"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#330099"><%=str_PrefixoNomeTecnico%></font> <input type="hidden" name="txtPrefixoNomeTecnico" value="<%=str_PrefixoNomeTecnico%>"> 
        <input type="text" name="txtNomeTecnico" size="20" maxlength="19"> <input type="hidden" name="txtAcao" value="C"> 
        <font face="Verdana" color="#330099" size="1">Máximo 19 caracteres</font> 
      </td>
    </tr>
    <tr> 
      <td width="27" height="25"></td>
      <td width="151" height="25" valign="top"> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Fun&ccedil;&atilde;o 
          : </b></font></div></td>
      <td height="25" valign="top" colspan="2"><b> 	  
	  <% 'response.Write("           000000000         ")
		  'response.Write(str_SQL_Fun_Neg) 
	  %>
        <select size="1" name="selFuncPrinc">
          <option value="0">== Selecione uma Função ==</option>
          <% 
		  set rs10=db.execute(str_SQL_Fun_Neg)
		  Do until Rs10.eof = true
		  %>
          <option value="<%=rs10("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs10("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs10("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
               rs10.movenext
            Loop
        %>
        </select>
        </b></td>
    </tr>
    <tr> 
      <td width="27" height="83"></td>
      <td width="151" height="83" valign="top"> <div align="right"><font face="Verdana" size="2" color="#330099"><b> 
          </b></font> <font face="Verdana" size="2" color="#330099"><b>Descrição 
          : </b></font> 
          <input type="hidden" name="txtFuncSelec" size="20">
          <input type="hidden" name="txtImp" size="20">
        </div></td>
      <td height="83" valign="top" colspan="2"> <p align="left" style="margin-top: 0; margin-bottom: 0"> 
          <textarea rows="3" name="txtDescMacroPerfil" cols="49" onkeyup="javascript:pega_tamanho()"></textarea>
        <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#330099">Caracteres 
          digitados</font><font face="Verdana" size="2" color="#330099"><b> 
          <input type="text" name="txttamanho" size="5" value="0" maxlength="50">
          </b></font><font face="Verdana" color="#330099" size="1">(Máximo 61 
          caracteres)</font> </td>
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
