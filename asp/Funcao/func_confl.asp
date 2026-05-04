<%
if request("selMegaProcesso") <> "" then
   str_MegaProcesso=request("selMegaProcesso")
else
   str_megaprocesso=0
end if
if request("selFuncao") <> "" then
   str_Funcao=request("selFuncao")
else
   str_Funcao = 0
end if
if request("selSubModulo") <> "" then
   str_SubModulo=request("selSubModulo")
else
   str_SubModulo = 0
end if

'response.Write(str_megaprocesso)
'response.Write(str_SubModulo)
'response.Write(str_Funcao)

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
set rs_mega=db.execute(str_SQL_MegaProc)

IF str_SubModulo <> 0 THEN
	ssql=""
	ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
	ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
	ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
	ssql=ssql & "WHERE FUNCAO_NEGOCIO. MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso 
    ssql=ssql & " and FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA =" & str_SubModulo
	ssql=ssql+" ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "
else
	ssql=""
	ssql="SELECT DISTINCT FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	ssql=ssql+"FROM FUNCAO_NEGOCIO "
	ssql=ssql & "WHERE FUNCAO_NEGOCIO. MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso 
	ssql=ssql+" ORDER BY FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
end if

if str_MegaProcesso<>0 then
	set rs_funcao=db.execute(ssql)
end if

set rs_origem=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO ORDER BY FUNE_CD_FUNCAO_NEGOCIO")

set rs_destino=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_CONFL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " MEPR_CD_MEGA_PROCESSO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo & " WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_CD_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "
'response.write str_Sub_Modulo
set rs_SubModulo=db.execute(str_Sub_Modulo)
if rs_SubModulo.eof then
   'response.Write(" fimmmmmm  ")
end if

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">
<!-- 
function manda1()
{
window.location.href='func_confl.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value
}
function manda2()
{
window.location.href='func_confl.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&selFuncao='+document.frm1.selFuncao.value
}
function manda3()
{
window.location.href='func_confl.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value
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

if(document.frm1.selFuncao.selectedIndex == 0)
{
alert("É obrigatória a seleção de uma FUNÇÃO R/3!");
document.frm1.selFuncao.focus();
return;
}
else
{
carrega_txt2(document.frm1.list2)
document.frm1.submit();
}
}

</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif')">
<form method="POST" action="valida_func_confl.asp" name="frm1">
          <input type="hidden" name="txtImp" size="20">
          <input type="hidden" name="txtFuncSelec" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
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
            <td width="27"></td>  <td width="50"></td>
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
        <div align="center"><font face="Verdana" color="#330099">Função R/3 Conflitante</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="924" height="91">
    <tr> 
      <td width="18" height="25"></td>
      <td width="131" height="25" valign="top">&nbsp;</td>
      <td width="622" height="25">&nbsp; </td>
      <td width="135" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="131" height="1" valign="top"> <div align="right"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
          : </b></font></div></td>
      <td width="622" height="1"> <select size="1" name="selMegaProcesso" onChange="javascript:manda1()">
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
        </select> </td>
      <td width="135" height="1"> <p align="left"> </td>
    </tr>
    <tr>
      <td height="23"></td>
      <td height="23"><div align="right"><font face="Verdana" size="2" color="#330099"><b>Assunto 
          :</b></font></div></td>
      <td height="23" colspan="2"><select size="1" name="selSubModulo" onChange="javascript:manda3()">
          <option value="0">== Selecione o Assunto ==</option>
          <%do until rs_SubModulo.eof=true
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
      <td width="18" height="23"></td>
      <td width="131" height="23"> <p align="right"><font size="2" color="#330099"><font face="Verdana"><b> 
          Função R/3 :</b></font></font></td>
      <td height="23" colspan="2"><select size="1" name="selFuncao" onChange="javascript:manda2()">
          <option value="0">== Selecione a Funcao de Negócio ==</option>
          <%do until rs_funcao.eof=true
      if trim(str_funcao)=trim(rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")) then
      %>
          <option selected value=<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs_funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%ELSE%>
          <option value=<%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>><%=rs_funcao("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs_funcao("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
      END IF
      RS_funcao.MOVENEXT
      LOOP
      %>
        </select></td>
    </tr>
  </table>
  <table width="904" border="0" cellpadding="0" cellspacing="0" height="180">
    <tr> 
      <td width="351" height="4" bgcolor="#0099CC"></td>
      <td width="553" height="4" bgcolor="#0099CC"></td>
    </tr>
    <tr> 
      <td height="7" width="351">&nbsp;</td>
      <td height="7" width="553">&nbsp;</td>
    </tr>
    <tr> 
      <td height="7" colspan="2" width="904"> 
        <div align="center"><b><font face="Verdana" size="2" color="#330099">Fun&ccedil;ão(&otilde;es)
          R/3 Conflitantes </font></b><font face="Verdana" color="#330099" size="1">(Todos
          os Mega-Processos)</font></div>
      </td>
    </tr>
    <tr> 
      <td height="7" colspan="2" width="904"> 
        <div align="center"><font face="Verdana" size="2" color="#330099"><font color="#003366"></font></font></div>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="10" width="904"> 
        <table width="671" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="364"> 
              <div align="center"> <b> 
                <select size="6" name="list1" multiple>
              	<%DO UNTIL RS_ORIGEM.EOF=TRUE%>
       	      <option value=<%=rs_origem("FUNE_CD_FUNCAO_NEGOCIO")%>><%=rs_origem("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=rs_origem("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
				<%
				RS_ORIGEM.MOVENEXT
				LOOP
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
               	<%DO UNTIL RS_destino.EOF=TRUE
               	set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs_destino("FUNC_CD_FUNCAO_CONFL") & "'")
               	VALOR=TEMP("FUNE_TX_TITULO_FUNCAO_NEGOCIO")               	
               	%>
        	      <option value=<%=rs_destino("FUNC_CD_FUNCAO_CONFL")%>><%=rs_destino("FUNC_CD_FUNCAO_CONFL")%>-<%=valor%></option>
				<%
				RS_destino.MOVENEXT
				LOOP
				%>

                </select>
                </font></div>
            </td>
            <td width="31">&nbsp;</td>
          </tr>
          <tr> 
            <td width="364"> </td>
            <td width="26" align="center"> </td>
            <td width="242"> </td>
            <td width="31"></td>
          </tr>
          <tr> 
            <td width="364"> </td>
            <td width="26" align="center"> </td>
            <td width="242"> </td>
            <td width="31"></td>
          </tr>
          <tr> 
            <td width="364"> </td>
            <td width="26" align="center"> </td>
            <td width="242"> </td>
            <td width="31"></td>
          </tr>
          <tr> 
            <td colspan="3" width="636"> 
              <p style="margin-top: 0; margin-bottom: 0" align="center"> 
            </td>
            <td width="31">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3" width="636"> 
              <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use 
                a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o 
                ou para desmarcar um item selecionado.</font></div>
            </td>
            <td width="31">&nbsp;</td>
          </tr>
          <tr> 
            <td width="364"><font color="#000080">&nbsp; </font></td>
            <td width="26" align="center">&nbsp;</td>
            <td width="242">&nbsp; </td>
            <td width="31">&nbsp;</td>
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