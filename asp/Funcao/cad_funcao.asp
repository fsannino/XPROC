<%
'set fso = Server.CreateObject("Scripting.FileSystemObject")
'Caminho = Server.Mappath("../../Publico/Bloqueio/status.txt")
'set arquivo = fso.OpenTextFile(Caminho)
'situacao = arquivo.readline

'if situacao=1 then
'	block = "disabled"
'end if

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

str_Sub_Modulo = ""
str_Sub_Modulo = str_SQL_MegaProc & " SELECT DISTINCT "
str_Sub_Modulo = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_Sub_Modulo = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_Sub_Modulo = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
'str_Sub_Modulo = str_Sub_Modulo & " WHERE MEPR_CD_MEGA_PROCESSO =  (" & Session("AcessoUsuario") & ")"
str_Sub_Modulo = str_Sub_Modulo & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_CD_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_MegaProcesso,2) & "%'" 
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "

'response.Write(str_Sub_Modulo)
set rs_SubModulo=db.execute(str_Sub_Modulo)

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TIPO_QUALIFICACAO ORDER BY TPQU_TX_DESC_TIPO_QUALIFICACAO")
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")

'set pai=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_TX_INDICA_REFERENCIADA=0 AND MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " ORDER BY FUNE_TX_TITULO_FUNCAO_NEGOCIO")
set pai=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE (FUNE_CD_FUNCAO_NEGOCIO = FUNE_CD_FUNCAO_NEGOCIO_PAI) AND MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " ORDER BY FUNE_TX_TITULO_FUNCAO_NEGOCIO")

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--
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
</head>

<script language="JavaScript">
<!--

function manda1()
{
window.location.href='cad_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value
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

function carrega_txt3(fbox) 
{
document.frm1.txtAss.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtAss.value = document.frm1.txtAss.value + "," + fbox.options[i].value;
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
   if(document.frm1.listAss.options.length == 0 )
     {
     alert("É obrigatória a seleção de, pelo menos, um ASSUNTO!");
     document.frm1.selAssunto.focus();
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
	carrega_txt3(document.frm1.listAss)
    if(document.frm1.txtImp.value == "")
      {
      if(confirm("Se não for selecionado nenhuma Áre de Abrangência o sistema entenderá que serão relacionada a todas as áreas listadas. Deseja continuar ou volta para seleção?"))
        {
        document.frm1.submit();
        }
      }
	else
	  {
	  document.frm1.submit();
	  }	 
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

function verifica_valor()
{
if (document.frm1.selFuncaoPai.value!=0)
{
document.frm1.pai.value=1;
}
else
{
document.frm1.pai.value=0;
}
}

</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif')">
<form method="POST" action="valida_cad_funcao.asp" name="frm1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60"><font color="#FFFFFF" size="1" face="Arial Narrow">&nbsp;</font></td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20"> 
        <table width="625" border="0" align="center">
          <tr> 
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="confirma_f02.gif"></a></td>
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
        <div align="center"><font face="Verdana" color="#330099" size="3">Cadastro 
          de Fun&ccedil;&atilde;o R/3</font></div>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
  <table width="818" height="65" border="0">
    <tr>
      <td width="19" height="24"></td>
      <td width="121" height="24" valign="middle"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo</b></font></td>
      <td height="24" colspan="3" valign="middle" width="452"><select size="1" name="selMegaProcesso" onChange="javascript:manda1()">
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
      </select></td>
      <td width="208" height="24" valign="middle">
        <p align="left"><font face="Verdana" size="2" color="#330099"><b>
      <input type="checkbox" name="selGenerica" value="1">
      Fun&ccedil;&atilde;o Gen&eacute;rica</b></font> </td>
    </tr>
    <tr>
      <td height="21" width="19">&nbsp;</td>
      <td height="21" valign="middle" width="121">&nbsp;</td>
      <td width="178" height="21" valign="middle">&nbsp;</td>
      <td width="67" height="21" valign="middle">&nbsp;</td>
      <td width="199" height="21" valign="middle">&nbsp;</td>
      <td height="21" valign="middle" width="208"><font face="Verdana" size="2" color="#330099"><b>
      <input name="chkEmUso" type="checkbox" value="1" checked>
      Fun&ccedil;&atilde;o em Uso</b></font></td>
    </tr>
    <tr>
      <td height="21" width="19"></td>
      <td height="21" valign="middle" width="121">&nbsp;</td>
      <td width="178" height="21" valign="middle"><div align="center"></div></td>
      <td width="67" height="21" valign="middle"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><font color="#31009C"></font></font></div></td>
      <td width="199" height="21" valign="middle"><div align="center"></div></td>
      <td height="21" valign="middle" width="208"><font face="Verdana" size="2" color="#330099"><b>
      <input name="chkAntec" type="checkbox" value="1" <%=block%>> Função </b></font><b><font face="Verdana" size="2" color="#330099">Antecipada</font></b></td>
    </tr>
    <%
    if block = "disabled" then
    %>
    <tr>
      <td height="21" width="19">&nbsp;</td>
      <td height="21" valign="middle" width="121">&nbsp;</td>
      <td width="178" height="21" valign="middle">&nbsp;</td>
      <td width="67" height="21" valign="middle">&nbsp;</td>
      <td width="199" height="21" valign="middle">&nbsp;</td>
      <td height="21" valign="middle" width="208"><font face="Verdana" size="1" color="#FF0000">As Funções Antecipadas não fazem parte do Escopo atual</font></td>
    </tr>
    <%end if%>
    </table>
  <table border="0" width="771" height="277">
    <% 'If str_MegaProcesso = 11 then
	'if InStrRev("11/10", Right("00" & str_MegaProcesso, 2)) <> 0 then	 
	'if rs_mega("MEPR_CD_MEGA_PROCESSO") = 11 then
	%><font size="1" color="#FF0000"> </font>
    <tr> 
      <td width="19" height="25" rowspan="5"></td>
      <td width="121" height="25" rowspan="5" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><font face="Verdana" size="2" color="#330099">&nbsp;</font></td>
          </tr>
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Assunto</b></font></td>
          </tr>
        </table></td>
      <td width="179" height="25" rowspan="5" align="center" valign="middle">        <p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assuntos Dispon&iacute;veis
              </font></p>
        <p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <select name="selAssunto" size="5" multiple>
            <%do until rs_SubModulo.eof=true%>
            <option value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
            <%
		rs_SubModulo.movenext
		loop
		%>
                </select>
        </font></p></td>
      <td width="66" height="21" align="center" valign="middle"><div align="left"></div></td>
      <td width="199" rowspan="5" align="center" valign="middle"><p align="left"><font color="#31009C" size="1" face="Verdana, Arial, Helvetica, sans-serif">Assuntos Selecionados</font></p>
        <p align="left">
            <select name="listAss" size="5" multiple id="listAss">
              <%do until rs_SubModulo.eof=true%>
              <%
					rs_SubModulo.movenext
					loop
					%>
                    </select>
        </p></td><td width="161" height="25" rowspan="5"><div align="left"><font face="Verdana" size="2" color="#330099"><b> 
        </b></font></div></td>
    </tr>
    <tr>
      <td height="26" align="center" valign="middle"><div align="center"></div></td>
    </tr>
    <tr>
      <td height="21" align="center" valign="middle"><div align="center"><img src="continua_F01.gif" width="24" height="24" onClick="move(document.frm1.selAssunto,document.frm1.listAss,1)"></div></td>
    </tr>
    <tr>
      <td height="26" align="center" valign="middle"><div align="center"><img src="continua2_F01.gif" width="24" height="24" onClick="move(document.frm1.listAss,document.frm1.selAssunto,1)"></div></td>
    </tr>
    <tr>
      <td height="21" align="center" valign="middle"><div align="left"></div></td>
    </tr>
    <%
	'end if	
	%>
    <tr> 
      <td width="19" height="25"></td>
      <td width="121" height="25" valign="top"><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o 
        R/3</b></font></td>
      <td height="25" colspan="4"> <div align="left">
        <input type="text" name="txtfuncao" size="58" maxlength="100"> 
      </div></td>
    </tr>
    <tr> 
      <td width="19" height="121"></td>
      <td width="121" height="121" valign="top"><font face="Verdana" size="2" color="#330099"><b> 
        </b></font> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Descrição da</b></font></td>
          </tr>
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><b>Fun&ccedil;&atilde;o 
              R/3</b></font></td>
          </tr>
        </table>
        <input type="hidden" name="txtQua" size="20"> 
		<input type="hidden" name="txtImp" size="20"> 
		<input type="hidden" name="txtAss" size="20"> 
      </td>
      <td valign="top" colspan="4"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td> <textarea rows="4" name="txtdescfuncao" cols="49" onKeyUp="javascript:pega_tamanho()"></textarea> 
            </td>
          </tr>
          <tr> 
            <td><font face="Verdana" size="2" color="#330099"><font size="1">Caracteres 
              digitados&nbsp;</font><b> 
              <input type="text" name="txttamanho" size="5" value="0" maxlength="50">
              </b></font><font face="Verdana" color="#330099" size="1">(Máximo 
              500 caracteres)</font> </td>
          </tr>
        </table></td>
    </tr>
  </table>
  <table width="644" border="0" align="center" cellpadding="0" cellspacing="0" height="219">
    <tr> 
      <td colspan="3"> 
        <div align="center"><font face="Verdana" size="2" color="#330099"><font color="#003366"><b>Qualificação 
          não R/3</b></font></font></div>
      </td>
      <td width="4">&nbsp;</td>
    </tr>
    <tr> 
      <td width="364"> 
        <div align="center"> <b> 
          <select size="6" name="list1" multiple>
            <%do until rs1.eof=true%>
            <option value="<%=rs1("TPQU_CD_TIPO_QUALIFICACAO")%>"><%=rs1("TPQU_TX_DESC_TIPO_QUALIFICACAO")%></option>
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
            <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1611','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="continua_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr> 
            <td height="25">&nbsp;</td>
          </tr>
          <tr> 
            <td height="25"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img0151111','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="continua2_F01.gif" width="24" height="24"></a></td>
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
      <td width="636" colspan="4"> 
        <p style="margin-top: 0; margin-bottom: 0" align="center"><b><font color="#003366" face="Verdana" size="2">&Aacute;rea 
          de abrangência</font></b> 
      </td>
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
      <td width="364"> 
        <div align="center"> 
          <select size="9" name="list3" multiple>
            <%do until rs2.eof=true%>
            <option value="<%=rs2("AGLU_CD_AGLUTINADO")%>"><%=rs2("AGLU_SG_AGLUTINADO")%></option>
            <%
           			rs2.movenext
           			loop
                    %>
          </select>
        </div>
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
          </select>
        </div>
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
      <td width="364"><font color="#000080">&nbsp; </font></td>
      <td width="26" align="center">&nbsp;</td>
      <td width="242">&nbsp; </td>
      <td width="4">&nbsp;</td>
    </tr>
  </table>
  <table width="666" border="0" cellpadding="0" cellspacing="0" align="center" height="2">
    <tr> 
      <td width="351" height="1" bgcolor="#FFFFFF"></td>
      <td width="315" height="1" bgcolor="#FFFFFF"></td>
    </tr>
  </table>
  <table border="0" cellpadding="0" cellspacing="0" width="724">
    <tr> 
      <td width="40" height="19" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="144" height="19" bgcolor="#D3D3D3">&nbsp;</td>
      <td width="644" height="19" bgcolor="#D3D3D3">&nbsp;</td>
    </tr>
    <tr> 
      <td width="40" height="19" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="144" height="19" bgcolor="#D3D3D3">&nbsp;</td>
      <td width="644" height="19" bgcolor="#D3D3D3"><font face="Verdana" size="1" color="#330099"><b>Se 
        você deseja criar funções que possuam as mesmas transações e devam estar 
        associadas a um único macro perfil, indique a Função de Referência no 
        campo abaixo:</b></font></td>
    </tr>
    <tr> 
      <td width="40" height="19" bgcolor="#FFFFFF"> 
        <input type="hidden" name="pai" size="6" value="0">
      </td>
      <td width="144" height="19" bgcolor="#D3D3D3"> 
        <div align="center"><font face="Verdana" size="2" color="#330099"><b>Função 
          de Referência</b></font> </div>
      </td>
      <td width="644" height="19" bgcolor="#D3D3D3"> 
        <select size="1" name="selFuncaoPai" onchange="javascript:verifica_valor()">
          <option value="0">== Selecione a Funcao de Referência ==</option>
          <%do until pai.eof=true%>
          <option value=<%=pai("FUNE_CD_FUNCAO_NEGOCIO")%>><%=pai("FUNE_CD_FUNCAO_NEGOCIO")%>-<%=pai("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
          pai.movenext
          loop
          %>
        </select>
      </td>
    </tr>
    <tr> 
      <td width="40" height="13" bgcolor="#FFFFFF">&nbsp;</td>
      <td width="144" height="13" bgcolor="#D3D3D3">&nbsp;</td>
      <td width="644" height="13" bgcolor="#D3D3D3"><font face="Verdana" size="1" color="#330099">(Este 
        campo somente ser&aacute; utilizado caso esta Função seja filha de uma 
        outra Função.) - FACULTATIVO</font> </td>
    </tr>
  </table>
  <p align="center"><img border="0" src="../../imagens/fluxograma.jpg"></p>
</form>
</body>

</html>