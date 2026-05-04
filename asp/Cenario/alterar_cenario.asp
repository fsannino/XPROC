<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=0
str_proc=0
str_Assunto=0

valor2=request("ID")

str_Assunto=request("selAssunto")

'response.write str_Assunto

if len(request("ID2"))>0 then
	valor2=request("ID2")
end if

IF LEN(VALOR2)<11 THEN
	RESPONSE.REDIRECT "ALTERA_CENARIO.ASP?sem=1"
END IF

SSQL30="select * from " & Session("PREFIXO") & "CENARIO where CENA_CD_CENARIO='" & valor2 & "'"
set cen_atual=db.execute(SSQL30)

IF cen_atual.eof=true THEN
	RESPONSE.REDIRECT "ALTERA_CENARIO.ASP?sem=1"
END IF

str_mega=request("selMegaProcesso")
str_proc=request("selProcesso")
str_sub=request("selSubProcesso")
str_onda=request("selOnda")

IF str_mega=0 or str_proc=0 then
	str_mega=cen_atual("MEPR_CD_MEGA_PROCESSO")
	str_proc=cen_atual("PROC_CD_PROCESSO")
	str_sub=cen_atual("SUPR_CD_SUB_PROCESSO")
	str_onda=cen_atual("ONDA_CD_ONDA")
end if
str_Assunto = cen_atual("SUMO_NR_CD_SEQUENCIA")

if session("MegaProcesso")<>0 and str_mega=0 then
	str_mega=session("MegaProcesso")
end if

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " WHERE MEPR_CD_MEGA_PROCESSO IN (" & Session("AcessoUsuario") & ")"
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
set rs_mega=db.execute(str_SQL_MegaProc)
set rs_onda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA ORDER BY ONDA_TX_DESC_ONDA")

if str_mega<>0 then
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " ORDER BY PROC_TX_DESC_PROCESSO")
	set rs_class=db.execute("SELECT * FROM " & Session("PREFIXO") & "CLASSE_CENARIO_MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
else
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY PROC_TX_DESC_PROCESSO")
	set rs_class=db.execute("SELECT * FROM " & Session("PREFIXO") & "CLASSE_CENARIO_MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega<>0 and str_proc<>0 then
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO")
else
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY SUPR_TX_DESC_SUB_PROCESSO ")
end if

if str_mega<>0 and str_proc=0 then
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " ORDER BY PROC_TX_DESC_PROCESSO")
end if

SQL_Assunto=""
SQL_Assunto = SQL_Assunto & " SELECT SUMO_NR_CD_SEQUENCIA"
SQL_Assunto = SQL_Assunto & " ,SUMO_TX_DESC_SUB_MODULO"
SQL_Assunto = SQL_Assunto & " ,MEPR_CD_MEGA_PROCESSO_TODOS "
SQL_Assunto = SQL_Assunto & " FROM " & Session("PREFIXO") & "SUB_MODULO"
if str_mega <> 0 then
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_Mega,2) & "%'" 
else
	SQL_Assunto=SQL_Assunto + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS = '9999'"
end if
SQL_Assunto=SQL_Assunto + " ORDER BY SUMO_TX_DESC_SUB_MODULO"

set rs_assunto=db.execute(SQL_Assunto)
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

function carrega_txt(fbox) {
document.frm1.txtEmpresa.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtEmpresa.value = document.frm1.txtEmpresa.value + "," + fbox.options[i].value;
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
function enviar()
{
window.location.href='alterar_cenario.asp?ID='+document.frm1.ID.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value
}

function Confirma()
{
if(document.frm1.selProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um PROCESSO!");
document.frm1.selProcesso.focus();
return;
}
if(document.frm1.selAssunto.selectedIndex == 0)
{
alert("É obrigatória a seleção de um Assunto!");
document.frm1.selAssunto.focus();
return;
}
if(document.frm1.selSubProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um SUB-PROCESSO!");
document.frm1.selSubProcesso.focus();
return;
}
if(document.frm1.selClasse.selectedIndex == 0)
{
alert("É obrigatória a seleção da CLASSE DO CENÁRIO!");
document.frm1.selClasse.focus();
return;
}
if(document.frm1.list2.options.length == 0)
{
alert("É obrigatória a seleção de pelo menos uma EMPRESA RELACIONADA!");
document.frm1.list1.focus();
return;
}
if(document.frm1.txtTitulo.value == "")
{
alert("É obrigatório o preenchimento do TÍTULO DO CENÁRIO!");
document.frm1.txtTitulo.focus();
return;
}
if(document.frm1.txtDescricao.value == "")
{
alert("É obrigatório o preenchimento da DESCRIÇÃO DO CENÁRIO!");
document.frm1.txtDescricao.focus();
return;
}
else
{
carrega_txt(document.frm1.list2);
document.frm1.submit();
}
}

</script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../imagens/continua_F02.gif','../../imagens/continua2_F02.gif')">
<form method="POST" action="valida_altera_cad_cenario.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
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
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"><img border="0" src="confirma_f02.gif" onclick="javascript:Confirma()"></td>
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
  <table width="88%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
      <td>
        <div align="center"><font face="Verdana" color="#330099" size="3">Alteração
          de Cenário - <%=valor2%></font></div>
      </td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="914" height="593" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="6" height="27"></td>
      <td width="386" height="27"><b><font face="Verdana" size="2" color="#330099">Selecione 
        o Mega-Processo&nbsp;</font></b></td>
      <td height="27" colspan="4" valign="baseline"><font face="Verdana" size="2" color="#330099">&nbsp;</font> 
        <input type="hidden" name="txtEmpresa" size="20"> </td>
    </tr>
    <tr> 
      <td height="27"></td>
      <td height="27"><select size="1" name="selMegaProcesso">
          <%do until rs_mega.eof=true
       if trim(str_mega)=trim(rs_mega("MEPR_CD_MEGA_PROCESSO")) then
       %>
          <option selected value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
		end if
		rs_mega.movenext
		loop
		%>
        </select></td>
      <td height="27" colspan="4" valign="baseline">&nbsp;</td>
    </tr>
    <tr> 
      <td height="27"></td>
      <td height="27"><b><font face="Verdana" size="2" color="#330099">Assunto 
        do Cenário</font></b> </td>
      <td height="27" colspan="4" valign="baseline">&nbsp;</td>
    </tr>
    <tr> 
      <td width="6" height="27"></td>
      <td width="386" height="27"><select size="1" name="selAssunto">
          <option value="0">Selecione um Assunto</option>
          <%do until rs_assunto.eof=true
          if trim(str_Assunto)=trim(rs_assunto("SUMO_NR_CD_SEQUENCIA")) then
          %>
          <option selected value="<%=rs_assunto("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_assunto("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%else%>
          <option value="<%=rs_assunto("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_assunto("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
          end if
          rs_assunto.movenext
          loop
          %>
        </select> </td>
      <%
    SET RSONDA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE onda_cd_onda=" & cen_atual("onda_cd_onda"))
    valor_onda=rsonda("ONDA_TX_DESC_ONDA")
    %>
      <td height="27" colspan="4" valign="baseline"><table width="51%" border="0">
          <tr> 
            <td colspan="2"><div align="center"><b><font face="Verdana" size="2" color="#330099">Empresas 
                Relacionadas</font></b> </div></td>
          </tr>
          <tr> 
            <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">N&atilde;o 
                selecionadas</font></div></td>
            <td><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Selecionadas</font></div></td>
          </tr>
        </table>
        
      </td>
    </tr>
    <tr> 
      <td width="6" height="21"></td>
      <td height="21" colspan="2"><div align="left"><b><font face="Verdana" size="2" color="#330099">Selecione 
          o Processo&nbsp;</font></b></div></td>
      <td width="127" height="115" rowspan="8" valign="top"> <p style="margin-top: 0; margin-bottom: 0"> 
          <select size="8" name="list1">
            <%
        set emp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & valor2 & "' AND CENA_TX_EMPRESA_RELAC LIKE '%PAPER COMPANIES%'")
        if emp.eof=TRUE then
        %>
            <option value="PAPER COMPANIES">PAPER COMPANIES</option>
            <%end if%>
            <%
        set emp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & valor2 & "' AND CENA_TX_EMPRESA_RELAC LIKE '%PETROBRAS%'")
        if emp.eof=TRUE then
        %>
            <option value="PETROBRAS">PETROBRAS</option>
            <%end if%>
            <%
        set emp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & valor2 & "' AND CENA_TX_EMPRESA_RELAC LIKE '%REFAP%'")
        if emp.eof=TRUE then
        %>
            <option value="REFAP">REFAP</option>
            <%end if%>
          </select>
      </td>
      <td width="24" height="61" rowspan="3"> <div align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2,0)"><img name="Image16" border="0" src="../../imagens/continua_F01.gif" width="24" height="24"></a></div></td>
      <td width="370" height="89" rowspan="8" valign="top"> <select size="8" name="list2">
          <%
        set emp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & valor2 & "' AND CENA_TX_EMPRESA_RELAC LIKE '%PAPER COMPANIES%'")
        if emp.eof=false then
        %>
          <option value="PAPER COMPANIES">PAPER COMPANIES</option>
          <%end if%>
          <%
        set emp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & valor2 & "' AND CENA_TX_EMPRESA_RELAC LIKE '%PETROBRAS%'")
        if emp.eof=false then
        %>
          <option value="PETROBRAS">PETROBRAS</option>
          <%end if%>
          <%
        set emp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & valor2 & "' AND CENA_TX_EMPRESA_RELAC LIKE '%REFAP%'")
        if emp.eof=false then
        %>
          <option value="REFAP">REFAP</option>
          <%end if%>
        </select> </td>
    </tr>
    <tr> 
      <td width="6" height="25"></td>
      <td height="25" colspan="2"> <select size="1" name="selProcesso" onchange="javascript:enviar()">
          <option value="0">== Selecione o Processo ==</option>
          <%do until rs_proc.eof=true
        if trim(str_proc)=trim(rs_proc("PROC_CD_PROCESSO")) then
        %>
          <option selected value=<%=rs_proc("PROC_CD_PROCESSO")%>><%=rs_proc("PROC_TX_DESC_PROCESSO")%></option>
          <%else%>
          <option value=<%=rs_proc("PROC_CD_PROCESSO")%>><%=rs_proc("PROC_TX_DESC_PROCESSO")%></option>
          <%
        end if
        rs_proc.movenext
        loop
        %>
        </select> </td>
    </tr>
    <tr> 
      <td width="6" height="3" rowspan="2"></td>
      <td height="3" colspan="2" rowspan="2"><b><font face="Verdana" size="2" color="#330099">Selecione 
        o Sub-Processo&nbsp; </font></b></td>
    </tr>
    <tr> 
      <td width="24" height="1" rowspan="2" align="center" valign="middle"> <p align="center"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img01511" border="0" src="../../imagens/continua2_F01.gif" width="24" height="24"></a> 
      </td>
    </tr>
    <tr> 
      <td width="6" height="1"></td>
      <td width="386" height="1"><b><font face="Verdana" size="2" color="#330099">
        <select size="1" name="selSubProcesso">
          <option value="0">== Selecione o Sub-Processo ==</option>
          <%do until rs_sub.eof=true
       if trim(str_sub)=trim(rs_sub("SUPR_CD_SUB_PROCESSO")) then
       %>
          <option selected value=<%=rs_sub("SUPR_CD_SUB_PROCESSO")%>><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%
		ELSE
		%>
          <option value=<%=rs_sub("SUPR_CD_SUB_PROCESSO")%>><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></option>
          <%END IF
		rs_sub.movenext
		loop
		%>
        </select>
        </font></b> </td>
    </tr>
    <tr> 
      <td width="6" height="5"></td>
      <td height="5" colspan="2"><b><font face="Verdana" size="2" color="#330099">Classe 
        do Cenário</font></b></td>
      <td width="24" height="19" rowspan="2" align="center">&nbsp; </td>
    </tr>
    <tr> 
      <td width="6" height="26" rowspan="2"></td>
      <td height="26" colspan="2" rowspan="2"> <select size="1" name="selClasse">
          <option value="0">== Selecione a Classe ==</option>
          <%
       do until rs_class.eof=true
       set atual=db.execute("SELECT * FROM " & Session("PREFIXO") & "CLASSE_CENARIO WHERE CLCE_CD_NR_CLASSE_CENARIO=" & rs_class("CLCE_CD_NR_CLASSE_CENARIO"))
       valor=atual("CLCE_TX_DESC_CLASSE_CENARIO")
       IF trim(cen_atual("CLCE_CD_NR_CLASSE_CENARIO"))=trim(rs_class("CLCE_CD_NR_CLASSE_CENARIO")) then
       %>
          <option selected value=<%=rs_class("CLCE_CD_NR_CLASSE_CENARIO")%>><%=valor%></option>
          <%
       else
       %>
          <option value=<%=rs_class("CLCE_CD_NR_CLASSE_CENARIO")%>><%=valor%></option>
          <%
       end if
		rs_class.movenext
		loop
		%>
        </select> &nbsp;&nbsp;&nbsp; <input type="hidden" name="lotus" size="20" value="<%=cen_atual("CENA_TX_SITUACAO_LOTUS")%>"> 
        <input type="hidden" name="ID" size="20" value=<%=VALOR2%>> 
      </td>
    </tr>
    <tr> 
      <td width="24" height="52"></td>
    </tr>
    <tr>
      <td width="2" height="21"> </td>
      <td width="2" height="21"> </td>
      <td height="21" colspan="2"> &nbsp;</td>
    </tr>
    <tr>
      <td width="2" height="21"> </td>
      <td width="2" height="21"> <b><font face="Verdana" size="2" color="#330099">Responsável</font></b></td>
      <td height="21" colspan="2"> &nbsp;</td>
    </tr>
    <tr>
      <td width="2" height="21"> </td>
      <td width="2" height="21"> <input type="text" name="txtResp" size="20" maxlength="4" value="<%=cen_atual("CENA_TX_RESPONSAVEL")%>"></td>
      <td height="21" colspan="2"> &nbsp;</td>
    </tr>
    <tr> 
      <td width="6" height="23">&nbsp;</td>
      <td width="386" height="23">&nbsp;</td>
    </tr>
    <tr> 
      <td width="6" height="23"></td>
      <td width="386" height="23"><b><font face="Verdana" size="2" color="#330099">Título 
        do Cenário</font></b></td>
    </tr>
    <tr> 
      <td width="6" height="26"></td>
      <td height="26" colspan="2"> <input type="text" name="txtTitulo" size="50" value="<%=cen_atual("CENA_TX_TITULO_CENARIO")%>"> 
      </td>
      <td width="127" height="26">&nbsp;</td>
    </tr>
    <tr> 
      <td height="2"></td>
      <td height="2" colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="4%" height="20"><img src="../../../xproc/imagens/b021.gif" width="16" height="16"></td>
            <td width="96%" height="20"><font size="2" face="Arial, Helvetica, sans-serif"><font color="#CC6600" face="Geneva, Arial, Helvetica, san-serif">O 
              t&iacute;tulo do cen&aacute;rio dever&aacute; ser cadastrado no 
              substantivo.</font></font></td>
          </tr>
          <tr> 
            <td width="4%" height="20">&nbsp;</td>
            <td width="96%" height="20"><font size="2" face="Arial, Helvetica, sans-serif" color="#CC6600">Exemplo: 
              &quot;Importa&ccedil;&atilde;o de material ...&quot;</font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td width="6" height="2"></td>
      <td height="2" colspan="5"><b><font face="Verdana" size="2" color="#330099">Descrição 
        do Cenário</font></b></td>
    </tr>
    <tr> 
      <td width="6" height="100"></td>
      <td height="100" colspan="5"> <textarea rows="4" name="txtDescricao" cols="68"><%=cen_atual("CENA_TX_DESC_CENARIO")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td width="6" height="21"></td>
      <td height="21" colspan="5"><b><font face="Verdana" size="2" color="#330099">Entrada</font></b></td>
    </tr>
    <tr> 
      <td width="6" height="25"></td>
      <td height="25" colspan="5"> <input type="text" name="txtEntrada" size="69" value="<%=cen_atual("CENA_TX_ENTRADA")%>"> 
      </td>
    </tr>
    <tr> 
      <td width="6" height="21"></td>
      <td height="21" colspan="5"><b><font face="Verdana" size="2" color="#330099">Saída</font></b></td>
    </tr>
    <tr> 
      <td width="6" height="25"></td>
      <td height="25" colspan="5"> <input type="text" name="txtSaida" size="69" value="<%=cen_atual("CENA_TX_SAIDA")%>"> 
      </td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
</body>

</html>