<!--#include file="conecta.asp" -->
<%
set objUSR = server.createobject("Seseg.Usuario")

if objUSR.GetUsuario then
	chave=objUSR.sei_chave
	lotacao=objUSR.sei_lotacao
	nome=objUSR.sei_nome
	set objUSR = nothing
else
	response.redirect "erro.asp?op=3"
end if

set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation=3

set db2 = server.createobject("ADODB.CONNECTION")
db2.open "Provider=Microsoft.Jet.Oledb.4.0;data source=" & server.mappath("base.mdb")
db2.CursorLocation=3

set is_cli = db.execute("SELECT * FROM CLI WHERE USMA_CD_USUARIO='" & chave & "'")

if is_cli.eof=false then
	set cli = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM CLI_ORGAO WHERE USMA_CD_USUARIO='" & chave & "'")
	
	if len(cli("ORME_CD_ORG_MENOR"))=2 then
		orgao_cli=cli("ORME_CD_ORG_MENOR")
	end if

	if right(cli("ORME_CD_ORG_MENOR"),8)="00000000" then
		orgao_cli=left(cli("ORME_CD_ORG_MENOR"),7)
	else
		if right(cli("ORME_CD_ORG_MENOR"),5)="00000" then
			orgao_cli=left(cli("ORME_CD_ORG_MENOR"),10)		
		else
			if right(cli("ORME_CD_ORG_MENOR"),2)="00" then
				orgao_cli=left(cli("ORME_CD_ORG_MENOR"),13)
			else
				orgao_cli=cli("ORME_CD_ORG_MENOR")
			end if
		end if
	end if

	if request("selOrgao")=0 then
		set users = db.execute("SELECT DISTINCT USMA_CD_USUARIO, USMA_TX_NOME_USUARIO FROM USUARIO_MAPEAMENTO WHERE ORME_CD_ORG_MENOR LIKE '" & orgao_cli & "%' AND USMA_CD_USUARIO<>'" & chave & "' ORDER BY USMA_TX_NOME_USUARIO")
	else	
		set users = db.execute("SELECT DISTINCT USMA_CD_USUARIO, USMA_TX_NOME_USUARIO FROM USUARIO_MAPEAMENTO WHERE ORME_CD_ORG_MENOR LIKE '" & request("selOrgao") & "%' AND USMA_CD_USUARIO<>'" & chave & "' ORDER BY USMA_TX_NOME_USUARIO")
	end if
else
	set temp = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_SG_ORG_MENOR='" & lotacao & "'")
	orgao_cli="0"
	set users = db.execute("SELECT * FROM USUARIO_MAPEAMENTO WHERE ORME_CD_ORG_MENOR LIKE '" & temp("ORME_CD_ORG_MENOR") & "%' AND PERF_CD_PERFIL<>2 AND USMA_CD_USUARIO<>'" & chave & "' ORDER BY USMA_TX_NOME_USUARIO")
end if

if orgao_cli="0" then
	ssql="SELECT DISTINCT ORME_CD_ORG_MENOR, ORME_SG_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR LIKE '" & temp("ORME_CD_ORG_MENOR") & "%' AND ORME_CD_STATUS='A' ORDER BY ORME_SG_ORG_MENOR"
	org_ind = left(temp("ORME_CD_ORG_MENOR"),7)
else
	ssql="SELECT DISTINCT ORME_CD_ORG_MENOR, ORME_SG_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR LIKE '" & orgao_cli & "%' AND ORME_CD_STATUS='A' ORDER BY ORME_SG_ORG_MENOR"
	org_ind = left(orgao_cli,7)	
end if

set orgao = db.execute(ssql)

set mega = db.execute("SELECT * FROM MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

if request("selMega")<>0 and orgao_cli<>"0" then

	f_orgao = left(orgao_cli,2)

	pre_o=""	
	
	set temp = db.execute("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & f_orgao & "'")
	if temp.eof=false then
		pre_o = temp("AGLU_SG_AGLUTINADO")
	end if

	g_orgao = left(orgao_cli,7)
	
	pre_m=""
	
	set temp2 = db.execute("SELECT DISTINCT ORME_SG_ORG_MENOR FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & g_orgao & "00000000' AND ORME_CD_STATUS='A'")
	if temp2.eof=false then
		pre_m = temp2("ORME_SG_ORG_MENOR")
	end if
	
	set temp = db.execute("SELECT MEPR_TX_ABREVIA_CURSO FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & request("selMega"))
	pre_c = temp("MEPR_TX_ABREVIA_CURSO")
	
	if pre_m="" then
		ssql="SELECT DISTINCT CURSO AS CURS_CD_CURSO FROM [" & pre_o & "] WHERE CURSO LIKE '" & pre_c & "%' ORDER BY CURSO"
		qu=1
	else
		ssql="SELECT DISTINCT CURSO AS CURS_CD_CURSO FROM [" & pre_o & "] WHERE ORGAO='" & pre_m & "' AND CURSO LIKE '" & pre_c & "%' ORDER BY CURSO"	
		qu=2
	end if
	
	on error resume next
	set curso = db2.execute(ssql)
	
	if err.number=0 and curso.eof=true and qu=2 and left(g_orgao,2)=78 then
		ssql="SELECT DISTINCT CURSO AS CURS_CD_CURSO FROM [" & pre_o & "] WHERE CURSO LIKE '" & pre_c & "%' ORDER BY CURSO"
		on error resume next
		set curso = db2.execute(ssql)
	end if

	if err.number<>0 then
		set curso = db.execute("SELECT DISTINCT CURS_CD_CURSO FROM CURSO WHERE MEPR_CD_MEGA_PROCESSO=0")
		err.clear()
	end if
		
else

	set curso = db.execute("SELECT DISTINCT CURS_CD_CURSO FROM CURSO WHERE MEPR_CD_MEGA_PROCESSO=0")

end if

if f_orgao="87" then
	estado="disabled"
	func01="move_selFunc();"
	func02="move_list1();"
	mensagem = "  - UNIDADES DO E&P - Não é necessário efetuar a seleção dos cursos, apenas clique nas setas..."	
else
	estado=""
	func01=""
	func02=""
	mensagem = " "	
end if

%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<script language="javascript" src="troca_lista.js"></script>

<script>
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
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
</script>

<script>
function carrega_txt(fbox) {
document.frm1.txtcurso.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtcurso.value = document.frm1.txtcurso.value + "," + fbox.options[i].value;
}
}

function carrega_curso()
{
window.location.href = 'mapeamento.asp?selMega='+document.frm1.selMega.value+'&selFunc='+document.frm1.selFunc.value+'&selOrgao='+document.frm1.selOrgao.value
}

function carrega_curso2()
{
window.location.href = 'mapeamento.asp?selMega='+document.frm1.selMega.value+'&selFunc=0&selOrgao='+document.frm1.selOrgao.value
}

function envia()
{
if(document.frm1.selFunc.selectedIndex == 0)
{
alert("Você deve selecionar um EMPREGADO!");
document.frm1.selFunc.focus();
return;
}
if(document.frm1.selMega.selectedIndex == 0)
{
alert("Você deve selecionar um MEGA-PROCESSO!");
document.frm1.selMega.focus();
return;
}
if ((document.frm1.selCurso.options.length == 0)&&(document.frm1.list1.options.length == 0))
{
alert("Não existe nenhum CURSO INDICADO no MEGA-PROCESSO Selecionado...");
document.frm1.selMega.focus();
return;
}
else
{
carrega_txt(document.frm1.list1);
document.frm1.action='valida_mapeamento.asp'
document.frm1.submit();
}
}

function encontra_chave(e)
{
if(e.length==4)
	{
	MM_swapImage('loader','','aguarde.gif',1);
	window.location.href = 'mapeamento.asp?selMega='+document.frm1.selMega.value+'&selFunc='+e.toUpperCase()+'&selOrgao=0'
	}
}

function abre_janela()
{
window.open("perfil_curso.asp?selOrgao="+<%=orgao_cli%>+"&selMega="+document.frm1.selMega.value,"_blank","width=600,height=400,history=0,scrollbars=1,titlebar=0,resizable=0")
}

function move_selFunc()
{
document.frm1.selCurso.disabled = false;
document.frm1.list1.disabled = false;                         	
var a = document.frm1.selCurso.options.length;							
for(var i = 0; i<a ; i++)
{
	document.frm1.selCurso.options[i].selected = true;
}
document.frm1.selCurso.disabled = true;
document.frm1.list1.disabled = true;                         	
}

function move_list1()
{
document.frm1.selCurso.disabled = false;
document.frm1.list1.disabled = false;                         	
var a = document.frm1.list1.options.length;							
for(var i = 0; i<a ; i++)
{
	document.frm1.list1.options[i].selected = true;
}
document.frm1.selCurso.disabled = true;
document.frm1.list1.disabled = true;                         	
}
</script>

<body topmargin="0" leftmargin="0" link="#800000" vlink="#800000" alink="#800000" bgcolor="#FFFFFF">
<form name="frm1" action="">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="95%" id="AutoNumber2" height="472">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2"><img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="403" valign="top"><img border="0" src="lado.jpg" width="83" height="417"></td>
                      <td width="87%" height="403" valign="top"><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="98%" id="AutoNumber3" height="73">
                         <tr>
                                    <td width="100%" valign="top" colspan="4" height="4"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"> Responsável pela Indicação</font></b></td>
                         </tr>
                         <tr>
                                    <td width="15%" height="7"><b><font face="Verdana" size="2">Chave</font></b></td>
                                    <td width="24%" height="7"><font face="Verdana" size="2"><%=chave%></font></td>
                                    <td width="16%" height="7"><b><font face="Verdana" size="2">Lotação</font></b></td>
                                    <td width="45%" height="7"><font face="Verdana" size="2"><%=lotacao%></font></td>
                         </tr>
                         <tr>
                                    <td width="15%" height="3"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">Nome</font></b></td>
                                    <td width="85%" height="3" colspan="3"><font face="Verdana" size="2"><%=nome%></font></td>
                         </tr>
                         </table>

        <br>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="99%" id="AutoNumber4" height="1">
                         <tr>
                                    <td width="100%" height="19" colspan="4"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"> Selecione o Órgão de Lotação do Multiplicador</font></b></td>
                         </tr>
                         <tr>
                                    <td width="100%" height="19" colspan="4">
                                    <select size="1" name="selOrgao" style="font-family: Verdana; font-size: 8 pt" onChange="carrega_curso2()">
                                       <option value="0">== Selecione o Órgão ==</option>
                                       <%
                                       do until orgao.eof=true
                                       
                                       if request("selOrgao") = orgao("ORME_CD_ORG_MENOR") then
                                       	self="selected"
                                       else
                                       	self=""
                                       end if

                                       %>
                                       <option <%=self%> value="<%=orgao("ORME_CD_ORG_MENOR")%>"><%=orgao("ORME_SG_ORG_MENOR")%></option>
                                       <%
                                       orgao.movenext
                                       loop
                                       %>
                                       
                                       </select></td>
                         </tr>
                         <tr>
                                    
            <td width="100%" height="8" colspan="4"></td>
                         </tr>
                         <tr>
                                    <td width="100%" height="21" colspan="4"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"> Selecione o Multiplicador</font></b></td>
                         </tr>
                         <tr>
                                    <td width="100%" height="22" colspan="4"><p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">
                                    <select size="1" name="selFunc" style="font-family: Verdana; font-size: 8 pt" onChange="carrega_curso()">
                                       <option value="XXXX">== Selecione o Empregado ==</option>
                                       <%
                                       do until users.eof=true
                                       
                                       if request("selFunc") = users("USMA_CD_USUARIO") then
                                       	self="selected"
                                       else
                                       	self=""
                                       end if

                                       %>
                                       <option <%=self%> value="<%=users("USMA_CD_USUARIO")%>"><%=users("USMA_TX_NOME_USUARIO")%></option>
                                       <%
                                       users.movenext
                                       loop
                                       %>
                                       </select> </font></b></p>
                                       <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="2">
                                       ou digite a chave do Multiplicador : <input type="text" name="txtchave" size="9" maxlength="4" onKeyUp="encontra_chave(this.value)">
                                       <img border="0" src="branco.gif" name="loader"></font></b></td>
                         </tr>
                         <tr>
                                    
            <td width="100%" height="27" colspan="4"><img src="b2.gif" width="490" height="20" name="verifica"></td>
                         </tr>
                         <tr>
                                    <td width="50%" height="19" colspan="2"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"> Selecione o Mega-Processo</font></b></td>
                                    <td width="50%" height="57" colspan="2" rowspan="3"><p align="center">
                                    <%
                                    if curso.eof=false then
                                    %>
                                    <a href="#" onClick="abre_janela()">
                                    <img border="0" src="desc_mult.gif" align="left">
                                    </a>
                                    <%end if%>
                                    </td>
                         </tr>
                         <tr>
                                    <td width="50%" height="19" colspan="2"><b><font face="Verdana" size="2">
                                    <select size="1" name="selMega" style="font-family: Verdana; font-size: 8 pt" onChange="carrega_curso()">
                                       <option value="0">== Selecione o Mega Processo ==</option>                                       
                                       <%
                                       do until mega.eof=true
                                       
                                       if trim(request("selMega")) = trim(mega("MEPR_CD_MEGA_PROCESSO")) then
                                       	selm="selected"
                                       else
                                       	selm=""
                                       end if
                                                                              
                                       %>
                                       <option <%=selm%> value="<%=mega("MEPR_CD_MEGA_PROCESSO")%>"><%=mega("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
                                       <%
                                       mega.movenext
                                       loop
                                       %>
                                       </select></font></b></td>
        	                 </tr>
							 <%
							 set tmp = db.execute("SELECT * FROM APOIO_LOCAL_MULT WHERE USMA_CD_USUARIO='" & request("selFunc") & "' AND APLO_NR_ATRIBUICAO=1 AND APLO_NR_SITUACAO=1")
							 if tmp.eof=false then
							 	set tmp2 = db.execute("SELECT * FROM APOIO_LOCAL_ORGAO WHERE USMA_CD_USUARIO='" & request("selFunc") & "' AND ORME_CD_ORG_MENOR LIKE '" & org_ind & "%' AND APLO_NR_ATRIBUICAO=1")
								if tmp2.eof=false then								
							 %>
							 <script>
							 {
							 MM_swapImage('verifica','','b3.gif',1);
							 }
							 </script>
							 <%
							 	end if
							 end if
							 %>
                         <tr>
                                    <td width="50%" height="19" colspan="2"><input type="hidden" name="txtcurso" size="20"></td>
                         </tr>
                         <tr>
                                    <td width="100%" height="19" colspan="4"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"> Cursos<font color="#FF0000"><%=mensagem%></font></font></b></td>
                         </tr>
                         <tr>
                                    <td width="3%" height="109" rowspan="4">&nbsp;</td>
                                    <td width="50%" height="109" rowspan="4">
                                    <select size="8" name="selCurso" multiple style="font-family: Verdana; font-size: 7 pt" <%=estado%>>
                                    <%
                                    do until curso.eof=true
                                    
                                    set temp = db.execute("SELECT * FROM APOIO_LOCAL_CURSO WHERE USMA_CD_USUARIO='" & request("selFunc") & "' AND CURS_CD_CURSO='" & curso("CURS_CD_CURSO") & "'")
                                    
                                    if temp.eof=true then
                                    
                                    set c = db.execute("SELECT * FROM CURSO WHERE CURS_CD_CURSO = '" & curso("CURS_CD_CURSO") & "'")
                                    %>
                                    
                                    <option value="<%=curso("CURS_CD_CURSO")%>"><%=c("CURS_TX_NOME_CURSO")%></option>
                                    <%
                                    end if
                                    curso.movenext
                                    loop
                                    %>
                                    </select></td>
                                    <td width="8%" height="13"></td>
                                    <td width="59%" height="109" rowspan="4">
                                    <select size="8" name="list1" multiple style="font-family: Verdana; font-size: 7 pt" <%=estado%>>
                                    <%
                                    on error resume next
                                    curso.movefirst
                                    err.clear
                                    
                                    do until curso.eof=true
                                    
                                    set temp = db.execute("SELECT * FROM APOIO_LOCAL_CURSO WHERE USMA_CD_USUARIO='" & request("selFunc") & "' AND CURS_CD_CURSO='" & curso("CURS_CD_CURSO") & "'")
                                    
                                    if temp.eof=false then
                                    set c = db.execute("SELECT * FROM CURSO WHERE CURS_CD_CURSO = '" & curso("CURS_CD_CURSO") & "'")
                                    %>
                                    <option value="<%=curso("CURS_CD_CURSO")%>"><%=c("CURS_TX_NOME_CURSO")%></option>
                                    <%
                                    end if
                                    curso.movenext
                                    loop
                                    
                                    %>

                                    </select></td>
                         </tr>
                         
                         <tr>
                                    <td width="8%" height="40"><p align="center"><img border="0" src="seta_d.jpg" onClick="<%=func01%>move(document.frm1.selCurso,document.frm1.list1,0)" alt="Adicionar os Cursos Selecionados..."></td>
                         </tr>
                         <tr>
                                    <td width="8%" height="41"><p align="center"><img border="0" src="seta_e.jpg" onClick="<%=func02%>move(document.frm1.list1,document.frm1.selCurso,0)" alt="Retirar os Cursos Selecionados...">                                    </td>
                         </tr>
                         <tr>
                                    <td width="8%" height="15"></td>
                         </tr>
                         </table>
                         <p align="center"><a href="menu.asp"><img border="0" src="voltar.gif"></a>&nbsp;&nbsp;&nbsp; <a href="#" onClick="envia()"> <img border="0" src="enviar.gif"></a></td>
           </tr>
</table>
</form>
</body>

</html>
<script>
document.title = 'Indicação de Multiplicadores'
</script>