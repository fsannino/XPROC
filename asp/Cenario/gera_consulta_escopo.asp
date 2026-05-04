<%
Server.Scripttimeout=99999999

mega=request("selMegaProcesso")
onda=request("selOnda")
situacao=request("selStatus")
assunto=request("selAssunto")

if mega<>0 then
	compl=compl+" MEPR_CD_MEGA_PROCESSO=" & mega & " AND"
end if

if onda<>0 then
	compl=compl+" ONDA_CD_ONDA=" & onda & " AND"
end if

if assunto<>0 then
	compl=compl+" SUMO_NR_SEQUENCIA=" & assunto & " AND"
end if

if situacao<>"0" then
	compl=compl+" CENA_TX_SITUACAO='" & situacao & "' AND"
end if

data_cons=request("data01")

dia=left(data_cons, 2)
mes=left((right(data_cons,7)),2)
ano=right(data_cons, 4)

if len(trim(mes))=1 then
	mes="0" & mes
end if

if len(trim(dia))=1 then
	dia="0" & dia
end if

dia=dia+1

data_final=ano & "-" & mes & "-" & dia

'Session("Conn_String_Cogest_Gravacao")="Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

'db.CursorLocation = adUseClient

if len(compl)>0 then
	compl=left(compl,(len(compl))-4)
	compl=" WHERE" + compl
end if

set rs=db.execute("SELECT DISTINCT * FROM " & Session("PREFIXO") & "CENARIO" & compl & " ORDER BY CENA_CD_CENARIO")
%>
<html>
<head>
<STYLE type=text/css>
BODY{
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>

<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
   var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value+"&selMegaProcesso2="+document.frm1.selMegaProcesso2.value+"&selProcesso="+document.frm1.selProcesso.value+"&selSubProcesso="+document.frm1.selSubProcesso.value+"&selAtividadeCarga="+document.frm1.selAtividadeCarga.value+"'");
}
function MM_goToURL2() { //v3.0
   var i, args=MM_goToURL2.arguments; document.MM_returnValue = false;
   for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value+"'");
}
function MM_goToURL5() { //v3.0
  
  var i,x,args=MM_goToURL5.arguments; document.MM_returnValue = false;
  
  //for (i=0; i<(args.length-1); i+=4) eval(args[i]+".location='"+args[i+1]+"?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selProcesso="+document.frm1.txtProcesso.value+"&selSubProcesso="+document.frm1.txtSubProcesso.value+"&selCenario="+document.frm1.txtCenario.value+"&p_CenarioChSequencia="+args[3]+"'");
  //alert(document.frm1.imgMarca1.src);
  // NÃO CONSIGO TESTAR EM DESENV OU PRODUÇÃO
  
	x=MM_findObj(args[4])
  
	var a = document.URL;
	var n=0;

	for (var i = 1 ; i < 1000; i++)
	{
	var final=a.slice(0,i)
	var t=a.slice(i-1,i);
	if (t=='/')
	{
	n = n + 1;
	}
	if(n == 4)
	{
	i = 1000;
	}
	}
	var tam=final.length;
	var caminho = final.slice(0,tam-1);

  var valor=caminho + "/imagens/func_tran_nao_marcada.gif"
  
 if(x.src == valor) {
	 window.open("inc_Funcao_Trans.asp?"+args[3],"_blank","width=100,height=100,history=0,scrollbars=0,titlebar=0,resizable=0")
     MM_swapImage(x.name,'','../../imagens/func_tran_marcada.gif',1);
    // window.open("exibe_funcao_trans.asp?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value,"_blank","width=100,height=100,history=0,scrollbars=1,titlebar=0,resizable=0")

	}
	else 
	{
  // if(document.frm1.imgMarca1.src == "http://S6000WS10.corp.petrobras.biz/xproc/imagens/b03.gif") 
	 window.open("exc_Funcao_Trans.asp?"+args[3],"_blank","width=100,height=100,history=0,scrollbars=0,titlebar=0,resizable=0")	
    MM_swapImage(x.name,'','../../imagens/func_tran_nao_marcada.gif',1);

    }
  //for (i=0; i<(args.length-1); i+=3) eval(args[i]+".location='"+args[i+1]+"?"+args[3]+"'");
}
function carrega_txt(fbox){
   document.frm1.txtTranSelecionada.value = "";
   for(var i=0; i<fbox.options.length; i++) 
     {
     document.frm1.txtTranSelecionada.value = document.frm1.txtTranSelecionada.value + "," + fbox.options[i].value;
     }
}
function Confirma2(){ 
	  document.frm1.submit();
}
function Confirma(){ 
   if (document.frm1.selMegaProcesso2.selectedIndex == 0) { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
   if (document.frm1.selProcesso.selectedIndex == 0) { 
	 alert("Selecione um Proceso.");
     document.frm1.selProcesso.focus();
     return;
     }	 
   if (document.frm1.selSubProcesso.selectedIndex == 0) { 
	 alert("Selecione um Sub Proceso.");
     document.frm1.selSubProcesso.focus();
     return;
     }	 
   if (document.frm1.selAtividadeCarga.selectedIndex == 0) { 
	 alert("Selecione uma Atividasde de Carga.");
     document.frm1.selAtividadeCarga.focus();
     return;
     }	 
	 else
     {
	 document.frm1.txtDescMegaProcesso2.value = document.frm1.selMegaProcesso2.options[document.frm1.selMegaProcesso2.selectedIndex].text;	 
	 document.frm1.txtDescProcesso.value = document.frm1.selProcesso.options[document.frm1.selProcesso.selectedIndex].text;	 
	 document.frm1.txtDescSubProcesso.value = document.frm1.selSubProcesso.options[document.frm1.selSubProcesso.selectedIndex].text;	 
	 document.frm1.txtDescAtividadeCarga.value = document.frm1.selAtividadeCarga.options[document.frm1.selAtividadeCarga.selectedIndex].text;	 	 
 	 carrega_txt(document.frm1.list2);
	 document.frm1.submit();
	 }
}
function Limpa(){
	document.frm1.reset();
}
function exibe_transacao(){
	window.open("exibe_funcao_trans.asp?selMegaProcesso="+document.frm1.txtMegaProcesso.value+"&selFuncao="+document.frm1.txtFuncao.value,"_blank","width=620,height=240,history=0,scrollbars=1,titlebar=0,resizable=0")
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
//-->
</script>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="gera_consulta_escopo.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>
              </div>
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
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20">&nbsp; 
      
    </td>
  </tr>
</table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" width="88%">
    <tr>
      <td width="71%"><p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#000080" face="Verdana" size="3">Consulta de
  Escopo de Cenário</font></p>
        <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#000080" face="Verdana" size="2">Data
  de Consulta : <%=request("data01")%></font><font color="#000080" face="Verdana" size="1"></font></b></p>
      </td>
      <td width="29%"><img border="0" src="../../Flash/preloader.gif" width="190" height="50" name="loader" id="loader"></td>
    </tr>
  </table>
  <p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" width="95%">
    <tr>
      <td width="62%" bgcolor="#000080"><font size="2" face="Verdana" color="#FFFFFF"><b>Cenário</b></font></td>
      <td width="13%" bgcolor="#000080" align="center"><font size="2" face="Verdana" color="#FFFFFF"><b>Status
        do Cenário</b></font></td>
    </tr>
   <%
	ssql=""
	ssql="SELECT CENA_CD_CENARIO, MAX(ATUA_DT_ATUALIZACAO) AS DATA, CEVA_TX_SITUACAO FROM " & session("prefixo") & "CENARIO_VALIDACAO"
	ssql=ssql + " GROUP BY CENA_CD_CENARIO, CEVA_TX_SITUACAO"
	ssql=ssql + " HAVING (MAX(ATUA_DT_ATUALIZACAO) < CONVERT(DATETIME, '" & data_final & " 00:00:00', 102))"
	'RESPONSE.WRITE SSQL & "<P>"
	set temp2 = db.execute(ssql)
	
	tem=0
	
	if temp2.eof = false then
	
	do until rs.eof=true
	
	VALOR5=0
   
	ssql="" 
	ssql="SELECT CENA_CD_CENARIO, (ATUA_DT_ATUALIZACAO) AS DATA, CEVA_TX_SITUACAO FROM " & session("prefixo") & "CENARIO_VALIDACAO"
	ssql=ssql + " GROUP BY CENA_CD_CENARIO, CEVA_TX_SITUACAO, ATUA_DT_ATUALIZACAO"
	ssql=ssql + " HAVING (CENA_CD_CENARIO = '" & RS("CENA_CD_CENARIO") & "') AND (ATUA_DT_ATUALIZACAO < CONVERT(DATETIME, '" & data_final & " 00:00:00', 102)) ORDER BY ATUA_DT_ATUALIZACAO"
	
	'RESPONSE.WRITE SSQL & "<P>"
	
	set temp=db.execute(ssql)
	
	if temp.eof=true then
		valor1="NENHUM REGISTRO ENCONTRADO"
		valor5=4
	else
		do until temp.eof=true	
			valor1=DAY(temp("DATA")) & "/" & MONTH(TEMP("DATA")) & "/" & YEAR(TEMP("DATA"))
			situacao=temp("CEVA_TX_SITUACAO")
			temp.movenext
		loop
	end if
	
	valor2=rs("CENA_TX_SITUACAO")
	
	select case valor2
		case "EE"
			valor3="EM ELABORAÇÃO"
		case "DS"
			valor3="DESENHADO"
		case "DF"
			valor3="DEFINIDO"
		case "PT"
			valor3="PRONTO PARA TESTE"
		case "TD"
			valor3="TESTADO NO PED"
		case "TQ"
			valor3="TESTADO NO PEQ"
		case else
			valor3="SEM STATUS DEFINIDO"
	end select
	
	if situacao<>0 AND VALOR5<>4 then
	%>
    <tr>
      <td width="62%" bgcolor="#D8D8D8"><font face="Verdana" size="1"><B><a href="gera_rel_geral.asp?id=<%=rs("CENA_CD_CENARIO")%>"><%=rs("CENA_CD_CENARIO")%></a></B>- <%=rs("CENA_TX_TITULO_CENARIO")%></font></td>
      <td width="13%" align="center" bgcolor="#CCD2C6"><font face="Verdana" size="1"><%=valor3%></font></td>
    </tr>
    <%
    tem=tem+1
    end if
    rs.movenext
	situacao=0
    loop
    %>
  </table>
  <%if tem<>0 then%>
  <P>
  <font face="Verdana" size="2"><b>Total de Cenários</b> : <%=tem%></font>
  <P>
  <%end if%>
  &nbsp;
  <%
  end if
  %>

  <%if tem=0 then%>
  <b><font color="#800000">Nenhum Registro Encontrado para a Seleção</font></b>
  <%end if%>
  </form>
</body>

<script>
MM_swapImage('loader','','../../Flash/branco.gif',1);
</script>

</html>