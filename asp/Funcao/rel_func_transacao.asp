<%
Response.Buffer=False

server.scripttimeout=99999999

str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")
if str_Uso = "" then
   str_Uso = 0
end if   
if str_Desuso = "" then
   str_Desuso = 0
end if   

if str_Uso = 1 and str_Desuso = 1 then
   str_usoDesuso =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	end if        	  
end if

mega1=request("selMegaFuncao")

valor=request("selMegaProcesso")
proc=request("selProcesso")
subproc=request("selSubProcesso")
str_selsubmodulo=request("selSubModulo")
tatual=0

if proc<>0 then
  complemento=" AND PROC_CD_PROCESSO=" & proc
end if

if subproc<>0 then
  complemento=complemento+" AND SUPR_CD_SUB_PROCESSO=" & subproc
end if

if str_selsubmodulo <> 0 then
 sub_modulo = "  and FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA =" & str_selsubmodulo & " "
else
 sub_modulo = ""
end if

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

db.cursorlocation=3

ssql=""
ssql="SELECT DISTINCT FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO_PAI, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
ssql=ssql+"WHERE FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & mega1 &  sub_modulo & " " & str_usodesuso
ssql=ssql+" ORDER BY FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "

set rs=db.execute(SSQL)

set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega1)

texto=temp("MEPR_TX_DESC_MEGA_PROCESSO")

TEM = 0

ssql01="SELECT DISTINCT MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & valor + complemento + " ORDER BY TRAN_CD_TRANSACAO"
set rs1=db.execute(ssql01)

reg = rs.RecordCount
regt = rs1.RecordCount

tem = 0
it = 0

do until it = regt
	if tem > 0 then
		exit do
	end if
	t=0
	do until t = reg		   
		if tem > 0 then
			exit do
		end if
		qtos=0
		set rstemp = db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RS1("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO_PAI") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")
		qtos = rstemp.RecordCount
		if qtos > 0 then
			tem = tem + 1	
		end if
		t = t + 1
		rs.movefirst
	loop
	it = it + 1
	rs1.movenext
loop

rs1.movefirst
%>	         
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript" type="text/JavaScript">
<!--


function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body link="#000000" vlink="#000000" alink="#000000" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('../../Anima/branco.gif')">

  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="93">
    <tr> 
      <td width="20%" height="64">&nbsp;</td>
      <td width="44%" height="64">&nbsp;</td>
      <td width="36%" valign="top" height="64"> 
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
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
      <td colspan="3" height="29">&nbsp; </td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
             <tr>
                        <td width="50%"><b><font face="Verdana" color="#000080">Relatório Geral de Função x Transação</font></b></td>
                        <td width="50%"><font face="Verdana" color="#330099" size="3"> <img src="preloader.gif" name="loader" width="190" height="50" border="0" id="loader"></font></td>
             </tr>
  </table>
<%if reg>0 and regt>0 and tem<>0 then%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="684" height="113">
           <tr>
           <td width="119" height="20" bgcolor="#000080">&nbsp;</td>
           <td width="218" height="20" bgcolor="#000080"><b><font face="Verdana" size="2" color="#FFFFFF">Função de Negócio ==&gt;</font></b></td>
           <%
            reg=rs.RecordCount
            i=0
            do until i=reg
           %>
           <td width="176" rowspan="2" style="font-family: Verdana; font-size: 7 pt; color: #800000" height="75" bgcolor="#C4D2E6"><p align="center"><%=rs("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></td>
           <%
    		i=i+1
    		rs.movenext
    		loop       
	       %>
           </tr>
           <tr>
           <td width="338" colspan="2" height="54" bgcolor="#000080"><b><font face="Verdana" size="2" color="#FFFFFF">Transação</font></b></td>
           </tr>
           <%
           regt=rs1.Recordcount
           it=0
           do until it=regt
           rs.movefirst
           %>
           <tr>
           <td width="338" colspan="2" height="37" bgcolor="#D6D9CC"><font face="Verdana" size="1"><%=rs1("TRAN_CD_TRANSACAO")%></font></td>
		   <%
			t=0
			do until t = reg		   
			qtos=0
		   	set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RS1("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO_PAI") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")
		   	qtos=rstemp.RecordCount
		   	if qtos>0 then
		   %>	           
           <td width="166" height="37"><p align="center"><b><font face="Verdana" size="5" color="#0000FF">X</font></b></td>
           <%else%>
           <td width="166" height="37"><p align="center"><b><font face="Verdana" size="5" color="#0000FF"></font></b></td>
		   <%end if
		   t=t+1
		   rs.movenext
		   loop
		   %>	           
           </tr>
           <%
           it=it+1
           rs1.movenext
           loop
	       %>
</table>
<%
else
%>
<p>&nbsp;</p>
  <b><font color="#800000">Nenhum Registro Encontrado para a Seleção!
<%end if%>
</font></b>
</body>
<script>
MM_swapImage('loader','','../../Flash/branco.gif',1);
  </script>
</html>