<%
server.scripttimeout=99999999

response.buffer=false

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
' response.Write(" sub modulo ")
' response.Write(sub_modulo)
ON ERROR RESUME NEXT

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

ssql=""
ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
ssql=ssql+"WHERE FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & mega1 &  sub_modulo & " " & str_usodesuso
ssql=ssql+" ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "

'SSQL1=""
'SSQL1="SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & mega1 &  sub_modulo & " " & str_usodesuso & " ORDER BY FUNE_CD_FUNCAO_NEGOCIO"

set rs=db.execute(SSQL1)

set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega1)

texto=temp("MEPR_TX_DESC_MEGA_PROCESSO")
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

<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
           <tr>
                      <td width="20%" height="20">&nbsp;</td>
                      <td width="44%" height="60">&nbsp;</td>
                      <td width="36%" valign="top"><table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
                         <tr>
                                    <td bgcolor="#330099" width="39" valign="middle" align="center">
                                       <div align="center">
                                                 <p align="center"><a href="JavaScript:history.back()"><img src="../../imagens/voltar.gif" name="Image1" border="0" id="Image1"></a></div>
                                    </td>
                                    <td bgcolor="#330099" width="36" valign="middle" align="center">
                                       <div align="center">
                                                 <a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
                                    </td>
                                    <td bgcolor="#330099" width="27" valign="middle" align="center">
                                       <div align="center">
                                                 <a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
                                    </td>
                         </tr>
                         <tr>
                                    <td bgcolor="#330099" height="12" width="39" valign="middle" align="center">
                                       <div align="center">
                                                 <a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
                                    </td>
                                    <td bgcolor="#330099" height="12" width="36" valign="middle" align="center">
                                       <div align="center">
                                                 <a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
                                    </td>
                                    <td bgcolor="#330099" height="12" width="27" valign="middle" align="center">
                                       <div align="center">
                                                 <a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
                                    </td>
                         </tr>
                         </table>
                      </td>
           </tr>
           <tr bgcolor="#00FF99">
                      <td colspan="3" height="20"><table width="625" border="0" align="center">
                         <tr>
                                    <td width="26"></td>
                                    <td width="50"><a href="javascript:print()"><img border="0" src="../../imagens/print.gif"></a></td>
                                    <td width="26">&nbsp;</td>
                                    <td width="195"></td>
                                    <td width="27"></td>
                                    <td width="50"><a href="rel_func_transacao_excel.asp?selMegaFuncao=<%=mega1%>&selMegaProcesso=<%=valor%>&selProcesso=<%=proc%>&selSubProcesso=<%=subproc%>&chkEmUso=<%=str_Uso%>&chkEmDesuso=<%=str_Desuso%>" target="blank">
                                       <img border="0" src="../../imagens/exp_excel.gif"></a></td>
                                    <td width="28"></td>
                                    <td width="26">&nbsp;</td>
                                    <td width="159"></td>
                         </tr>
                         </table>
                      </td>
           </tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<table width="91%" border="0">
           <tr>
                      <td width="62%" height="44"><font face="Verdana" color="#330099" size="3">Relatório de Associação de Funções de Negócio</font></td>
                      <td width="38%"><font face="Verdana" color="#330099" size="3"><img src="../../Flash/preloader.gif" name="loader" width="190" height="50" border="0" id="loader"></a> </font></td>
           </tr>
</table>
<font face="Verdana" color="#330099" size="3"></font><p style="margin-bottom: 0"><b><font face="Verdana" color="#330099" size="2">Mega-Processo selecionado : <%=valor%>-<%=texto%></font></b></p>
<%
cor=4

TEM=0

DO UNTIL RS.EOF=TRUE
	ssql1="SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & valor & " AND FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNE_CD_FUNCAO_NEGOCIO_PAI") & "'"+ complemento
	set ATUAL2=db.execute(ssql1)
	IF ATUAL2.EOF=FALSE THEN
		TEM=TEM+1
	END IF
RS.MOVENEXT	
LOOP

RS.MOVEFIRST

IF TEM<>0 THEN

set ATUAL=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & temp("MEPR_CD_MEGA_PROCESSO"))

if atual.eof=false then 

IF RS.EOF=FALSE THEN

tatual=1
%> <p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp; </font></p>
<p align="left" style="margin-top: 0; margin-bottom: 0"></p>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" id="AutoNumber1" width="635">
           <tr>
                      <td width="445" bgcolor="#330099" height="13" colspan="2" align="center"><p align="right" style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#FFFFFF">Função R/3 --&gt;</font></b></td>
                      <%
  do until rs.eof=true
  set functemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "'")
  tit_funcao=functemp("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
  %> <td width="98" height="35" rowspan="2" bgcolor="#FFFFFF" valign="middle"><p align="center" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1">
                         <a href="#" onclick="javascript:window.open(&quot;exibe_funcao.asp?selFuncao=<%=rs("FUNE_CD_FUNCAO_NEGOCIO")%>&quot;,&quot;&quot;,&quot;width=550,height=340,status=0,toolbar=0&quot;)"><%=tit_funcao%> </font></a></td>
                      <%
  rs.movenext
  loop

  valor_sql="SELECT DISTINCT MEPR_CD_MEGA_PROCESSO, TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & valor + complemento
  'response.write valor_sql
  set rs1=db.execute(valor_sql)
  %> </tr>
           <tr>
                      <td width="281" bgcolor="#330099" height="20" align="center"><p style="margin-top: 0; margin-bottom: 0" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Mega-Processo</font></b></td>
                      <td width="162" bgcolor="#330099" height="20" align="center"><p style="margin-top: 0; margin-bottom: 0" align="center"><b><font size="1" face="Verdana" color="#FFFFFF">Transação</font></b></td>
           </tr>
           <%
  MEGA_ANTERIOR=""
  
  DO UNTIL RS1.EOF=TRUE
  RS.MOVEFIRST
  EXISTE=0
    
  DO UNTIL RS.EOF=TRUE
  set rst=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RS1("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")
  IF RST.EOF=FALSE THEN
  EXISTE=EXISTE+1
  END IF
  RS.MOVENEXT 
  LOOP
  
  IF EXISTE<>0 THEN 
  %> <tr>
                      <td width="281" height="24" bgcolor="#FFFFFF">
                         <%
      VALOR_ATUAL=""
      SET MEGA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs1("MEPR_CD_MEGA_PROCESSO"))
      VALOR_MEGA=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")
      IF TRIM(MEGA_ANTERIOR)=TRIM(VALOR_MEGA)THEN
      VALOR_ATUAL=""
      ELSE
      VALOR_ATUAL=VALOR_MEGA
      END IF
      MEGA_ANTERIOR=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")
      %> <p style="margin-top: 0; margin-bottom: 0" align="center"><font size="1" face="Verdana"><%=VALOR_ATUAL%></font></td>
                      <td width="162" height="24" bgcolor="#FFFFFF"><p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" size="1"><%=RS1("TRAN_CD_TRANSACAO")%></font></td>
                      <%
    IF COR=1 THEN
		COR=4
	ELSE
		COR=1
	END IF
	
	SELECT CASE COR
		CASE 1
			COLOR="#FAF4D8"
		CASE 4
			COLOR="#C0C0C0"
	END SELECT	
	
   rs.movefirst 
	
	DO UNTIL RS.EOF=TRUE
	set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & RS1("MEPR_CD_MEGA_PROCESSO") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & rs("FUNE_CD_FUNCAO_NEGOCIO") & "' AND TRAN_CD_TRANSACAO='" & RS1("TRAN_CD_TRANSACAO") & "'")
	IF RSTEMP.EOF=TRUE THEN
	%> <td width="98" height="24" bgcolor="<%=color%>"><p style="margin-top: 0; margin-bottom: 0" align="center"></td>
                      <%
	ELSE
	%> <td width="85" height="24" bgcolor="<%=color%>"><p align="center" style="margin-top: 0; margin-bottom: 0"><b><font size="4" color="#330099" face="Verdana, Arial, Helvetica, sans-serif">X</font></b></td>
                      <%
	 END IF
	 
	 RS.MOVENEXT
    LOOP
    
    END IF
    RS1.MOVENEXT
	 LOOP
	
	 END IF
	 END IF
	 END IF
    %> </tr>
</table>
<%IF TATUAL=0 THEN%> <p><b><font face="Arial Unicode MS" color="#663300">Nenhum Registro Encontrado</font></b></p>
<%END IF%>
</body>

<script>
MM_swapImage('loader','','../../Flash/branco.gif',1);
</script>

</html>
