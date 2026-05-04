<%
response.Buffer=false
Server.ScriptTimeOut=99990000

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_MegaProcesso = request("selMegaProcesso")
str_SubModulo = request("selSubModulo") 

'response.Write(" mega " & str_MegaProcesso)
'response.Write(" sub " & str_SubModulo)


str_Sql = " SELECT distinct "
str_Sql = str_Sql & " dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO "
str_Sql = str_Sql & " , dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
str_Sql = str_Sql & " , dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO "
str_Sql = str_Sql & " , dbo.TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_Sql = str_Sql & " , dbo.FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO AS Mega_Dono "
str_Sql = str_Sql & " FROM  dbo.FUN_NEG_TRANSACAO INNER JOIN"
str_Sql = str_Sql & " dbo.FUNCAO_NEGOCIO ON "
str_Sql = str_Sql & " dbo.FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO = dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO INNER JOIN"
str_Sql = str_Sql & " dbo.MEGA_PROCESSO ON dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO INNER JOIN"
str_Sql = str_Sql & " dbo.TRANSACAO ON dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO.TRAN_CD_TRANSACAO INNER JOIN"
str_Sql = str_Sql & " dbo.FUNCAO_NEGOCIO_SUB_MODULO ON "
str_Sql = str_Sql & " dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO" 
str_Sql = str_Sql & " WHERE dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO > 0 "
if str_MegaProcesso <> "0" then
	str_Sql = str_Sql & " AND dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
end if
if str_SubModulo <> "0" then
	str_Sql = str_Sql & " AND dbo.FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA = " & str_SubModulo
end if
str_Sql = str_Sql & " ORDER BY dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " , dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO "

'RESPONSE.Write(str_Sql)
'RESPONSE.End()
SET rds_Func_Tran=db.EXECUTE(str_Sql)

%>
<html>
<head>
<script>
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
</script>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" link="#800000" vlink="#800000" alink="#800000">
<form method="POST" action="" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
      <tr>
        <td bgcolor="#330099" width="39" valign="middle" align="center">
          <div align="center">
            <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>    
        </div></td>
        <td bgcolor="#330099" width="36" valign="middle" align="center">
          <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div></td>
        <td bgcolor="#330099" width="27" valign="middle" align="center">
          <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div></td>
      </tr>
      <tr>
        <td bgcolor="#330099" height="12" width="39" valign="middle" align="center">
          <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div></td>
        <td bgcolor="#330099" height="12" width="36" valign="middle" align="center">
          <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div></td>
        <td bgcolor="#330099" height="12" width="27" valign="middle" align="center">
          <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div></td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"></td>
          <td width="50"></td>
          <td width="26"></td>
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
        
  <table width="100%" height="38" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="71%">&nbsp;</td>
      <td width="29%">&nbsp;</td>
    </tr>
    <tr>
      <td>
        <div align="center">
          <p align="center" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório
          de Transa&ccedil;&atilde;o n&atilde;o associadas ao curso </font>
        </div>
      </td>
      <td><font face="Verdana" color="#330099" size="3"><img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></font></td>
    </tr>
  </table>
  <table width="100%" height="38" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>
        <div align="center"><font face="Verdana" color="#330099" size="3"><%=str_Tit_Curso%></font>      </div></td>
    </tr>
  </table>
  <table border="0" width="81%">
          <tr>
            <td width="21%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Mega Processo </font></b></td>
            <td width="21%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Fun&ccedil;&atilde;o R/3 </font></b></td>
            <td width="45%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Transa&ccedil;&atilde;o</font></b></td>
          </tr>
          <%
		tem=0
		int_Tot_Reg_Processado = 0
		cd_Func_Ant=""
		cd_Func_Atu=" "			
		int_Conta_Existe = 0
		do until rds_Func_Tran.eof=true
		'response.Write(int_Tot_Reg_Processado & "<p>")
			if int_Tot_Reg_Processado > 10 then
				'response.Flush()
				int_Tot_Reg_Processado = 0
			else
				int_Tot_Reg_Processado = int_Tot_Reg_Processado + 1
			end if	
			cd_Mega_Ant = rds_Func_Tran("MEPR_CD_MEGA_PROCESSO")
			cd_Func_Ant = rds_Func_Tran("FUNE_CD_FUNCAO_NEGOCIO")
			
			str_Sql = ""
			str_Sql = str_Sql & " SELECT distinct "
			str_Sql = str_Sql & " CURS_CD_CURSO, TRAN_CD_TRANSACAO"
			str_Sql = str_Sql & " FROM  dbo.CURSO_TRANSACAO"
			str_Sql = str_Sql & " WHERE (dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO = '" & rds_Func_Tran("TRAN_CD_TRANSACAO") & "')"					
			response.Write(str_Sql)
			set rds_Existe = db.execute(str_Sql)
			if not rds_Existe.Eof then
				int_Conta_Existe = int_Conta_Existe + 1
			end if
		
          %>
          <tr>
			<%
			if int_Conta_Existe = 0 then
				if cd_Mega_Ant <> cd_Mega_Atu then
					cd_Mega = rds_Func_Tran("MEPR_CD_MEGA_PROCESSO") & " - " & rds_Func_Tran("MEPR_TX_DESC_MEGA_PROCESSO")
				else
					cd_Mega = ""
				end if
				
				if cd_Func_Ant <> cd_Func_Atu then
					cd_Func = rds_Func_Tran("FUNE_CD_FUNCAO_NEGOCIO") & " - " & rds_Func_Tran("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
				else
					cd_Func = ""
				end if
				
				if nome1="" then
					cor="white"
				else
					cor="#CCCCCC"	
				end if			
				%>
            <td width="21%" bgcolor="<%=cor%>"><font size="1" face="Verdana"><%=cd_Mega%></font></td>				
            <td width="21%" bgcolor="<%=cor%>"><font size="1" face="Verdana"><%=cd_Func%></font></td>
            <td width="45%" bgcolor="#FFFFEA"><font size="1" face="Verdana"><%=rds_Func_Tran("TRAN_CD_TRANSACAO")%>-<%=rds_Func_Tran("TRAN_TX_DESC_TRANSACAO")%></font></td>
          </tr>
          <%
		  	end if
			tem=tem+1			
			cd_Mega_Ant = rds_Func_Tran("MEPR_CD_MEGA_PROCESSO")
			cd_Func_Ant = rds_Func_Tran("FUNE_CD_FUNCAO_NEGOCIO")
			rds_Func_Tran.movenext		
			if not rds_Func_Tran.Eof then
				cd_Mega_Atu = rds_Func_Tran("MEPR_CD_MEGA_PROCESSO")
				cd_Func_Atu = rds_Func_Tran("FUNE_CD_FUNCAO_NEGOCIO")
			end if
          loop
          %>
          
  </table>
<b>
<%if tem=0 then%>
<font face="Verdana" size="2" color="#800000">Nenhum Registro encontrado</font></b>
<%end if%>
</form>
</body>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>
