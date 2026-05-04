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
str_Sql = str_Sql & " dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " , dbo.MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " FROM dbo.FUNCAO_NEGOCIO INNER JOIN"
str_Sql = str_Sql & " dbo.MEGA_PROCESSO ON dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = dbo.MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " WHERE dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO > 0 "
if str_MegaProcesso <> "0" then
	str_Sql = str_Sql & " AND dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso
end if
if str_SubModulo <> "0" then
	str_Sql = str_Sql & " AND dbo.FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA = " & str_SubModulo
end if
str_Sql = str_Sql & " ORDER BY dbo.FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " , dbo.FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO"

'RESPONSE.Write(str_Sql)
'RESPONSE.End()
SET rds_Func=db.EXECUTE(str_Sql)

'SET rds_Func_Tran=db.EXECUTE(str_Sql)

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
  <table border="0" width="99%">
          <tr>
            <td width="21%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Mega Processo </font></b></td>
            <td width="36%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Fun&ccedil;&atilde;o R/3 </font></b></td>
            <td width="43%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Transa&ccedil;&atilde;o</font></b></td>
          </tr>
          <%
		tem=0
		int_Tot_Reg_Processado = 0
		tem_Fun = 1
		tem_Fun_Tran = 0
		tem_Fun_Geral = 1
		tem_Fun_Tran_Geral = 0
		cd_Func_Ant=""
		cd_Func_Atu=" "	
		cd_Mega_Ant = ""
		cd_Mega_Atu = " "
		int_Conta_Existe = 0
		boo_Primeiro = true
		do until rds_Func.eof=true
			
			str_Sql = ""
			str_Sql = str_Sql & " SELECT distinct "
			str_Sql = str_Sql & " dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO"
			str_Sql = str_Sql & " , dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO"
			str_Sql = str_Sql & " , dbo.TRANSACAO.TRAN_TX_DESC_TRANSACAO"
			str_Sql = str_Sql & " FROM dbo.FUN_NEG_TRANSACAO INNER JOIN"
			str_Sql = str_Sql & " dbo.TRANSACAO ON dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO.TRAN_CD_TRANSACAO"
			str_Sql = str_Sql & " WHERE dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO = '" & rds_Func("FUNE_CD_FUNCAO_NEGOCIO") & "'"
			'response.Write(str_Sql)
			set rds_Fun_Tran = db.execute(str_Sql)
			if not rds_Fun_Tran.Eof then
				'cd_Mega_Ant = rds_Func("MEPR_CD_MEGA_PROCESSO")
				'cd_Func_Ant = rds_Func("FUNE_CD_FUNCAO_NEGOCIO")
				do while not rds_Fun_Tran.Eof
			   		str_Sql = ""
					str_Sql = str_Sql & " SELECT distinct "
					str_Sql = str_Sql & " TRAN_CD_TRANSACAO"
					str_Sql = str_Sql & " FROM  dbo.CURSO_TRANSACAO"
					str_Sql = str_Sql & " WHERE dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO <> ''"
					str_Sql = str_Sql & " and  (dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO = '" & rds_Fun_Tran("TRAN_CD_TRANSACAO") & "')"					
					'response.Write(str_Sql)
					set rds_Existe = db.execute(str_Sql)
					if rds_Existe.Eof then
					
						'response.Write(" Func Ant " & cd_Func_Ant & " - Atu " & cd_Func_Atu & "<p>")
						'response.Write(" Mega Ant " & cd_Mega_Ant & " - Atu " & cd_Mega_Atu & "<p>")

						if cor="#CCCCCC" then
							cor = "white"
						else
							cor = "#EBEBEB"	
						end if			
						
						if cd_Mega_Ant <> cd_Mega_Atu then
							if tem_Fun = -1 then 								
						%>
							  <tr>
								<td width="21%"><%=cd_Mega_Ant%></td>
								<td width="36%"><b><font face="Verdana" size="2">Total da Função: <%=tem_Fun%></font></b></td>
								<td width="43%"><%=cd_Mega_Atu%></td>
							  </tr>			
						<%
								tem_Fun = 0
							end if								
							cd_Mega = rds_Func("MEPR_CD_MEGA_PROCESSO") & " - " & rds_Func("MEPR_TX_DESC_MEGA_PROCESSO")
						else
							cd_Mega = ""
							cor = "white"
						end if
												
						if cd_Func_Ant <> cd_Func_Atu then
							if tem_Fun_Tran > 0 then 								
						%>
							  <tr>
								<td width="21%"><b></b></td>
								<td width="36%"><b></b></td>
								<td width="43%"><b><font face="Verdana" size="2">Total de Transação: <%=tem_Fun_Tran%></font></b></td>
							  </tr>			
						<%
								tem_Fun_Tran = 0
								tem_Fun = tem_Fun + 1				
								tem_Fun_Geral = tem_Fun_Geral + 1				
							end if
							cd_Func = rds_Func("FUNE_CD_FUNCAO_NEGOCIO") & " - " & rds_Func("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
						else
							cd_Func = ""
							cor = "white"
						end if

						if boo_Primeiro then
							cd_Mega = rds_Func("MEPR_CD_MEGA_PROCESSO") & " - " & rds_Func("MEPR_TX_DESC_MEGA_PROCESSO")
							cd_Func = rds_Func("FUNE_CD_FUNCAO_NEGOCIO") & " - " & rds_Func("FUNE_TX_TITULO_FUNCAO_NEGOCIO")							
							boo_Primeiro = false
						end if

						cd_Tran = ""
						cd_Tran =  rds_Fun_Tran("TRAN_CD_TRANSACAO") & " - " & rds_Fun_Tran("TRAN_TX_DESC_TRANSACAO") 

						'cd_Mega = rds_Func("MEPR_CD_MEGA_PROCESSO") & " - " & rds_Func("MEPR_TX_DESC_MEGA_PROCESSO")
						'cd_Func = rds_Func("FUNE_CD_FUNCAO_NEGOCIO") & " - " & rds_Func("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
												
						%>
          <tr>
            <td width="21%" valign="top" bgcolor="<%=cor%>"><font size="1" face="Verdana"><%=cd_Mega%></font></td>				
            <td width="36%" bgcolor="<%=cor%>"><font size="1" face="Verdana"><%=cd_Func%></font></td>
            <td width="43%" bgcolor="#FBFEDA"><font size="1" face="Verdana"><%=cd_Tran%></font></td>
          </tr>
			<%
						tem_Fun_Tran = tem_Fun_Tran + 1
						tem_Fun_Tran_Geral = tem_Fun_Tran_Geral + 1
					end if					
					rds_Existe.close					
					rds_Fun_Tran.movenext
					cd_Mega_Ant = rds_Func("MEPR_CD_MEGA_PROCESSO")
					cd_Func_Ant = rds_Func("FUNE_CD_FUNCAO_NEGOCIO")					
				Loop	
				rds_Fun_Tran.close				
		  	end if
			cd_Mega_Ant = rds_Func("MEPR_CD_MEGA_PROCESSO")
			cd_Func_Ant = rds_Func("FUNE_CD_FUNCAO_NEGOCIO")			
			rds_Func.movenext		
			if not rds_Func.Eof then
				cd_Mega_Atu = rds_Func("MEPR_CD_MEGA_PROCESSO")
				cd_Func_Atu = rds_Func("FUNE_CD_FUNCAO_NEGOCIO")
			end if
		loop
          %> 
	  <tr>
		<td width="21%"><b></b></td>
		<td width="36%"><b></b></td>
		<td width="43%"><b><font face="Verdana" size="2">Total de Transação: <%=tem_Fun_Tran%></font></b></td>
	  </tr>		
	  <% a= 1
	  if a = 2 then %>			  
	  <tr>
		<td width="21%"><b></b></td>
		<td width="36%"><b><font face="Verdana" size="2">Total da Função: <%=tem_Fun%></font></b></td>
		<td width="43%">&nbsp;</td>
	  </tr>	
	  <% end if %>		
  </table>
  <p><b>
  <%if tem_Fun_Tran=0 then%>
  <font face="Verdana" size="2" color="#800000">Nenhum Registro encontrado</font></b>
    <% else %>
</p>
  <table width="984" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="205">&nbsp;</td>
    <td width="354"><b><font face="Verdana" size="2">Total Geral Fun&ccedil;&atilde;o: <%=tem_Fun_Geral%></font></b></td>
    <td width="425"><b><font face="Verdana" size="2">Total Geral Transa&ccedil;&atilde;o: <%=tem_Fun_Tran_Geral%></font></b></td>
  </tr>
</table>
<%end if%>
</form>
</body>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>
