<%
Response.Expires=0

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

dim rdsMaxPlano, strPlano, intCDPlanoGeral

strGravado = 0

strAcao				= Trim(Request("pAcao"))
if strAcao <> "" then
	strPlano 			= Request("pPlano") ' 
	intPlano			= Request("pintPlano") ' 
	intIdTaskProject	= Request("idTaskProject")
	intCdSeqFunc   	    = Request("pCdSeqFunc")
else
	strAcao				= Trim(Request("pAcao2"))
	strPlano 			= Request("pPlano2") ' 
	intPlano			= Request("pintPlano2") ' 
	intIdTaskProject	= Request("idTaskProject2")
	intCdSeqFunc   	    = Request("pCdSeqFunc2")
end if
strNomeAtividade	= Request("pNomeAtividade")
strDtInicioAtiv 	= Formatdatetime(Request("pDtInicioAtiv"), 2)
strDtFimAtiv 		= Formatdatetime(Request("pDtFimAtiv"), 2)

strMSG =  ""

'response.Write("Ação : " & strAcao & "<p>")
'response.Write("Ds_Plano : " & strPlano & "<p>")
'response.Write("Nr_Palno : " & intPlano & "<p>")
'response.Write("Nr_Tarefa : " & intIdTaskProject & "<p>")
'response.Write("Nr_Seq_Func : " & intCdSeqFunc & "<p>")

'************************************** INCLUSÃO ************************************************
if strAcao = "I" then

	blnNaoCadastraPlano = False
											
	'*** PLANO DE PARADA OPERACIONAL - INCLUSÃO
	if strPlano = "PDS" then
				
		str_FuncDesat = Request("txtFuncDesat")
		dat_DtDesliga = Request("txtDtDesliga")
		hor_HrDesliga = Right("00" & (Request("txtHrDesliga")),2)
		hor_MnDesliga = Right("00" & (Request("txtmnDesliga")),2)
		str_ProcDesl = Request("txtProcDesl")
		str_DestDados = Request("txtDestDados")

		'Response.write "aaa = FuncDesat=" & str_FuncDesat & "<br>"
		'Response.write "DtDesliga" & dat_DtDesliga & "<br>"
		'Response.write "HrDesliga" & hor_HrDesliga & "<br>"
		'Response.write "MnDesliga" & hor_MnDesliga & "<br>"
		'Response.write "ProcDesl" & str_ProcDesl & "<br>"
		'Response.write "DestDados" & str_DestDados & "<br>"
				
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtDataParada = 	split(dat_DtDesliga,"/")	
		strDia = vetDtDataParada(0)
		strMes = vetDtDataParada(1)
		strAno = vetDtDataParada(2)	
		dat_DtDesliga = strMes & "/" & strDia & "/" & strAno 					

		'Response.write "DtDesliga" & dat_DtDesliga & "<br>"
		'Response.end	

		'*** Seleciona o cod para a Nova Seq para Sub-Atividade de PDS
		intCdSeqFunc = 0	
		str_SQL = ""
		str_SQL = str_SQL & " SELECT "
		str_SQL = str_SQL & " MAX(PPDS_NR_SEQUENCIA_FUNC) AS int_Max_SeqFunc "
		str_SQL = str_SQL & " FROM XPEP_PLANO_TAREFA_PDS_FUNC "
		str_SQL = str_SQL & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		str_SQL = str_SQL & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject
		Set rdsMaxFunc = db_Cogest.Execute(str_SQL)			
		if isnull(rdsMaxFunc("int_Max_SeqFunc")) then
			intCdSeqFunc = 1
		else
			intCdSeqFunc = rdsMaxFunc("int_Max_SeqFunc") + 1
		end if
		rdsMaxFunc.Close	
		set rdsMaxFunc = nothing
													
		strSQL_Nova_Funcionalidade = ""
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & " INSERT INTO XPEP_PLANO_TAREFA_PDS_FUNC ( "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & " PLAN_NR_SEQUENCIA_PLANO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PLTA_NR_SEQUENCIA_TAREFA "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_NR_SEQUENCIA_FUNC "		
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_FUNC_DESATIVADAS "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_DT_DESLIGAMENTO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_HR_DESLIGAMENTO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_PROC_DESLIGAMENTO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", PPDS_TX_DEST_DD_TEMPO_RETENCAO "			
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", ATUA_TX_OPERACAO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", ATUA_CD_NR_USUARIO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", ATUA_DT_ATUALIZACAO "
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & " ) Values( " 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & intPlano 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", " & intIdTaskProject 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", " & intCdSeqFunc
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & UCase(str_FuncDesat) & "'"
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & dat_DtDesliga & "'"
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '"  & hor_HrDesliga & ":" & hor_MnDesliga & "'" 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & UCase(str_ProcDesl) & "'" 
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", '" & UCase(str_DestDados) & "'"			
		strSQL_Nova_Funcionalidade = strSQL_Nova_Funcionalidade & ", 'I','" & Session("CdUsuario") & "',GETDATE())" 		
	
		'Response.write strSQL_Nova_Funcionalidade
		'Response.end
	
		on error resume next
			db_Cogest.Execute(strSQL_Nova_Funcionalidade)
	
	'*** PLANO DE COMUNICAÇÃO - INCLUSÃO
	elseif strPlano = "PCM" then			
		
											
	end if

	if err.number = 0 then
	
		strMSG = "Registro incluido com sucesso."
		strGravado = 1
		'set correio = server.CreateObject("Persits.MailSender")
		'correio.host = "harpia.petrobras.com.br"
		 
		'correio.from="cursos@xproc.com"
	
		'correio.AddAddress "robson_28.infotec@petrobras.com.br"	     			
		'correio.AddAddress "gustavogomes.bearingpoint@petrobras.com.br"
		'correio.AddAddress "sergio.salomao.bearingpoint@petrobras.com.br"
						
		'correio.Subject="Inclusão de Novo Curso"
					
		'data_Atual=day(date) &"/"& month(date) &"/"& year(date)
						
		'correio.Body=" O Curso '" & UCASE(valor_cod) & "' foi INCLUÍDO em  " & DATA_ATUAL & " / POR : " & Session("CdUsuario")					
		'correio.send
	else
		strMSG = "Houve um erro no cadastro do registro."
	end if	 

'************************************** ALTERAÇÃO ************************************************	
elseif strAcao = "A" then
			
	'*** PLANO DE PARADA OPERACIONAL - ALTERAÇÃO
	if strPlano = "PDS" then
				
		str_FuncDesat = Request("txtFuncDesat")
		dat_DtDesliga = Request("txtDtDesliga")
		hor_HrDesliga = Right("00" & (Request("txtHrDesliga")),2)
		hor_MnDesliga = Right("00" & (Request("txtmnDesliga")),2)
		str_ProcDesl = Request("txtProcDesl")
		str_DestDados = Request("txtDestDados")

		'Response.write "FuncDesat=" & str_FuncDesat & "<br>"
		'Response.write "DtDesliga" & dat_DtDesliga & "<br>"
		'Response.write "HrDesliga" & hor_HrDesliga & "<br>"
		'Response.write "MnDesliga" & hor_MnDesliga & "<br>"
		'Response.write "ProcDesl" & str_ProcDesl & "<br>"
		'Response.write "DestDados" & str_DestDados & "<br>"
				
		strDia = ""		
		strMes = ""
		strAno = ""
		vetDtDataParada = 	split(dat_DtDesliga,"/")	
		strDia = vetDtDataParada(0)
		strMes = vetDtDataParada(1)
		strAno = vetDtDataParada(2)	
		dat_DtDesliga = strMes & "/" & strDia & "/" & strAno 					
													
		strSQL_AltPlanoPDS_Func = ""
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " UPDATE XPEP_PLANO_TAREFA_PDS_FUNC SET"		
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " PPDS_TX_FUNC_DESATIVADAS = '" & UCase(str_FuncDesat) & "'"	
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_DT_DESLIGAMENTO = '" & dat_DtDesliga & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_TX_HR_DESLIGAMENTO = '" & hor_HrDesliga & ":" & hor_MnDesliga & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_TX_PROC_DESLIGAMENTO ='" & UCase(str_ProcDesl) & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", PPDS_TX_DEST_DD_TEMPO_RETENCAO = '" & UCase(str_DestDados) & "'"		
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", ATUA_TX_OPERACAO = 'A'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", ATUA_CD_NR_USUARIO = '" & Session("CdUsuario") & "'"
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & ", ATUA_DT_ATUALIZACAO = GETDATE() "
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject 
		strSQL_AltPlanoPDS_Func = strSQL_AltPlanoPDS_Func & " AND PPDS_NR_SEQUENCIA_FUNC = " & intCdSeqFunc 
		'response.Write(strSQL_AltPlanoPDS_Func)
		'Response.end
			
		on error resume next
			db_Cogest.Execute(strSQL_AltPlanoPDS_Func)			
		
	'*** PLANO DE COMUNICAÇÃO- ALTERAÇÃO
	elseif strPlano = "PCM" then								
			
	end if

	if err.number = 0 then		
		strMSG = "Detalhamento alterado com sucesso."
		strGravado = 1
		'set correio = server.CreateObject("Persits.MailSender")
		'correio.host = "harpia.petrobras.com.br"
		 
		'correio.from="cursos@xproc.com"
	
		'correio.AddAddress "robson_28.infotec@petrobras.com.br"	     			
		'correio.AddAddress "gustavogomes.bearingpoint@petrobras.com.br"
		'correio.AddAddress "sergio.salomao.bearingpoint@petrobras.com.br"
						
		'correio.Subject="Inclusão de Novo Curso"
					
		'data_Atual=day(date) &"/"& month(date) &"/"& year(date)
						
		'correio.Body=" O Curso '" & UCASE(valor_cod) & "' foi INCLUÍDO em  " & DATA_ATUAL & " / POR : " & Session("CdUsuario")					
		'correio.send
	else
		strMSG = "Houve um erro na alteração do detalhamento."
	end if	 

elseif strAcao = "E" then
			
	'*** PLANO DE PARADA OPERACIONAL - ALTERAÇÃO
	if strPlano = "PDS" then
													
		strSQL_ExcPlanoPDS_Func = ""
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " DELETE FROM XPEP_PLANO_TAREFA_PDS_FUNC "		
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " WHERE PLAN_NR_SEQUENCIA_PLANO = " & intPlano
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " AND PLTA_NR_SEQUENCIA_TAREFA = " & intIdTaskProject 
		strSQL_ExcPlanoPDS_Func = strSQL_ExcPlanoPDS_Func & " AND PPDS_NR_SEQUENCIA_FUNC = " & intCdSeqFunc 
		'response.Write(strSQL_ExcPlanoPDS_Func)
		'Response.end
			
		on error resume next
			db_Cogest.Execute(strSQL_ExcPlanoPDS_Func)			
		
	'*** PLANO DE COMUNICAÇÃO- ALTERAÇÃO
	elseif strPlano = "PCM" then								
			
	end if

	if err.number = 0 then		
		strMSG = "Detalhamento excluído com sucesso."
		strGravado = 1
		'set correio = server.CreateObject("Persits.MailSender")
		'correio.host = "harpia.petrobras.com.br"
		 
		'correio.from="cursos@xproc.com"
	
		'correio.AddAddress "robson_28.infotec@petrobras.com.br"	     			
		'correio.AddAddress "gustavogomes.bearingpoint@petrobras.com.br"
		'correio.AddAddress "sergio.salomao.bearingpoint@petrobras.com.br"
						
		'correio.Subject="Inclusão de Novo Curso"
					
		'data_Atual=day(date) &"/"& month(date) &"/"& year(date)
						
		'correio.Body=" O Curso '" & UCASE(valor_cod) & "' foi INCLUÍDO em  " & DATA_ATUAL & " / POR : " & Session("CdUsuario")					
		'correio.send
	else
		strMSG = "Houve um erro na alteração do detalhamento."
	end if	 

end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333; text-decoration: none}
a:hover {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333;  text-decoration: underline}
-->
</style>
<link href="/css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="../../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->

<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
<script language="JavaScript">	

</script>
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../../img/000003.gif" width="19" height="21"></td>
			    <td width="202" height="21">
					<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						<strong>
						</strong>
					</font>
			    </td>
			    <td>&nbsp;</td>
		      </tr>
			</table>
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../../indexA_xpep.asp"><img src="../../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" -->   
	<table width="849" height="207" border="0" cellpadding="5" cellspacing="5">
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
			    <td height="29"></td>
			    <td height="29" valign="middle" align="left"></td>
			    <td height="29" valign="middle" align="left" colspan="2">&nbsp;</td>
      </tr>
			  <tr>
				
		  <td width="117" height="29"></td>
				
		  <td width="53" height="29" valign="middle" align="left"></td>
				
		  <td height="29" valign="middle" align="left" colspan="2"> 
		  <%if err.number=0 then%>
		  <b><font face="Verdana" color="#330099" size="2"><%=strMSG%></font></b> 
		  </td>				
			  </tr>
		  <%else%>    
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  <b><font face="Verdana" size="2" color="#800000"><%=strMSG%> - <%=err.description%></font></b> 
		  </td>
			  </tr>
			  <%end if%>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" width="32"> 
			<a href="../../../indexA_xpep.asp">
			<img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
				
		  <td height="1" valign="middle" align="left" width="629"> 
			<font face="Verdana" color="#330099" size="2">Retornar para Tela Principal</font></td>
			  </tr>
			  <% if strPlano = "PDS" then 
			        if str_Acao = "I" then %>
			  <tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="inclui_altera_plano_pds_func.asp?pAcao=I&pPlano=<%=intPlano%>&pIdTaskProject=<%=intIdTaskProject%>"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"> <font face="Verdana" color="#330099" size="2">Retornar - Cadastro de de mais uma Funcionalidade</font></td>
	  </tr>
	  <% end if %>
			  <tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="inclui_altera_plano_pds.asp?pAcao=A&pPlano=<%=intPlano%>&pTArefa=<%=intIdTaskProject%>"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"><font face="Verdana" color="#330099" size="2">Retornar - Detalhamento de PDS </font></td>
      </tr>
	  <% else %>
			  <tr>
			    <td height="1"></td>
			    <td height="1" valign="middle" align="left"></td>
			    <td width="32" height="1" align="left" valign="middle"><a href="seleciona_plano.asp"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
			    <td height="1" valign="middle" align="left"><font face="Verdana" color="#330099" size="2">Retornar - Cadastramento de mais um PCM para este Plano </font></td>
      </tr>
	  <% end if %>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" width="32"> 
			<a href="seleciona_plano.asp">
			</a><a href="seleciona_plano.asp"><img src="../../../imagens/download_01.gif" width="18" height="18" border="0" align="right"></a></td>
				
		  <td height="1" valign="middle" align="left" width="629"> 
			<font face="Verdana" color="#330099" size="2">Retornar - Seleção para Detalhamento das Atividades</font></td>
			  </tr>
			  <tr>
				
		  <td width="117" height="1"></td>
				
		  <td width="53" height="1" valign="middle" align="left"></td>
				
		  <td height="1" valign="middle" align="left" colspan="2"> 
		  </td>
			  </tr>
			</table>
  <table width="614" border="0">
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td class="campob">&nbsp;</td>
      <td>&nbsp;</td>
    </tr>	
	<%			
	if (strAcao = "I" and strPlano <> "PAC" and strGravado = 1) and (strAcao = "I" and strPlano <> "PCM" and strGravado = 1) then
	%>	
		<tr>
		  <td width="2">&nbsp;</td>
		  <td width="271">&nbsp;</td>
		  <td width="45">&nbsp;</td>
		  <td width="235" valign="top" class="campob">&nbsp;</td>
		  <td width="39">&nbsp;</td>
		</tr>
	<%
	end if
	%>
  </table>
  <p>&nbsp;</p>
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>
