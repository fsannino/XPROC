<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " USUA_CD_USUARIO "
str_RespLegado = str_RespLegado & " , USUA_TX_NOME_USUARIO "
str_RespLegado = str_RespLegado & " FROM dbo.USUARIO "
str_RespLegado = str_RespLegado & " ORDER BY USUA_TX_NOME_USUARIO "
set rds_RespLegado = db_Cogest.Execute(str_RespLegado)
%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<table width="98%" border="0">
  <tr> 
    <td colspan="3"><table width="89%" border="0">
        <tr> 
          <td width="6%">&nbsp;</td>
          <td width="94%" class="campob">Respons&aacute;vel Sinergia</td>
        </tr>
      </table></td>
    <td width="28%">&nbsp;</td>
    <td width="32%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1%" valign="top" class="campo">&nbsp;</td>
    <td width="15%" valign="top"><div align="right"><span class="campob">T&eacute;cnico</span>:</div></td>
    <td colspan="3">
		<select name="selRespTecSinGeral" size="1" class="listResponsavel" id="lstRespTecSineGeral">
			<option value="0">== Selecione um Responsável Sinergia - Técnico ==</option>
			<%'contRegistro = 0
			  rds_RespLegado.movefirst
			  Do While not rds_RespLegado.Eof 'and contRegistro < 10
				  if str_txtRespSinergiaTec	 = trim(rds_RespLegado("USUA_CD_USUARIO")) then%>		  
					<option value=<%=rds_RespLegado("USUA_CD_USUARIO")%> selected><%=rds_RespLegado("USUA_TX_NOME_USUARIO")%></option>
				<%else%>
					<option value=<%=rds_RespLegado("USUA_CD_USUARIO")%>><%=rds_RespLegado("USUA_TX_NOME_USUARIO")%></option>
				<%end if
				  'contRegistro = contRegistro + 1
				  rds_RespLegado.movenext			  
			  Loop	
			%>
		  </select>
	  </td>
  </tr>
  <tr> 
    <td valign="top">&nbsp;</td>
    <td valign="top"> <div align="right"><span class="campob">Funcional</span>:</div></td>
    <td colspan="3">
		<select name="selRespFunSinGeral" size="1" class="listResponsavel">
			<option value="0">== Selecione um Responsável Sinergia - Funcional ==</option>
			<%'contRegistro = 0
			  rds_RespLegado.movefirst								 
			  Do While not rds_RespLegado.Eof' and contRegistro < 10 
				  if str_txtRespSinergiaFunc = trim(rds_RespLegado("USUA_CD_USUARIO")) then%>
					<option value=<%=rds_RespLegado("USUA_CD_USUARIO")%> selected><%=rds_RespLegado("USUA_TX_NOME_USUARIO")%></option>
				<%else%>
					<option value=<%=rds_RespLegado("USUA_CD_USUARIO")%>><%=rds_RespLegado("USUA_TX_NOME_USUARIO")%></option>
				<%end if	
				  'contRegistro = contRegistro + 1 
				  rds_RespLegado.movenext
			  Loop
			  rds_RespLegado.close
			  set rds_RespLegado = Nothing%>
		  </select>
	</td>
  </tr>
</table>
