<%
set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

str_RespLegado = ""
str_RespLegado = str_RespLegado & " SELECT  TOP 100 PERCENT "
str_RespLegado = str_RespLegado & " USMA_CD_USUARIO "
str_RespLegado = str_RespLegado & " , USMA_TX_NOME_USUARIO "
str_RespLegado = str_RespLegado & " FROM dbo.USUARIO_MAPEAMENTO "
str_RespLegado = str_RespLegado & " Where USMA_TX_MATRICULA <> 0"
str_RespLegado = str_RespLegado & " ORDER BY USMA_TX_NOME_USUARIO "
set rds_RespLegado = db_Cogest.Execute(str_RespLegado)

%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<table width="88%" border="0">
  <tr> 
    <td height="25" colspan="3"> <table width="89%" border="0">
        <tr> 
          <td width="6%">&nbsp;</td>
          <td width="94%" class="campob">Respons&aacute;vel Legado </td>
        </tr>
      </table></td>
    <td width="29%">&nbsp;</td>
    <td width="29%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1%" valign="top" class="campo">&nbsp;</td>
    <td width="17%" valign="top"> <div align="right"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="25"> <div align="right"><span class="campob">T&eacute;cnico</span>:</div></td>
          </tr>
        </table>
      </div></td>
    <td colspan="3">
		<select name="selRespTecLegGeral" size="1" class="listResponsavel" id="select2">
			<option value="0">== Selecione um Responsável Legado - Técnico ==</option>
			<% 
			'contRegistro = 0
			rds_RespLegado.movefirst
			Do While not rds_RespLegado.Eof 'and contRegistro < 10
				if str_txtRespLegadoTec = trim(rds_RespLegado("USMA_CD_USUARIO")) then%>
					<option value=<%=rds_RespLegado("USMA_CD_USUARIO")%> selected><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
				<%else%>
					<option value=<%=rds_RespLegado("USMA_CD_USUARIO")%>><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
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
    <td valign="top"> <div align="right"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="25"> <div align="right"><span class="campob">Funcional</span>: 
              </div></td>
          </tr>
        </table>
      </div></td>
    <td colspan="3">
	<select name="selRespFunLegGeral" size="1" class="listResponsavel" id="lstRespFuncLegGeral">
        <option value="0">== Selecione um Responsável Legado - Funcional ==</option>
		  <% 
		  'contRegistro = 0
		  rds_RespLegado.MoveFirst				  
		  Do While not rds_RespLegado.Eof 'and contRegistro < 10
		  if str_txtRespLegadoFunc = trim(rds_RespLegado("USMA_CD_USUARIO")) then%>
        	<option value=<%=rds_RespLegado("USMA_CD_USUARIO")%> selected><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
          <%else%>
		  	<option value=<%=rds_RespLegado("USMA_CD_USUARIO")%>><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
		  <%end if
				  'contRegistro = contRegistro + 1
			  	rds_RespLegado.movenext			  
		      Loop
		rds_RespLegado.close
		set rds_RespLegado = Nothing
		%>
      </select></td>
  </tr>
</table>
