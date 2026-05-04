<%	
if Request("pCdSeqPCD") <> "" then
	intDesenvolvimento = Request("pCdSeqPCD") 
else
	intDesenvolvimento = 0
end if

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

if str_Acao <> "C" then
	str_DesenvAssoc = ""
	str_DesenvAssoc = str_DesenvAssoc & " SELECT DESE_CD_DESENVOLVIMENTO "
	str_DesenvAssoc = str_DesenvAssoc & " , DESE_TX_DESC_DESENVOLVIMENTO "
	str_DesenvAssoc = str_DesenvAssoc & " FROM DESENVOLVIMENTO "
	str_DesenvAssoc = str_DesenvAssoc & " ORDER BY DESE_TX_DESC_DESENVOLVIMENTO "
	set rds_DesenvAssoc = db_Cogest.Execute(str_DesenvAssoc)
end if

str_DesenvAssocSel = ""
str_DesenvAssocSel = str_DesenvAssocSel & " SELECT DESE_CD_DESENVOLVIMENTO"
str_DesenvAssocSel = str_DesenvAssocSel & " , DESE_TX_DESC_DESENVOLVIMENTO"
str_DesenvAssocSel = str_DesenvAssocSel & " FROM DESENVOLVIMENTO"
str_DesenvAssocSel = str_DesenvAssocSel & " WHERE DESE_CD_DESENVOLVIMENTO IN"
str_DesenvAssocSel = str_DesenvAssocSel & 		" (SELECT DESE_CD_DESENVOLVIMENTO"
str_DesenvAssocSel = str_DesenvAssocSel & 		" FROM XPEP_TAREFA_DESENVOLVIMENTO"
str_DesenvAssocSel = str_DesenvAssocSel & 		" WHERE PLAN_NR_SEQUENCIA_PLANO = '" & int_Plano & "'"
str_DesenvAssocSel = str_DesenvAssocSel & 		" AND PLTA_NR_SEQUENCIA_TAREFA  = '" & int_Id_TarefaProject & "'"
str_DesenvAssocSel = str_DesenvAssocSel & 		" AND PPCD_NR_SEQUENCIA_FUNC  = '" & intDesenvolvimento & "')"
str_DesenvAssocSel = str_DesenvAssocSel & " ORDER BY DESE_CD_DESENVOLVIMENTO"
'Response.WRITE str_DesenvAssocSel
'Response.END
set rds_DesenvAssocSel = db_Cogest.Execute(str_DesenvAssocSel)
%>  
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">

<table width="98%" border="0">
  <tr> 
    <td height="25" colspan="3"> <table width="89%" border="0">
        <tr> 
          <td width="7%">&nbsp;</td>
          <td width="93%" class="campob">Desenvolvimentos Associados</td>
        </tr>
      </table></td>
    <td width="33%">&nbsp;</td>
    <td width="33%">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" class="campo">&nbsp;</td>
    <td valign="top" class="campo">&nbsp;</td>
    <td colspan="3"><table width="100%"  cellpadding="0" cellspacing="0" border="0">
      <tr>
	  	<%if str_Acao <> "C" then%>
			<td width="39%" class="campo">Desenvolvimentos Existentes:</td>
			<td width="5%">&nbsp;</td>			
		<%end if%>
        <td width="40%" class="campo">Desenvolvimentos Selecionados: </td>
        <td width="16%">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr> 
    <td width="1%" valign="top" class="campo">&nbsp;</td>
    <td width="6%" valign="top" class="campo"> <div align="right"> </div></td>
    <td colspan="3"><table width="798" border="0">
        <tr> 
		  <%if str_Acao <> "C" then%>			
			  <td width="350"> 
				<select name="lstDesenvAssociados" multiple size="5" class="listResponsavel">
				  <%
				  if not rds_DesenvAssoc.bof and not rds_DesenvAssoc.eof then
					  rds_DesenvAssoc.movefirst
						Do While not rds_DesenvAssoc.Eof %>
						   <option value=<%=rds_DesenvAssoc("DESE_CD_DESENVOLVIMENTO")%>><%=rds_DesenvAssoc("DESE_CD_DESENVOLVIMENTO") & " - " & rds_DesenvAssoc("DESE_TX_DESC_DESENVOLVIMENTO")%></option>
						   <% rds_DesenvAssoc.movenext
						 Loop
				  end if
				  %>
				</select>
			  </td>
			  <td width="32"> <table width="30" border="0">
				  <tr> 
					<td width="24"><img src="../img/000030_1.gif" alt="Seleciona desenvolvimento" name="imgSetaDireita1" width="20" height="20" id="imgSetaDireita1" onClick="move(document.frm_Plano_PCD.lstDesenvAssociados,document.frm_Plano_PCD.lstDesenvAssociadosSel,1)" onMouseOver="mOvr(this,'../../../imagens/continua_F02.gif');" onMouseOut="mOut(this,'../../../imagens/continua_F01.gif');"></td>
				  </tr>
				</table></td>
		  <%end if%>
		  
          <td width="354">	  
		  	<select name="lstDesenvAssociadosSel" size="5" class="listResponsavel" multiple id="select">			
            <%
			  'if str_DesenvAssociados <> "" then
				  'vetSistemas = split(str_DesenvAssociados,",")					  
				  'rds_DesenvAssoc.movefirst
				   'Do While not rds_DesenvAssoc.Eof
				  if not rds_DesenvAssocSel.bof and not rds_DesenvAssocSel.eof then
					  rds_DesenvAssocSel.movefirst
					  Do While not rds_DesenvAssocSel.Eof 					  
						  'for i = lbound(vetSistemas) to ubound(vetSistemas)
							'if trim(vetSistemas(i)) = trim(rds_DesenvAssoc("DESE_CD_DESENVOLVIMENTO")) then%>
								<!--<option value=<%'=rds_DesenvAssoc("DESE_CD_DESENVOLVIMENTO")%>><%'=rds_DesenvAssoc("DESE_TX_DESC_DESENVOLVIMENTO")%></option>-->
								<option value=<%=rds_DesenvAssocSel("DESE_CD_DESENVOLVIMENTO")%>><%=rds_DesenvAssocSel("DESE_TX_DESC_DESENVOLVIMENTO")%></option>
							<%'end if
						  'next
						  rds_DesenvAssocSel.movenext
					  Loop				  
				   end if
			  'end if
			%>		
			</select>
		  </td>		  
			  <td width="354"><table width="30" border="0">
				<tr>
					<%if str_Acao <> "C" then%>
				  		<td width="24"><img src="../img/botao_deletar_on_03.gif" alt="Apaga desenvolvimento" name="imgSetaDireita1" width="20" height="20" id="imgSetaDireita1" onClick="deleta(document.frm_Plano_PCD.lstDesenvAssociadosSel,1)" onMouseOver="mOvr(this,'../../../imagens/continua_F02.gif');" onMouseOut="mOut(this,'../../../imagens/continua_F01.gif');"></td>
					<%else%>
						<td>&nbsp;</td>
					<%end if%>
				</tr>			
          </table></td>
        </tr>
      </table></td>
  </tr>
</table>
<%

rds_DesenvAssocSel.close
set rds_DesenvAssocSel = nothing

if str_Acao <> "C" then
	rds_DesenvAssoc.close
	set rds_DesenvAssoc = Nothing	
end if
%>
