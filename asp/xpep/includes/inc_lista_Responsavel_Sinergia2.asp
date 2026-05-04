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

<table width="98%" border="0">
  <tr> 
    <td colspan="3"><table width="89%" border="0">
        <tr> 
          <td width="15%">&nbsp;</td>
          <td width="85%" class="campob">Respons&aacute;vel Sinergia</td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td valign="top" class="campo">&nbsp;</td>
    <td valign="top" class="campo"> <div align="right">T&eacute;cnico:</div></td>
    <td colspan="3"><table width="750" border="0">
        <tr> 
          <td width="198"> <select name="lstRespTecSinGeral" size="5" class="listResponsavel" id="lstRespTecSineGeral">
            <% 
			  contRegistro = 0
			  rds_RespLegado.movefirst
			  Do While not rds_RespLegado.Eof and contRegistro < 10%>
            <option value=<%=rds_RespLegado("USUA_CD_USUARIO")%>><%=rds_RespLegado("USUA_TX_NOME_USUARIO")%></option>
            <% 
			  contRegistro = contRegistro + 1
			  rds_RespLegado.movenext			  
		      Loop
		'rds_RespLegado.close
		'set rds_RespLegado = Nothing
		%>
          </select></td>
          <td width="30"> <table width="30" border="0">
              <tr> 
                <td width="24"><img src="../../../imagens/continua_F01.gif" name="imgSetaDireita3" width="24" height="24" id="imgSetaDireita3" onClick="move(document.frm1.lstRespTecSinGeral,document.frm1.lstRespTecSinSel,1)" onmouseover="mOvr(this,'../../imagens/continua_F02.gif');" onmouseout="mOut(this,'../../imagens/continua_F01.gif');"></td>
              </tr>
              <tr> 
                <td><img src="../../../imagens/continua2_F01.gif" name="imgSetaEsquerda1" width="24" height="24" id="imgSetaEsquerda1" onClick="move(document.frm1.lstRespTecSinSel,document.frm1.lstRespTecSinGeral,1)" onmouseover="mOvr(this,'../../imagens/continua2_F02.gif');" onmouseout="mOut(this,'../../imagens/continua2_F01.gif');"></td>
              </tr>
            </table></td>
          <td width="309"> <select name="lstRespTecSinSel" size="5" class="listResponsavel">
            </select></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td valign="top">&nbsp;</td>
    <td valign="top"> <div align="right"><span class="campo">Funcional</span>:</div></td>
    <td colspan="3"><table width="750" border="0">
        <tr> 
          <td width="198"> 
		  	<select name="lstRespFunSinGeral" size="5" class="listResponsavel">
			  <% 
			  contRegistro = 0
			  rds_RespLegado.movefirst								 
			  Do While not rds_RespLegado.Eof and contRegistro < 10 
				  %>
					<option value=<%=rds_RespLegado("USUA_CD_USUARIO")%>><%=rds_RespLegado("USUA_TX_NOME_USUARIO")%></option>
				  <%
				  contRegistro = contRegistro + 1 
				  rds_RespLegado.movenext
			  Loop
			  rds_RespLegado.close
			  set rds_RespLegado = Nothing
			%>
            </select>
			</td>
          <td width="30"> <table width="30" border="0">
              <tr> 
                <td width="24"><img src="../../../imagens/continua_F01.gif" name="imgSetaDireita4" width="24" height="24" id="imgSetaDireita4" onClick="move(document.frm1.lstRespFunSinGeral,document.frm1.lstRespFunSinSel,1)" onmouseover="mOvr(this,'../../imagens/continua_F02.gif');" onmouseout="mOut(this,'../../imagens/continua_F01.gif');"></td>
              </tr>
              <tr> 
                <td><img src="../../../imagens/continua2_F01.gif" name="imgSetaEsquerda4" width="24" height="24" id="imgSetaEsquerda4" onClick="move(document.frm1.lstRespFunSinSel,document.frm1.lstRespFunSinGeral,1)" onmouseover="mOvr(this,'../../imagens/continua2_F02.gif');" onmouseout="mOut(this,'../../imagens/continua2_F01.gif');"></td>
              </tr>
            </table></td>
          <td width="309"> <select name="lstRespFunSinSel" size="5" class="listResponsavel">
            </select></td>
        </tr>
      </table></td>
  </tr>
</table>
