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

<table width="98%" border="0">
  <tr> 
    <td height="25" colspan="3" class="campo"> <table width="89%" border="0">
        <tr> 
          <td width="15%">&nbsp;</td>
          <td width="85%" class="campob">Respons&aacute;vel Legado </td>
        </tr>
      </table></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td valign="top" class="campo">&nbsp;</td>
    <td valign="top" class="campo"> <div align="right"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="25" class="campo"> <div align="right">T&eacute;cnico:</div></td>
          </tr>
        </table>
      </div></td>
    <td colspan="3"><table width="750" border="0">
        <tr> 
          <td width="198"> <select name="lstRespTecLegGeral" size="5" class="listResponsavel" id="select">
            <% 
			  contRegistro = 0
			  rds_RespLegado.movefirst
			  Do While not rds_RespLegado.Eof and contRegistro < 10%>
            	<option value=<%=rds_RespLegado("USMA_CD_USUARIO")%>><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
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
                <td width="24"><img src="../../../imagens/continua_F01.gif" name="imgSetaDireita1" width="24" height="24" id="imgSetaDireita1" onClick="move(document.frm1.lstRespTecLegGeral,document.frm1.lstRespTecLegSel,1)" onMouseOver="mOvr(this,'../../../imagens/continua_F02.gif');" onMouseOut="mOut(this,'../../../imagens/continua_F01.gif');"></td>
              </tr>
              <tr> 
                <td><img src="../../../imagens/continua2_F01.gif" name="imgSetaEsquerda1" width="24" height="24" id="imgSetaEsquerda1" onClick="move(document.frm1.lstRespTecLegSel,document.frm1.lstRespTecLegGeral,1)" onmouseover="mOvr(this,'../../../imagens/continua2_F02.gif');" onmouseout="mOut(this,'../../../imagens/continua2_F01.gif');"></td>
              </tr>
            </table></td>
          <td width="309"> <select name="lstRespTecLegSel" size="5" class="listResponsavel">
            </select></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td valign="top">&nbsp;</td>
    <td valign="top"> <div align="right"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="25" class="campo"> <div align="right"><span class="campo">Funcional</span>: 
              </div></td>
          </tr>
        </table>
      </div></td>
    <td colspan="3"><table width="750" border="0">
        <tr> 
          <td width="198"> 
		  <select name="lstRespFunLegGeral" size="5" class="listResponsavel" id="lstRespFuncLegGeral">
				<% 
				  contRegistro = 0
				  rds_RespLegado.MoveFirst				  
				  Do While not rds_RespLegado.Eof and contRegistro < 10%>
					 <option value=<%=rds_RespLegado("USMA_CD_USUARIO")%>><%=rds_RespLegado("USMA_TX_NOME_USUARIO")%></option>
				  <% 
				  contRegistro = contRegistro + 1
			  	rds_RespLegado.movenext			  
		      Loop
		rds_RespLegado.close
		set rds_RespLegado = Nothing
		%>
            </select></td>
          <td width="30"> <table width="30" border="0">
              <tr> 
                <td width="24"><img src="../../../imagens/continua_F01.gif" name="imgSetaDireita2" width="24" height="24" id="imgSetaDireita2" onClick="move(document.frm1.lstRespFunLegGeral,document.frm1.lstRespFunLegSel,1)" onmouseover="mOvr(this,'../../../imagens/continua_F02.gif');" onmouseout="mOut(this,'../../../imagens/continua_F01.gif');"></td>
              </tr>
              <tr> 
                <td><img src="../../../imagens/continua2_F01.gif" name="imgSetaEsquerda2" width="24" height="24" id="imgSetaEsquerda2" onClick="move(document.frm1.lstRespFunLegSel,document.frm1.lstRespFunLegGeral,1)" onmouseover="mOvr(this,'../../../imagens/continua2_F02.gif');" onmouseout="mOut(this,'../../../imagens/continua2_F01.gif');"></td>
              </tr>
            </table></td>
          <td width="309"> <select name="lstRespFunLegSel" size="5" class="listResponsavel">
            </select></td>
        </tr>
      </table></td>
  </tr>
</table>
