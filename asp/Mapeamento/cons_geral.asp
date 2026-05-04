<!--#include file="conecta.asp" -->
<%
'set objUSR = server.createobject("Seseg.Usuario")

chave = request("selFunc")
mega = request("selMega")

'if objUSR.GetUsuario then
'	chave=objUSR.sei_chave
'	lotacao=objUSR.sei_lotacao
'	nome=objUSR.sei_nome
'	set objUSR = nothing
'else
'	response.redirect "erro.asp?op=3"
'end if

set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")

set db2 = server.createobject("ADODB.CONNECTION")
db2.open "Provider=Microsoft.Jet.Oledb.4.0;data source=" & server.mappath("base.mdb")

tipo_cons="CONSULTA DE TODOS OS MAPEADOS POR ÓRGÃO"
legenda = "SELECIONE O ÓRGÃO DESEJADO"
legenda2 = "SELECIONE O CURSO DESEJADO"
pagina = "gera_cons_geral.asp"

set rs = db.execute("SELECT AGLU_CD_AGLUTINADO, AGLU_SG_AGLUTINADO FROM ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")
set rs2 = db.execute("SELECT CURS_CD_CURSO, CURS_TX_NOME_CURSO FROM CURSO ORDER BY CURS_TX_NOME_CURSO")

if len(request("selItem")) < 1 then
	Item = 88
end if

	ssql="SELECT DISTINCT AGLU_SG_AGLUTINADO AS ORME_SG_ORG_MENOR FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & Item
	
	response.write ssql

	set temp = db.execute(ssql)
	TABELA = "[" & temp("ORME_SG_ORG_MENOR") &"]"
	
	set rs1 = db2.execute("SELECT DISTINCT CURSO FROM " & TABELA & " ORDER BY CURSO")

%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<script>
function envia()
{
window.location = 'cons_geral.asp?selOrgao='+document.frm1.selItem.value
}
</script>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body topmargin="0" leftmargin="0" link="#000080" vlink="#000080" alink="#000080">
<form name="frm1" method="post" action="<%=pagina%>">

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="89%" id="AutoNumber2" height="514">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2"><img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="445" valign="top"><img border="0" src="lado.jpg" width="83" height="429"></td>
                      <td width="87%" height="445" valign="top">
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="93%" id="AutoNumber3" height="65">
          <tr> 
            <td width="19%" height="130" align="center" colspan="2"><img border="0" src="mult_c.jpg" align="right"></td>
            <td width="81%" height="130" align="left"><font face="Verdana" color="#800000"><b><%=tipo_cons%></b></font></td>
          </tr>
          <tr> 
            <td width="11%" height="29" align="center">&nbsp;</td>
            <td width="8%" height="29" align="center"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
            <td width="81%" height="29" align="left"><b><font face="Verdana" size="2"><%=legenda%></font></b></td>
          </tr>
          <tr> 
            <td width="11%" height="24" align="center">&nbsp;</td>
            <td width="8%" height="24" align="center">&nbsp;</td>
            <td width="81%" height="24" align="left"> 
              <select size="1" name="selItem" style="font-family: Verdana; font-size: 7 pt" onChange="envia()">
                <%
                                       do until rs.eof=true
                                       %>
                <option value="<%=rs.fields(0).value%>"><%=rs.fields(1).value%></option>
                <%                                    
                                       rs.movenext
                                       loop
                                       %>
              </select>
            </td>
          </tr>
          <tr> 
            <td width="11%" height="26" align="center">&nbsp;</td>
            <td width="8%" height="26" align="center"><b><font face="Verdana" size="2"><img border="0" src="../../imagens/flecha.gif"></font></b></td>
            <td width="81%" height="26" align="left"><b><font face="Verdana" size="2"><%=legenda2%></font></b></td>
          </tr>
          <tr> 
            <td width="11%" height="26" align="center">&nbsp;</td>
            <td width="8%" height="26" align="center">&nbsp;</td>
            <td width="81%" height="26" align="left"> 
              <select size="1" name="selCurso" style="font-family: Verdana; font-size: 7 pt">
			  <option value="XXXX">=== TODOS ===</option>
                <%
                                       do until rs1.eof=true
                                       %>
                <option <%=sel%> value="<%=rs1.fields(0).value%>"><%=rs1.fields(0).value%></option>
                <%                                    
                                       rs1.movenext
                                       loop
                                       %>
              </select>
            </td>
          </tr>
          <tr> 
            <td width="100%" height="45" align="center" colspan="3"> 
              <p><a href="javascript:history.go(-1)"><img border="0" src="voltar.gif"></a>&nbsp;&nbsp; 
                <a href="#" onClick="document.frm1.submit()"><img border="0" src="enviar.gif"></a>
            </td>
          </tr>
        </table>
                      </td>
           </tr>
</table>
</form>
</body>

</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>