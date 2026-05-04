<!--#include file="conecta.asp" -->
<%
'set objUSR = server.createobject("Seseg.Usuario")

set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")

lotacao = request("selItem")

ssql=""
ssql="SELECT DISTINCT"
ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO, USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, ORGAO_MENOR.ORME_SG_ORG_MENOR, USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR," 
ssql=ssql+" APOIO_LOCAL_CURSO.CURS_CD_CURSO, CURSO.CURS_TX_NOME_CURSO"
ssql=ssql+" FROM APOIO_LOCAL_CURSO"
ssql=ssql+" INNER JOIN USUARIO_MAPEAMENTO ON"
ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO = USUARIO_MAPEAMENTO.USMA_CD_USUARIO"
ssql=ssql+" INNER JOIN CURSO ON"
ssql=ssql+" APOIO_LOCAL_CURSO.CURS_CD_CURSO = CURSO.CURS_CD_CURSO"
ssql=ssql+" INNER JOIN ORGAO_MENOR ON"
ssql=ssql+" USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR = ORGAO_MENOR.ORME_CD_ORG_MENOR"
ssql=ssql+" WHERE USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & lotacao & "%' ORDER BY ORGAO_MENOR.ORME_SG_ORG_MENOR, USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, APOIO_LOCAL_CURSO.CURS_CD_CURSO"

set rs = db.execute(ssql)
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body topmargin="0" leftmargin="0" link="#000080" vlink="#000080" alink="#000080">
<form>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="94%" id="AutoNumber2" height="514">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2"><img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="445" valign="top"><img border="0" src="lado.jpg" width="83" height="429"></td>
                      <td width="87%" height="445" valign="top">
                         <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="76%" id="AutoNumber3" height="24">
                         <tr>
                                    <td width="92%" height="1" align="center"><p align="right"><a href="javascript:history.go(-1)"><img border="0" src="voltar.gif" align="left"></a></td>
                                    <%
                                    IF RS.EOF=FALSE THEN
                                    %>
                                    <td width="8%" height="1" align="center"><a href="javascript:print()"><img border="0" src="../Apoio/impressao.jpg" width="29" height="30" align="middle" alt="Imprimir Consulta"></a></td>
                                    <%
                                    END IF
                                    %>
                         </tr>
                         </table>
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
                         
        <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#800000">Consulta 
          de todos os Mapeados por Órgão</font></b></p>
<%if rs.eof=true then%>
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
                         <p style="margin-top: 0; margin-bottom: 0"><b><font color="#FF0000">Nenhum Registro encontrado para a Seleção
<%else%>
</font></b> 
</p>
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
                         <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#D7D5CC" width="763" id="AutoNumber4" height="43">
                                    <tr>
                                               <td width="177" height="22" bgcolor="#666633"><b><font face="Verdana" size="2" color="#FFFFFF">Lotação</font></b></td>
                                               <td width="235" height="22" bgcolor="#666633"><b><font size="2" face="Verdana" color="#FFFFFF">Multiplicador</font></b></td>
                                               <td width="347" height="22" bgcolor="#666633"><b><font size="2" face="Verdana" color="#FFFFFF">Curso</font></b></td>
                                    </tr>
                                    <%
                                    atual=""
                                    anterior=""
                                    do until rs.eof=true
                                    atual=rs.fields(2).value
                                                                        
                                    if atual<>anterior then
                                    	lot_user = rs.fields(2).value	
                                    else
                                    	lot_user = " "
                                    end if
                                    
                                    usuario=rs("USMA_CD_USUARIO") & " - " &  rs("USMA_TX_NOME_USUARIO")
                                    %>
                                    <tr>
                                               <td width="177" height="20"><font size="1" face="Verdana"><b><%=lot_user%></b></font></td>
                                               <td width="235" height="20"><font size="1" face="Verdana"><%=usuario%></font></td>
                                               <td width="347" height="20"><font size="1" face="Verdana"><%=rs("CURS_TX_NOME_CURSO")%></font></td>
                                    </tr>
                                    <%
                                    anterior=rs.fields(2).value
                                    rs.movenext
                                    loop
                                    %>
                                    
                         </table>
                      </td>
           </tr>
</table>
<%end if%>
</form>
</body>

</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>