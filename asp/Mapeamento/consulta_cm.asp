<!--#include file="conecta.asp" -->
<%
set objUSR = server.createobject("Seseg.Usuario")

curso = request("selItem")

if objUSR.GetUsuario then
	resp=objUSR.sei_chave
	lotacao=objUSR.sei_lotacao
	set objUSR = nothing
else
	response.redirect "erro.asp?op=3"
end if

set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")

set db2 = server.createobject("ADODB.CONNECTION")
db2.open "Provider=Microsoft.Jet.Oledb.4.0;data source=" & server.mappath("base.mdb")
db2.CursorLocation=3

set temp = db.execute("SELECT DISTINCT ORME_CD_ORG_MENOR FROM CLI_ORGAO WHERE USMA_CD_USUARIO='" & resp & "'")

curso = request("selItem")

if temp.eof=false then
if len(temp("ORME_CD_ORG_MENOR"))<7 then
	lotacao = temp("ORME_CD_ORG_MENOR")
else	
	lotacao = left(temp("ORME_CD_ORG_MENOR"),7)
end if
else
	lotacao=""
end if


ssql=""
ssql="SELECT DISTINCT"
ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO, USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR," 
ssql=ssql+" APOIO_LOCAL_CURSO.CURS_CD_CURSO, CURSO.CURS_TX_NOME_CURSO"
ssql=ssql+" FROM APOIO_LOCAL_CURSO"
ssql=ssql+" INNER JOIN USUARIO_MAPEAMENTO ON"
ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO = USUARIO_MAPEAMENTO.USMA_CD_USUARIO"
ssql=ssql+" INNER JOIN CURSO ON"
ssql=ssql+" APOIO_LOCAL_CURSO.CURS_CD_CURSO = CURSO.CURS_CD_CURSO"
if curso<>"XXXX" then
	ssql=ssql+" WHERE APOIO_LOCAL_CURSO.CURS_CD_CURSO='" & curso & "' AND USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & lotacao & "%' ORDER BY APOIO_LOCAL_CURSO.CURS_CD_CURSO, USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"
else
	ssql=ssql+" WHERE USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & lotacao & "%' ORDER BY APOIO_LOCAL_CURSO.CURS_CD_CURSO, USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO"	
end if

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
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="89%" id="AutoNumber2" height="514">
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
                         <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" color="#800000">Consulta Curso x Multiplicador </font></b></p>
<%if rs.eof=true then%>
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
                         <p style="margin-top: 0; margin-bottom: 0"><b><font color="#FF0000">Nenhum Registro encontrado para a Seleção
<%else%>
</font></b> 
</p>
                         <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
                         <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#D7D5CC" width="97%" id="AutoNumber4" height="43">
                                    <tr>
                                               <td width="60%" height="22" bgcolor="#666633"><b><font face="Verdana" size="2" color="#FFFFFF">Curso</font></b></td>
                                               <td width="40%" height="22" bgcolor="#666633"><b><font face="Verdana" size="2" color="#FFFFFF">Multiplicador</font></b></td>
                                    </tr>
                                    <%
                                    atual=""
                                    anterior=""
                                    do until rs.eof=true
									
									set is_cli = db.execute("SELECT * FROM CLI_ORGAO WHERE USMA_CD_USUARIO='" & resp & "'")
	
										if is_cli.eof=false then
											if len(is_cli("ORME_CD_ORG_MENOR"))=2 then
												org_cli = is_cli("ORME_CD_ORG_MENOR")
												set temp = db.execute("SELECT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO=" & org_cli)
												tabela=temp("ORGAO")			
											else
												org_cli = left(is_cli("ORME_CD_ORG_MENOR"),7)
												set temp = db.execute("SELECT ORME_SG_ORG_MENOR AS ORGAO FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & org_cli &"00000000' AND ORME_CD_STATUS='A'")
												tabela=temp("ORGAO")
											end if
										else
											tabela="XX"
										end if
									
									on error resume next
									
									ssql="SELECT * FROM [" & tabela & "] WHERE CURSO='" & rs("CURS_CD_CURSO") & "'"
									
									set tem_curso = db2.execute(ssql)
									
										if tem_curso.eof=false and err.number=0 then
									
										atual=rs("CURS_CD_CURSO")
        	                            
                                    %>
                                    <tr>
                                               <td width="60%" height="20"><font size="1" face="Verdana"><b><%=rs("CURS_TX_NOME_CURSO")%></b></font></td>
                                               <td width="40%" height="20"><font size="1" face="Verdana"><%=rs("USMA_CD_USUARIO")%> - <%=rs("USMA_TX_NOME_USUARIO")%></font></td>
                                    </tr>
                                    <%
										tem=tem+1
                                    
										anterior=rs("CURS_CD_CURSO")
										end if
									
									err.clear
                                    rs.movenext
                                    loop
                                    if tem=0 then
									%>
									<tr>
                                               
            <td width="60%" height="20"><font size="1" face="Verdana"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#990000">Nenhum 
              Registro Encontrado para a Seleção</font></b></font></td>
                                    </tr>									
									<%
									end if
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