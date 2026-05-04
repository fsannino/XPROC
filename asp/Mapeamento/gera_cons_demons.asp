<%
orgao = request("selOrgao")
atual = request("selCurso")

if orgao="87" then
	response.redirect "gera_cons_demons_ep.asp?selOrgao=" & orgao & "&selCurso=" & atual
end if
%>
<!--#include file="conecta.asp" -->
<%
Response.Buffer=false

set db = server.createobject("ADODB.CONNECTION")
db.open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation=3

set db2 = server.createobject("ADODB.CONNECTION")
db2.open "Provider=Microsoft.Jet.Oledb.4.0;data source=" & server.mappath("base.mdb")
db2.CursorLocation=3

set temp = db.execute("SELECT DISTINCT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & orgao & "'")

tabela = "[" & temp("ORGAO") & "]"

sigl_orgao = temp("ORGAO")

ssql="SELECT DISTINCT ORGAO FROM " & TABELA & " ORDER BY ORGAO"

on error resume next
set rs = db2.execute(ssql)

if err.number<>0 then
	erro = 1
	err.clear
else
	erro = 0
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="pt-br">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>

<script>
function grafico()
{
if(document.frm1.indicados.value!=0)
{
window.open("r03.asp?orgao="+document.frm1.orgao.value+"&resto="+document.frm1.resto.value+"&mapeados="+document.frm1.mapeados.value+"&indicados="+document.frm1.indicados.value,"_blank","width=460,height=330,history=0,scrollbars=1,titlebar=0,resizable=0")
}
else
{
alert('Gráfico não disponível!');
return;
}
}
</script>

<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

<body topmargin="0" leftmargin="0" link="#000080" vlink="#000080" alink="#000080">
<form name="frm1">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="94%" id="AutoNumber2" height="514">
           <tr>
                      <td width="100%" height="68" valign="top" colspan="2"><img border="0" src="topo.jpg"></td>
           </tr>
           <tr>
                      <td width="13%" height="445" valign="top">&nbsp;</td>
                      <td width="87%" height="445" valign="top">
                         
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber4" height="22">
          <tr> 
            <td width="50%"><a href="javascript:history.go(-1)"><img border="0" src="voltar.gif"></a></td>
            <td width="20%"> 
              <p align="center"><a href="javascript:grafico()"><img border="0" src="graf.jpg" width="35" height="26" alt="Visualizar Gráfico"></a></td>
            <td width="30%"><a href="javascript:print()"><img border="0" src="../Apoio/impressao.jpg" width="29" height="29" alt="Imprimir Relatório"></a></td>
          </tr>
        </table>
                         <p><b><font face="Verdana" color="#800000">Consulta Demonstrativa por Curso</font></b></p>
                         
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#D7D5CC" width="812" id="AutoNumber3" height="39">
          <tr>
                                               
            <td width="132" height="19" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Órgão</font></b></td>
                                               
            <td width="428" height="19" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Curso</font></b></td>
                                               
            <td width="70" height="19" align="center" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Vagas 
              Indicadas</font></b></td>
                                               
            <td width="64" height="19" align="center" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Vagas 
              Mapeadas</font></b></td>
                                               
            <td width="106" height="19" align="center" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Pend&ecirc;ncias 
              em Vagas Indicadas</font></b></td>
                                    </tr>
                                    
                                    <%
                                    if erro = 0 then
                                    
                                    do until rs.eof=true
                                    
                                    if request("selCurso")="XXXX" then
	                                    set temp = db2.execute("SELECT DISTINCT CURSO, INDICADOS FROM " & TABELA & " WHERE ORGAO='" & rs("orgao") & "' ORDER BY CURSO")
	                                else
	                                    set temp = db2.execute("SELECT DISTINCT CURSO, INDICADOS FROM " & TABELA & " WHERE CURSO = '" & request("selCurso") & "' AND ORGAO='" & rs("orgao") & "' ORDER BY CURSO")	                                
	                                end if
                                    
                                    orgao=rs("orgao")
                                    
									orgao_2=orgao
                                    
                                    do until temp.eof=true
                                    
                                    set temp2 = db.execute("SELECT * FROM CURSO WHERE CURS_CD_CURSO='" & temp("curso") & "'")
                                    
                                    nome_curso = temp2("CURS_TX_NOME_CURSO")
                                    
                                    INDICADOS = temp("INDICADOS")
                                    
                                    %>
                                    <tr>
                                               
            <td width="132" height="19"><font face="Verdana" size="1"><b><%=orgao%></b></font></td>
                                               
            <td width="428" height="19"><b><font face="Verdana" size="1"><%=nome_curso%></font></b></td>
                                               
            <td width="70" height="19" align="center"><font face="Verdana" size="2"><%=INDICADOS%></font></td>
												<%
												
												INDICA_GRAF = INDICA_GRAF + INDICADOS
												
												set org = db.execute("SELECT * FROM ORGAO_MENOR WHERE ORME_SG_ORG_MENOR='" & trim(orgao_2) & "' AND ORME_CD_STATUS='A'")
												
												org_pesquisa = left(org("ORME_CD_ORG_MENOR"),7)
												
												if len(org_pesquisa)=0 then
													org_pesquisa = request("selOrgao")
												end if
												
												if left(org_pesquisa,2)="55" then
													org_pesquisa = left(org_pesquisa,2)
												end if
												                                               
												ssql=""
												ssql="SELECT DISTINCT"
												ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO,"
												ssql=ssql+" APOIO_LOCAL_CURSO.CURS_CD_CURSO "
												ssql=ssql+" FROM APOIO_LOCAL_CURSO"
												ssql=ssql+" INNER JOIN USUARIO_MAPEAMENTO ON"
												ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO = USUARIO_MAPEAMENTO.USMA_CD_USUARIO"
												ssql=ssql+" INNER JOIN CURSO ON"
												ssql=ssql+" APOIO_LOCAL_CURSO.CURS_CD_CURSO = CURSO.CURS_CD_CURSO"
												ssql=ssql+" WHERE APOIO_LOCAL_CURSO.CURS_CD_CURSO='" & temp("curso") & "' AND USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & org_pesquisa & "%' ORDER BY APOIO_LOCAL_CURSO.CURS_CD_CURSO"
												
												set mapea = db.execute(ssql)
												
												MAPEADOS = mapea.recordcount
												
												MAPEA_GRAF = MAPEA_GRAF + MAPEADOS
												
												%>

                                               
            <td width="64" height="19" align="center"><font face="Verdana" size="2"><%=MAPEADOS%></font></td>
                                               
                                               <%
                                               RESTO = INDICADOS - MAPEADOS
                                               
                                               IF RESTO <= 0 THEN
                                               	RESTO = 0
                                               END IF
                                               
                                               RESTO_GRAF = RESTO_GRAF + RESTO
                                               %>
                                               
                                               
            <td width="106" height="19" align="center"><font face="Verdana" size="2"><%=RESTO%></font></td>
                                    </tr>
                                    <%
                                    orgao=""
                                    tem = tem + 1

                                    temp.movenext
                                    loop

                                    rs.movenext
                                    loop
                                    %>
                                    <tr>
                                               
            <td width="132" height="19" bgcolor="WHITE"><b><font face="Verdana" size="1" color="#FFFFFF">Total</font></b></td>
                                    </tr>

                                    <%
                                    end if
                                    if tem=0 then
                                    indica_graf=0
                                    %>
                                    <tr>
                                               
            <td height="19" colspan="5"><font color="#800000"><b>Nenhum Registro 
              Encontrado para a Seleção</b></font></td>
                                    </tr>
                                    <%
					                end if
					                
					                
                                    %>
                                    </table>
        <p><input type="hidden" name="orgao" size="20" value=<%=replace(sigl_orgao," E ", "_E_")%>></p>
        <p><input type="hidden" name="indicados" size="20" value=<%=indica_graf%>></p>
        <p><input type="hidden" name="mapeados" size="20" value=<%=mapea_graf%>></p>
        <p><input type="hidden" name="resto" size="20" value=<%=resto_graf%>></td>
           </tr>

</table>
</form>
</body>


</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>