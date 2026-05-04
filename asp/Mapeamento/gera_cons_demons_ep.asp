<%
orgao = request("selOrgao")
atual = request("selCurso")
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

set temp5 = db.execute("SELECT DISTINCT AGLU_SG_AGLUTINADO AS ORGAO FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & orgao & "'")

tabela = "[" & temp5("ORGAO") & "]"

sigl_orgao = temp5("ORGAO")

'on error resume next
set rs = db2.execute("SELECT DISTINCT ORGAO FROM [MEGA_E&P] ORDER BY ORGAO")

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
<title>Em Manutenção...</title>
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
                         <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber4">
                                    <tr>
                                               <td width="50%"><a href="javascript:history.go(-1)"><img border="0" src="voltar.gif"></a></td>
                                               <td width="25%"><p align="center"><a href="javascript:grafico()"><img border="0" src="graf.jpg" width="35" height="26" alt="Visualizar Gráfico"></a></td>
                                               <td width="25%"><a href="javascript:print()"><img border="0" src="../Apoio/impressao.jpg" width="29" height="29" alt="Imprimir Relatório"></a></td>
                                    </tr>
                         </table>
                         <p><b><font face="Verdana" color="#800000">Consulta Demonstrativa por Mega-Processo - E &amp; P</font></b></p>
                         
        <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#D7D5CC" width="804" id="AutoNumber3" height="39">
          <tr>
                                               
            <td width="132" height="19" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Órgão</font></b></td>
                                               
            <td width="428" height="19" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Mega 
              - Processo</font></b></td>
                                               
            <td width="70" height="19" align="center" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Vagas 
              Indicadas</font></b></td>
                                               
            <td width="64" height="19" align="center" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Vagas 
              Mapeadas</font></b></td>
                                               
            <td width="98" height="19" align="center" bgcolor="#666633"><b><font face="Verdana" size="1" color="#FFFFFF">Pend&ecirc;ncias 
              em Vagas Indicadas</font></b></td>
                                    </tr>
                                    
                                    <%
                                    if erro = 0 then
                                    
                                    do until rs.eof=true
                                    
                                    if atual="XXXX" then
		                               ssql="SELECT DISTINCT ORGAO, MEGA, INDICADOS FROM [MEGA_E&P] WHERE ORGAO='" & rs("orgao") & "' ORDER BY MEGA"
		                            else
	    	                           set m = db.execute("SELECT MEPR_TX_DESC_MEGA_PROCESSO FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & atual)
	    	                           ssql="SELECT DISTINCT ORGAO, MEGA, INDICADOS FROM [MEGA_E&P] WHERE MEGA = '" & m("MEPR_TX_DESC_MEGA_PROCESSO") & "' AND ORGAO='" & rs("orgao") & "' ORDER BY MEGA"
	        	                    end if
	        	                    
	        	                    set temp = db2.execute(ssql)
                                    
                                    if temp.eof=false then
	                                    orgao_2 = temp("orgao")
	                                end if
                                    
        	                        do until temp.eof=true
            	                        
                    	            	nome_mega = temp("MEGA")
                                    
                        	        	INDICADOS = temp("INDICADOS")
                        	        	
                        	        	INDICA_GRAF = INDICA_GRAF + INDICADOS

												set org = db.execute("SELECT * FROM ORGAO_MENOR WHERE ORME_SG_ORG_MENOR='" & trim(temp("orgao")) & "'")
												org_pesquisa = left(org("ORME_CD_ORG_MENOR"),7)
												
												set mc = db.execute("SELECT * FROM MEGA_PROCESSO WHERE MEPR_TX_DESC_MEGA_PROCESSO='" & trim(nome_mega) & "'")
												pre_mega = mc("MEPR_TX_ABREVIA_CURSO")
																								
												set cursos_f = db2.execute("SELECT DISTINCT CURSO FROM " & tabela & " WHERE ORGAO='" & trim(temp("orgao")) & "' AND CURSO LIKE '" & pre_mega & "%'")
												
												total_cursos = ""
												
												do until cursos_f.eof=true
													total_cursos = total_cursos & "'" & cursos_f("CURSO") & "',"
													cursos_f.movenext
												loop
												
												on error resume next
												total_cursos = left(total_cursos,len(total_cursos)-1)
												err.clear
												
												ssql=""
												ssql="SELECT DISTINCT"
												ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO"
												ssql=ssql+" FROM APOIO_LOCAL_CURSO"
												ssql=ssql+" INNER JOIN USUARIO_MAPEAMENTO ON"
												ssql=ssql+" APOIO_LOCAL_CURSO.USMA_CD_USUARIO = USUARIO_MAPEAMENTO.USMA_CD_USUARIO"
												ssql=ssql+" INNER JOIN CURSO ON"
												ssql=ssql+" APOIO_LOCAL_CURSO.CURS_CD_CURSO = CURSO.CURS_CD_CURSO"
												ssql=ssql+" WHERE APOIO_LOCAL_CURSO.CURS_CD_CURSO IN (" & total_cursos & ") AND USUARIO_MAPEAMENTO.ORME_CD_ORG_MENOR LIKE '" & org_pesquisa & "%'"
												
												set mapea = db.execute(ssql)
												
												MAPEADOS = mapea.recordcount
												
												MAPEA_GRAF = MAPEA_GRAF + MAPEADOS
												
                                               RESTO = INDICADOS - MAPEADOS
                                               
                                               
                                               IF RESTO <= 0 THEN
                                               	RESTO = 0
                                               END IF
                                               
                                               RESTO_GRAF = RESTO_GRAF + RESTO
                                                                                              
                                               IF INDICADOS<>0 THEN
                                               %>
                                	    <tr>
                                               
            <td width="132" height="19"><font face="Verdana" size="1"><b><%=orgao_2%></b></font></td>
                                               
            <td width="428" height="19"><b><font face="Verdana" size="1"><%=nome_mega%></font></b></td>
                                               
            <td width="70" height="19" align="center"><font face="Verdana" size="2"><%=INDICADOS%></font></td>
                                               
            <td width="64" height="19" align="center"><font face="Verdana" size="2"><%=MAPEADOS%></font></td>                                               
                                               
            <td width="98" height="19" align="center"><font face="Verdana" size="2"><%=RESTO%></font></td>
	                                    </tr>
                                    <%
                                    orgao_2=" "                                    
                                    tem = tem + 1
                                    END IF

                                    temp.movenext
                                    loop

                                    rs.movenext
                                    loop
                                    
                                    end if
                                    if tem=0 then
                                    %>
                                    <tr>
                                               
            <td height="19" colspan="5"><font color="#800000"><b>Nenhum Registro 
              Encontrado para a Seleção</b></font></td>
                                    </tr>
                                    <%
                                    end if
                                    %>
                                    </table>
                      </td>
           </tr>
        <p><input type="hidden" name="orgao" size="20" value=<%=replace(sigl_orgao,"&","e")%>></p>
        <p><input type="hidden" name="indicados" size="20" value=<%=indica_graf%>></p>
        <p><input type="hidden" name="mapeados" size="20" value=<%=mapea_graf%>></p>
        <p><input type="hidden" name="resto" size="20" value=<%=resto_graf%>></td>

</table>
</form>
</body>

</html>

<script>
document.title = 'Indicação de Multiplicadores'
</script>