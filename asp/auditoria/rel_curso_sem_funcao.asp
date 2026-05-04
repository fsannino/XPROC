<%
'RESPONSE.Write(Session("Conn_String_Cogest_Gravacao"))
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=REQUEST("selMegaProcesso")
str_onda=request("selOnda")

if str_mega > 0 then
	compl=" and  " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO=" + str_mega
else
	compl=""
end if
if str_onda >0 then
	compl2=" and  " & Session("PREFIXO") & "CURSO.ONDA_CD_ONDA = " + str_onda
else
	compl2=""
end if

SSQL1="SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO AS MEGA"
SSQL1 = SSQL1 & ", " & Session("PREFIXO") & "CURSO.* "
SSQL1 = SSQL1 & " FROM " & Session("PREFIXO") & "CURSO INNER JOIN " & Session("PREFIXO") & "MEGA_PROCESSO ON " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
SSQL1 = SSQL1 & " where MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO > 0 " 
SSQL1 = SSQL1 & COMPL & COMPL2  
SSQL1 = SSQL1 & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "CURSO.CURS_CD_CURSO"

str_Sql = ""
str_Sql = str_Sql & " SELECT "
str_Sql = str_Sql & " dbo.CURSO.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " , dbo.CURSO.CURS_TX_NOME_CURSO"
str_Sql = str_Sql & " , dbo.CURSO.CURS_CD_CURSO"
str_Sql = str_Sql & " , dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO"
str_Sql = str_Sql & " FROM dbo.CURSO LEFT OUTER JOIN"
str_Sql = str_Sql & " dbo.CURSO_FUNCAO ON dbo.CURSO.CURS_CD_CURSO = dbo.CURSO_FUNCAO.CURS_CD_CURSO"
str_Sql = str_Sql & " WHERE (dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO IS NULL) and ONDA_CD_ONDA<>4 "
str_Sql = str_Sql & COMPL & COMPL2  
str_Sql = str_Sql & " ORDER BY " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO, " & Session("PREFIXO") & "CURSO.CURS_CD_CURSO"


'RESPONSE.Write(str_Sql)
SET RS=DB.EXECUTE(str_Sql)

SET MEGA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
set rs_onda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ABRANGENCIA_CURSO WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_TX_DESC_ONDA")

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>
<SCRIPT>
function envia()
{
this.location.href='relat_geral_curso.asp?mega='+document.frm1.selMegaProcesso.value+'&selOnda='+document.frm1.selOnda.value
}
</SCRIPT>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" link="#800000" vlink="#800000" alink="#800000">
<form method="POST" action="" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top">      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
      <tr>
        <td bgcolor="#330099" width="39" valign="middle" align="center">
          <div align="center">
            <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a>    
        </div></td>
        <td bgcolor="#330099" width="36" valign="middle" align="center">
          <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div></td>
        <td bgcolor="#330099" width="27" valign="middle" align="center">
          <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div></td>
      </tr>
      <tr>
        <td bgcolor="#330099" height="12" width="39" valign="middle" align="center">
          <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div></td>
        <td bgcolor="#330099" height="12" width="36" valign="middle" align="center">
          <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div></td>
        <td bgcolor="#330099" height="12" width="27" valign="middle" align="center">
          <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div></td>
      </tr>
    </table></td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"></td>
          <td width="50"></td>
          <td width="26"></td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        
  <table width="100%" height="38" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>
        <div align="center">
          <p align="center" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório
          de Cursos sem fun&ccedil;&atilde;o associada </font>
        </div>
      </td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <table border="0" width="81%">
          <tr>
            <td width="21%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Mega-Processo</font></b></td>
            <td width="45%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Curso</font></b></td>
          </tr>
          <%
          tem=0
   		  int_Tot_Mega = 0
			atual1=""
			ant1=""			

          do until rs.eof=true
          atual1=rs("mepr_cd_mega_processo")
          %>
          <tr>
				<%
             	SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
             	if atual1<>ant1 then
					if tem > 0 then 						
					%>
				
    <tr>
					<td>Total do Mega:<%=int_Tot_Mega%></td>
					<td bgcolor="#FFFFEA">&nbsp;</td>
    </tr>
				
				<%  int_Tot_Mega = 0
					end if
				
					NOME1=RS1("MEPR_TX_DESC_MEGA_PROCESSO")            
				else
					nome1=""
				end if
				
				if nome1="" then
					cor="white"
				else
					cor="#CCCCCC"	
				end if			
				%>
            <td width="21%" bgcolor="<%=cor%>"><font size="1" face="Verdana"><%=nome1%></font></td>
            <td width="45%" bgcolor="#FFFFEA"><font size="1" face="Verdana"><a href="../Curso/exibe_curso.asp?curso=<%=rs("CURS_CD_CURSO")%>"><%=rs("CURS_TX_NOME_CURSO")%></a></font></td>
          </tr>
          <%
          tem=tem+1
          int_Tot_Mega = int_Tot_Mega + 1
          ant1=rs("mepr_cd_mega_processo")
          
          rs.movenext
          
          on error resume next
          atual1=rs("mepr_cd_mega_processo")
          
          loop
          %>

    <tr>
					<td>Total do Mega:<%=int_Tot_Mega%></td>
					<td bgcolor="#FFFFEA">&nbsp;</td>
    </tr>
          
  </table>
  <p>&nbsp;</p>
  <p align="center"><b>
    <%if tem=0 then%>
      <font face="Verdana" size="2" color="#800000">Nenhum Registro encontrado</font></b>
    <% else %>
      Total Geral: <%=tem%>
    <%end if%>
  </p>
  <p>&nbsp;      </p>
</form>

</body>

</html>
