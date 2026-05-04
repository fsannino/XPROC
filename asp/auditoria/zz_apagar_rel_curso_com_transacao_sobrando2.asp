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

str_Sql = ""
str_Sql = str_Sql & " SELECT  DISTINCT "
str_Sql = str_Sql & " dbo.CURSO.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " , dbo.CURSO.CURS_TX_NOME_CURSO"
str_Sql = str_Sql & " , dbo.CURSO.CURS_CD_CURSO"
str_Sql = str_Sql & " , dbo.CURSO.ONDA_CD_ONDA"

'str_Sql = str_Sql & " FROM dbo.CURSO_FUNCAO INNER JOIN"
'str_Sql = str_Sql & " dbo.FUN_NEG_TRANSACAO ON "
'str_Sql = str_Sql & " dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN"
'str_Sql = str_Sql & " dbo.CURSO ON dbo.CURSO_FUNCAO.CURS_CD_CURSO = dbo.CURSO.CURS_CD_CURSO"

str_Sql = str_Sql & " FROM dbo.CURSO_FUNCAO INNER JOIN"
str_Sql = str_Sql & " dbo.FUN_NEG_TRANSACAO ON "
str_Sql = str_Sql & " dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN"
str_Sql = str_Sql & " dbo.CURSO ON dbo.CURSO_FUNCAO.CURS_CD_CURSO = dbo.CURSO.CURS_CD_CURSO INNER JOIN"
str_Sql = str_Sql & " dbo.CURSO_TRANSACAO ON dbo.CURSO.CURS_CD_CURSO = dbo.CURSO_TRANSACAO.CURS_CD_CURSO"

str_Sql = str_Sql & " WHERE dbo.CURSO.MEPR_CD_MEGA_PROCESSO > 0 "
str_Sql = str_Sql & COMPL & COMPL2 
str_Sql = str_Sql & " Order by dbo.CURSO.MEPR_CD_MEGA_PROCESSO, dbo.CURSO.CURS_CD_CURSO "
response.Write(str_Sql)
set rds_Curso = db.execute(str_Sql)

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
          de Cursos com algumas transa&ccedil;&otilde;es  associadas - 2 </font>
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
	atual1=""
	ant1=""			
	If not rds_Curso.Eof then
		do until rds_Curso.eof=true
            atual1=rds_Curso("mepr_cd_mega_processo")
			 
			str_Sql = ""
			str_Sql = str_Sql & " SELECT  DISTINCT "
			str_Sql = str_Sql & " dbo.CURSO_FUNCAO.CURS_CD_CURSO"
			str_Sql = str_Sql & " , dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO"
			str_Sql = str_Sql & " FROM dbo.CURSO_FUNCAO INNER JOIN"
            str_Sql = str_Sql & " dbo.FUN_NEG_TRANSACAO ON dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO"
			str_Sql = str_Sql & " WHERE dbo.CURSO_FUNCAO.CURS_CD_CURSO = '" & rds_Curso("CURS_CD_CURSO") & "'"
			'response.Write(str_Sql)
			set rds_Curso_Fun_Tran = db.execute(str_Sql)
			int_Conta_Existe = 0
			int_Conta_Nao_Existe = 0
			if not rds_Curso_Fun_Tran.Eof then
				do while not rds_Curso_Fun_Tran.Eof
			   		str_Sql = ""
					str_Sql = str_Sql & " SELECT distinct "
					str_Sql = str_Sql & " CURS_CD_CURSO, TRAN_CD_TRANSACAO"
					str_Sql = str_Sql & " FROM  dbo.CURSO_TRANSACAO"
					str_Sql = str_Sql & " WHERE (dbo.CURSO_TRANSACAO.CURS_CD_CURSO = '" & rds_Curso_Fun_Tran("CURS_CD_CURSO") & "')" 
					str_Sql = str_Sql & " AND (dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO = '" & rds_Curso_Fun_Tran("TRAN_CD_TRANSACAO") & "')"					
					'response.Write(str_Sql)
					set rds_Existe = db.execute(str_Sql)
					if not rds_Existe.Eof then
						int_Conta_Existe = int_Conta_Existe + 1
					else
						int_Conta_Nao_Existe = int_Conta_Nao_Existe + 1
						exit do
					end if
					rds_Curso_Fun_Tran.movenext
				Loop	
				'response.Write(" - " & int_Conta_Nao_Existe & " - ")
				if int_Conta_Nao_Existe > 0 then
          %>
          <tr>
				<%
					SET RS1=DB.EXECUTE("SELECT * FROM MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rds_Curso("MEPR_CD_MEGA_PROCESSO"))
					if atual1<>ant1 then
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
            <td width="45%" bgcolor="#FFFFEA"><font size="1" face="Verdana"><a href="rel_funcao_transacao_sobrando.asp?Cdcurso=<%=rds_Curso("CURS_CD_CURSO")%>&TlCurso=<%=rds_Curso("CURS_TX_NOME_CURSO")%>"><%=rds_Curso("CURS_TX_NOME_CURSO")%></a></font></td>
          </tr>
          <%		  
				end if
			end if						  
			tem=tem+1			
			ant1 = rds_Curso("mepr_cd_mega_processo")			
			rds_Curso.movenext
			if not rds_Curso.Eof then
				atual1 = rds_Curso("mepr_cd_mega_processo")
          	end if
		loop
          %>          
  </table>
<b>
<% else %>
<font face="Verdana" size="2" color="#800000">Nenhum Registro encontrado</font></b>
<% end if %>
</form>

</body>

</html>
