<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
num_mega=request("selMegaProcesso")
num_processo=request("selProcesso")
num_sub=request("selSubProcesso")
num_atividade=request("selAtividade")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")


if num_processo<>0 and num_sub<>0 and num_atividade<>0 then
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo &" AND SUPR_CD_SUB_PROCESSO=" & num_sub & " AND ATIV_CD_ATIVIDADE=" & num_atividade
		query_=4
else
if num_processo<>0 and num_sub<>0 then
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo &" AND SUPR_CD_SUB_PROCESSO=" & num_sub
		query_=3
	else
	if num_processo<>0 then
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo 
		query_=2		
	else
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega
		query_=1
	end if
end if
end if

select case query_
	case 1
		ssql="DELETE FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
		desc="Não foi possível excluir o Mega-Processo, pois o mesmo possui Processos relacionados à ele"
	case 2
		ssql="DELETE FROM " & Session("PREFIXO") & "PROCESSO "
		desc="Não foi possível excluir o Processo, pois o mesmo possui Sub-Processos relacionados à ele"
	case 3
		ssql="DELETE FROM " & Session("PREFIXO") & "SUB_PROCESSO "
		desc="Não foi possível excluir o Sub-Processo, pois o mesmo possui Atividade ou Empresa relacionadas à ele. Deseja excluir o Registro mesmo assim?( TODOS os registros relacionados à ele também serão excluídos!)"
	case 4
		ssql="DELETE FROM " & Session("PREFIXO") & "ATIVIDADE "
		desc="Não foi possível excluir a Atividade, pois a mesma possui Transações relacionadas à ela"
END SELECT	

SSQL=SSQL + SQL_COMPL

'RESPONSE.WRITE SSQL

on error resume next
db.execute(ssql)

if err.number<>0 then
	select case err.number
			
	end select
end if	

%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>


</head>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p><%'=ssql%></p>
<p>&nbsp;</p>
<%if err.number=0 then%>
<p align="center"><b><font color="#000080" face="Verdana" size="2">Registro
Excluído com Sucesso!</font></b></p>
<%else%>
<p align="center"><b><font face="Arial" size="3" color="#800000"><%=desc%></font></b></p>
<div align="center">
  <center>
  <table border="0" width="33%">
    <tr>
      <td width="50%" align="center">
        <p align="center"><b><a href="excluir_tudo_sub.asp?mega=<%=num_mega%>&proc=<%=num_processo%>&sub=<%=num_sub%>"><font face="Verdana" size="4" color="#008000">SIM</font></a></b></td>
      <td width="50%" align="center"><b><a href="../indexA.asp"><font face="Verdana" size="4" color="#FF0000">NÃO</font></a></b></td>
    </tr>
  </table>
  </center>
</div>
<%end if%>

<p>&nbsp;</p>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%
num_mega=request("selMegaProcesso")
num_processo=request("selProcesso")
num_sub=request("selSubProcesso")
num_atividade=request("selAtividade")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")


if num_processo<>0 and num_sub<>0 and num_atividade<>0 then
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo &" AND SUPR_CD_SUB_PROCESSO=" & num_sub & " AND ATIV_CD_ATIVIDADE=" & num_atividade
		query_=4
else
if num_processo<>0 and num_sub<>0 then
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo &" AND SUPR_CD_SUB_PROCESSO=" & num_sub
		query_=3
	else
	if num_processo<>0 then
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega &" AND PROC_CD_PROCESSO="& num_processo 
		query_=2		
	else
		sql_compl="WHERE MEPR_CD_MEGA_PROCESSO="& num_mega
		query_=1
	end if
end if
end if

select case query_
	case 1
		ssql="DELETE FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
		desc="Não foi possível excluir o Mega-Processo, pois o mesmo possui Processos relacionados à ele"
	case 2
		ssql="DELETE FROM " & Session("PREFIXO") & "PROCESSO "
		desc="Não foi possível excluir o Processo, pois o mesmo possui Sub-Processos relacionados à ele"
	case 3
		ssql="DELETE FROM " & Session("PREFIXO") & "SUB_PROCESSO "
		desc="Não foi possível excluir o Sub-Processo, pois o mesmo possui Atividade ou Empresa relacionadas à ele. Deseja excluir o Registro mesmo assim?( TODOS os registros relacionados à ele também serão excluídos!)"
	case 4
		ssql="DELETE FROM " & Session("PREFIXO") & "ATIVIDADE "
		desc="Não foi possível excluir a Atividade, pois a mesma possui Transações relacionadas à ela"
END SELECT	

SSQL=SSQL + SQL_COMPL

'RESPONSE.WRITE SSQL

on error resume next
db.execute(ssql)

if err.number<>0 then
	select case err.number
			
	end select
end if	

%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>


</head>


<body topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<p><%'=ssql%></p>
<p>&nbsp;</p>
<%if err.number=0 then%>
<p align="center"><b><font color="#000080" face="Verdana" size="2">Registro
Excluído com Sucesso!</font></b></p>
<%else%>
<p align="center"><b><font face="Arial" size="3" color="#800000"><%=desc%></font></b></p>
<div align="center">
  <center>
  <table border="0" width="33%">
    <tr>
      <td width="50%" align="center">
        <p align="center"><b><a href="excluir_tudo_sub.asp?mega=<%=num_mega%>&proc=<%=num_processo%>&sub=<%=num_sub%>"><font face="Verdana" size="4" color="#008000">SIM</font></a></b></td>
      <td width="50%" align="center"><b><a href="../indexA.asp"><font face="Verdana" size="4" color="#FF0000">NÃO</font></a></b></td>
    </tr>
  </table>
  </center>
</div>
<%end if%>

<p>&nbsp;</p>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
