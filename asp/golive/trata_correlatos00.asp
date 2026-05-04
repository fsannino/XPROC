<%@LANGUAGE="VBSCRIPT"%> 
<%
Server.ScriptTimeOut=30000

Dim dat_Inicio 
Dim dat_Fim 

Dim int_NumReg_UsuAprov 
Dim int_NumReg_Correlatos 

'strCnnCogest = "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest002;uid=cogestadm;database=cogest"

set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
'conn_Cogest.Open strCnnCogest
conn_Cogest.cursorlocation = 3

Set rstCorrelatos = CreateObject("ADODB.Recordset")
rstCorrelatos.CursorLocation = 3

Set rstUsuAprov = CreateObject("ADODB.Recordset")
rstUsuAprov.CursorLocation = 3

Set rstUsuJaAprov = CreateObject("ADODB.Recordset")
rstUsuJaAprov.CursorLocation = 3

dat_Inicio = Now

str_SQL = ""
str_SQL = str_SQL & " select "
str_SQL = str_SQL & " CURS_CD_CURSO, CURS_CD_CURSO_CORRELATO"
str_SQL = str_SQL & " FROM dbo.CURSO_CORRELATO"
str_SQL = str_SQL & " ORDER BY CURS_CD_CURSO"
rstCorrelatos.Open str_SQL, conn_Cogest, , , adCmdText

int_Controle_loop = 0

int_NumReg_Correlatos = rstCorrelatos.RecordCount
If Not rstCorrelatos.EOF Then
	Do While Not rstCorrelatos.EOF
'		RESPONSE.Write("-----1 - FIM-----")
'		RESPONSE.End()
		str_SQL = ""
		str_SQL = str_SQL & " select distinct "
		str_SQL = str_SQL & " USAP_CD_USUARIO"
		str_SQL = str_SQL & " FROM dbo.USUARIO_APROVADO"
		str_SQL = str_SQL & " WHERE  CURS_CD_CURSO = '" & rstCorrelatos("CURS_CD_CURSO") & "'"
		str_SQL = str_SQL & " AND (USAP_TX_APROVEITAMENTO = 'AP')"
		rstUsuAprov.Open str_SQL, conn_Cogest, , , adCmdText
		int_NumReg_UsuAprov = rstUsuAprov.RecordCount
		If Not rstUsuAprov.EOF Then
			strCursoPrincAnterior = rstCorrelatos("CURS_CD_CURSO")
			Do While Not rstCorrelatos.EOF And strCursoPrincAnterior = rstCorrelatos("CURS_CD_CURSO")
				Do While Not rstUsuAprov.EOF
				int_Controle_loop = int_Controle_loop + 1
					IF int_Controle_loop = 100 THEN
						'RESPONSE.Write("----- 2 - FIM-----")
						'RESPONSE.End()
					END IF
					If rstCorrelatos("CURS_CD_CURSO") <> "MES201" Then
						'MsgBox ("A")
					End If
					If rstUsuAprov("USAP_CD_USUARIO") = "ZNZY" Then
						'MsgBox ("B")
					End If
					str_SQL = ""
					str_SQL = str_SQL & " select distinct "
					str_SQL = str_SQL & " USAP_CD_USUARIO, USAP_TX_APROVEITAMENTO"
					str_SQL = str_SQL & " FROM dbo.USUARIO_APROVADO"
					str_SQL = str_SQL & " WHERE"
					str_SQL = str_SQL & " USAP_CD_USUARIO = '" & rstUsuAprov("USAP_CD_USUARIO") & "'"
					str_SQL = str_SQL & " AND CURS_CD_CURSO = '" & rstCorrelatos("CURS_CD_CURSO_CORRELATO") & "'"
					rstUsuJaAprov.Open str_SQL, conn_Cogest, , , adCmdText
					'RESPONSE.Write(str_SQL)
					'RESPONSE.End()
					If rstUsuJaAprov.EOF Then
						str_SQL = ""
						str_SQL = str_SQL & " Insert into USUARIO_APROVADO("
						str_SQL = str_SQL & " USAP_CD_USUARIO"
						str_SQL = str_SQL & " , CURS_CD_CURSO"
						str_SQL = str_SQL & " , USAP_TX_APROVEITAMENTO"
						str_SQL = str_SQL & " , MOTI_NR_CD_MOTIVO "
						str_SQL = str_SQL & " , USAP_DT_LIBERADO_MANUAL"
						str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
						str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
						str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
						str_SQL = str_SQL & " ) Values ("
						str_SQL = str_SQL & "'" & rstUsuAprov("USAP_CD_USUARIO") & "',"
						str_SQL = str_SQL & "'" & rstCorrelatos("CURS_CD_CURSO_CORRELATO") & "',"
						str_SQL = str_SQL & "'LM',"
						str_SQL = str_SQL & "1,"
						str_SQL = str_SQL & "GETDATE(),"
						str_SQL = str_SQL & "'C' ,'" & "XK45" & "' ,GETDATE())"
						conn_Cogest.Execute str_SQL
					Else
						If rstUsuJaAprov("USAP_TX_APROVEITAMENTO") = "  " Then
							str_SQL = ""
							str_SQL = str_SQL & " UPDATE USUARIO_APROVADO SET "
							str_SQL = str_SQL & "  USAP_TX_APROVEITAMENTO = 'LM'"
							str_SQL = str_SQL & " ,MOTI_NR_CD_MOTIVO = 1"
							str_SQL = str_SQL & " , USAP_DT_LIBERADO_MANUAL = GETDATE()"
							str_SQL = str_SQL & " , ATUA_TX_OPERACAO = 'A'"
							str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO = 'XK45'"
							str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO = GETDATE()"
							str_SQL = str_SQL & " WHERE"
							str_SQL = str_SQL & " USAP_CD_USUARIO = '" & rstUsuAprov("USAP_CD_USUARIO") & "'"
							str_SQL = str_SQL & " AND CURS_CD_CURSO = '" & rstCorrelatos("CURS_CD_CURSO_CORRELATO") & "'"
							conn_Cogest.Execute str_SQL
						End If
					End If
					rstUsuJaAprov.Close
					rstUsuAprov.MoveNext
				Loop
				rstCorrelatos.MoveNext
				If rstCorrelatos.EOF Then
					Exit Do
				End If
				rstUsuAprov.MoveFirst
			Loop
		Else
			rstCorrelatos.MoveNext
		End If
		If rstCorrelatos.EOF Then
			Exit Do
		End If
		rstUsuAprov.Close
	Loop
End If

dat_Fim = Now

conn_Cogest.close
set conn_Cogest = Nothing

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="POST" action="">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
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
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
  <table width="100%"  border="0" cellspacing="0" cellpadding="1">
    <tr>
      <td>Treina 1 </td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
<%
'response.Redirect("importa_usuario_treinasin2.asp")
%>
<script language="javascript">	
	function fechar()
	{
		window.top.close();	
	}	
		
	setTimeout('fechar()',1);	
	//window.top.frame2.focus();
	//window.top.frame2.print();
</script>
</html>
