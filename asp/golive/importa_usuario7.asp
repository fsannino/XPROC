<%@LANGUAGE="VBSCRIPT"%> 
<%
 Server.ScriptTimeOut=30000
%>
<%
set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
conn_Cogest.cursorlocation = 3

Conn_String_Treinasin_Leitura = "Provider=SQLOLEDB.1;server=S5200DB01\DB01;pwd=sinergiacogest;uid=usr_cogest;database=IntranetSinergia"

set conn_Treinasin=server.createobject("ADODB.CONNECTION")
conn_Treinasin.Open Conn_String_Treinasin_Leitura
conn_Treinasin.cursorlocation = 3

    str_SQL = ""
    str_SQL = str_SQL & " Select "
    str_SQL = str_SQL & " USMA_CD_USUARIO "
    str_SQL = str_SQL & " , FUNE_CD_FUNCAO_NEGOCIO "
    str_SQL = str_SQL & " , CURS_CD_CURSO "
    str_SQL = str_SQL & " from USU_CUR_FUN "
    str_SQL = str_SQL & " Where USMA_CD_USUARIO > '0'"
    str_SQL = str_SQL & " AND FUUS_IN_VALIDADO = 'S'"	
	str_SQL = str_SQL & " AND SUBSTRING(USMA_CD_USUARIO,1,1) IN ('W')"
	
    str_SQL = str_SQL & " order by USMA_CD_USUARIO , FUNE_CD_FUNCAO_NEGOCIO  "
    Set rds_Usu_Curso = conn_Cogest.Execute(str_SQL)
    int_NumReg_Usu_Curso = rds_Usu_Curso.RecordCount
    int_Loop_Usu_Curso = 0
    If int_NumReg_Usu_Curso > 0 Then
        'LOOP EM TODOS OS REGISTROS
        Do Until int_NumReg_Usu_Curso = int_Loop_Usu_Curso
            str_Cd_Usu_Anterior = rds_Usu_Curso("USMA_CD_USUARIO")
            'LOOP EM TODOS OS REGISTROS ENQUANTO MESMO USUARIO
            'Do While int_NumReg_Usu_Curso > int_Loop_Usu_Curso And _
            'str_Cd_Usu_Anterior = rds_Usu_Curso("USMA_CD_USUARIO")
            Do Until int_NumReg_Usu_Curso = int_Loop_Usu_Curso Or _
            str_Cd_Usu_Anterior <> rds_Usu_Curso("USMA_CD_USUARIO")
                str_Cd_Fun_Anterior = rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")
                int_Nao_Aprovado = 0
                'LOOP EM TODOS OS REGISTROS ENQUANTO MESMO USUARIO E FUNCÃO
                'Do While int_NumReg_Usu_Curso > int_Loop_Usu_Curso And _
                'str_Cd_Usu_Anterior = rds_Usu_Curso("USMA_CD_USUARIO") And _
                'str_Cd_Fun_Anterior = rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")
                Do Until int_NumReg_Usu_Curso = int_Loop_Usu_Curso Or _
                str_Cd_Usu_Anterior <> rds_Usu_Curso("USMA_CD_USUARIO") Or _
                str_Cd_Fun_Anterior <> rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")
                    str_SQL = " SELECT "
                    str_SQL = str_SQL & " APROV "
                    str_SQL = str_SQL & " FROM  TabInscritos "
                    str_SQL = str_SQL & " WHERE CHAVE ='" & str_Cd_Usu_Anterior & "'"
                    str_SQL = str_SQL & " And COD_DISCIPLINA ='" & rds_Usu_Curso("CURS_CD_CURSO") & "'"
                    Set rds_TabIncr = conn_Treinasin.Execute(str_SQL)
                    int_NumReg_TabIncr = rds_TabIncr.RecordCount
                    int_Loop_TabIncr = 0
                    int_Aprovado = 0
                    If int_NumReg_TabIncr > 0 Then
                        Do Until int_NumReg_TabIncr = int_Loop_TabIncr
                            If rds_TabIncr("APROV") = "AP" Then
                                int_Aprovado = int_Aprovado + 1
                                Exit Do
                            End If
                            int_Loop_TabIncr = int_Loop_TabIncr + 1
                            rds_TabIncr.MoveNext
                        Loop
						If int_Aprovado > 0 Then
							Call f_grava_registro_curso(Trim(str_Cd_Usu_Anterior), rds_Usu_Curso("CURS_CD_CURSO"), "AP")
						Else
							Call f_grava_registro_curso(Trim(str_Cd_Usu_Anterior), rds_Usu_Curso("CURS_CD_CURSO"), "")
						End If
					Else
						Call f_grava_registro_curso(Trim(str_Cd_Usu_Anterior), rds_Usu_Curso("CURS_CD_CURSO"), "")												
                    End If
                    rds_Usu_Curso.MoveNext
                    int_Loop_Usu_Curso = int_Loop_Usu_Curso + 1
                    If rds_Usu_Curso.EOF Then
                       Exit Do
                    End If
                    rds_TabIncr.Close
                Loop
                If rds_Usu_Curso.EOF Then
                   Exit Do
                End If
            Loop
            If rds_Usu_Curso.EOF Then
               Exit Do
            End If
        Loop
    Else
    
    End If

Sub f_grava_registro_curso(str_Cd_Usu_Anterior, str_Cd_Curso, str_Status)

    str_SQL = ""
    str_SQL = str_SQL & " select "
    str_SQL = str_SQL & " USAP_CD_USUARIO"
    str_SQL = str_SQL & " , CURS_CD_CURSO"
    str_SQL = str_SQL & " , USAP_TX_APROVEITAMENTO"
    str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
    str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
    str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
    str_SQL = str_SQL & " FROM dbo.USUARIO_APROVADO"
    str_SQL = str_SQL & " WHERE USAP_CD_USUARIO = '" & str_Cd_Usu_Anterior & "'"
    str_SQL = str_SQL & " AND CURS_CD_CURSO = '" & str_Cd_Curso & "'"
    
	set rstRepeticao = conn_Cogest.Execute(str_SQL)
	
    If rstRepeticao.EOF Then

        str_SQL = ""
        str_SQL = str_SQL & " Insert into USUARIO_APROVADO("
        str_SQL = str_SQL & " USAP_CD_USUARIO"
        str_SQL = str_SQL & " , CURS_CD_CURSO"
        str_SQL = str_SQL & " , USAP_TX_APROVEITAMENTO"
        str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
        str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
        str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
        str_SQL = str_SQL & " ) Values ("
        str_SQL = str_SQL & "'" & str_Cd_Usu_Anterior & "',"
        str_SQL = str_SQL & "'" & str_Cd_Curso & "',"
        str_SQL = str_SQL & "'" & str_Status & "',"
		str_SQL = str_SQL & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"		
        conn_Cogest.Execute str_SQL
	else
		if str_Status = "AP" then
			if Trim(rstRepeticao("USAP_TX_APROVEITAMENTO")) = "" then
				str_SQL = ""
				str_SQL = str_SQL & " update USUARIO_APROVADO set "
				str_SQL = str_SQL & " USAP_TX_APROVEITAMENTO = '" & str_Status & "'"
				str_SQL = str_SQL & " , ATUA_TX_OPERACAO = 'A'"
				str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO = '" & Session("CdUsuario")   & "'"
				str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO = GETDATE()"
				str_SQL = str_SQL & " WHERE USAP_CD_USUARIO = '" & str_Cd_Usu_Anterior & "'"
				str_SQL = str_SQL & " AND CURS_CD_CURSO = '" & str_Cd_Curso & "'"
 		        conn_Cogest.Execute str_SQL
			end if
		end if	
	end if
	rstRepeticao.close
	set rstRepeticao = Nothing
end sub
conn_Treinasin.close
set conn_Treinasin = Nothing
conn_Cogest.close
set conn_Cogest = Nothing
%>
<html>
<head>
<script>
function Confirma() 
{ 
if (document.frm1.selAtividade.selectedIndex == 0)
     { 
	 alert("A seleção de uma Atividade é obrigatório!");
     document.frm1.selAtividade.focus();
     return;
     }
	 else
     {
 	  window.location.href.href='altera_Atividade1.asp?selAtiv='+document.frm1.selAtividade.value
	 }
 }
</SCRIPT>
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
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>Ok importa&ccedil;&atilde;o conclu&iacute;da </td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>
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
