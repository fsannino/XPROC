<%@LANGUAGE="VBSCRIPT"%> 
<%
set con_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
conn_Cogest.cursorlocation = 3

if request("str_Tipo_Saida")="Excel" then
	Response.Buffer = true
	Response.ContentType = "application/vnd.ms-excel"
end if

set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")

dim int_Num_Lote
dim boo_Criado_Lote

boo_Criado_Lote = False

    str_SQL = ""
    str_SQL = str_SQL & " Select "
    str_SQL = str_SQL & " USMA_CD_USUARIO "
    str_SQL = str_SQL & " , FUNE_CD_FUNCAO_NEGOCIO "
    str_SQL = str_SQL & " , CURS_CD_CURSO "
    str_SQL = str_SQL & " from USU_CUR_FUN "
    str_SQL = str_SQL & " Where USMA_CD_USUARIO > '0'"
    str_SQL = str_SQL & " AND FUUS_IN_VALIDADO = 'S'"
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
                int_Aprovado = 0
                'LOOP EM TODOS OS REGISTROS ENQUANTO MESMO USUARIO E FUNCÃO
                'Do While int_NumReg_Usu_Curso > int_Loop_Usu_Curso And _
                'str_Cd_Usu_Anterior = rds_Usu_Curso("USMA_CD_USUARIO") And _
                'str_Cd_Fun_Anterior = rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")
                Do Until int_NumReg_Usu_Curso = int_Loop_Usu_Curso Or _
                str_Cd_Usu_Anterior <> rds_Usu_Curso("USMA_CD_USUARIO") Or _
                str_Cd_Fun_Anterior <> rds_Usu_Curso("FUNE_CD_FUNCAO_NEGOCIO")
                    str_SQL = " SELECT "
                    str_SQL = str_SQL & " USAP_TX_APROVEITAMENTO "
                    str_SQL = str_SQL & " FROM  USUARIO_APROVADO "
                    str_SQL = str_SQL & " WHERE USAP_CD_USUARIO ='" & str_Cd_Usu_Anterior & "'"
                    str_SQL = str_SQL & " And CURS_CD_CURSO ='" & rds_Usu_Curso("CURS_CD_CURSO") & "'"
                    Set rds_TabIncr = conn_Cogest.Execute(str_SQL)
                    int_NumReg_TabIncr = rds_TabIncr.RecordCount
                    int_Loop_TabIncr = 0
                    If int_NumReg_TabIncr > 0 Then
                        Do Until int_NumReg_TabIncr = int_Loop_TabIncr
                            If rds_TabIncr("USAP_TX_APROVEITAMENTO") = "AP" OR rds_TabIncr("USAP_TX_APROVEITAMENTO") = "LM" Then
                                int_Aprovado = int_Aprovado + 1
                            End If
                            int_Loop_TabIncr = int_Loop_TabIncr + 1
                            rds_TabIncr.MoveNext
                        Loop
                    Else
                        int_Nao_Aprovado = int_Nao_Aprovado + 1
                    End If
                    rds_Usu_Curso.MoveNext
                    int_Loop_Usu_Curso = int_Loop_Usu_Curso + 1
                    'If int_Nao_Aprovado <> 0 Then
                    '    Exit Do
                    'End If
                    If rds_Usu_Curso.EOF Then
                       Exit Do
                    End If
                    rds_TabIncr.Close
                Loop
                If int_Nao_Aprovado = 0 Then
                    Call f_grava_registro(str_Cd_Usu_Anterior, str_Cd_Fun_Anterior, "AP")
                Else
                    'Call f_grava_registro(str_Cd_Usu_Anterior, str_Cd_Fun_Anterior, "")
                   'Exit Do
                End If
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

Sub f_grava_registro (str_Cd_Usu_Anterior,str_Cd_Fun_Anterior,str_Status)

    str_SQL = ""
    str_SQL = str_SQL & " select "
    str_SQL = str_SQL & " FUNE_CD_FUNCAO_NEGOCIO "
    str_SQL = str_SQL & " FROM GOLI_FUNCAO_USUARIO "
    str_SQL = str_SQL & " WHERE USMA_CD_USUARIO = '" & str_Cd_Usu_Anterior & "'"
    str_SQL = str_SQL & " AND FUNE_CD_FUNCAO_NEGOCIO = '" & str_Cd_Fun_Anterior & "'"

    'rstRepeticao.Open str_SQL, conn_Cogest, , , adCmdText
    
	set rstRepeticao = conn_Cogest.Execute(str_SQL)
	
    If rstRepeticao.EOF Then

		if boo_Criado_Lote = False then
		   int_Num_Lote = f_Cria_Lote()
		   boo_Criado_Lote = True
		end if
		
		str_SQL = ""
		str_SQL = str_SQL & " Insert into GOLI_FUNCAO_USUARIO("
		str_SQL = str_SQL & " USMA_CD_USUARIO"
		str_SQL = str_SQL & " , FUNE_CD_FUNCAO_NEGOCIO"
		str_SQL = str_SQL & " , USFU_TX_APRO_TREINA"
		str_SQL = str_SQL & " , USFU_TX_APRO_XPROC"
		str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
		str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
		str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
		str_SQL = str_SQL & " )Values("
		str_SQL = str_SQL & "'" & str_Cd_Usu_Anterior & "',"
		str_SQL = str_SQL & "'" & str_Cd_Fun_Anterior & "',"
		str_SQL = str_SQL & "'" & str_Status & "',"
		str_SQL = str_SQL & "'" & str_Status & "',"	
		str_SQL = str_SQL & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
		Set rdsNovo = conn_Cogest.Execute(str_SQL)

	end if
	
end sub

function f_Cria_Lote()

	Dim int_Num_Lote_Anterior
	
	str_SQL = ""
	str_SQL = str_SQL & " SELECT "
	str_SQL = str_SQL & " LOTE_NR_SEQ_LOTE"
	str_SQL = str_SQL & " , LOTE_DT_ENVIO"
	str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
	str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
	str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
	str_SQL = str_SQL & " FROM  dbo.GOLI_LOTE"
	
	str_SQL = ""
	str_SQL = str_SQL & " SELECT MAX(LOTE_NR_SEQ_LOTE)AS NUM_LOTE FROM GOLI_LOTE "
	
	int_Num_Lote = 0
	
	set rs=db.execute(str_SQL)
	
	if not isnull(rs("NUM_LOTE")) then
		int_Num_Lote_Anterior = rs("NUM_LOTE")
	end if
	
	if int_Num_Lote=0 then
		int_Num_Lote=1
	else
		int_Num_Lote=int_Num_Lote+1
	end if



end function
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
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td colspan="3" height="20">&nbsp; </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="50%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="50%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alteração 
        de Atividade</font></td>
      <td width="26%">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%"><b><font face="Verdana, Arial, Helvetica, sans-serif" color="#000099" size="2">&nbsp;Selecione 
        a Atividade que deseja alterar</font></b></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <p> 
          <select size="1" name="selAtividade" onchange="javascript:Confirma()">
            <option value="0">== Selecione a Atividade ==</option>
            <%do while not rs.EOF %>
            <option value=<%=rs("ATCA_CD_ATIVIDADE_CARGA")%>><%=rs("ATCA_TX_DESC_ATIVIDADE")%></option>
            <% rs.movenext
  Loop
  %>
          </select>
        </p>
        </font></td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%"></td>
      <td width="12%"></td>
      <td width="63%">&nbsp;</td>
      <td width="16%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="9%">&nbsp;</td>
      <td width="12%">&nbsp;</td>
      <td width="63%">&nbsp;</td>
    </tr>
  </table> 
</form>
</body>
</html>
