<%

curso=UCase(request("curso"))
mega=request("mega")

prer=request("txtTrans")
vet_Pre_Selecionada = split(prer,",")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

dim vet_Pre_Cadastrada(10)

vet_Pre_Cadastrada(0) = ""

'***********************************************************************************************
' CARREGA VETOR COMOS CADASTRADOS
'***********************************************************************************************
str_Sql = ""
str_Sql = str_Sql & " SELECT "
str_Sql = str_Sql & " CURS_PRE_REQUISITO"
str_Sql = str_Sql & " , CURS_CD_CURSO"
str_Sql = str_Sql & " FROM dbo.CURSO_PRE_REQUISITO"
str_Sql = str_Sql & " WHERE CURS_CD_CURSO='" & curso & "'"
'response.Write(str_Sql)
'response.End()
set rds_Curso_Pre=db.execute(str_Sql)
if not rds_Curso_Pre.Eof then
	int_Indice = 0 
	do while not rds_Curso_Pre.Eof
		vet_Pre_Cadastrada(int_Indice) = rds_Curso_Pre("CURS_PRE_REQUISITO")
		int_Indice = int_Indice + 1
		rds_Curso_Pre.movenext
	loop
end if	
rds_Curso_Pre.close
set rds_Curso_Pre = Nothing

'***********************************************************************************************
' VERIFICA QUEM SERÁ EXCLUIDO
'***********************************************************************************************
int_Indice = 0 
for int_Indice = LBound(vet_Pre_Cadastrada) to UBound(vet_Pre_Cadastrada)
	lng_Posicao = SequentialSearchStringArray(vet_Pre_Selecionada, vet_Pre_Cadastrada(int_Indice))
	if lng_Posicao = -1 then
		deletar_Pre(vet_Pre_Cadastrada(int_Indice))
	end if
	'response.Write("<P>" & "----" & int_Indice  &  vet_Pre_Cadastrada(int_Indice))
next			
'response.Write("<P>" & "----FINAL CADA----INDICE - " & int_Indice  & "VET - " & vet_Pre_Cadastrada(int_Indice-1))
'***********************************************************************************************
' VERIFICA QUEM SERÁ INCLUÍDO
'***********************************************************************************************
int_Indice = 0 
for int_Indice = LBound(vet_Pre_Selecionada) to UBound(vet_Pre_Selecionada)
	lng_Posicao = SequentialSearchStringArray(vet_Pre_Cadastrada, vet_Pre_Selecionada(int_Indice))
	if lng_Posicao = -1 then
		incluir_Pre(vet_Pre_Selecionada(int_Indice))
	end if
	'response.Write("<P>" & "----" & int_Indice  &  vet_Pre_Selecionada(int_Indice))	
next			
'response.Write("<P>" & "----FINAL SELE----INDICE - " & int_Indice  & "VET - " &  vet_Pre_Selecionada(int_Indice-1))
'RESPONSE.End()

sub deletar_Pre(pCursoPre)

	'response.Write("<p> deletar")

	str_Sql = ""
	str_Sql = str_Sql & " DELETE FROM "
	str_Sql = str_Sql & " CURSO_PRE_REQUISITO "
	str_Sql = str_Sql & " WHERE "
	str_Sql = str_Sql & " CURS_CD_CURSO = '" & curso & "'"
	str_Sql = str_Sql & " and CURS_PRE_REQUISITO = '" & pCursoPre & "'"

	db.execute(str_Sql)

	'*** VERIFICA AS FUNÇÕES LIGADAS AO CURSO E DESASSOCIA AO PRE 
	str_Sql = ""
	str_Sql = str_Sql & " SELECT DISTINCT "
	str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO "
	str_Sql = str_Sql & " FROM "
	str_Sql = str_Sql & " CURSO_FUNCAO "
	str_Sql = str_Sql & " WHERE "
	str_Sql = str_Sql & " CURS_CD_CURSO='" & curso & "'"
	
	set fonte_funcao_todas=db.execute(str_Sql)	
	do until fonte_funcao_todas.eof=true
		str_Cd_Funcao = fonte_funcao_todas("FUNE_CD_FUNCAO_NEGOCIO")
		Call deleta_funcao_prepre(pCursoPre, str_Cd_Funcao,curso)
		fonte_funcao_todas.movenext
	loop
	fonte_funcao_todas.close
	
end sub

Sub deleta_funcao_prepre(pCursoPrePre, pFuncao, pCursoPre)

	str_Sql = ""
	str_Sql = str_Sql & " SELECT DISTINCT "
	str_Sql = str_Sql & " CURS_CD_CURSO"
	str_Sql = str_Sql & " FROM dbo.CURSO_PRE_REQUISITO"
	str_Sql = str_Sql & " WHERE "
	str_Sql = str_Sql & " CURS_PRE_REQUISITO = '" & pCursoPrePre & "'"
    str_Sql = str_Sql & " AND CURS_CD_CURSO <> '" & pCursoPre & "'"
	
	set rds_Curso_Pre_Pre=db.execute(str_Sql)

	int_Qtd_Relacao = 0
	if not rds_Curso_Pre_Pre.Eof then		
		do while not rds_Curso_Pre_Pre.Eof
			str_Cd_Curso = rds_Curso_Pre_Pre("CURS_CD_CURSO")
			str_Sql = ""
			str_Sql = str_Sql & " SELECT DISTINCT "
			str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO "
			str_Sql = str_Sql & " FROM CURSO_FUNCAO "
			str_Sql = str_Sql & " WHERE CURS_CD_CURSO='" & str_Cd_Curso & "'"
			str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & pFuncao & "'"
			
			set rds_Func_Curso_f2=db.execute(str_Sql)

			if not rds_Func_Curso_f2.Eof then		
				int_Qtd_Relacao = int_Qtd_Relacao + 1
				exit do
			end if
			rds_Func_Curso_f2.close
			rds_Curso_Pre_Pre.movenext
		loop
	end if		

    If int_Qtd_Relacao = 0 Then
        str_Sql = ""
        str_Sql = str_Sql & " DELETE FROM "
        str_Sql = str_Sql & " CURSO_FUNCAO "
        str_Sql = str_Sql & " WHERE "
        str_Sql = str_Sql & " CURS_CD_CURSO='" & pCursoPrePre & "'"
        str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & pfuncao & "'"
    
        db.Execute (str_Sql)
        
        str_Sql = ""
        str_Sql = str_Sql & " SELECT DISTINCT "
        str_Sql = str_Sql & " CURS_PRE_REQUISITO"
        str_Sql = str_Sql & " FROM dbo.CURSO_PRE_REQUISITO"
        str_Sql = str_Sql & " WHERE "
        str_Sql = str_Sql & " CURS_CD_CURSO = '" & pCursoPrePre & "'"
        
        Set rds_Curso_Pre_Pre_Pre = db.Execute(str_Sql)
        If Not rds_Curso_Pre_Pre_Pre.EOF Then
            pCursoPrePre2 = rds_Curso_Pre_Pre_Pre("CURS_PRE_REQUISITO")
            'response.Write("<p> Curso2 = " & pCursoPrePre2)
            'response.Write("<p> Funcao2 = " & pFuncao)
            ' CRIA UAM RECURSIVIDADE
            Call deleta_funcao_prepre(pCursoPrePre2, pfuncao, pCursoPrePre)
        End If
        
    End If
	
end sub

Sub incluir_Pre(pCursoPre)

	'response.Write("<p> Incluir")

	seq=0

	str_Sql = ""
	str_Sql = str_Sql & " SELECT "
	str_Sql = str_Sql & " MAX(CUPR_NR_SEQUENCIA) AS CODIGO "
	str_Sql = str_Sql & " FROM CURSO_PRE_REQUISITO "
	str_Sql = str_Sql & " WHERE "
	str_Sql = str_Sql & " CURS_CD_CURSO='" & curso & "'"
	
	set rs_seq=db.execute(str_Sql)	
	if not isnull(rs_seq("CODIGO")) then
		seq=rs_seq("CODIGO")+1
	else
		seq=1
	end if	
	rs_seq.close
	
	'**** INSERE O NOVO PRE REQUISITO
	str_Sql = ""
	str_Sql = str_Sql & " INSERT INTO "
	str_Sql = str_Sql & " CURSO_PRE_REQUISITO ("
	str_Sql = str_Sql & " CURS_CD_CURSO"
	str_Sql = str_Sql & " ,CUPR_NR_SEQUENCIA"
	str_Sql = str_Sql & " ,CURS_PRE_REQUISITO"
	str_Sql = str_Sql & " ,ATUA_TX_OPERACAO"
	str_Sql = str_Sql & " ,ATUA_CD_NR_USUARIO"
	str_Sql = str_Sql & " ,ATUA_DT_ATUALIZACAO"
	str_Sql = str_Sql & " ) VALUES("
	str_Sql = str_Sql & " '" & curso & "'"
	str_Sql = str_Sql & " ," & seq 
	str_Sql = str_Sql & " ,'" & pCursoPre & "'"
	str_Sql = str_Sql & " ,'I'"
	str_Sql = str_Sql & " ,'" & Session("CdUsuario") & "'"
	str_Sql = str_Sql & " ,GETDATE())"
	db.execute(str_Sql)

	'*** VERIFICA AS FUNÇÕES LIGADAS AO CURSO E ASSOCIA AO NOVO PRE
	str_Sql = ""
	str_Sql = str_Sql & " SELECT DISTINCT "
	str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO "
	str_Sql = str_Sql & " FROM "
	str_Sql = str_Sql & " CURSO_FUNCAO "
	str_Sql = str_Sql & " WHERE "
	str_Sql = str_Sql & " CURS_CD_CURSO='" & curso & "'"
	
	set fonte_funcao_todas=db.execute(str_Sql)	
	do until fonte_funcao_todas.eof=true
		str_Cd_Funcao = fonte_funcao_todas("FUNE_CD_FUNCAO_NEGOCIO")
		Call grava_funcao_prepre(pCursoPre, str_Cd_Funcao)
		fonte_funcao_todas.movenext
	loop
	fonte_funcao_todas.close

end sub

sub grava_funcao_prepre(pCursoPrePre, pFuncao)

	'response.Write("<p> Curso = " & pCursoPrePre)
	'response.Write("<p> Funcao = " & pFuncao)	
	

		str_Sql = ""
		str_Sql = str_Sql & " SELECT DISTINCT "
		str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO "
		str_Sql = str_Sql & " FROM CURSO_FUNCAO "
		str_Sql = str_Sql & " WHERE CURS_CD_CURSO='" & pCursoPrePre & "'"
		str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & pFuncao & "'"
		
		set rds_Func_Curso_f2=db.execute(str_Sql)
		if rds_Func_Curso_f2.Eof then		

			'response.Write("<p> Curso1 = " & pCursoPrePre)
			'response.Write("<p> Funcao1 = " & pFuncao)	

			str_Sql = ""
			str_Sql = str_Sql & " INSERT INTO "
			str_Sql = str_Sql & " CURSO_FUNCAO ("
			str_Sql = str_Sql & " CURS_CD_CURSO"
			str_Sql = str_Sql & " ,FUNE_CD_FUNCAO_NEGOCIO"
			str_Sql = str_Sql & " ,CUFU_TX_INDICA_MOSTRA"		
			str_Sql = str_Sql & " ,ATUA_TX_OPERACAO"
			str_Sql = str_Sql & " ,ATUA_CD_NR_USUARIO"
			str_Sql = str_Sql & " ,ATUA_DT_ATUALIZACAO"
			str_Sql = str_Sql & " ) VALUES("
			str_Sql = str_Sql & " '" & pCursoPrePre & "'"
			str_Sql = str_Sql & " ,'" & pFuncao & "'"
			str_Sql = str_Sql & " ,'N'"
			str_Sql = str_Sql & " ,'I'"
			str_Sql = str_Sql & " ,'" & Session("CdUsuario") & "'"
			str_Sql = str_Sql & " ,GETDATE())"			
			db.execute(str_Sql)

		end if
		rds_Func_Curso_f2.close
		
		str_Sql = ""
		str_Sql = str_Sql & " SELECT DISTINCT "
		str_Sql = str_Sql & " CURS_PRE_REQUISITO"
		str_Sql = str_Sql & " FROM dbo.CURSO_PRE_REQUISITO"
		str_Sql = str_Sql & " WHERE "
		str_Sql = str_Sql & " CURSO_PRE_REQUISITO.CURS_CD_CURSO = '" & pCursoPrePre & "'"
		set rds_Curso_Pre_Pre=db.execute(str_Sql)
		if not rds_Curso_Pre_Pre.Eof then		
			pCursoPrePre2 = rds_Curso_Pre_Pre("CURS_PRE_REQUISITO")

			'response.Write("<p> Curso2 = " & pCursoPrePre2)
			'response.Write("<p> Funcao2 = " & pFuncao)	
			
			' CRIA UAM RECURSIVIDADE 
			Call grava_funcao_prepre(pCursoPrePre2, pFuncao)
		end if
		rds_Curso_Pre_Pre.close
		
end sub

sub grava_funcao_prepre_ORIGINAL(pCursoPrePre, pFuncao)

		str_Sql = ""
		str_Sql = str_Sql & " SELECT DISTINCT "
		str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO "
		str_Sql = str_Sql & " FROM CURSO_FUNCAO "
		str_Sql = str_Sql & " WHERE CURS_CD_CURSO='" & pCursoPre & "'"
		'str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & fonte_funcao_todas("FUNE_CD_FUNCAO_NEGOCIO") & "'"
		str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & pFuncao & "'"
		
		set rds_Func_Curso_f2=db.execute(str_Sql)
		if rds_Func_Curso_f2.Eof then		
			str_Sql = ""
			str_Sql = str_Sql & " INSERT INTO "
			str_Sql = str_Sql & " CURSO_FUNCAO ("
			str_Sql = str_Sql & " CURS_CD_CURSO"
			str_Sql = str_Sql & " ,FUNE_CD_FUNCAO_NEGOCIO"
			str_Sql = str_Sql & " ,CUFU_TX_INDICA_MOSTRA"		
			str_Sql = str_Sql & " ,ATUA_TX_OPERACAO"
			str_Sql = str_Sql & " ,ATUA_CD_NR_USUARIO"
			str_Sql = str_Sql & " ,ATUA_DT_ATUALIZACAO"
			str_Sql = str_Sql & " ) VALUES("
			str_Sql = str_Sql & " '" & pCursoPre & "'"
			str_Sql = str_Sql & " ,'" & fonte_funcao_todas("FUNE_CD_FUNCAO_NEGOCIO") & "'"
			str_Sql = str_Sql & " ,'N'"
			str_Sql = str_Sql & " ,'I'"
			str_Sql = str_Sql & " ,'" & Session("CdUsuario") & "'"
			str_Sql = str_Sql & " ,GETDATE())"			
			db.execute(ssql)

			Dim vet_Pre_Pre(3)

			str_Sql = ""
			str_Sql = str_Sql & " SELECT DISTINCT "
			str_Sql = str_Sql & " CURSO_PRE_REQUISITO_1.CURS_PRE_REQUISITO"
			str_Sql = str_Sql & " , CURSO_PRE_REQUISITO_2.CURS_PRE_REQUISITO AS Expr1 "
			str_Sql = str_Sql & " , CURSO_PRE_REQUISITO_3.CURS_PRE_REQUISITO AS Expr2"
			str_Sql = str_Sql & " FROM dbo.CURSO_PRE_REQUISITO CURSO_PRE_REQUISITO_2 RIGHT OUTER JOIN"
			str_Sql = str_Sql & " dbo.CURSO_PRE_REQUISITO CURSO_PRE_REQUISITO_1 ON "
			str_Sql = str_Sql & " CURSO_PRE_REQUISITO_2.CURS_CD_CURSO = CURSO_PRE_REQUISITO_1.CURS_PRE_REQUISITO LEFT OUTER JOIN"
			str_Sql = str_Sql & " dbo.CURSO_PRE_REQUISITO CURSO_PRE_REQUISITO_3 ON "
			str_Sql = str_Sql & " CURSO_PRE_REQUISITO_2.CURS_PRE_REQUISITO = CURSO_PRE_REQUISITO_3.CURS_CD_CURSO"
			str_Sql = str_Sql & " WHERE "
			str_Sql = str_Sql & " CURSO_PRE_REQUISITO_1.CURS_CD_CURSO = '" & a & "'"
			set rds_Curso_Pre_Pre=db.execute(str_Sql)
			if rds_Curso_Pre_Pre.Eof then		
				vet_Pre_Pre(0) = rds_Curso_Pre_Pre("CURS_PRE_REQUISITO")
				vet_Pre_Pre(1) = rds_Curso_Pre_Pre("Expr1")
				vet_Pre_Pre(2) = rds_Curso_Pre_Pre("Expr2")
				For i = LBound(vet_Pre_Pre) + 1 To UBound(vet_Pre_Pre)
					if not IsNull(vet_Pre_Pre(i)) then
					
					end if
				next
			end if
		end if
		fonte_funcao_todas.movenext
		rds_Func_Curso_f2.close

end sub

'Function SequentialSearchStringArray(ByRef sArray() As String, ByVal sFind As String) As Long
Function SequentialSearchStringArray(sArray(), sFind) 

	Dim i       'As Long
	Dim iLBound 'As Long
	Dim iUBound 'As Long
	
	iLBound = LBound(sArray)
	iUBound = UBound(sArray)

	For i = iLBound To iUBound
		'response.Write("<P>" & "----PROCURA----" & i &  sArray(i) & " < - > " & sFind  )   		
	  If sArray(i) = sFind Then SequentialSearchStringArray = i: Exit Function
	Next 'i
	
	SequentialSearchStringArray = -1
End Function

%>

<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="../Planilhas/js/troca_lista.js"></script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_cad_curso.asp" name="frm1">
        <input type="hidden" name="txtImp" size="20"><input type="hidden" name="txtQua" size="20">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif" width="30" height="30"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif" width="30" height="30"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
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
        
  <table width="847" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="845">
      </td>
    </tr>
    <tr>
      <td width="845">
        <p align="center"><font face="Verdana" color="#330099" size="3">Relação
          Curso x Pré - Requisito</font>
      </td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="57">
    <tr> 
      <td width="124" height="29"></td>
      <td width="56" height="29" valign="middle" align="left"></td>
      <td height="29" valign="middle" align="left" colspan="2"> <%if err.number=0 then%> <b><font face="Verdana" color="#330099" size="2">O Curso 
        e seus Pré-Requisitos foram relacionados com Sucesso</font></b> </td>
    </tr>
    <%else%>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> <b><font face="Verdana" size="2" color="#800000">Houve 
        um erro no cadastro do registro - <%=err.description%></font></b> </td>
    </tr>
    <%end if%>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>

    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="109"> <a href="../../indexA.asp"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right" width="22" height="20"></a></td>
      <td height="1" valign="middle" align="left" width="542"> <font face="Verdana" color="#330099" size="2">Retornar 
        para Tela Principal</font></td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="109"> <a href="rel_curso_pre_requisitos.asp?mega=<%=mega%>&amp;curso=<%=curso%>"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right" width="22" height="20"></a></td>
      <td height="1" valign="middle" align="left" width="542"> <font face="Verdana" color="#330099" size="2">Retornar 
        para Tela de Relacionar Curso x Pré-Requisitos</font></td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="109"> </td>
      <td height="1" valign="middle" align="left" width="542"> </td>
    </tr>
    <tr> 
      <td width="124" height="1"></td>
      <td width="56" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" colspan="2"> </td>
    </tr>
  </table>
  </form>

</body>

</html>
