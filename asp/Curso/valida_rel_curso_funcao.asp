<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

Set rds_Curso_Pre_Requisito_Duplo = CreateObject("ADODB.Recordset")
rds_Curso_Pre_Requisito_Duplo.CursorLocation = 3

dim vet_Func_Cadastrada(2000)
'dim vet_Func_Selecionada(10)
'vet_Func_Cadastrada(0) = ""
'vet_Func_Selecionada(1) = ""

curso=Ucase(request("curso"))
mega=request("mega")

func=request("txtTrans")
vet_Func_Selecionada = split(func,",")

'strCdFuncao = vet_Func_Selecionada(1)
'response.Write(strCdFuncao)
'response.End()

'***********************************************************************************************
' CARREGA VETOR COMOS CADASTRADOS
'***********************************************************************************************
str_Sql = ""
str_Sql = str_Sql & " SELECT DISTINCT "
str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO "
str_Sql = str_Sql & " FROM CURSO_FUNCAO "
str_Sql = str_Sql & " WHERE CURS_CD_CURSO='" & curso & "'"
set rds_Func_Curso=db.execute(str_Sql)
int_Indice = 0
if not rds_Func_Curso.Eof then
	do while not rds_Func_Curso.Eof
		vet_Func_Cadastrada(int_Indice) = rds_Func_Curso("FUNE_CD_FUNCAO_NEGOCIO")
		int_Indice = int_Indice + 1
		rds_Func_Curso.movenext
	loop
end if	

rds_Func_Curso.close
set rds_Func_Curso = Nothing

'response.Write("<p>" & UBound(vet_Func_Cadastrada))
'response.Write("<p>" & UBound(vet_Func_Selecionada))

'int_Indice = 0 
'for int_Indice = LBound(vet_Func_Cadastrada) to UBound(vet_Func_Cadastrada)
'	response.Write("<p>Indice-Cad " & int_Indice & " : " & vet_Func_Cadastrada(int_Indice))
'next
'int_Indice = 0 
'for int_Indice = LBound(vet_Func_Selecionada) to UBound(vet_Func_Selecionada)
'	response.Write("<p>Indice-Sel " & int_Indice & " : " & vet_Func_Selecionada(int_Indice))
'next
'response.End()

'***********************************************************************************************
' VERIFICA QUEM SERÁ EXCLUIDO
'***********************************************************************************************
int_Indice = 0 
for int_Indice = LBound(vet_Func_Cadastrada) to UBound(vet_Func_Cadastrada)
	lng_Posicao = SequentialSearchStringArray(vet_Func_Selecionada, vet_Func_Cadastrada(int_Indice))
	if lng_Posicao = -1 then
		deletar_funcao(vet_Func_Cadastrada(int_Indice))
	end if
	response.Write("<P>" & "----" & int_Indice  &  vet_Func_Cadastrada(int_Indice))
next			
response.Write("<P>" & "----FINAL CADA----INDICE - " & int_Indice  & "VET - " & vet_Func_Cadastrada(int_Indice-1))
'***********************************************************************************************
' VERIFICA QUEM SERÁ INCLUÍDO
'***********************************************************************************************
int_Indice = 0 
for int_Indice = LBound(vet_Func_Selecionada) to UBound(vet_Func_Selecionada)
	lng_Posicao = SequentialSearchStringArray(vet_Func_Cadastrada, vet_Func_Selecionada(int_Indice))
	if lng_Posicao = -1 then
		incluir_funcao(vet_Func_Selecionada(int_Indice))
	end if
	response.Write("<P>" & "----" & int_Indice  &  vet_Func_Selecionada(int_Indice))	
next			
response.Write("<P>" & "----FINAL SELE----INDICE - " & int_Indice  & "VET - " &  vet_Func_Selecionada(int_Indice-1))
'RESPONSE.End()

sub deletar_funcao(pFuncao)
	
	str_Sql = ""
	str_Sql = str_Sql & " SELECT     "
	str_Sql = str_Sql & " CURS_CD_CURSO"
	str_Sql = str_Sql & " , CURS_PRE_REQUISITO"
	str_Sql = str_Sql & " , CUPR_NR_SEQUENCIA"
	str_Sql = str_Sql & " FROM  dbo.CURSO_PRE_REQUISITO"
	str_Sql = str_Sql & " WHERE CURS_CD_CURSO = '" & curso & "'"
	'response.Write(str_Sql)
	'response.End()
	set rds_Curso_Pre_Requisito_f = db.execute(str_Sql)
	if not rds_Curso_Pre_Requisito_f.Eof then
		do while not rds_Curso_Pre_Requisito_f.Eof					
			str_Sql = ""
			str_Sql = str_Sql & " SELECT  DISTINCT"
			str_Sql = str_Sql & " dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO"
			str_Sql = str_Sql & " , dbo.CURSO_FUNCAO.CURS_CD_CURSO"
			str_Sql = str_Sql & " , dbo.CURSO_FUNCAO.CUFU_TX_INDICA_MOSTRA"
			str_Sql = str_Sql & " , dbo.CURSO_PRE_REQUISITO.CURS_PRE_REQUISITO"
			str_Sql = str_Sql & " , dbo.CURSO_FUNCAO.ATUA_DT_ATUALIZACAO"
			str_Sql = str_Sql & " FROM dbo.CURSO_FUNCAO INNER JOIN"
			str_Sql = str_Sql & " dbo.CURSO_PRE_REQUISITO ON dbo.CURSO_FUNCAO.CURS_CD_CURSO = dbo.CURSO_PRE_REQUISITO.CURS_CD_CURSO"
			str_Sql = str_Sql & " WHERE "
			str_Sql = str_Sql & " dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO = '" & pFuncao & "'"
			str_Sql = str_Sql & " AND dbo.CURSO_PRE_REQUISITO.CURS_PRE_REQUISITO = '" & rds_Curso_Pre_Requisito_f("CURS_PRE_REQUISITO") & "'"
			'response.Write(str_Sql)
			'response.End()
			rds_Curso_Pre_Requisito_Duplo.Open str_Sql, db, 3, 4						
			'set rds_Curso_Pre_Requisito_Duplo=db.execute(str_Sql)
			if not rds_Curso_Pre_Requisito_Duplo.Eof then
				if rds_Curso_Pre_Requisito_Duplo.Recordcount < 2 then				
					str_Sql = ""
					str_Sql = str_Sql & " DELETE "
					str_Sql = str_Sql & " FROM CURSO_FUNCAO "
					str_Sql = str_Sql & " WHERE "
					str_Sql = str_Sql & " CURS_CD_CURSO='" & rds_Curso_Pre_Requisito_f("CURS_PRE_REQUISITO") & "'"
					str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & pFuncao & "'"
					str_Sql = str_Sql & " AND CUFU_TX_INDICA_MOSTRA = 'N'"
					'response.Write(str_Sql)
					'response.End()
					db.execute(str_Sql)	
				end if
			end if			
			rds_Curso_Pre_Requisito_Duplo.close
			rds_Curso_Pre_Requisito_f.movenext
		loop
	end if

	str_Sql = ""
	str_Sql = str_Sql & " DELETE "
	str_Sql = str_Sql & " FROM CURSO_FUNCAO "
	str_Sql = str_Sql & " WHERE "
	str_Sql = str_Sql & " CURS_CD_CURSO='" & curso & "'"
	str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & pFuncao & "'"
	'response.Write(str_Sql)
	'response.End()
	db.execute(str_Sql)
	
	set rds_Curso_Pre_Requisito_Duplo = Nothing
	rds_Curso_Pre_Requisito_f.Close
	set rds_Curso_Pre_Requisito_f = Nothing

end sub

Sub incluir_funcao(pFuncao)

	str_Sql = ""
	str_Sql = str_Sql & " INSERT INTO CURSO_FUNCAO ( "
	str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO"
	str_Sql = str_Sql & " , CURS_CD_CURSO"
	str_Sql = str_Sql & " , CUFU_TX_INDICA_MOSTRA"
	str_Sql = str_Sql & " , ATUA_TX_OPERACAO"
	str_Sql = str_Sql & " , ATUA_CD_NR_USUARIO"
	str_Sql = str_Sql & " , ATUA_DT_ATUALIZACAO"
	str_Sql = str_Sql & " ) VALUES ("
	str_Sql = str_Sql & "'" & pFuncao & "'"
	str_Sql = str_Sql & ",'" & curso & "'"
	str_Sql = str_Sql & ",'S'"
	str_Sql = str_Sql & ",'I','" & Session("CdUsuario") & "',GETDATE())"
	'response.Write(str_Sql)
	'response.End()
	db.execute(str_Sql)
		
	str_Sql = ""
	str_Sql = str_Sql & " SELECT     "
	str_Sql = str_Sql & " CURS_CD_CURSO"
	str_Sql = str_Sql & " , CURS_PRE_REQUISITO"
	str_Sql = str_Sql & " , CUPR_NR_SEQUENCIA"
	str_Sql = str_Sql & " FROM  dbo.CURSO_PRE_REQUISITO"
	str_Sql = str_Sql & " WHERE CURS_CD_CURSO = '" & curso & "'"
	'response.Write(str_Sql)
	'response.End()
	set rds_Curso_Pre_Requisito = db.execute(str_Sql)
	if not rds_Curso_Pre_Requisito.Eof then
		do while not rds_Curso_Pre_Requisito.Eof
			call Grava_Func_CursoPreRequisito(pFuncao,rds_Curso_Pre_Requisito("CURS_PRE_REQUISITO"))
			rds_Curso_Pre_Requisito.movenext
		loop
	end if
end sub

Sub Grava_Func_CursoPreRequisito(pFuncao,pCurso)

	str_Sql = ""
	str_Sql = str_Sql & " SELECT DISTINCT "
	str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO "
	str_Sql = str_Sql & " FROM CURSO_FUNCAO "
	str_Sql = str_Sql & " WHERE CURS_CD_CURSO='" & pCurso & "'"
	str_Sql = str_Sql & " AND FUNE_CD_FUNCAO_NEGOCIO='" & pFuncao & "'"
	
	set rds_Func_Curso_f2=db.execute(str_Sql)
	if rds_Func_Curso_f2.Eof then
		str_Sql = ""
		str_Sql = str_Sql & " INSERT INTO CURSO_FUNCAO ("
		str_Sql = str_Sql & " FUNE_CD_FUNCAO_NEGOCIO"
		str_Sql = str_Sql & " , CURS_CD_CURSO"
		str_Sql = str_Sql & " , CUFU_TX_INDICA_MOSTRA"
		str_Sql = str_Sql & " , ATUA_TX_OPERACAO"
		str_Sql = str_Sql & " , ATUA_CD_NR_USUARIO"
		str_Sql = str_Sql & " , ATUA_DT_ATUALIZACAO"
		str_Sql = str_Sql & " ) VALUES ("
		str_Sql = str_Sql & "'" & pFuncao & "'"
		str_Sql = str_Sql & ",'" & pCurso & "'"
		str_Sql = str_Sql & ",'N'"
		str_Sql = str_Sql & ",'I','" & Session("CdUsuario") & "',GETDATE())"
		'on error resume next	
		db.execute(str_Sql)
		'if err.number=0 then
			'call grava_log(ucase(curso),"" & Session("PREFIXO") & "CURSO_FUNCAO","I",1)
		'end if
	end if	
	
end sub

'Function SequentialSearchStringArray(ByRef sArray() As String, ByVal sFind As String) As Long
Function SequentialSearchStringArray(sArray(), sFind) 

	Dim i       'As Long
	Dim iLBound 'As Long
	Dim iUBound 'As Long
	
	iLBound = LBound(sArray)
	iUBound = UBound(sArray)

	For i = iLBound To iUBound
		response.Write("<P>" & "----PROCURA----" & i &  sArray(i) & " < - > " & sFind  )   		
	  If sArray(i) = sFind Then SequentialSearchStringArray = i: Exit Function
	Next 'i
	
	SequentialSearchStringArray = -1
End Function

'response.End()

response.redirect "rel_curso_func_transacao.asp?curso=" & curso

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
          Curso x Furção de Negócio</font>
      </td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="87">
          <tr>
            
      <td width="205" height="29"></td>
            
      <td width="93" height="29" valign="middle" align="left"></td>
            
      <td width="531" height="29" valign="middle" align="left" colspan="2"> 
      <%if err.number=0 then%>
      <b><font face="Verdana" color="#330099" size="2">O Curso e Funções de
      Negócio foram
      relacionados com
      Sucesso</font></b> 
      </td>
            
          </tr>
      <%else%>    
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      <b><font face="Verdana" size="2" color="#800000">Houve um erro no cadastro
      do registro - <%=err.description%></font></b> 
      </td>
          </tr>
          <%end if%>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
        <a href="../../indexA.asp"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right" width="22" height="20"></a></td>
            
      <td height="1" valign="middle" align="left" width="439"> 
        <font face="Verdana" color="#330099" size="2">Retornar para Tela
        Principal</font></td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
        <a href="rel_curso_funcao.asp?mega=<%=mega%>&amp;curso=<%=curso%>"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right" width="22" height="20"></a></td>
            
      <td height="1" valign="middle" align="left" width="439"> 
        <font face="Verdana" color="#330099" size="2">Retornar para Tela de Relacionar Curso x
        Fun&ccedil;&atilde;o R/3</font></td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
      </td>
            
      <td height="1" valign="middle" align="left" width="439"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
      </td>
            
      <td height="1" valign="middle" align="left" width="439"> 
      </td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> 
      </td>
          </tr>
        </table>
  </form>

</body>

</html>
