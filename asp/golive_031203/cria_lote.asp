<%@LANGUAGE="VBSCRIPT"%> 
<%
response.Buffer=false
server.ScriptTimeout=3600
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
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
      <td width="5%">&nbsp;</td>
      <td width="69%">&nbsp;</td>
      <td width="26%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td><img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  
  <%
set conn_Cogest=server.createobject("ADODB.CONNECTION")
conn_Cogest.Open Session("Conn_String_Cogest_Gravacao")
conn_Cogest.cursorlocation = 3

dim vet_Fun_Antecipada(16)
dim vet_Fun_Definitiva(16)

vet_Fun_Antecipada(0) = "MM.51"
vet_Fun_Definitiva(0) = "MM.12"
vet_Fun_Antecipada(1) = "MM.66"
vet_Fun_Definitiva(1) = "MM.100"
vet_Fun_Antecipada(2) = "MM.59"
vet_Fun_Definitiva(2) = "MM.112"
vet_Fun_Antecipada(3) = "MM.61"
vet_Fun_Definitiva(3) = "MM.114"
vet_Fun_Antecipada(4) = "MM.57"
vet_Fun_Definitiva(4) = "MM.104"
vet_Fun_Antecipada(5) = "PM.023"
vet_Fun_Definitiva(5) = "PM.034"
vet_Fun_Antecipada(6) = "PM.024"
vet_Fun_Definitiva(6) = "PM.046"
vet_Fun_Antecipada(7) = "PM.027"
vet_Fun_Definitiva(7) = "PM.035"
vet_Fun_Antecipada(8) = "PM.027"
vet_Fun_Definitiva(8) = "PM.037"
vet_Fun_Antecipada(9) = "PM.026"
vet_Fun_Definitiva(9) = "PM.035"
vet_Fun_Antecipada(10) = "PM.022"
vet_Fun_Definitiva(10) = "PM.038"
vet_Fun_Antecipada(11) = "PM.025"
vet_Fun_Definitiva(11) = "PM.041"
vet_Fun_Antecipada(12) = "PM.028"
vet_Fun_Definitiva(12) = "PM.043"

vet_Fun_Antecipada(13) = "PM.029"
vet_Fun_Definitiva(13) = "PM.031"
vet_Fun_Antecipada(14) = "PM.029"
vet_Fun_Definitiva(14) = "PM.032"
vet_Fun_Antecipada(15) = "PM.029"
vet_Fun_Definitiva(15) = "PM.033"

str_Lotacao = 0

if request("txtDescLote") <> "" then
   str_DescLote = request("txtDescLote")
else
   str_DescLote = ""
end if

if request("txtOrgSel") <> "" then
   str_OrgSel = request("txtOrgSel")
else
   str_OrgSel = ""
end if

if request("txtDescOrgao") <> "" then
   str_DescOrgao = request("txtDescOrgao")
else
   str_DescOrgao = ""
end if

'response.Write(str_OrgSel & "<p>")
'response.Write(str_DescOrgao & "<p>")
'response.End()

if request("Str01") <> 0 then
   str_Str01 = request("Str01")
   'str_Lotacao = Right("00" & str_Str01,2)   
   str_Lotacao = str_Str01
   int_inicio_String = 1
   int_fim_String = 2
else
   str_Str01 = 0
end if

if request("Str02") <> 0 then
   str_Str02 = request("Str02")
   'str_Lotacao = Right("000" & str_Str02, 3)   
   str_Lotacao = str_Str02
   int_inicio_String = 3
   int_fim_String = 3   
else
   str_Str02 = 0
end if
if request("Str03") <> 0 then
   str_Str03 = request("Str03")
   'str_Lotacao = Right("00" & str_Str03,2)   
   str_Lotacao = str_Str03
   int_inicio_String = 1
   int_fim_String = 2         
else
   str_Str03 = 0
end if
'response.Write(request("Str01") & "<p>")
'response.Write(request("Str02") & "<p>")
'response.Write(request("Str03") & "<p>")
'response.Write(str_Lotacao & "<p>")
'response.End()
if request("selMegaProcesso") <> 0 then
   str_MegaProcesso = request("selMegaProcesso")
else
   str_MegaProcesso = 0
end if

if request("selOnda") <> 0 then
   str_Onda = request("selOnda")
else
   str_Onda = 0
end if

'response.Write(request("txtFuncSel") & "<p>")
'response.End()
str_FuncSel = ""
int_Tamanho = Len(Request("txtFuncSel"))
if request("txtFuncSel") <> ",0" and request("txtFuncSel") <> "" then
	vet_FuncSelec = split(Mid(Request("txtFuncSel"),2,int_Tamanho-1), ",")
	'response.Write(UBound(vet_FuncSelec))
	'response.End()
	for i=0 to UBound(vet_FuncSelec) 
	    if i < UBound(vet_FuncSelec) then
			str_Virgula = ","
		else
			str_Virgula = ""
		end if
		str_FuncSel =  str_FuncSel & "'" & vet_FuncSelec(i) & "'" & str_Virgula
	next   
else
   str_FuncSel = "0"
end if

'response.Write(request("txtFuncSel") & "<p>")
'response.End()
int_Tamanho = Len(Request("txtOrgSel"))
if request("txtOrgSel") <> ",0" and request("txtOrgSel") <> "" then
	str_OrgSel = "("
	vet_OrgSelec = split(Mid(Request("txtOrgSel"),2,int_Tamanho-1), ",")
	'response.Write(UBound(vet_FuncSelec))
	'response.End()
	for i=0 to UBound(vet_OrgSelec) 
	    if i < UBound(vet_OrgSelec) then
			str_Operador =  "OR"
		else
			str_Operador = ""
		end if
		str_OrgSel =  str_OrgSel & " ORME_CD_ORG_MENOR like '" & vet_OrgSelec(i) & "%' " & str_Operador 
	next   
	str_OrgSel =  str_OrgSel & ")"
else
   str_OrgSel = ""
end if

'response.Write(str_OrgSel)
'response.End()

dim int_Num_Lote
dim boo_Criado_Lote

boo_Criado_Lote = False
int_Num_Lote = 0

str_SQL = ""
str_SQL = str_SQL & " Select "
str_SQL = str_SQL & " ORME_CD_ORG_MENOR"
str_SQL = str_SQL & " , USMA_CD_USUARIO "
str_SQL = str_SQL & " , FUNE_CD_FUNCAO_NEGOCIO "
str_SQL = str_SQL & " , CURS_CD_CURSO "
str_SQL = str_SQL & " from USU_CUR_FUN "
str_SQL = str_SQL & " Where USMA_CD_USUARIO > '0'"
str_SQL = str_SQL & " AND FUUS_IN_VALIDADO = 'S'"

'if str_Lotacao <> 0 then
'	str_SQL = str_SQL & " AND ORME_CD_ORG_MENOR like '" & str_Lotacao & "%'"
'	'str_SQL = str_SQL & " AND Substring(ORME_CD_ORG_MENOR," & int_inicio_String & "," & int_fim_String  & ") = '" & str_Lotacao & "'"
'end if
if str_FuncSel <> "0" then
	str_SQL = str_SQL & " AND FUNE_CD_FUNCAO_NEGOCIO in (" & str_FuncSel & ")"
end if
if str_MegaProcesso <> 0 then
	str_SQL = str_SQL & " AND MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso 
end if
if str_OrgSel <> "" then
	str_SQL = str_SQL & " and " & str_OrgSel
end if

str_SQL = str_SQL & " order by USMA_CD_USUARIO , FUNE_CD_FUNCAO_NEGOCIO  "

'response.Write(str_SQL & "<p>")
'response.End()

Set rds_Usu_Curso = conn_Cogest.Execute(str_SQL)
int_NumReg_Usu_Curso = rds_Usu_Curso.RecordCount
'response.Write(str_SQL & "<p>")
'response.Write(int_NumReg_Usu_Curso)
'response.End()
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
			lng_Posicao = SequentialSearchStringArray(vet_Fun_Definitiva, str_Cd_Fun_Anterior)
			if lng_Posicao <> - 1 then
				str_Cd_Fun_Anterior = vet_Fun_Antecipada(lng_Posicao)
			end if
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
						If rds_TabIncr("USAP_TX_APROVEITAMENTO") = "AP" or rds_TabIncr("USAP_TX_APROVEITAMENTO") = "LM" Then
							int_Aprovado = int_Aprovado + 1
						ELSE
							int_Nao_Aprovado = int_Nao_Aprovado + 1							
						End If
						int_Loop_TabIncr = int_Loop_TabIncr + 1
						rds_TabIncr.MoveNext
					Loop
				Else
					int_Nao_Aprovado = int_Nao_Aprovado + 1
				End If
				'if int_Aprovado = 0 then
				'   int_Nao_Aprovado = int_Nao_Aprovado + 1
				'end if
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
	if int_Num_Lote > 0 then
		str_Msg = "Criado Lote = " & int_Num_Lote & " - " & UCase(str_DescLote)
	else
		str_Msg = "Não existem registros para exportação"
	end if		
Else
	str_Msg = "Não existem registros para exportação"
End If

Sub f_grava_registro (str_Cd_Usu_Anterior,str_Cd_Fun_Anterior,str_Status)

    str_SQL = ""
    str_SQL = str_SQL & " select "
    str_SQL = str_SQL & " FUNE_CD_FUNCAO_NEGOCIO "
    str_SQL = str_SQL & " FROM GOLI_FUNCAO_USUARIO "
    str_SQL = str_SQL & " WHERE USMA_CD_USUARIO = '" & str_Cd_Usu_Anterior & "'"
    str_SQL = str_SQL & " AND FUNE_CD_FUNCAO_NEGOCIO = '" & str_Cd_Fun_Anterior & "'"
	str_SQL = str_SQL & " AND USFU_TX_INDICA_GERA_SAIDA = 'S'"
    
	set rstRepeticao = conn_Cogest.Execute(str_SQL)
	
    If rstRepeticao.EOF Then

        If boo_Criado_Lote = False Then
           int_Num_Lote = f_Cria_Lote()
           boo_Criado_Lote = True
        End If
		
		str_SQL = " SELECT DISTINCT "
		str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO, dbo.FUNCAO_USUARIO_PERFIL.USMA_CD_USUARIO"
		str_SQL = str_SQL & " FROM dbo.FUNCAO_USUARIO_PERFIL INNER JOIN"
		str_SQL = str_SQL & " dbo.MICRO_PERFIL_R3 ON "
		str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.MCPR_NR_SEQ_MACRO_PERFIL = dbo.MICRO_PERFIL_R3.MCPR_NR_SEQ_MACRO_PERFIL AND "
		str_SQL = str_SQL & " dbo.FUNCAO_USUARIO_PERFIL.MIPE_NR_SEQ_MICRO_PERFIL = dbo.MICRO_PERFIL_R3.MIPE_NR_SEQ_MICRO_PERFIL"
		str_SQL = str_SQL & " WHERE dbo.FUNCAO_USUARIO_PERFIL.FUNE_CD_FUNCAO_NEGOCIO = '" & str_Cd_Fun_Anterior & "'"
		str_SQL = str_SQL & " AND dbo.FUNCAO_USUARIO_PERFIL.USMA_CD_USUARIO = '" & str_Cd_Usu_Anterior & "'"
		str_SQL = str_SQL & " AND FUUP_IN_VALIDADO = 'S'"
		'response.Write(str_SQL)
		'response.End()
		set rds_Perfil = conn_Cogest.Execute(str_SQL)
		if not rds_Perfil.Eof then
		   str_Saida = "S"
		else
		   str_Saida = "N"		
		end if
		str_SQL = ""
		str_SQL = str_SQL & " Insert into GOLI_FUNCAO_USUARIO("
		str_SQL = str_SQL & " LOTE_NR_SEQ_LOTE"
		str_SQL = str_SQL & " , USMA_CD_USUARIO"
		str_SQL = str_SQL & " , FUNE_CD_FUNCAO_NEGOCIO"
		str_SQL = str_SQL & " , USFU_TX_APRO_TREINA"
		str_SQL = str_SQL & " , USFU_TX_APRO_XPROC"
		str_SQL = str_SQL & " , USFU_TX_INDICA_GERA_SAIDA"		
		str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
		str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
		str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
		str_SQL = str_SQL & " )Values("
		str_SQL = str_SQL & "'" & int_Num_Lote & "',"	
		str_SQL = str_SQL & "'" & str_Cd_Usu_Anterior & "',"
		str_SQL = str_SQL & "'" & str_Cd_Fun_Anterior & "',"
		str_SQL = str_SQL & "'" & str_Status & "',"
		str_SQL = str_SQL & "'" & str_Status & "',"	
		str_SQL = str_SQL & "'" & str_Saida & "',"	
		str_SQL = str_SQL & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
		Set rdsNovo = conn_Cogest.Execute(str_SQL)

	end if
	
end sub

Function f_Cria_Lote()
        
    str_SQL = ""
    str_SQL = str_SQL & " SELECT MAX(LOTE_NR_SEQ_LOTE)AS NUM_LOTE FROM GOLI_LOTE "
        
    Set rs = conn_Cogest.Execute(str_SQL)
    
    If Not IsNull(rs("NUM_LOTE")) Then
        int_Num_Lote = rs("NUM_LOTE")
    Else
        int_Num_Lote = 0
    End If
    
    If int_Num_Lote = 0 Then
        int_Num_Lote = 1
    Else
        int_Num_Lote = int_Num_Lote + 1
    End If
        
    str_SQL = ""
    str_SQL = str_SQL & " Insert into GOLI_LOTE("
    str_SQL = str_SQL & " LOTE_NR_SEQ_LOTE"
    str_SQL = str_SQL & " , LOTE_TX_DESCRICAO"	
    str_SQL = str_SQL & " , LOTE_DT_ENVIO"
	str_SQL = str_SQL & " , LOTE_NR_QTD_EXPORTACAO"
	str_SQL = str_SQL & " , LOTE_TX_ORGAO_SELEC"
	str_SQL = str_SQL & " , LOTE_TX_FUNCAO_SELEC"	
    str_SQL = str_SQL & " , ATUA_TX_OPERACAO"
    str_SQL = str_SQL & " , ATUA_CD_NR_USUARIO"
    str_SQL = str_SQL & " , ATUA_DT_ATUALIZACAO"
    str_SQL = str_SQL & " ) Values ("
    str_SQL = str_SQL & int_Num_Lote & ","
	if str_DescLote <> "" then
		str_DescLote = UCase(str_DescLote)
	end if	
	str_SQL = str_SQL & "'" & str_DescLote & "'," 
    str_SQL = str_SQL & "GETDATE(),"
	str_SQL = str_SQL & "0,"
	str_SQL = str_SQL & "'" & Left(str_DescOrgao,300) & "'," 
	str_SQL = str_SQL & "'" & Left(Request("txtFuncSel"),150) & "'," 		
    str_SQL = str_SQL & "'C' ,'" & Session("CdUsuario")  & "' ,GETDATE())"
	'response.Write(str_SQL)
	'response.End()
    conn_Cogest.Execute str_SQL
    
    f_Cria_Lote = int_Num_Lote
    
End Function

Function SequentialSearchStringArray(ByRef sArray() As String, ByVal sFind As String) As Long
   Dim i       As Long
   Dim iLBound As Long
   Dim iUBound As Long

   iLBound = LBound(sArray)
   iUBound = UBound(sArray)

   For i = iLBound To iUBound
      If sArray(i) = sFind Then SequentialSearchStringArray = i: Exit Function
   Next i

   SequentialSearchStringArray = -1
End Function
%>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="66%">&nbsp;</td>
      <td width="10%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="66%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Msg%></font> </td>
      <td width="10%">&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
		<% if int_Num_Lote > 0 then %>
      <td><font face="Verdana, Arial, Helvetica, sans-serif"><a href="consulta_lote_usu_func.asp?str_Tipo_Saida=Tela&pLote=<%=int_Num_Lote%>&pDescLote=<%=str_DescLote%>&pVezesImp=0"><font size="2">Prepara arquivo saida</font></a> </font></td>
		<% end if %>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><a href="javascript:history.go(-1)"><img src="../../imagens/selecao_F02_off.gif" width="22" height="20" border="0"></a> <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Retorna para tela de cria&ccedil;&atilde;o de lote para exporta&ccedil;&atilde;o </font></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
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
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>
