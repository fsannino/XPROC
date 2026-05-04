<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
int_Total_Tot_Empresa = 0

set Conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Atividade = request("selAtividade")
str_Empresa_Unid = Request("txtEmpSelecionada")
str_Empresa_Unid_Original = str_Empresa_Unid
if Len(str_Empresa_Unid) > 2 then    
   if InStr(str_Empresa_Unid,",10") <> 0 or InStr(str_Empresa_Unid,",11") <> 0 or InStr(str_Empresa_Unid,",12") <> 0 then      
      str_Empresa_Unid = str_Empresa_Unid + ","
	  if InStr(str_Empresa_Unid,",1,") <> 0 then
         msg = "tem e achou o 1"
	  else
	     msg = "tem e não achou o 1"
		 str_Empresa_Unid = str_Empresa_Unid_Original & ",1" 
	  end if
   else
      msg = "não tem 10/11/12"
   end if	  
end if

conn_db.execute("DELETE FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID WHERE ATCA_CD_ATIVIDADE_CARGA=" & str_Atividade)

Sub Grava_Nova_Atividade(strA, strE)
	str_SQL_Nova_Sub_Empr = ""
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID ( "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ATCA_CD_ATIVIDADE_CARGA, EMPR_CD_NR_EMPRESA, "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ATUA_DT_ATUALIZACAO) VALUES ("
   	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strA & "," & strE & ", "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	
   	'response.write str_SQL_Nova_Sub_Empr
   	
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Sub_Empr)
	strChave = CStr(strA) & " " & CStr(strE) 
	'call grava_log(strChave,"ATIVIDADE_CARGA_EMPRESA_UNID","I",0)
	
end sub

'guarda o conteúdo da String
str_valor = str_Empresa_Unid

'Coloca uma virgula no fim de string, se não houver
if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if

'Pega o tamanho da string
tamanho = Len(str_valor)

'Retira a vírgula do início da string, se houver
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If

'Atualiza o tamanho da string
tamanho = Len(str_valor)

'Inicializa o Contador
contador = 1

'início da Rotina
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
	' Aqui entra o que vc quer fazer com o caracter em questão!
	   call Grava_Nova_Atividade(str_Atividade, str_atual)
	   
	   valor_total=valor_total+1

        quantos = 0
    End If
    contador = contador + 1
Loop

conn_db.Close
set conn_db = Nothing

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frm1.selMegaProcesso.value+"'");
}
function Confirma2() 
{ 
	  document.frm1.submit();
}
function Confirma() 
{ 
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
	 else
     {
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top" height="65"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a>
            </div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27">&nbsp;</td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="78%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registros gravados :</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=valor_total%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Relação
    entre Atividade x Empresa realizada com sucesso!</b></font> 
      
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="relacao_ativ_emp_.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Rela&ccedil;&atilde;o Atividade / Empresa Unid</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
int_Total_Tot_Empresa = 0

set Conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_Atividade = request("selAtividade")
str_Empresa_Unid = Request("txtEmpSelecionada")
str_Empresa_Unid_Original = str_Empresa_Unid
if Len(str_Empresa_Unid) > 2 then    
   if InStr(str_Empresa_Unid,",10") <> 0 or InStr(str_Empresa_Unid,",11") <> 0 or InStr(str_Empresa_Unid,",12") <> 0 then      
      str_Empresa_Unid = str_Empresa_Unid + ","
	  if InStr(str_Empresa_Unid,",1,") <> 0 then
         msg = "tem e achou o 1"
	  else
	     msg = "tem e não achou o 1"
		 str_Empresa_Unid = str_Empresa_Unid_Original & ",1" 
	  end if
   else
      msg = "não tem 10/11/12"
   end if	  
end if

conn_db.execute("DELETE FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID WHERE ATCA_CD_ATIVIDADE_CARGA=" & str_Atividade)

Sub Grava_Nova_Atividade(strA, strE)
	str_SQL_Nova_Sub_Empr = ""
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE_CARGA_EMPRESA_UNID ( "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ATCA_CD_ATIVIDADE_CARGA, EMPR_CD_NR_EMPRESA, "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO, "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & " ATUA_DT_ATUALIZACAO) VALUES ("
   	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & strA & "," & strE & ", "
	str_SQL_Nova_Sub_Empr = str_SQL_Nova_Sub_Empr & "'C' ,'" & Session("CdUsuario") & "' ,GETDATE())"
   	
   	'response.write str_SQL_Nova_Sub_Empr
   	
	Set rdsNovo = conn_db.Execute(str_SQL_Nova_Sub_Empr)
	strChave = CStr(strA) & " " & CStr(strE) 
	'call grava_log(strChave,"ATIVIDADE_CARGA_EMPRESA_UNID","I",0)
	
end sub

'guarda o conteúdo da String
str_valor = str_Empresa_Unid

'Coloca uma virgula no fim de string, se não houver
if right(str_valor,1)<>"," then
    str_valor = str_valor + ","
end if

'Pega o tamanho da string
tamanho = Len(str_valor)

'Retira a vírgula do início da string, se houver
If Left(str_valor, 1) = "," Then
    tamanho = tamanho - 1
    str_valor = Right(str_valor, tamanho)
End If

'Atualiza o tamanho da string
tamanho = Len(str_valor)

'Inicializa o Contador
contador = 1

'início da Rotina
Do Until contador = tamanho + 1
    str_atual = Left(str_valor, contador)
    quantos = quantos + 1
    str_temp = Right(str_atual, 1)
    tamanho_atual = Len(str_atual)
    If str_temp = "," Then
        str_atual = Right(str_atual, quantos)
        str_atual = Left(str_atual, quantos - 1)
        
	' Aqui entra o que vc quer fazer com o caracter em questão!
	   call Grava_Nova_Atividade(str_Atividade, str_atual)
	   
	   valor_total=valor_total+1

        quantos = 0
    End If
    contador = contador + 1
Loop

conn_db.Close
set conn_db = Nothing

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<script language="JavaScript">
<!--
function MM_goToURL1() { //v3.0
  var i, args=MM_goToURL1.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+document.frm1.selMegaProcesso.value+"'");
}
function Confirma2() 
{ 
	  document.frm1.submit();
}
function Confirma() 
{ 
if (document.frm1.selMegaProcesso.selectedIndex == 0)
     { 
	 alert("A seleção de um Mega Processo é obrigatório!");
     document.frm1.selMegaProcesso.focus();
     return;
     }
	 else
     {
	  document.frm1.submit();
	 }
 }

function Limpa(){
	document.frm1.reset();
}

//-->
</script>

</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr> 
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top" height="65"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center"> 
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a>
            </div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
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
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27">&nbsp;</td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="78%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp; </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Total 
      de registros gravados :</font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><%=valor_total%></font></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%">&nbsp;</td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      
    <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366"><b>Relação
    entre Atividade x Empresa realizada com sucesso!</b></font> 
      
    </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"></td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
      <table width="90%" border="0" cellpadding="0" cellspacing="0" align="center">
        <tr> 
          <td height="41"><a href="relacao_ativ_emp_.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela de Rela&ccedil;&atilde;o Atividade / Empresa Unid</font></td>
        </tr>
        <tr> 
          <td height="41"><a href="../indexA.asp"><img src="../imagens/selecao_F02.gif" width="22" height="20" border="0"></a></td>
          <td height="41"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Volta 
            para tela Principal</font></td>
        </tr>
      </table>
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="14%">&nbsp;</td>
    <td width="70%"> 
     </td>
    <td width="16%">&nbsp;</td>
  </tr>
</table>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
