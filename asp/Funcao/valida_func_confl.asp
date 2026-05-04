<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_funcao=request("selFuncao")
str_mega=request("selMegaProcesso")

db.execute("DELETE FROM " & Session("PREFIXO") & "FUN_NEG_CONFL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

Sub Grava_Funcao(Funcao,mega,valor)

   str_SQL = ""
   str_SQL = str_SQL & " insert into " & Session("PREFIXO") & "FUN_NEG_CONFL "
   str_SQL = str_SQL & " (FUNE_CD_FUNCAO_NEGOCIO, FUNC_CD_FUNCAO_CONFL"
   str_SQL = str_SQL & " , ATUA_TX_OPERACAO, ATUA_CD_NR_USUARIO , ATUA_DT_ATUALIZACAO)"
   str_SQL = str_SQL & " VALUES( '" & Funcao & "','" & VALOR & "', "
   str_SQL = str_SQL & "'I','" & Session("CdUsuario") & "',GETDATE())"

   'RESPONSE.WRITE STR_SQL
   db.execute(str_SQL)
   
   ''call grava_log(str_funcao,"" & Session("PREFIXO") & "FUN_NEG_CONFL","I",1)

end sub

str_valor=request("txtImp")

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
'response.write str_valor
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
        
        call Grava_Funcao(str_funcao,str_mega,str_atual)
	    
	    valor_total=valor_total+1
	    
        quantos = 0
    End If
    contador = contador + 1
Loop

db.Close
set db = Nothing

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="../MacroPerfil/js/troca_lista.js"></script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
      <td colspan="3" height="20">&nbsp;</td>
  </tr>
</table>
        
  <table width="847" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="845">
      </td>
    </tr>
    <tr>
      <td width="845">
        <p align="center"><font face="Verdana" color="#330099">Fun&ccedil;&atilde;o R/3 Conflitante</font></td>
    </tr>
    <tr>
      <td width="845">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="81">
          <tr>
            
      <td width="205" height="29"></td>
            
      <td width="93" height="29" valign="middle" align="left"></td>
            
      <td width="531" height="29" valign="middle" align="left" colspan="2"> 
      <%if err.number=0 then%>
        <b><font face="Verdana" color="#330099" size="2">Opera&ccedil;&atilde;o 
        realizada com Sucesso - <%=str_funcao%></font></b> 
      </td>
            
          </tr>
      <%else%>    
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531" colspan="2"> <b><font face="Verdana" size="2" color="#800000">Houve 
        um erro na opera&ccedil;&atilde;o - <%=err.description%></font></b> 
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
        <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
            
      <td height="1" valign="middle" align="left" width="439"> 
        <font face="Verdana" color="#330099" size="2">Retornar para Tela
        Principal</font></td>
          </tr>
          <tr>
            
      <td width="205" height="1"></td>
            
      <td width="93" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="93"> 
        <a href="func_confl.asp?selMegaProcesso=<%=str_mega%>&selFuncao=<%=str_funcao%>"> 
        <img border="0" src="../../imagens/selecao_F02.gif" align="right"></a></td>
            
      <td height="1" valign="middle" align="left" width="439"> 
        <font face="Verdana" size="2" color="#330099">Retornar para cadastro de 
        Função Conflitante</font></td>
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