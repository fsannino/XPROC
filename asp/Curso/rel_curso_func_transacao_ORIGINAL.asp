<%

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

curso=request("curso")

str_Sql = ""
str_Sql = str_Sql & " Delete from CURSO_TRANSACAO WHERE TRAN_CD_TRANSACAO IN ("
str_Sql = str_Sql & " SELECT  dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO  "
str_Sql = str_Sql & " FROM dbo.CURSO_FUNCAO (nolock) INNER JOIN  dbo.FUN_NEG_TRANSACAO (nolock) ON  "
str_Sql = str_Sql & " dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO RIGHT OUTER JOIN "
str_Sql = str_Sql & " dbo.CURSO_TRANSACAO (nolock) ON dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO AND "
str_Sql = str_Sql & " dbo.CURSO_FUNCAO.CURS_CD_CURSO = dbo.CURSO_TRANSACAO.CURS_CD_CURSO "
str_Sql = str_Sql & " WHERE dbo.CURSO_TRANSACAO.CURS_CD_CURSO = '" & curso & "'"
str_Sql = str_Sql & " AND dbo.CURSO_FUNCAO.CURS_CD_CURSO IS NULL )"
db.execute(str_Sql)

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & curso & "'")
valor1=rs("CURS_CD_CURSO") & " - " & rs("CURS_TX_NOME_CURSO")

' errado porque está olhando transação associadas a ela propria. Prolema quando é Função de Referencia
str_Sql = ""
str_Sql = str_Sql & " SELECT DISTINCT TOP 100 PERCENT dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO, dbo.TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_Sql = str_Sql & " FROM dbo.CURSO_FUNCAO INNER JOIN "
str_Sql = str_Sql & " dbo.FUN_NEG_TRANSACAO ON " 
str_Sql = str_Sql & " dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO = dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN "
str_Sql = str_Sql & " dbo.TRANSACAO ON dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO.TRAN_CD_TRANSACAO "
str_Sql = str_Sql & " WHERE dbo.CURSO_FUNCAO.CURS_CD_CURSO = '" & curso & "'" 
str_Sql = str_Sql & " AND dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO NOT IN ( "
str_Sql = str_Sql & " SELECT DISTINCT dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO "
str_Sql = str_Sql & " FROM   dbo.CURSO_TRANSACAO INNER JOIN dbo.TRANSACAO ON dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO.TRAN_CD_TRANSACAO"
str_Sql = str_Sql & " WHERE  dbo.CURSO_TRANSACAO.CURS_CD_CURSO = '" & curso & "')"
str_Sql = str_Sql & " ORDER BY dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO "

' certo olhando a função PAI
str_Sql = ""
str_Sql = str_Sql & " SELECT DISTINCT TOP 100 PERCENT dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO, dbo.TRANSACAO.TRAN_TX_DESC_TRANSACAO "
str_Sql = str_Sql & " FROM dbo.FUNCAO_NEGOCIO INNER JOIN "
str_Sql = str_Sql & " dbo.CURSO_FUNCAO ON dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO = dbo.CURSO_FUNCAO.FUNE_CD_FUNCAO_NEGOCIO INNER JOIN "
str_Sql = str_Sql & " dbo.TRANSACAO INNER JOIN "
str_Sql = str_Sql & " dbo.FUN_NEG_TRANSACAO ON dbo.TRANSACAO.TRAN_CD_TRANSACAO = dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO ON  "
str_Sql = str_Sql & " dbo.FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO_PAI = dbo.FUN_NEG_TRANSACAO.FUNE_CD_FUNCAO_NEGOCIO "
str_Sql = str_Sql & " WHERE dbo.CURSO_FUNCAO.CURS_CD_CURSO = '" & curso & "'" 
str_Sql = str_Sql & " AND dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO NOT IN ( "
str_Sql = str_Sql & " SELECT DISTINCT dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO "
str_Sql = str_Sql & " FROM   dbo.CURSO_TRANSACAO INNER JOIN dbo.TRANSACAO ON dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO.TRAN_CD_TRANSACAO"
str_Sql = str_Sql & " WHERE  dbo.CURSO_TRANSACAO.CURS_CD_CURSO = '" & curso & "')"
str_Sql = str_Sql & " ORDER BY dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO "

'RESPONSE.WRITE str_Sql
set rst_TransNaoAssociadas = db.execute(str_Sql)

str_Sql = ""
str_Sql = str_Sql & " SELECT DISTINCT dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO, dbo.TRANSACAO.TRAN_TX_DESC_TRANSACAO"
str_Sql = str_Sql & " FROM   dbo.CURSO_TRANSACAO INNER JOIN dbo.TRANSACAO ON dbo.CURSO_TRANSACAO.TRAN_CD_TRANSACAO = dbo.TRANSACAO.TRAN_CD_TRANSACAO"
str_Sql = str_Sql & " WHERE  dbo.CURSO_TRANSACAO.CURS_CD_CURSO = '" & curso & "'"
str_Sql = str_Sql & " ORDER BY dbo.TRANSACAO.TRAN_TX_DESC_TRANSACAO "
'RESPONSE.WRITE str_Sql
set rst_TransAssociadas = db.execute(str_Sql)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="javascript" src="../js/troca_toda_lista.js"></script>

<script>
function Confirma() 
   {
   carrega_txt(document.frm1.list2)
   document.frm1.submit();
   }

function carrega_txt(fbox) 
   {
   document.frm1.txtTrans.value = "";
   for(var i=0; i<fbox.options.length; i++) 
      {
      document.frm1.txtTrans.value = document.frm1.txtTrans.value + "," + fbox.options[i].value;
      }
   }
</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_rel_curso_transacao.asp" name="frm1">
        <input type="hidden" name="txtTrans">
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
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../Funcao/confirma_f02.gif"></a></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Enviar</b></font></td>
          <td width="26">&nbsp;</td>
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
        
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>
      </td>
    </tr>
    <tr>
      <td>
        <div align="center"><font face="Verdana" color="#330099" size="3">Relação
          Curso x Transação</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="47">
    <tr> 
      <td width="162" height="19"></td>
      <td width="136" height="19" valign="middle" align="left"><div align="right"><font face="Verdana" size="2" color="#330099"><b>Curso 
          :&nbsp;</b></font></div></td>
      <td width="531" height="19" valign="middle" align="left"> <font face="Verdana" size="2" color="#330099"><%=VALOR1%></font></td>
    </tr>
    <tr> 
      <td width="162" height="1"></td>
      <td width="136" height="19" valign="middle" align="left"></td>
      <td width="531" height="19" valign="middle" align="left"><input name="curso" type="hidden" id="curso" value="<%=curso%>"> 
      </td>
    </tr>
    <tr> 
      <td width="162" height="1"></td>
      <td width="136" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="531"> </td>
    </tr>
  </table>

  <p style="margin: 0" align="center"><font face="Verdana" size="2" color="#330099"></font></p>
        
  <table border="0" width="846" height="142">
    <tr>
            
      <td width="24" height="138" rowspan="5"></td>
            
      <td width="429" height="138" rowspan="5" valign="top"> 
        <table width="92%" border="0">
          <tr>
            <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Transa&ccedil;&otilde;es 
              n&atilde;o selecionadas</strong></font></td>
          </tr>
          <tr>
            <td><b>
              <select name="list1" size="8" multiple>
                <% While (NOT rst_TransNaoAssociadas.EOF)
%>
                <option value="<%=(rst_TransNaoAssociadas.Fields.Item("TRAN_CD_TRANSACAO").Value)%>" ><%=(rst_TransNaoAssociadas.Fields.Item("TRAN_CD_TRANSACAO").Value) & "-" & (rst_TransNaoAssociadas.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></option>
                <%
  rst_TransNaoAssociadas.MoveNext()
Wend
If (rst_TransNaoAssociadas.CursorType > 0) Then
  rst_TransNaoAssociadas.MoveFirst
Else
  rst_TransNaoAssociadas.Requery
End If
rst_TransNaoAssociadas.close
set rst_TransNaoAssociadas = Nothing
%>
              </select>
              </b></td>
          </tr>
        </table>
      </td>
            <td width="90" height="28" align="center">
              </td>
            
      <td width="403" height="138" rowspan="5" valign="top">
<table width="90%" border="0">
          <tr> 
            <td><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Transa&ccedil;&otilde;es 
              selecionadas</strong></font></td>
          </tr>
          <tr> 
            <td><font color="#000080">
              <select name="list2" size="8" multiple>
                <% While (NOT rst_TransAssociadas.EOF)
%>
                <option value="<%=(rst_TransAssociadas.Fields.Item("TRAN_CD_TRANSACAO").Value)%>"><%=(rst_TransAssociadas.Fields.Item("TRAN_CD_TRANSACAO").Value) & "-" & (rst_TransAssociadas.Fields.Item("TRAN_TX_DESC_TRANSACAO").Value)%></option>
                <%
  rst_TransAssociadas.MoveNext()
Wend
If (rst_TransAssociadas.CursorType > 0) Then
  rst_TransAssociadas.MoveFirst
Else
  rst_TransAssociadas.Requery
End If
rst_TransAssociadas.close
set rst_TransAssociadas = Nothing
%>
              </select>
              </font></td>
          </tr>
        </table>
      </td>
          </tr>
          <tr>
            <td width="90" height="28" align="center"><a href="#" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="../Funcao/continua_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr>
            
      <td width="90" height="28" align="center"><a href="#" onClick="movetudo(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="../../imagens/seta_dupla_direita.gif" width="24" height="24"></a></td>
          </tr>
          <tr>
            <td width="90" height="27" align="center"><a href="javascript:;"  onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="../Funcao/continua2_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr>
            
      <td width="90" height="27" align="center"><a href="javascript:;"  onClick="movetudo(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="../../imagens/seta_dupla_esquerda.gif" width="24" height="24"></a></td>
          </tr>
        </table>

<p style="margin: 0">&nbsp;</p>

<p style="margin: 0">&nbsp;</p>

<p style="margin: 0">&nbsp;</p>
  </form>

</body>

</html>
