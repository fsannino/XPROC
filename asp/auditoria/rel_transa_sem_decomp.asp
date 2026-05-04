<%
response.Buffer=false
Server.ScriptTimeOut=9990000

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MODULO=REQUEST("selModulo")
ATIVIDADE=REQUEST("selAtividade")
TRANSACAO=REQUEST("selTransacao")

IF MODULO<>0 THEN
COMPL1="dbo.MODU_ATIV_TRA_CARGA.MODU_CD_MODULO=" & MODULO
END IF

IF ATIVIDADE<>0 THEN
COMPL2="dbo.MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA=" & ATIVIDADE
END IF

IF TRANSACAO<>"0" THEN
COMPL3="dbo.MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO='" & TRANSACAO & "'"
END IF

IF COMPL1<>"" THEN
COMPLE = COMPL1
END IF

IF COMPL2<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE + " AND " + COMPL2
ELSE
COMPLE=COMPL2
END IF
END IF

IF COMPL3<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL3
ELSE
COMPLE=COMPL3
END IF
END IF

IF COMPLE<>"" THEN
CONECTA=" AND "
END IF

ORDENA=" ORDER BY dbo.MODU_ATIV_TRA_CARGA.MODU_CD_MODULO, dbo.MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA, dbo.MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO"

'SSQL="SELECT * FROM " & Session("PREFIXO") & "MODU_ATIV_TRA_CARGA " & CONECTA & COMPLE & ORDENA

str_Sql = ""
str_Sql = str_Sql & " SELECT dbo.MODU_ATIV_TRA_CARGA.MODU_CD_MODULO"
str_Sql = str_Sql & " , dbo.MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA"
str_Sql = str_Sql & " ,dbo.MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO"
str_Sql = str_Sql & " , dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_Sql & " FROM dbo.MODU_ATIV_TRA_CARGA LEFT OUTER JOIN"
str_Sql = str_Sql & " dbo.RELACAO_FINAL ON dbo.MODU_ATIV_TRA_CARGA.TRAN_CD_TRANSACAO = dbo.RELACAO_FINAL.TRAN_CD_TRANSACAO AND "
str_Sql = str_Sql & " dbo.MODU_ATIV_TRA_CARGA.ATCA_CD_ATIVIDADE_CARGA = dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA AND "
str_Sql = str_Sql & " dbo.MODU_ATIV_TRA_CARGA.MODU_CD_MODULO = dbo.RELACAO_FINAL.MODU_CD_MODULO"
str_Sql = str_Sql & " WHERE dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO IS NULL"
str_Sql = str_Sql & CONECTA & COMPLE & ORDENA
'response.Write(str_Sql)
set rs=db.execute(str_Sql)
%>

<html>
<head>
<script>
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
</script>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0">
<form method="POST" action="../gera_rel_modatca.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
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
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="65%"><div align="center"><b><font color="#330099" size="3" face="Verdana">Rela&ccedil;&atilde;o de Transa&ccedil;&otilde;es sem decomposi&ccedil;&atilde;o</font> </b></div></td>
      <td width="35%"><font face="Verdana" color="#330099" size="3"> <img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></font></td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3"></font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<%if rs.eof=true then%>  
<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#800000" size="2"><b>Nenhum
Registro Encontrado</b></font></p>
<%end if%>

<%if rs.eof=false then%>
<table border="0" width="100%" cellspacing="3">
  <tr>
    <td width="33%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Agrupamento
      das Atividades</b></font></td>
    <td width="33%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Atividade</b></font></td>
    <td width="34%" bgcolor="#A8C4C6"><font face="Verdana" size="2"><b>Transação</b></font></td>
  </tr>
  <%
  IF RS.EOF<>TRUE THEN
  VALOR_MODULO=RS("MODU_CD_MODULO")
  VALOR_ATIVIDADE=RS("ATCA_CD_ATIVIDADE_CARGA")
  VALOR_TRANSACAO=RS("TRAN_CD_TRANSACAO")
  END IF
  
  VALOR1=""
  VALOR2=""
  VALOR3=""

  DO UNTIL RS.EOF=TRUE
  
  'IF MODULO_ANTERIOR<>VALOR_MODULO THEN
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & VALOR_MODULO)
  VALOR1=RS1("MODU_TX_DESC_MODULO")
  'END IF
  
  IF MODULO_ANTERIOR<>VALOR_MODULO or ATIVIDADE_ANTERIOR <> VALOR_ATIVIDADE THEN
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & VALOR_ATIVIDADE)
  VALOR2=RS1("ATCA_TX_DESC_ATIVIDADE")
  END IF
  
  SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & VALOR_TRANSACAO & "'")
  VALOR3=RS("TRAN_CD_TRANSACAO") & "-"& RS1("TRAN_TX_DESC_TRANSACAO")
  
  %>
  <tr>
    <%if valor2<>"" then%>
    <td width="33%" bgcolor="#FFCC00"><font face="Verdana" size="1"><%=VALOR1%></font></td>
    <%else%>
    <td width="33%"><font face="Verdana" size="1"></font></td>
    <%end if%>
	
    <%if valor2<>"" then%>
    <td width="33%" bgcolor="#FFFF00"><font face="Verdana" size="1"><%=VALOR2%></font></td>
	<%else%>
	<td width="33%"><font face="Verdana" size="1"><%=VALOR2%></font></td>
    <%end if%>
    
    <%if valor3<>"" then%>
    <td width="34%" bgcolor="#FFCAB0"><font face="Verdana" size="1"><%=VALOR3%></font></td>
    <%else%>
    <td width="34%"><font face="Verdana" size="1"><%=VALOR3%></font></td>

    <%end if%>
  </tr>
  <%
  
  MODULO_ANTERIOR=RS("MODU_CD_MODULO")
  ATIVIDADE_ANTERIOR=RS("ATCA_CD_ATIVIDADE_CARGA")
  
  RS.MOVENEXT
  
  IF RS.EOF<>TRUE THEN
  VALOR_MODULO=RS("MODU_CD_MODULO")
  VALOR_ATIVIDADE=RS("ATCA_CD_ATIVIDADE_CARGA")
  VALOR_TRANSACAO=RS("TRAN_CD_TRANSACAO")
  END IF
  
  VALOR1=""
  VALOR2=""
  VALOR3=""
  
  LOOP
  %>
</table>
<%end if%>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
</form>

</body>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>
