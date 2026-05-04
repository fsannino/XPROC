<%
response.Buffer=false

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY MODU_TX_DESC_MODULO")
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO")
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
<form method="POST" action="rel_transa_sem_decomp.asp" name="frm1">
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
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
          <td width="26"><h3><a href="#"><img border="0" src="../../imagens/confirma_f02.gif" onclick="javascript:submit()"></a></h3></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Montar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
            <td width="27">&nbsp;</td>  
            <td width="50"><font color="#330099" face="Verdana" size="2">&nbsp;</font></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; </font></p>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="65%"><div align="center">
        <p style="margin-top: 0; margin-bottom: 0" align="center"></p>
        <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Sele&ccedil;&atilde;o para rela&ccedil;&atilde;o de transa&ccedil;&otilde;es sem decomposi&ccedil;&atilde;o </font></p>
      </div></td>
      <td width="35%"><font face="Verdana" color="#330099" size="3"> <img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></font></td>
    </tr>
  </table>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="88%" height="235">
  <tr>
    <td width="24%" height="21"></td>
    <td width="76%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        o Agrupamento ( Master List R/3 ) (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="24%" height="25"></td>
    <td width="76%" height="25"><select size="1" name="selModulo">
          <option value="0">== Selecione o Agrupamento ( Master List R/3 ) ==</option>
        <%DO UNTIL RS1.EOF=TRUE%>
                    <option value="<%=RS1("MODU_CD_MODULO")%>"><%=RS1("MODU_TX_DESC_MODULO")%></option>
          <%
          RS1.MOVENEXT
          LOOP
          %>
        </select></td>
  </tr>
  <tr>
    <td width="24%" height="17"></td>
    <td width="76%" height="17"></td>
  </tr>
  <tr>
    <td width="24%" height="21"></td>
    <td width="76%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        a Atividade (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="24%" height="25"></td>
    <td width="76%" height="25"><select size="1" name="selAtividade">
          <option value="0">== Selecione a Atividade ==</option>
          <%DO UNTIL RS2.EOF=TRUE%>
                    <option value="<%=RS2("ATCA_CD_ATIVIDADE_CARGA")%>"><%=left(RS2("ATCA_TX_DESC_ATIVIDADE"),75)%></option>
          <%
          RS2.MOVENEXT
          LOOP
          %>
        </select></td>
  </tr>
  <tr>
    <td width="24%" height="23"></td>
    <td width="76%" height="23"></td>
  </tr>
  <tr>
    <td width="24%" height="21"></td>
    <td width="76%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        a Transação (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="24%" height="25"></td>
    <td width="76%" height="25"><select size="1" name="selTransacao">
          <option value="0">== Selecione a Transação ==</option>
          <%DO UNTIL RS3.EOF=TRUE%>
                    <option value="<%=RS3("TRAN_CD_TRANSACAO")%>"><%=RS3("TRAN_CD_TRANSACAO")%>-<%=LEFT(RS3("TRAN_TX_DESC_TRANSACAO"),50)%></option>
          <%
          RS3.MOVENEXT
          LOOP
          %>
        </select></td>
  </tr>
  <tr>
    <td width="24%" height="21"></td>
    <td width="76%" height="21"></td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>
