 
<%
str_cenario=request("txtCenario")
opt=request("OPTION")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

'response.write  str_cenario

ssql="SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='"& trim(str_cenario) & "'"
set rs=db.execute(ssql)

valor=rs("CENA_TX_TITULO_CENARIO")

set rstrans=db.execute("SELECT * FROM " & Session("PREFIXO") & "HISTORICO_CENARIO WHERE CENA_CD_CENARIO='"& trim(str_cenario) & "' order by HICE_NR_SEQUENCIAL")
%>
<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">

function fechar()
{
window.close();
}

function mover()
{
window.moveTo(100	,200)
}

</script>

<body topmargin="0" leftmargin="0" onload="javascript:mover()">
<table border="0" width="79%">
  <tr> 
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td width="82%"><font face="Verdana" size="2" color="#000080"><b>Relação de 
      Hist&oacute;ricos cadastrados</b></font></td>
    <td width="18%"><b><a href="#" onClick="javascript:fechar()"><font face="Verdana" size="2" color="#800000">Fechar 
      Janela</font></a></b> </td>
  </tr>
  <tr> 
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="2"><font face="Verdana" size="2" color="#000080"><b>Cenário :</b> 
      <%=cenario%>-<%=valor%></font></td>
  </tr>
  <tr> 
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <%if rstrans.eof=true then%>
      <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Nenhum hist&oacute;rico 
      cadastrado para este Cenário</font></b> 
      <% else %>
      <%end if%>
    </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <table width="622" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="68" bgcolor="#0066CC"> 
            <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Seq</font></b></div>
          </td>
          <td width="529" bgcolor="#0066CC"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Texto</font></b></td>
          <td width="25">&nbsp;</td>
        </tr>
        <%DO UNTIL RSTRANS.EOF=TRUE%>
        <tr> 
          <td width="68" valign="top"> 
            <div align="center"><%=RSTRANS("HICE_NR_SEQUENCIAL")%></div>
          </td>
          <td width="529"><%=RSTRANS("HICE_TX_HISTORICO")%></td>
          <td width="25">&nbsp;</td>
        </tr>
        <%
         RSTRANS.MOVENEXT
         LOOP
         %>
        <tr> 
          <td width="68">&nbsp;</td>
          <td width="529">&nbsp;</td>
          <td width="25">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <p align="right"><b><a href="#" onClick="javascript:fechar()"><font face="Verdana" size="2" color="#800000">Fechar 
        Janela</font></a></b> 
    </td>
  </tr>
</table>

</body>

</html>
