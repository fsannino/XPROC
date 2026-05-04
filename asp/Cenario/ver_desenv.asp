<%@LANGUAGE="VBSCRIPT"%> 
 
<%
set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

trans=request("transacao")

set rs=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO_DESENV WHERE TRAN_CD_TRANSACAO='" & trans  & "' order by DESE_CD_DESENVOLVIMENTO" )

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Visualização de Desenvolvimentos</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#800000" vlink="#800000" alink="#800000">
<form name="frm1" method="post" action="http://its_server3/valida_status1.asp">
        <input type="hidden" name="txtOpc" value="1"><input type="hidden" name="INC" size="20" value="1"> 
        <table border="0" width="97%">
          <tr>
            <td width="50%"> 
        <p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="3">Visualização
        de Desenvolvimentos</font></p>
        <p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p align="left" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#330099" face="Verdana" size="2"><b>Transação
        Selecionada : <%=TRANS%></b> </font></p>
            </td>
            <td width="50%">
              <p align="right" style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#800000" size="3"><a href="javascript:window.close()">Fechar
              Janela</a></font></b></p>
            </td>
          </tr>
        </table>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <table border="0" width="98%">
          <tr>
            <td width="17%" bgcolor="#330099" align="center">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FFFFFF" size="2" face="Verdana"><b>Codigo
              Desenv</b></font></p>
            </td>
            <td width="42%" bgcolor="#330099" align="center">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FFFFFF" size="2" face="Verdana"><b>Descrição</b></font></p>
            </td>
            <td width="26%" bgcolor="#330099" align="center">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FFFFFF" size="2" face="Verdana"><b>Data
              Prevista para Conclusão</b></font></p>
            </td>
            <td width="39%" bgcolor="#330099" align="center">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#FFFFFF" size="2" face="Verdana"><b>Data
              de Conclusão</b></font></p>
            </td>
          </tr>
          <%
          tem=0
          do until rs.eof=true
          set temp=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "DESENVOLVIMENTO WHERE DESE_CD_DESENVOLVIMENTO='"& rs("DESE_CD_DESENVOLVIMENTO") &"'")
          if cor="#E0E0E0" then
          		cor="white"
          	else
          		cor="#E0E0E0"
          	end if
          %>
          <tr>
            <td width="17%" align="center" bgcolor=<%=cor%>>
              <font size="1" face="Verdana">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
              <%=rs("DESE_CD_DESENVOLVIMENTO")%></font></td>
            <td width="42%" align="center" bgcolor=<%=cor%>>
              <font face="Verdana" size="1">
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><%=temp("DESE_tx_desc_DESENVOLVIMENTO")%></font></td>
            <td width="26%" align="center" bgcolor=<%=cor%>>
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
              <%
              data1=day(temp("DESE_DT_PREVISTA_REALIZACAO")) &"/" & month(temp("DESE_DT_PREVISTA_REALIZACAO")) &"/" &year(temp("DESE_DT_PREVISTA_REALIZACAO"))
              if data1="//" then
              	data1=""
              end if

              %>
              <font face="Verdana" size="1"><%=data1%></font></td>
            <td width="39%" align="center" bgcolor=<%=cor%>>
            <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
            <%
              data2=day(temp("DESE_DT_CONCLUSAO")) &"/" & month(temp("DESE_DT_CONCLUSAO")) &"/" &year(temp("DESE_DT_CONCLUSAO"))
              if data2="//" then
              	data2=""
              end if
              %>
              <font face="Verdana" size="1"><%=data2%></font></td>
          </tr>
          <%
          tem = tem + 1
          rs.movenext
          loop
          %>
         </table>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">
        <%if tem=0 then%>
        <font color="#800000"><b>
        Nenhum Registro Encontrado!</b></font>
        <%end if%>
        <input type="hidden" name="txtcaminho" size="20">
  </form>
</body>
</html>
