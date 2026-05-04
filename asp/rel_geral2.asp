<<<<<<< HEAD
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=0
str_proc=0

str_mega=request("selMegaProcesso")
str_proc=request("selProcesso")

if session("MegaProcesso")<>0 and str_mega=0 then
	str_mega=session("MegaProcesso")
end if

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

if str_mega<>0 then
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
else
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega<>0 and str_proc<>0 then
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc)
else
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega<>0 and str_proc=0 then
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
end if

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY MODU_TX_DESC_MODULO")
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO")
%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script>
function redefine()
{
window.location.href='rel_geral2.asp'
}
function manda1()
{
window.location.href='rel_geral2.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value
}
function manda2()
{
window.location.href='rel_geral2.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value
}
</script>

<body topmargin="0" leftmargin="0">
<form method="POST" action="gera_rel_geral_back_com_agrup.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:submit()"></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Montar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
            <td width="27"><img src="../imagens/desiste_F02.gif" width="24" height="24" onclick="javascript:redefine()"></td>  <td width="50"><font color="#330099" face="Verdana" size="2"><b>Redefinir</b></font></td>
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
  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Relatório
  Geral de Relações Definidas</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="88%" height="198">
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Mega-Processo (Opcional)</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="25"></td>
    <td width="73%" height="25"><select size="1" name="selMegaProcesso" onchange="javascript:manda1()">
        <option value="0">== Selecione o Mega-Processo ==</option>
       <%do until rs_mega.eof=true
       if trim(str_mega)=trim(rs_mega("MEPR_CD_MEGA_PROCESSO")) then
       %>
       <option selected value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
		<%ELSE%>
       <option value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
		<%
		end if
		rs_mega.movenext
		loop
		%>	        
       </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Processo (Opcional)</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="25"></td>
    <td width="73%" height="25"><select size="1" name="selProcesso" onchange="javascript:manda2()">
        <option value="0">== Selecione o Processo ==</option>
        <%do until rs_proc.eof=true
        if trim(str_proc)=trim(rs_proc("PROC_CD_PROCESSO")) then
        %>
        <option selected value=<%=rs_proc("PROC_CD_PROCESSO")%>><%=rs_proc("PROC_TX_DESC_PROCESSO")%></option>
        <%else%>
        <option value=<%=rs_proc("PROC_CD_PROCESSO")%>><%=rs_proc("PROC_TX_DESC_PROCESSO")%></option>
        <%
        end if
        rs_proc.movenext
        loop
        %>
        </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Sub-Processo (Opcional)</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="25"></td>
    <td width="73%" height="25"><select size="1" name="selSubProcesso">
        <option value="0">== Selecione o Sub-Processo ==</option>
       <%do until rs_sub.eof=true%>
       <option value=<%=rs_sub("SUPR_CD_SUB_PROCESSO")%>><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></option>
		<%
		rs_sub.movenext
		loop
		%>
      </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        o Agrupamento ( Master List R/3 ) (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><select size="1" name="selModulo">
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
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        a Atividade (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><select size="1" name="selAtividade">
          <option value="0">== Selecione a Atividade ==</option>
          <%DO UNTIL RS2.EOF=TRUE%>
                    <option value="<%=RS2("ATCA_CD_ATIVIDADE_CARGA")%>"><%=RS2("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%
          RS2.MOVENEXT
          LOOP
          %>
        </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        a Transação (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><select size="1" name="selTransacao">
          <option value="0">== Selecione a Transação ==</option>
          <%DO UNTIL RS3.EOF=TRUE%>
                    <option value="<%=RS3("TRAN_CD_TRANSACAO")%>"><%=RS3("TRAN_CD_TRANSACAO")%>-<%=LEFT(RS3("TRAN_TX_DESC_TRANSACAO"),50)%></option>
          <%
          RS3.MOVENEXT
          LOOP
          %>
        </select></td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>
=======
 
<!--#include file="../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=0
str_proc=0

str_mega=request("selMegaProcesso")
str_proc=request("selProcesso")

if session("MegaProcesso")<>0 and str_mega=0 then
	str_mega=session("MegaProcesso")
end if

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

if str_mega<>0 then
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
else
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega<>0 and str_proc<>0 then
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc)
else
	set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

if str_mega<>0 and str_proc=0 then
	set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega)
end if

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 ORDER BY MODU_TX_DESC_MODULO")
set rs2=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA ORDER BY ATCA_TX_DESC_ATIVIDADE")
set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO ORDER BY TRAN_CD_TRANSACAO")
%>

<html>

<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script>
function redefine()
{
window.location.href='rel_geral2.asp'
}
function manda1()
{
window.location.href='rel_geral2.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value
}
function manda2()
{
window.location.href='rel_geral2.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selProcesso='+document.frm1.selProcesso.value
}
</script>

<body topmargin="0" leftmargin="0">
<form method="POST" action="gera_rel_geral_back_com_agrup.asp" name="frm1">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
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
            <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"><img border="0" src="../imagens/confirma_f02.gif" onclick="javascript:submit()"></td>
          <td width="50"><font color="#330099" face="Verdana" size="2"><b>Montar</b></font></td>
          <td width="26">&nbsp;</td>
          <td width="195"></td>
            <td width="27"><img src="../imagens/desiste_F02.gif" width="24" height="24" onclick="javascript:redefine()"></td>  <td width="50"><font color="#330099" face="Verdana" size="2"><b>Redefinir</b></font></td>
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
  <p style="margin-top: 0; margin-bottom: 0" align="center"><font face="Verdana" color="#330099" size="3">Relatório
  Geral de Relações Definidas</font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
<table border="0" width="88%" height="198">
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Mega-Processo (Opcional)</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="25"></td>
    <td width="73%" height="25"><select size="1" name="selMegaProcesso" onchange="javascript:manda1()">
        <option value="0">== Selecione o Mega-Processo ==</option>
       <%do until rs_mega.eof=true
       if trim(str_mega)=trim(rs_mega("MEPR_CD_MEGA_PROCESSO")) then
       %>
       <option selected value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
		<%ELSE%>
       <option value=<%=RS_MEGA("MEPR_CD_MEGA_PROCESSO")%>><%=RS_MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
		<%
		end if
		rs_mega.movenext
		loop
		%>	        
       </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Processo (Opcional)</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="25"></td>
    <td width="73%" height="25"><select size="1" name="selProcesso" onchange="javascript:manda2()">
        <option value="0">== Selecione o Processo ==</option>
        <%do until rs_proc.eof=true
        if trim(str_proc)=trim(rs_proc("PROC_CD_PROCESSO")) then
        %>
        <option selected value=<%=rs_proc("PROC_CD_PROCESSO")%>><%=rs_proc("PROC_TX_DESC_PROCESSO")%></option>
        <%else%>
        <option value=<%=rs_proc("PROC_CD_PROCESSO")%>><%=rs_proc("PROC_TX_DESC_PROCESSO")%></option>
        <%
        end if
        rs_proc.movenext
        loop
        %>
        </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
      o Sub-Processo (Opcional)</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="25"></td>
    <td width="73%" height="25"><select size="1" name="selSubProcesso">
        <option value="0">== Selecione o Sub-Processo ==</option>
       <%do until rs_sub.eof=true%>
       <option value=<%=rs_sub("SUPR_CD_SUB_PROCESSO")%>><%=rs_sub("SUPR_TX_DESC_SUB_PROCESSO")%></option>
		<%
		rs_sub.movenext
		loop
		%>
      </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        o Agrupamento ( Master List R/3 ) (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><select size="1" name="selModulo">
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
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        a Atividade (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><select size="1" name="selAtividade">
          <option value="0">== Selecione a Atividade ==</option>
          <%DO UNTIL RS2.EOF=TRUE%>
                    <option value="<%=RS2("ATCA_CD_ATIVIDADE_CARGA")%>"><%=RS2("ATCA_TX_DESC_ATIVIDADE")%></option>
          <%
          RS2.MOVENEXT
          LOOP
          %>
        </select></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><b><font face="Verdana" size="2" color="#330099">Selecione
        a Transação (Opcional )</font></b></td>
  </tr>
  <tr>
    <td width="27%" height="21"></td>
    <td width="73%" height="21"><select size="1" name="selTransacao">
          <option value="0">== Selecione a Transação ==</option>
          <%DO UNTIL RS3.EOF=TRUE%>
                    <option value="<%=RS3("TRAN_CD_TRANSACAO")%>"><%=RS3("TRAN_CD_TRANSACAO")%>-<%=LEFT(RS3("TRAN_TX_DESC_TRANSACAO"),50)%></option>
          <%
          RS3.MOVENEXT
          LOOP
          %>
        </select></td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
</form>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>

</body>

</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
