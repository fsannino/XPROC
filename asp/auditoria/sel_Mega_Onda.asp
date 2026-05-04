<%
'RESPONSE.Write(Session("Conn_String_Cogest_Gravacao"))
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_mega=request("MEGA")
str_onda=request("selOnda")
str_OPT = request("pOpt") 
If str_OPT = 1 then
	str_Titulo = "Seleção de Mega para relatório de Cursos sem Função associada"
elseif str_OPT = 2 then
	str_Titulo = "Seleção de Mega para relatório de Cursos sem Transação associada"
elseif str_OPT = 3 then
	str_Titulo = "Seleção de Mega para relatório de Cursos com transações associadas a Função e não associadas a Curso"
end if

if str_mega > 0 then
	compl=" and  " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO=" + str_mega
else
	compl=""
end if
if str_onda >0 then
	compl2=" and  " & Session("PREFIXO") & "CURSO.ONDA_CD_ONDA = " + str_onda
else
	compl2=""
end if

'SSQL="SELECT * FROM " & Session("PREFIXO") & "CURSO ORDER BY MEPR_CD_MEGA_PROCESSO, CURS_CD_CURSO"
'SSQL = ""
'SSQL= SSQL & " SELECT * "
'SSQL= SSQL & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO AS MEGA "
'SSQL= SSQL & "," & Session("PREFIXO") & "CURSO.* FROM " & Session("PREFIXO") & "CURSO INNER JOIN " & Session("PREFIXO") & "MEGA_PROCESSO ON " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO where MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO > 0 " & COMPL & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "CURSO.CURS_CD_CURSO"

SSQL1="SELECT " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO AS MEGA, " & Session("PREFIXO") & "CURSO.* FROM " & Session("PREFIXO") & "CURSO INNER JOIN " & Session("PREFIXO") & "MEGA_PROCESSO ON " & Session("PREFIXO") & "CURSO.MEPR_CD_MEGA_PROCESSO = " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO where MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO > 0 " & COMPL & COMPL2 & " ORDER BY " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO, " & Session("PREFIXO") & "CURSO.CURS_CD_CURSO"
'RESPONSE.Write(SSQL1)
SET RS=DB.EXECUTE(SSQL1)

SET MEGA=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
set rs_onda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ABRANGENCIA_CURSO WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_TX_DESC_ONDA")

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<style>
a {text-decoration:none;}
a:hover {text-decoration:underline;}
</style>

</head>
<SCRIPT>
function envia()
{
this.location.href='sel_Mega_Onda.asp?mega='+document.frm1.selMegaProcesso.value+'&selOnda='+document.frm1.selOnda.value+'&pOpt='+document.frm1.txtOPT.value
}

function Confirma()
{
	//alert(document.frm1.txtOPT.value)
	 if(document.frm1.txtOPT.value == 1)
	   {
	   document.frm1.action="rel_curso_sem_funcao.asp";
	   //document.frm1.target="corpo";
	   document.frm1.submit();
	   }
	 if(document.frm1.txtOPT.value == 2)
	   {
	   document.frm1.action="rel_curso_sem_transacao.asp";
	   //document.frm1.target="corpo";
	   document.frm1.submit();
	   }
	 if(document.frm1.txtOPT.value == 3)
	   {
	   document.frm1.action="rel_curso_com_transacao_sobrando2.asp";
	   //document.frm1.target="corpo";
	   document.frm1.submit();
	   }

}


</SCRIPT>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF" link="#800000" vlink="#800000" alink="#800000">
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
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
          <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table></td>
  </tr>
</table>
        
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>
        <div align="center">
          <p align="center" style="margin-top: 0; margin-bottom: 0"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Titulo%></font>                  </div>
      </td>
    </tr>
  </table>
  <p><b><font face="Verdana" color="#330099" size="2"> </font></b></p>
  <table width="75%" border="0" cellpadding="5" cellspacing="0">
    <tr>
      <td width="37%"><div align="right"><b><font face="Verdana" color="#330099" size="2">Mega-Processo :</font></b></div></td>
      <td width="61%"><b><font face="Verdana" color="#330099" size="2">
        <select size="1" name="selMegaProcesso" onChange="javascript:envia()">
          <option value="0">== TODOS ==</option>
          <%DO UNTIL MEGA.EOF=TRUE
  if trim(str_mega)=trim(MEGA("MEPR_CD_MEGA_PROCESSO")) then
  %>
          <option selected value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%else%>
          <option value="<%=MEGA("MEPR_CD_MEGA_PROCESSO")%>"><%=MEGA("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
  end if
  MEGA.MOVENEXT
  LOOP
  %>
        </select>
        </font></b></td>
      <td width="2%">&nbsp;</td>
    </tr>
    <tr>
      <td><div align="right"><font face="Verdana" size="2" color="#330099"><b>Onda :</b></font></div></td>
      <td><select size="1" name="selOnda"  onChange="javascript:envia()">
          <option value="0">== Selecione a Onda ==</option>
          <%DO UNTIL RS_ONDA.EOF=TRUE
      IF TRIM(str_onda)=trim(rs_onda("ONDA_CD_ONDA")) then
      %>
          <option selected value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> 
          - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%else%>
          <option value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> 
          - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%
		END IF
		RS_ONDA.MOVENEXT
		LOOP
		%>
        </select></td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input type="hidden" name="txtOPT" value="<%=str_OPT%>"></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <b></b>
</form>

</body>

</html>
