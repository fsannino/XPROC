<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
str_MegaProcesso = Request("selMegaProcesso")
str_Processo = Request("selProcesso")

set db = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.recordset")

db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation = 3

ssql="SELECT DISTINCT SUPR_TX_IMPACTO AS IMPACTO FROM SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_MegaProcesso & " AND PROC_CD_PROCESSO=" & str_Processo

set verifica=db.execute(ssql)

qtos=verifica.recordcount

if qtos=1 then

val_atual = verifica("IMPACTO")

select case val_atual
	case 1
		valor1="Checked"
		valor2=""
		valor3=""
	case 2
		valor1=""
		valor2="Checked"
		valor3=""
	case 3
		valor1=""
		valor2=""
		valor3="Checked"
	case else
		valor1=""
		valor2=""
		valor3=""
end select
else
	valor1=""
	valor2=""
	valor3=""
end if

SSQL1="SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso

set rs10=db.execute(SSQL1)

valor_mega=rs10("MEPR_TX_DESC_MEGA_PROCESSO")

SSQL="SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO = " & str_MegaProcesso & " AND PROC_CD_PROCESSO = " & str_Processo

'RESPONSE.WRITE SSQL

set rs=db.execute(SSQL)
%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function ver_historico(mega, processo)
{
var a=mega;
var b=processo;
window.open("ver_historico.asp?mega=" + a + "&processo=" + b + "","_blank","width=600,height=260,history=0,scrollbars=1,titlebar=0,resizable=0")
}
</script>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000000" vlink="#000000" alink="#000000">
<form method="POST" name="frm1" action="valida_altera_processo.asp">
  <table width="993" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="175" height="66" colspan="2">&nbsp;</td>
      <td width="429" height="66" colspan="2">&nbsp;</td>
      <td width="383" valign="top" colspan="2" height="66"> 
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
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="146">&nbsp; </td>
      <td height="20" width="27"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0" onclick="javascript:submit()"></td>
      <td height="20" width="422"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Alterar</b></font></td>
      <td colspan="2" height="20" width="5">&nbsp;</td>
      <td height="20" width="383">&nbsp;</td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Alteração 
        de Processos&nbsp;</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="778" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="21"></td>
      <td width="181"></td>
      <td width="446" colspan="6"></td>
      <td width="104"></td>
    </tr>
    <tr> 
      <td width="21"></td>
      <td width="181"></td>
      <td width="446" colspan="6"></td>
      <td width="104"></td>
    </tr>
    <tr> 
      <td width="21"></td>
      <td width="181"></td>
      <td width="48" colspan="3"></td>
      <td width="414" colspan="3"></td>
      <td width="104"></td>
    </tr>
    <tr> 
      <td width="21"></td>
      <td width="181"></td>
      <td width="48" colspan="3"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">&nbsp;</font></td>
      <td width="414" colspan="3">&nbsp;</td>
      <td width="104"></td>
    </tr>
    <tr>
      <td width="21">&nbsp;</td>
      <td width="181">&nbsp;</td>
      <td width="446" colspan="6"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Mega-Processo 
        :&nbsp;</b><%=valor_mega%></font></td>
      <td width="104">&nbsp;</td>
    </tr>
    <tr> 
      <td width="21">&nbsp;</td>
      <td width="181">&nbsp;</td>
      <td width="446" colspan="6"></td>
      <td width="104">&nbsp;</td>
    </tr>
    <tr> 
      <td width="21">&nbsp;</td>
      <td width="181"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Descrição 
        do Processo</b></font></td>
      <td width="446" colspan="6"> 
        <input type="text" name="AlteraProcesso" size="59" value="<%=RS("PROC_TX_DESC_PROCESSO")%>">
      </td>
      <td width="104">
        <p style="margin-top: 0; margin-bottom: 0"><input type="text" name="AlteraSeq" size="9" value="<%=RS("PROC_NR_SEQUENCIA")%>"></p>
      </td>
    </tr>
    <tr> 
      <td width="21">&nbsp;</td>
      <td width="181">&nbsp;</td>
      <td width="446" colspan="6">&nbsp;</td>
      <td width="104">&nbsp;</td>
    </tr>
    <tr> 
      <td width="21"></td>
      <td width="181"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Impacto</b></font></td>
      <td width="20"> 
        <input type="hidden" name="mega" size="8" value="<%=str_MegaProcesso%>">
        <input type="hidden" name="proc" size="8" value="<%=str_Processo%>">
        
		<input name="selImpacto" type="radio" value="1" <%=valor1%> text="Alto">
      </td>
      <td width="31"> 
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Alto</b></font>
      </td>
      <td width="42"> 
        <p align="right">
        <input type="radio" value="2" name="selImpacto" text="Alto" <%=valor2%>>
      </td>
      <td width="31"> 
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Médio</b></font>
      </td>
      <td width="44"> 
        <p align="right">
        <input type="radio" value="3" name="selImpacto"  <%=valor3%> text="Alto">
      </td>
      <td width="296"> 
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Baixo</b></font>
      </td>
      <td width="104"><a href="#"><img src="../imagens/b04.gif" alt="Clique aqui para Visualizar os Sub-Processos" width="21" height="20" border="0" onClick="ver_historico(<%=str_MegaProcesso%>,<%=str_Processo%>)"></a></td>
    </tr>
    <tr> 
      <td width="21">&nbsp;</td>
      <td width="181">&nbsp;</td>
      <td width="446" colspan="6">&nbsp;</td>
      <td width="104">&nbsp;</td>
    </tr>
    <tr> 
      <td width="21">&nbsp;</td>
      <td width="181">&nbsp;</td>
      <td width="446" colspan="6">&nbsp;</td>
      <td width="104">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>