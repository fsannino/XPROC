 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

curso=request("curso")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & curso & "'")
str_onda = rs("ONDA_CD_ONDA")
set rs_onda=db.execute("SELECT * FROM " & Session("PREFIXO") & "ONDA WHERE ONDA_CD_ONDA<>4 ORDER BY ONDA_TX_DESC_ONDA")

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="JavaScript">


</script>

<script language="javascript" src="../Planilhas/js/troca_lista.js"></script>

<script>
function Confirma()
{
if(document.frm1.selMegaProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MEGA-PROCESSO!");
document.frm1.selMegaProcesso.focus();
return;
}

if(document.frm1.txtnomecurso.value == "")
{
alert("É obrigatória a definição do NOME DO CURSO!");
document.frm1.txtnomecurso.focus();
return;
}

if(document.frm1.txtcargacurso.value == "")
{
alert("É obrigatória a CARGA HORÁRIA DO CURSO!");
document.frm1.txtcargacurso.focus();
return;
}
if(document.frm1.selMetodo.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MÉTODO!");
document.frm1.selMetodo.focus();
return;
}

else
{
document.frm1.submit();
}
}

function ver_conteudo(fbox)
{
valor=fbox.value;
tamanho=valor.length;
str1=valor.slice(tamanho-1,tamanho);
if (str1!=0 && str1!=1 && str1!=2 && str1!=3 && str1!=4 && str1!=5 && str1!=6 && str1!=7 && str1!=8 && str1!=9){
	fbox.value="";
	str2=valor.slice(0,tamanho-1)
	fbox.value=str2;
}
}

</script>
<script language="javascript" src="../js/troca_toda_lista.js"></script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_altera_curso.asp" name="frm1">
        <input type="hidden" name="txtImp" size="20"><input type="hidden" name="txtQua" size="20">
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
        <div align="center"><font face="Verdana" color="#330099" size="3">Alteração
          de Cursos - <b><%=curso%></b></font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="893" height="87">
    <tr> 
      <td width="18" height="29"> <input type="hidden" name="selMegaProcesso" size="20" value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"> 
      </td>
      <td width="137" height="29" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo 
        :</b></font></td>
      <td width="674" height="29" valign="middle" align="left"> <%
      set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
      %> <font face="Verdana" size="2" color="#330099"><%=rs("MEPR_CD_MEGA_PROCESSO")%> - <%=rstemp("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
    </tr>
    <tr> 
      <td width="18" height="1"><input type="hidden" name="txtCurso" size="20" value="<%=curso%>"></td>
      <td width="137" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Nome 
        do Curso :</b></font></td>
      <td height="1" valign="middle" align="left" width="674"> <input type="text" name="txtnomecurso" size="58" maxlength="100" value="<%=RS("CURS_TX_NOME_CURSO")%>"></td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
    <tr> 
      <td height="1"></td>
      <td height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Onda 
        :</b></font></td>
      <td height="1" valign="middle" align="left"><select size="1" name="selOnda">
          <option value="0">== Selecione a Onda ==</option>
          <%DO UNTIL RS_ONDA.EOF=TRUE
      IF TRIM(str_onda)=trim(rs_onda("ONDA_CD_ONDA")) then
      %>
          <option selected value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%else%>
          <option value=<%=rs_onda("ONDA_CD_ONDA")%>><%=rs_onda("ONDA_TX_ABREV_ONDA")%> - <%=rs_onda("ONDA_TX_DESC_ONDA")%></option>
          <%
		END IF
		RS_ONDA.MOVENEXT
		LOOP
		%>
        </select></td>
    </tr>
    <tr>
      <td height="1"></td>
      <td height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Abrang&ecirc;ncia 
        :</b></font></td>
      <td height="1" valign="middle" align="left"><table border="0" width="351">
          <tr> 
            <td width="153" height="138" rowspan="4" valign="top"> <table width="138" border="0" align="right">
                <tr> 
                  <td width="132"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    n&atilde;o selecionadas</strong></font></td>
                </tr>
                <tr> 
                  <td><b> 
                    <select name="list1" size="5" multiple>
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
              </table></td>
            <td width="24" align="center"><a href="#" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="../Funcao/continua_F01.gif" width="24" height="24"></a></td>
            <td width="160" height="138" rowspan="4" valign="top"> <table width="138" border="0">
                <tr> 
                  <td width="95"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    selecionadas</strong></font></td>
                </tr>
                <tr> 
                  <td><font color="#000080"> 
                    <select name="list2" size="5" multiple>
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
              </table></td>
          </tr>
          <tr> 
            <td width="24" align="center"><a href="#" onClick="movetudo(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="../../imagens/seta_dupla_direita.gif" width="24" height="24"></a></td>
          </tr>
          <tr> 
            <td width="24" align="center"><a href="javascript:;"  onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="../Funcao/continua2_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr> 
            <td width="24" align="center"><a href="javascript:;"  onClick="movetudo(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="../../imagens/seta_dupla_esquerda.gif" width="24" height="24"></a></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Carga 
        Horária (h):</b></font></td>
      <td height="1" valign="middle" align="left" width="674"> <input type="text" name="txtcargacurso" size="14" onkeyup="javascript:ver_conteudo(txtcargacurso)" value="<%=RS("CURS_NUM_CARGA_CURSO")%>"></td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Método 
        :</b></font></td>
      <td height="1" valign="middle" align="left" width="674"> <select size="1" name="selMetodo">
          <%
          SELECT CASE RS("CURS_TX_METODO_CURSO")
          CASE "À DISTÂNCIA"
          %>
          <option value="0">== Selecione o Método ==</option>
          <option selected value="À DISTÂNCIA">À DISTÂNCIA</option>
          <option value="Presencial">PRESENCIAL</option>
          <%
          CASE "PRESENCIAL"
          %>
          <option value="0">== Selecione o Método ==</option>
          <option value="À DISTÂNCIA">À DISTÂNCIA</option>
          <option selected value="Presencial">PRESENCIAL</option>
          <%case else%>
          <option value="0">== Selecione o Método ==</option>
          <option value="À DISTÂNCIA">À DISTÂNCIA</option>
          <option value="Presencial">PRESENCIAL</option>
          <%
          end select
          %>
        </select></td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Público 
        Alvo : </b></font></td>
      <td height="1" valign="middle" align="left" width="674"> <textarea rows="4" name="txtPublicoAlvo" cols="50"><%=rs("CURS_TX_PUBLICO_ALVO")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="top" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Requisitos 
        n&atilde;o R/3:</b></font></td>
      <td height="1" valign="middle" align="left" width="674"> <textarea rows="4" name="txtPreRequisitos" cols="50"><%=rs("CURS_TX_PRE_REQUISITOS")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="top" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Conteúdo 
        Programático :</b></font></td>
      <td height="1" valign="middle" align="left" width="674"> <textarea rows="4" name="txtConteudo" cols="50"><%=rs("CURS_TX_CONTEUDO_PROGRAM")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="top" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="top" align="left"><font face="Verdana" size="2" color="#330099"><b>Objetivo 
        :</b></font></td>
      <td height="1" valign="middle" align="left" width="674"> <textarea rows="4" name="txtObjetivo" cols="50"><%=rs("CURS_TX_OBJETIVO")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td width="18" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"></td>
      <td height="1" valign="middle" align="left" width="674"> </td>
    </tr>
  </table>
  </form>

</body>

</html>
