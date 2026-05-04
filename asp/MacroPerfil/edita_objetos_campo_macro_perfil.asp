<%
str_Opt = Request("txtOPT")
if request("selMacro_Perfil") <> 0 then
   str_cd_Macro_Perfil = request("selMacro_Perfil")
else
   str_cd_Macro_Perfil = 0
end if

if request("selDescMacro_Perfil") <> "0" then
   str_DescMacro_Perfil = request("selDescMacro_Perfil")
else
   str_DescMacro_Perfil = "não achou"
end if

if request("selTransacao") <> "0" then
   str_Transacao = request("selTransacao")
else
   str_Transacao = 0
end if

if request("selDescTransacao") <> "0" then
   str_DescTransacao = request("selDescTransacao")
else
   str_DescTransacao = 0
end if

if request("txtFuncao") <> "0" then
   str_Funcao = request("txtFuncao")
else
   str_Funcao = 0
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL = ""
str_SQL = str_SQL & " SELECT TROB_TX_OBJETO, TROB_TX_CAMPO, MPTO_TX_SIT_ALTERACAO_VALOR, "
str_SQL = str_SQL & " MPTO_TX_VALORES, TROB_TX_CRITICO, MPTO_TX_SIT_ALTERACAO_VALOR1 "
str_SQL = str_SQL & " FROM " & Session("PREFIXO") & "MAC_PER_TRAN_OBJETO "
str_SQL = str_SQL & " WHERE MCPR_NR_SEQ_MACRO_PERFIL = " &  str_cd_Macro_Perfil
str_SQL = str_SQL & " and TRAN_CD_TRANSACAO = '" & str_Transacao & "'"
str_SQL = str_SQL & " order by TROB_TX_OBJETO, TROB_TX_CAMPO "
set rs_Objeto=db.execute(str_SQL)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>
<script language="JavaScript">

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_changePropOO(objName,x,theProp,theValue) { //v3.0
  var obj = MM_findObj(objName);
  var obj2 = MM_findObj(theValue);
  //alert("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
  if (obj && obj2 && (theProp.indexOf("style.")==-1 || obj.style &&  obj2.style )) eval("obj."+theProp+"="+"obj."+theProp+"+"+ "'  '+"+"obj2."+theProp);
}

function Confirma()
{
document.frm1.submit();
}

</script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="grava_alteracao_objetos.asp" name="frm1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr> 
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../Funcao/voltar.gif"></a> 
              </div>
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
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <div align="center"><font face="Verdana" color="#330099" size="3">Edita 
          objetos de Macro Perfil</font></div>
      </td>
    </tr>
    <tr> 
      <td><%'=opt : str_Opt%></td>
    </tr>
  </table>
  <table width="93%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="26%"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Macro 
          Perfil :</font></div>
      </td>
      <td width="59%"> <font size="2"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
        <b> <font color="#330099"> 
        <input type="hidden" name="txtMacroPerfil" value="<%=str_cd_Macro_Perfil%>">
        <%=str_DescMacro_Perfil%></font></b></font></font></td>
      <td width="15%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="26%">&nbsp;</td>
      <td width="59%"><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
        <input type="hidden" name="txtFuncao" value="<%=str_Funcao%>">
        </font></b></font></font></td>
      <td width="15%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="26%"> 
        <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Transa&ccedil;&atilde;o 
          :</font></div>
      </td>
      <td width="59%"> <font size="2"> <font face="Verdana, Arial, Helvetica, sans-serif"> 
        <b> <font color="#330099"> 
        <input type="hidden" name="txtTransacao" value="<%=str_Transacao%>">
        <%=str_Transacao%> - <%=str_DescTransacao%></font></b></font></font></td>
      <td width="15%">&nbsp;</td>
    </tr>
  </table>
  <table width="75%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td width="25%">&nbsp;</td>
      <td colspan="4"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000"><b>os 
        valores poss&iacute;veis dever&atilde;o ser separados por &quot;/&quot; 
        ou &quot;-&quot;. Ex 10/20 ou AAB-CCV</b></font> </td>
    </tr>
    <tr> 
      <td width="25%" bgcolor="#0000FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Objeto</font></b></td>
      <td width="34%" bgcolor="#0000FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Campo</font></b></td>
      <td width="8%" bgcolor="#0000FF">&nbsp;</td>
      <td width="22%" bgcolor="#0000FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Valor</font></b></td>
      <td width="11%" bgcolor="#0000FF"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">Cr&iacute;tico</font></b></td>
    </tr>
    <% if not rs_Objeto.EOF then
	      int_sequencia = 0
	      Do While not rs_Objeto.EOF  
		     int_sequencia = int_sequencia + 1
		     IF ls_Cor_Linha="#FFFFFF" THEN
                ls_Cor_Linha="#F7F7F7"    'CINZA CLARO
             ELSE
        	    ls_Cor_Linha="#FFFFFF"    'BRANCO
             END IF
         %>
    <tr bgcolor="<%=ls_Cor_Linha%>"> 
      <td width="25%" height="25"><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"><%=rs_Objeto("TROB_TX_OBJETO")%> 
        <input type="hidden" name="txtObj<%=int_sequencia%>" value="<%=rs_Objeto("TROB_TX_OBJETO")%>">
        </font></b></font></font></td>
      <td width="34%" height="25"><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"><%=rs_Objeto("TROB_TX_CAMPO")%> </font><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
        <input type="hidden" name="txtCampo<%=int_sequencia%>" value="<%=rs_Objeto("TROB_TX_CAMPO")%>">
        </font></b></font></font></b></font></font></td>
      <td width="8%" height="25"> 
        <div align="left"><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
          <input type="hidden" name="txtValorPadrao<%=int_sequencia%>" value="<%=rs_Objeto("MPTO_TX_VALORES")%>">
          </font><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
          </font></b></font></font></b></font></font><font color="#330099"> 
          <% If rs_Objeto("MPTO_TX_SIT_ALTERACAO_VALOR1") = "1" then
		      ls_Imagem1 = "aprova_03.gif"
		   elseIf rs_Objeto("MPTO_TX_SIT_ALTERACAO_VALOR1") = "0" then
		      ls_Imagem1 = "func_tran_nao_marcada.gif"
           end if %>
          <img src="../../imagens/<%=ls_Imagem1%>" width="14" height="12" alt="Indica que este valor foi alterado pelo Validador">	
          <% If rs_Objeto("MPTO_TX_SIT_ALTERACAO_VALOR") = "1" then
		      ls_Imagem2 = "aprova_02.gif"
		   elseIf rs_Objeto("MPTO_TX_SIT_ALTERACAO_VALOR") = "0" then
		      ls_Imagem2 = "func_tran_nao_marcada.gif"
           end if %>
          <img src="../../imagens/<%=ls_Imagem2%>" width="14" height="12" alt="Indica que este valor foi alterado pelo Elaborador">	

          </font></b></font></font></b></font></font></div>
      </td>
      <td width="22%" height="25"> 
        <input type="text" name="txtValor<%=int_sequencia%>" value="<%=rs_Objeto("MPTO_TX_VALORES")%>">
        <font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"> 
        </font></b></font></font></b></font></font> </td>
      <td width="11%" height="25"> 
        <div align="center"><font size="2"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font color="#330099"><%=rs_Objeto("TROB_TX_CRITICO")%></font></b></font></font></div>
      </td>
    </tr>
    <%       rs_Objeto.movenext
	      Loop
	else
	%>
    <tr> 
      <td width="25%" height="18">&nbsp; </td>
      <td width="34%" height="18">&nbsp;</td>
      <td width="8%" height="18">&nbsp;</td>
      <td width="22%" height="18">&nbsp;</td>
      <td width="11%" height="18">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="5"> 
        <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FF0000">N&atilde;o 
          encontrado objetos para esta transa&ccedil;&atilde;o.</font></div>
      </td>
    </tr>
    <% end if %>
    <tr> 
      <td width="25%"> 
        <input type="hidden" name="txtQtdObj" value="<%=int_sequencia%>">
        <input type="hidden" name="txtOPT" value="<%=str_Opt%>">
      </td>
      <td width="34%">&nbsp;</td>
      <td width="8%">&nbsp;</td>
      <td width="22%">&nbsp;</td>
      <td width="11%">&nbsp;</td>
    </tr>
  </table>
</form>
</body>

</html>
<%


%>