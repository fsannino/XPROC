 
<!--#include file="../../asp/protege/protege.asp" -->
<%
str_cenario=request("ID")

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

SET RS=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CENARIO WHERE CENA_CD_CENARIO='" & str_cenario & "' AND CENA_TX_SITUACAO='DS'")

if rs.eof=true then
	tem=0
else
	tem=1
end if

ON ERROR RESUME NEXT
valor_tipo=rs("CENA_TX_SITU_DESENHO_TIPO")
valor_desenv=rs("CENA_TX_SITU_DESENHO_DESE")
valor_conf=rs("CENA_TX_SITU_DESENHO_CONF")
valor_teste=rs("CENA_TX_SITU_DESENHO_TESTE")

IF ERR.NUMBER<>0 THEN
	valor_tipo=0
	valor_desenv=0
	valor_conf=0
	valor_teste=0
END IF

%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function Confirma()
{
  if(document.frm1.txtTipo.value==0)
    {
    alert('Você deve especificar se o Cenário possui ou não DESENVOLVIMENTO')
    return;
    }
  else
    {
    document.frm1.submit();
    }
}

function verifica_status()
{
   if(document.frm1.txtTemStatus.value==0)
     {
	 //alert('não tem!');
	 document.frm1.Desenv1.enabled=false;
     }
}

function preenche_txt1()
{
document.frm1.txtTipo.value = 1;
document.frm1.Desenv1.checked=false;
document.frm1.Conf1.checked=false;
document.frm1.Conf2.checked=false;
document.frm1.Teste1.checked=false;
document.frm1.Teste2.checked=false;
document.frm1.txtDesenv.value=0;
document.frm1.txtConf.value=0
document.frm1.txtTeste.value=0
}

function preenche_txt2()
{
document.frm1.txtTipo.value = 2;
document.frm1.Desenv1.checked=false;
document.frm1.Conf1.checked=false;
document.frm1.Conf2.checked=false;
document.frm1.Teste1.checked=false;
document.frm1.Teste2.checked=false;
document.frm1.txtDesenv.value=0;
document.frm1.txtConf.value=0
document.frm1.txtTeste.value=0
}

function carrega_desenv()
{
   if(document.frm1.Desenv1.checked==true)
     {
     document.frm1.txtDesenv.value=document.frm1.Desenv1.value;
     }
   else
     {
     document.frm1.txtDesenv.value=0;
     }
}

function carrega_conf()
{
if(document.frm1.Conf1.checked==true)
{
document.frm1.txtConf.value=document.frm1.Conf1.value;
}
else
{
if(document.frm1.Conf2.checked==true)
{
document.frm1.txtConf.value=document.frm1.Conf2.value;
}
else
{
document.frm1.txtConf.value=0
}
}
}

function carrega_teste()
{
   if(document.frm1.Teste1.checked==true)
     {
     document.frm1.txtTeste.value=document.frm1.Teste1.value;
     }
   else
     {
     if(document.frm1.Teste2.checked==true)
       {
       document.frm1.txtTeste.value=document.frm1.Teste2.value;
       }
     else
       {
       document.frm1.txtTeste.value=0
       }
   }
}
</script>

<SCRIPT LANGUAGE="JavaScript">
function addbookmark()
{
bookmarkurl="http://S6000WS10.corp.petrobras.biz/xproc/index.htm"
bookmarktitle="Sinergia - Cadastro"
if (document.all)
window.external.AddFavorite(bookmarkurl,bookmarktitle)
}
//  End -->
</script>


<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="frm1" method="post" action="valida_altera_status.asp">
  <input type="hidden" name="INC" size="20" value="1"><input type="hidden" name="txtOpc" value="1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
    <tr>
      <td width="20%" height="20">&nbsp;</td>
      <td width="44%" height="60">&nbsp;</td>
      <td width="36%" valign="top"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif" width="30" height="30"></a>
              </div>
            </td>
            <td bgcolor="#330099" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" width="27" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif" width="30" height="30"></a></div>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
              <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
              <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif" width="30" height="30"></a></div>
            </td>
            <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
              <div align="center"><a href="../../indexA.asp"><img src="home.gif" border="0" width="19" height="20"></a>&nbsp;</div>
            </td>
          </tr>
        </table>
      </td>
  </tr>
  <tr bgcolor="#00FF99"> 
    <td colspan="3" height="20"> 
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
          <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Envia</font></b></font></td>
          <td width="26"></td>
          <td width="195"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28">&nbsp;</td>
          <td width="26">&nbsp;</td>
          <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <table border="0" width="100%">
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3"> 
        <p align="center"><font face="Verdana" color="#330099" size="3">Alteração de Status de Cenário - Questionário -
        <b><%=str_cenario%></b></font></p>
      </td>
    </tr>
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="16%"> 
        <p align="right"><font face="Verdana" size="2" color="#330099"><b>&nbsp;
        <%'OPTION BOX DE TIPO
        IF VALOR_TIPO=1 THEN
        %>
        <input type="radio" value="1" name="Tipo1" onclick="javascript:preenche_txt1()" checked>
        <%ELSE
        %>
        <input type="radio" value="1" name="Tipo1" onclick="javascript:preenche_txt1()">
        <%END IF%>
        </b></font></p>
      </td>
      <td colspan="2"><font face="Verdana" size="2" color="#330099"><b>Com Desenvolvimento</b></font></td>
    </tr>
    <tr>
      <td width="16%"><font face="Verdana" size="2" color="#330099"><b>
<input type="HIDDEN" name="txtCenario" size="20" value="<%=str_cenario%>"></b></font></td>
      <td width="13%"></td>
      <td width="37%"></td>
    </tr>
    <tr>
      <td width="16%"><font face="Verdana" size="2" color="#330099"><b>
<input type="HIDDEN" name="txtTipo" size="20" value="<%=VALOR_TIPO%>"></b></font></td>
      <td width="13%"> 
        <p align="right">
        <%'CHECKBOX DE CONFERENCIA
        IF VALOR_TIPO=1 AND VALOR_CONF=1 THEN
        %>
        <input type="checkbox" name="Conf1" value="1" onclick="javascript:carrega_conf()" checked>
        <%ELSE%>
        <input type="checkbox" name="Conf1" value="1" onclick="javascript:carrega_conf()">
        <%END IF%>
        </td>
      <td width="37%"><font face="Verdana" size="2" color="#330099"><b>Configuração 
        Concluída</b></font></td>
    </tr>
    <tr>
      <td width="16%"><font face="Verdana" size="2" color="#330099"><b>
<input type="HIDDEN" name="txtDesenv" size="20" value="<%=VALOR_DESENV%>"></b></font></td>
      <td width="13%" align="right"> 
        <%'CHECKBOX DE DESENVOLVIMENTO
      IF VALOR_TIPO=1 AND VALOR_DESENV=1 THEN
      %>
      <input type="checkbox" name="Desenv1" value="1" onclick="javascript:carrega_desenv()" checked>
      <%ELSE%>
      <input type="checkbox" name="Desenv1" value="1" onclick="javascript:carrega_desenv()">
      <%END IF%>
      </td>
      <td width="37%"><font face="Verdana" size="2" color="#330099"><b>Desenvolvimento 
        Concluído</b></font></td>
    </tr>
    <tr>
      <td width="16%"><font face="Verdana" size="2" color="#330099"><b>
<input type="HIDDEN" name="txtConf" size="20" value="<%=VALOR_CONF%>"></b></font></td>
      <td width="13%" align="right"> 
        <%
      IF VALOR_TIPO=1 AND VALOR_TESTE=1 THEN
      %>
      <input type="checkbox" name="Teste1" value="1" onclick="javascript:carrega_teste()" checked>
      <%ELSE%>
      <input type="checkbox" name="Teste1" value="1" onclick="javascript:carrega_teste()">
      <%END IF%>
      </td>
      <td width="37%"><font face="Verdana" size="2" color="#330099"><b>Dados de 
        Teste Carregados (PED 220)</b></font></td>
    </tr>
    <tr>
      <td width="16%"> </td>
      <td colspan="2" align="right"> </td>
    </tr>
    <tr>
      <td width="16%"> </td>
      <td colspan="2" align="right"> </td>
    </tr>
    <tr>
      <td width="16%"> 
        <p align="right"><font face="Verdana" size="2" color="#330099"><b>&nbsp;
        <%'OPTION BOX DE TIPO
        IF VALOR_TIPO=2 THEN
        %>
        <input type="radio" value="2" name="Tipo1" onclick="javascript:preenche_txt2()" checked>
        <%else%>
        <input type="radio" value="2" name="Tipo1" onclick="javascript:preenche_txt2()">
        <%end if%>
        </b></font></p>
      </td>
      <td colspan="2" align="right"> 
        <p align="left"><font face="Verdana" size="2" color="#330099"><b>Sem
        Desenvolvimento</b></font></td>
    </tr>
    
    <tr>
      <td width="16%">
<input type="HIDDEN" name="txtmega" size="5" value="<%=request("selMegaProcesso")%>"><input type="HIDDEN" name="txtproc" size="5" value="<%=request("selProcesso")%>"><input type="HIDDEN" name="txtsub" size="5" value="<%=request("selSubProcesso")%>"></td>
      <td width="13%" align="right"> 
        <%'CHECKBOX DE CONFERENCIA
      IF VALOR_TIPO=2 AND VALOR_CONF=1 THEN
      %>
      <input type="checkbox" name="Conf2" value="1" onclick="javascript:carrega_conf()" checked></td>
      <%ELSE%>
      <input type="checkbox" name="Conf2" value="1" onclick="javascript:carrega_conf()">
      <%END IF%>
      <td width="34%"><font face="Verdana" size="2" color="#330099">&nbsp;<b>Configuração 
        Concluída</b></font></td>
    </tr>
    <tr>
      <td width="16%"><font face="Verdana" size="2" color="#330099"><b>&nbsp;
<input type="hidden" name="txtTemStatus" size="20" value="<%=TEM%>"><input type="HIDDEN" name="txtTeste" size="20" value=<%=valor_teste%>></b></font></td>
      <td width="13%"> 
        <p align="right">
      <%
      IF VALOR_TIPO=2 AND VALOR_TESTE=1 THEN
      %>
      <input type="checkbox" name="Teste2" value="1" onclick="javascript:carrega_teste()" checked>
      <%ELSE%>
      <input type="checkbox" name="Teste2" value="1" onclick="javascript:carrega_teste()">
      <%END IF%>
      </p>
      </td>
      <td width="37%"><font face="Verdana" size="2" color="#330099"><b>Dados de 
        Teste Carregados</b></font></td>
    </tr>
  </table>
  </form>
<p></p>
</body>
</html>


