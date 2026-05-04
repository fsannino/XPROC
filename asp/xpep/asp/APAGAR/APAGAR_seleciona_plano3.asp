<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%

set db_Cronograma = Server.CreateObject("ADODB.Connection")
db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html><!-- InstanceBegin template="/Templates/BASICO_XPEP_03.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Untitled Document</title>
<!-- InstanceEndEditable -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- InstanceBeginEditable name="head" -->
<!-- InstanceEndEditable -->
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
a {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333; text-decoration: none}
a:hover {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; font-weight: normal; color: #333333;  text-decoration: underline}
-->
</style>
<link href="/css/biblioteca.css" rel="stylesheet" type="text/css">
<link href="../../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBeginEditable name="Head01" -->

<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaScri" -->
<script language="JavaScript">

function chamapagina()
{
	//alert("entrei");
	//window.location.href='seleciona_plano.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+"&chkEmUso="+document.frm1.chkEmUso.checked+"&chkEmDesuso="+document.frm1.chkEmDesuso.checked
	document.frm1.action="seleciona_plano.asp";
	//document.frm1.target="corpo";
	document.frm1.submit();
}

function Confirma()
{
   if(document.frm1.selOnda.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de uma Onda!");
      document.frm1.selOnda.focus();
      return;
      }
   if(document.frm1.selPlano.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um Plano!");
      document.frm1.selPlano.focus();
      return;
      }

   if(document.frm1.selTask1.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de uma Atividade!");
      document.frm1.selTask1.focus();
      return;
      }

           document.frm1.action="encaminha_plano.asp";
           //document.frm1.target="corpo";
           document.frm1.submit();

}
function Confirma2()
{
   if(document.frm1.selMegaProcesso.selectedIndex == 0)
      {
      alert("É obrigatória a seleção de um MEGA-PROCESSO!");
      document.frm1.selMegaProcesso.focus();
      return;
      }
    if((document.frm1.txtOPT.value != 6)&&(document.frm1.txtOPT.value != 7)&&(document.frm1.txtOPT.value != 8))
	  { 
      if(document.frm1.selFuncao.selectedIndex == 0)
        {
        alert("É obrigatória a seleção de uma FUNÇÃO DE NEGÓCIO!");
        document.frm1.selFuncao.focus();
        return;
        }
      else
        {
	     //alert(document.frm1.txtOPT.value);
         if(document.frm1.txtOPT.value == 1)
           {
           document.frm1.action="alterar_funcao.asp";
           //document.frm1.target="corpo";
           document.frm1.submit();
           }
         if(document.frm1.txtOPT.value == 2)
           {
           document.frm1.action="valida_exclui_funcao.asp";
           //document.frm1.target="corpo";
           document.frm1.submit();
           }
          if(document.frm1.txtOPT.value == 3)
            {
            document.frm1.action="cad_funcao_transacao2.asp";
            //document.frm1.target="corpo";
            document.frm1.submit();
            }
         if(document.frm1.txtOPT.value == 4)
           {
           document.frm1.action="cad_funcao_transacao2_outro.asp";
           //document.frm1.target="corpo";
           document.frm1.submit();
           }		
         if(document.frm1.txtOPT.value == 5)
           {
           document.frm1.action="gera_rel_mega_funcao.asp";
           //document.frm1.target="corpo";
           document.frm1.submit();
           }		
		}   
	 }  
   else
      {
	  //alert(document.frm1.txtOPT.value);
      if(document.frm1.txtOPT.value == 6)
        {
        document.frm1.action="altera_funcao_assunto_em_massa.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 7)
        {
        document.frm1.action="rel_funcao_sem_Assunto.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
      if(document.frm1.txtOPT.value == 8)
        {
        document.frm1.action="gera_rel_func_confl.asp";
        //document.frm1.target="corpo";
        document.frm1.submit();
        }		
     }
}

	function mOvr(src,clrOver) {
		if (!src.contains(event.fromElement)) {
			src.style.cursor = 'hand';
			src.bgColor = clrOver;
		}
	}
	function mOut(src,clrIn) {
		if (!src.contains(event.toElement)) {
			src.style.cursor = 'default';
			src.bgColor = clrIn;
		}
	}
	function mClk(src) {
		if(event.srcElement.tagName=='TD'){
			src.children.tags('A')[0].click();
		}
	}

</script>
<!-- InstanceEndEditable -->
<%=ls_Script%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<div id="Layer1" style="position:absolute; left:20px; top:10px; width:134px; height:53px; z-index:1"><img src="../../img/000005.gif" alt=":: Logo Sinergia" width="134" height="53" border="0" usemap="#Map2"> 
	  <map name="Map2">
	    <area shape="rect" coords="6,7,129,49">
	  </map>
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td><table width="780" height="44" border="0" cellpadding="0" cellspacing="0">
	        <tr>
	          <td width="583" height="44"><img src="../../img/_0.gif" width="1" height="1"></td>
	          <td width="197" height="44"><img src="../../../../imagens/000043.gif" width="95" height="44"></td>
	        </tr>
	      </table></td>
	  </tr>
</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td bgcolor="#6699CC">
			<table width="780" border="0" cellspacing="0" cellpadding="0">
			  <tr>
			    <td width="154" height="21"><img src="../../img/000002.gif" width="154" height="21"></td>
			    <td width="19" height="21"><img src="../../img/000003.gif" width="19" height="21"></td>
			    <td width="202" height="21">
					<font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
						<strong>
						</strong>
					</font>
			    </td>
			    <td>&nbsp;</td>
		      </tr>
			</table>
	    </td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="1" height="1" bgcolor="#003366"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td height="5"><img src="../../img/_0.gif" width="1" height="1"></td>
	  </tr>
	</table>
	<table width="780" height="58" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
        <td width="740" height="39" background="../../img/000006.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="1">
            <tr>
              <td width="11%">&nbsp;</td>
              <td width="13%">&nbsp;</td>
              <td width="61%"><font color="#666666" size="3" face="Verdana, Arial, Helvetica, sans-serif"><b>PLANO DE ENTRADA EM PRODU&Ccedil;&Atilde;O</b></font></td>
              <td width="15%"><a href="../../../../indexA_xpep.asp"><img src="../../img/botao_home_off_01.gif" alt="Ir para tela inicial" width="34" height="23" border="0"></a></td>
            </tr>
        </table></td>
        <td width="20" height="39"><img src="../../img/_0.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><div align="right"><span class="style8">
        <%=Session("CdUsuario") & "-" &  Session("NomeUsuario")%></span></div></td>
        <td>&nbsp;</td>
      </tr>
    </table>
	<!-- InstanceBeginEditable name="corpo" --><table width="98%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="10%">&nbsp;</td>
    <td width="77%">&nbsp;</td>
    <td width="13%">&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td class="subtitulo"><strong>Sele&ccedil;&atilde;o para Detalhamento das 
      Atividades</strong></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
<form action="" method="post" name="frm1" id="frm1">
  <table width="98%" border="0" cellspacing="10">
    <tr> 
      <td width="13%"> <div align="right" class="campob">Onda:</div></td><td width="85%">
      <table width="100%" border="0" cellpadding="0" cellspacing="0" class="cmb200">
        <tr> 
          <td><!--#include file="../includes/inc_Combo_Onda.asp"--></td>
          <td class="cmb200"> <div align="right"></div></td>
          <td></td></td>
        <td>&nbsp;</td>
        </tr>
      </table></td>
      <td width="2%">&nbsp;</td>
    </tr>
    <%if Request("selOnda") = "5" then%>
    <tr>
      <td class="campob"><div align="right">Fase:</div></td>
      <td><!--#include file="../includes/inc_combo_fases.asp" --></td>
      <td>&nbsp;</td>
    </tr>
    <%end if%>	
    <tr> 
      <td class="campob"><div align="right" class="campob">Plano:</div></td>
      <td> <!--#include file="../includes/inc_Combo_Plano.asp" --> </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="21" valign="top" class="campo"> <div align="right" class="campob">Atividades:</div></td>
      <td><!--#include file="../includes/inc_combo_tarefas_nivel1.asp" --></td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<%
  db_Cronograma.Close
  db_Cogest.close
  set db_Cronograma = Nothing
  set db_Cogest = Nothing
  %>
<!-- InstanceEndEditable -->
    <table width="200" border="0" align="center">
<tr>	
	<td height="10" width="780"></td>
</tr>
<tr>
	<td width="780">			
		<p width="780" align="center"><img src="../../../../img/000025.gif" width="467" height="1"></p>
		<p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
	</td>
</tr></table>
</body>
<!-- InstanceEnd --></html>
