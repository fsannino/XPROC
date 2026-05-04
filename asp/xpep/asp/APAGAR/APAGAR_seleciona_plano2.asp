<%
Session("Conn_String_Cogest_Gravacao")= "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest"
Session("Conn_String_Cronograma_Gravacao")= "Provider=SQLOLEDB.1;server=10.22.22.13;pwd=sinergia;uid=sinergia;database=sinergia"

set db_Cronograma = Server.CreateObject("ADODB.Connection")
db_Cronograma.Open Session("Conn_String_Cronograma_Gravacao")

set db_Cogest = Server.CreateObject("ADODB.Connection")
db_Cogest.Open Session("Conn_String_Cogest_Gravacao")

%>
<html>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceBegin template="/Templates/BASICO_XPEP_01.dwt" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>SINERGIA # XPROC # Processos de Negócio</title>
<link href="../../../css/objinterface.css" rel="stylesheet" type="text/css">
<link href="../../../css/texinterface.css" rel="stylesheet" type="text/css">
<!-- InstanceEndEditable -->
</head>
<!-- InstanceBeginEditable name="JavaSci" -->
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
           document.frm1.action="inclui_plano_pcd.asp";
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
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../Funcao/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../Funcao/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../Funcao/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../Funcao/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../Funcao/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../../indexA.asp"><img src="../../Funcao/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
      
    <td colspan="3" height="20"><!-- InstanceBeginEditable name="Botao_01" -->
      <table width="625" border="0" align="center">
        <tr> 
          <td width="26"><a href="javascript:Confirma()"><img src="../../../imagens/continua_F02.gif" width="24" height="24" border="0"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
          <td width="195"></td>
          <td width="27"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
      <!-- InstanceEndEditable --></td>
  </tr>
</table>     
<!-- InstanceBeginEditable name="Corpo_Princ" -->
<table width="98%" border="0" cellspacing="0" cellpadding="0">
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
    <td class="subtitulo">Seleciona</td>
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
    <td width="13%"> <div align="right" class="campo">Onda:</div></td>
    <td width="85%"> <!--#include file="../includes/inc_Combo_Onda.asp" --> </td>
    <td width="2%">&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="right" class="campo">Plano:</div></td>
    <td> <!--#include file="../includes/inc_Combo_Plano.asp" --> </td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><div align="right" class="campo">Projeto:</div></td>
    <td> <!--#include file="../includes/inc_combo_projeto.asp" --> </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="campo"><div align="right">Fases:</div></td>
    <td><!--#include file="../includes/inc_combo_fases.asp" --> </td></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td class="campo"><div align="right">Tarefa N&iacute;vel 1:</div></td>
    <td> <!--#include file="../includes/inc_combo_tarefas_nivel1.asp" --> </td>
    <td>&nbsp;</td>
  </tr>
  <tr>
      <td height="160">&nbsp;</td>
    <td><table width="100%" border="1" cellpadding="0" cellspacing="0">
          <tr class="titcoltabela"> 
            <td width="16%"><div align="center">Descri&ccedil;&atilde;o Parada</div></td>
            <td width="12%"><div align="center">Resp Sinergia</div></td>
            <td colspan="2"> <div align="center">Resp Legado</div></td>
            <td width="16%"><div align="center">Tempo de Parada</div></td>
            <td width="24%"><div align="center">Procedimentos para parada</div></td>
            <td width="14%"><div align="center">Data da Parada</div></td>
            <td width="6%">&nbsp;</td>
          </tr>
          <tr> 
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td width="12%" height="2" valign="top" class="titcoltabela"><div align="center">T&eacute;cnico</div></td>
            <td width="16%" valign="top" class="titcoltabela"><div align="center">Funcional</div></td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
          </tr>
          <tr> 
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="2" valign="top">&nbsp;</td>
            <td>&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
          </tr>
          <tr> 
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="3" valign="top">&nbsp;</td>
            <td>&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
            <td height="30">&nbsp;</td>
          </tr>
          <tr> 
            <td height="23">&nbsp;</td>
            <td height="23">&nbsp;</td>
            <td height="23" valign="top">&nbsp;</td>
            <td>&nbsp;</td>
            <td height="23">&nbsp;</td>
            <td height="23">&nbsp;</td>
            <td height="23">&nbsp;</td>
            <td height="23">&nbsp;</td>
          </tr>
        </table></td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
      <td> 
        <table width="75%" border="0">
          <tr>
            <td><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_ppo.asp">PPO 
              </a></td>
            <td><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_pcd.asp">PCD 
              </a></td>
            <td><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_pcm.asp">PCM 
              </a></td>
            <td><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_pai.asp">PAI</a></td>
            <td><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_pac.asp">PAC 
              </a></td>
            <td><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_pce.asp">PCE 
              </a></td>
            <td><a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_pds.asp">PDS 
              </a></td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>
        <br>
        <a href="http://s6000ws10.corp.petrobras.biz/xproc/asp/xpep/asp/inclui_plano_pds.asp"> 
        </a></td>
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
</body>

<!-- InstanceEnd --></html>
