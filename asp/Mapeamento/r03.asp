<%
orgao = request("orgao")

DIM Conta_Array(8)

Conta_Array(0) = cINT(Request("indicados"))
Conta_Array(1) = cINT(Request("mapeados"))
Conta_Array(2) = cINT(Request("resto"))

DIM Label_Array(8)

Label_Array(0) =  "  - Indicadas - " & Conta_Array(0)
Label_Array(1) =  "  - Mapeadas - " & Conta_Array(1)
Label_Array(2) =  "  - Pendentes - " & Conta_Array(2)
%>
<HTML>
<HEAD>
<meta http-equiv="Content-Language" content="pt-br">
<TITLE>Indicação de Multiplicadores</TITLE>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
</HEAD>

<BODY link="#800000" vlink="#800000" alink="#800000">
<table width="98%" border="0" height="21">
  <tr>
    <td width="61%"><font face="Verdana" size="2">Órgão Selecionado : <b><%=orgao%></b></font></td>
    <td width="39%"> 
      <div align="left">
        <p style="margin-top: 0; margin-bottom: 0" align="right"><font face="Verdana" size="2"><b><a href="javascript:window.close()">Fechar 
          Janela</a></b></font> 
      </div>
      </td>
  </tr>
</table>
<p align="left" style="margin-top: 0; margin-bottom: 0"><applet code="aspbr_pie_2d.class" width=376 height=210>
    <param name="NomeGrafico" value="">
    <param name="var01" value="0">
    <param name="var02" value="<%=Conta_Array(0)%>">
    <param name="var03" value="0">
    <param name="var04" value="0">
    <param name="var05" value="<%=Conta_Array(1)%>">
    <param name="var06" value="0">
    <param name="var07" value="0">
    <param name="var08" value="0">
    <param name="nome2" value="<%=Label_Array(0)%>">
    <param name="nome5" value="<%=Label_Array(1)%>">
    <!-- Mensagem mostrada se o usuário não está com o JAVA habilitado -->
    Por favor habilite o seu browser para permitir JAVA 
  </applet> </p>
<div align="left">
  <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font face="Arial, Helvetica, sans-serif">Pend&ecirc;ncias 
    em Vagas Indicadas - <%=Conta_Array(2)%></font></b></font> 
</div>
</BODY>
</HTML>