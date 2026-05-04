<%
DIM Conta_Array(8)

'Conta_Array(0) = cINT(Request.Form("indicados"))
'Conta_Array(1) = cINT(Request.Form("mapeados"))

Conta_Array(0) = 10
Conta_Array(1) = 20

DIM Label_Array(8)

Label_Array(0) =  " INDICADOS - " & Conta_Array(0)
Label_Array(1) =  " MAPEADOS - " & Conta_Array(1)

%>
<HTML>
<HEAD><TITLE>Gráfico</TITLE></HEAD>
<BODY>
<BR>
<p align="left">
   <br>
  <br>
   
  </p>
   
  <table width="554" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#F0F0FF">
    <tr align="center" valign="middle"> 
         
      <td width="52" height="25">&nbsp;</td>
         
      <td width="469" align="left" valign="top"><font face="Arial, Helvetica, sans-serif" color="#040481"><b><br>
            </b></font></td>
         
      <td width="33" height="25"><p align="left">&nbsp;</td>
      </tr>
      <tr align="center" valign="middle"> 
         
      <td width="52" height="38">&nbsp;</td>
         
      <td width="469" align="center" valign="middle" bgcolor="#CCCCCC" height="38"><applet code="aspbr_pie_2d.class" width=460 height=210>
          <param name="NomeGrafico" value="Demonstrativo de vendas das filiais.">
          <param name="var01" value="<%=Conta_Array(0)%>">
          <param name="var02" value="<%=Conta_Array(1)%>">
          <param name="nome1" value="<%=Label_Array(0)%>">
          <param name="nome2" value="<%=Label_Array(1)%>">Por favor habilite o seu browser para permitir JAVA 
        </applet></td>
      <td width="33" height="38">&nbsp;</td>
      </tr>
      <tr align="center" valign="middle"> 
         
      <td width="52">&nbsp;</td>
         
      <td width="469" align="left" valign="top">&nbsp;</td>
         
      <td width="33">&nbsp;</td>
      </tr>
   </table>

<div align="left"><BR>
  &nbsp;</div>
<p>&nbsp;</p>
</BODY>
</HTML>