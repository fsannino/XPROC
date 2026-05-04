 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

processo=0

mega=request("mega")
curso=request("curso")
processo=request("proc")
subproc=request("subproc")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & curso & "'")

valor1=rs("CURS_CD_CURSO") & " - " & rs("CURS_TX_NOME_CURSO")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rsproc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega)

if processo<>0 then
	set rssub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega & " AND PROC_CD_PROCESSO=" & processo)
else
	set rssub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=0")
end if

ssql=""
ssql="SELECT DISTINCT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO=" & mega 

if processo<>0 then
	ssql=ssql+ " AND PROC_CD_PROCESSO=" & processo
end if

if subproc<>0 then
	ssql=ssql+ " AND SUPR_CD_SUB_PROCESSO=" & subproc
end if

SSQL=SSQL+" ORDER BY TRAN_CD_TRANSACAO"

set rstrans=db.execute(ssql)

ssql=""
ssql="SELECT DISTINCT TRAN_CD_TRANSACAO FROM " & Session("PREFIXO") & "CURSO_TRANSACAO WHERE CURS_CD_CURSO='" & curso & "' AND MEPR_CD_MEGA_PROCESSO=" & mega 


if processo<>0 then
	ssql=ssql+ " AND PROC_CD_PROCESSO=" & processo
end if

if subproc<>0 then
	ssql=ssql+ " AND SUPR_CD_SUB_PROCESSO=" & subproc
end if

ssql_tem=ssql

ssql=ssql+ " ORDER BY TRAN_CD_TRANSACAO"

set rstranscurso=db.execute(ssql)

'RESPONSE.WRITE SSQL_tem
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="javascript" src="../js/troca_lista.js"></script>

<script>
function Confirma()
{
if(document.frm1.selProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um PROCESSO!");
document.frm1.selProcesso.focus();
return;
}
if(document.frm1.mega.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MEGA-PROCESSO!");
document.frm1.mega.focus();
return;
}
if(document.frm1.selSubProcesso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um SUB-PROCESSO!");
document.frm1.selSubProcesso.focus();
return;
}
else
{
carrega_txt(document.frm1.list2)
document.frm1.submit();
}
}

function envia1()
{
window.location.href="rel_curso_transacao.asp?mega="+document.frm1.mega.value+"&curso="+document.frm1.curso.value
}

function envia2()
{
window.location.href="rel_curso_transacao.asp?mega="+document.frm1.mega.value+"&curso="+document.frm1.curso.value+"&proc="+document.frm1.selProcesso.value
}

function envia3()
{
window.location.href="rel_curso_transacao.asp?mega="+document.frm1.mega.value+"&curso="+document.frm1.curso.value+"&proc="+document.frm1.selProcesso.value+"&subproc="+document.frm1.selSubProcesso.value
}

function carrega_txt(fbox) 
{
document.frm1.txtTrans.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtTrans.value = document.frm1.txtTrans.value + "," + fbox.options[i].value;
}
}
</script>
<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_rel_curso_transacao.asp" name="frm1">
        <input type="hidden" name="txtTrans">
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
        <div align="center"><font face="Verdana" color="#330099" size="3">Relação
          Curso x Transação</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="64">
          <tr>
            
      <td width="162" height="19"></td>
            
      <td width="136" height="19" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Curso
        :&nbsp;</b></font></td>
            
      <td width="531" height="19" valign="middle" align="left"> 
        <font face="Verdana" size="2" color="#330099"><%=VALOR1%></font></td>
            
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="19" valign="middle" align="left"></td>
            
      <td width="531" height="19" valign="middle" align="left"> 
      </td>
            
          </tr>
          <tr>
            
      <td width="162" height="1"><input type="hidden" name="curso" size="10" value="<%=curso%>"></td>
            
      <td width="136" height="19" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo
        :</b></font></td>
            
      <td width="531" height="19" valign="middle" align="left"> 
        <font face="Verdana" size="2" color="#330099">
      <select size="1" name="mega" onchange="javascript:envia1()">
      <option value="0">== Selecione o Processo ==</option>
      <%DO UNTIL RS.EOF=TRUE
        if trim(mega)=trim(RS("MEPR_CD_MEGA_PROCESSO")) then
        %>
        <option selected value="<%=RS("MEPR_CD_MEGA_PROCESSO")%>"><%=RS("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
        <%else%>
        <option value="<%=RS("MEPR_CD_MEGA_PROCESSO")%>"><%=RS("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
        <%
        end if
        RS.MOVENEXT
        LOOP
        %>        
      </select> 
        </font></td>
            
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Processo
        : </b></font></td>
            
      <td height="1" valign="middle" align="left" width="531"> 
      <select size="1" name="selProcesso" onchange="javascript:envia2()">
      <option value="0">== Selecione o Processo ==</option>
      <%DO UNTIL RSPROC.EOF=TRUE
        if trim(processo)=trim(RSPROC("PROC_CD_PROCESSO")) then
        %>
        <option selected value="<%=RSPROC("PROC_CD_PROCESSO")%>"><%=RSPROC("PROC_TX_DESC_PROCESSO")%></option>
        <%else%>
        <option value="<%=RSPROC("PROC_CD_PROCESSO")%>"><%=RSPROC("PROC_TX_DESC_PROCESSO")%></option>
        <%
        end if
        RSPROC.MOVENEXT
        LOOP
        %>        
      </select> 
      </td>
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Sub-Processo
        :</b></font></td>
            
      <td height="1" valign="middle" align="left" width="531"> 
        <select size="1" name="selSubProcesso" onchange="javascript:envia3()">
        <option value="0">== Selecione o Sub-Processo ==</option>
        <%DO UNTIL RSSUB.EOF=TRUE
        if trim(subproc)=trim(RSSUB("SUPR_CD_SUB_PROCESSO")) then
        %>
        <option selected value="<%=RSSUB("SUPR_CD_SUB_PROCESSO")%>"><%=RSSUB("SUPR_TX_DESC_SUB_PROCESSO")%></option>
        <%else%>
        <option value="<%=RSSUB("SUPR_CD_SUB_PROCESSO")%>"><%=RSSUB("SUPR_TX_DESC_SUB_PROCESSO")%></option>
        <%
        end if
        RSSUB.MOVENEXT
        LOOP
        %>        
        </select></td>
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"></td>
            
      <td height="1" valign="middle" align="left" width="531"> 
      </td>
          </tr>
        </table>

<p style="margin: 0" align="center"><font face="Verdana" size="2" color="#330099"><b>Transações</b></font></p>
        <table border="0" width="964" height="142">
          <tr>
            <td width="300" height="138" rowspan="5"></td>
            <td width="300" height="138" rowspan="5">
              <p style="margin: 0"><select size="7" name="list1" multiple>
               <%do until rstrans.eof=true
               set rstem=db.execute( ssql_tem +  " AND TRAN_CD_TRANSACAO='" & rstrans("TRAN_CD_TRANSACAO") & "'")
               if rstem.eof=true then
               set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & rstrans("TRAN_CD_TRANSACAO") & "'")
               VALOR_TRANS=RSTEMP("TRAN_TX_DESC_TRANSACAO")
               %>
               <option value="<%=rstrans("TRAN_CD_TRANSACAO")%>"><%=rstrans("TRAN_CD_TRANSACAO")%>-<%=VALOR_TRANS%></option>
               <%
				  end if
				  RSTRANS.MOVENEXT               
                LOOP
               %>
              </select></td>
            <td width="117" height="28" align="center">
              <p style="margin: 0"></td>
            <td width="526" height="138" rowspan="5">
              <p style="margin: 0"><select size="7" name="list2" multiple>
               <%do until rstranscurso.eof=true
               set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & rstranscurso("TRAN_CD_TRANSACAO") & "'")
               VALOR_TRANS=RSTEMP("TRAN_TX_DESC_TRANSACAO")
               %>
               <option value="<%=rstranscurso("TRAN_CD_TRANSACAO")%>"><%=rstranscurso("TRAN_CD_TRANSACAO")%>-<%=VALOR_TRANS%></option>
               <%
               RSTRANSCURSO.MOVENEXT               
               LOOP
               %>
              </select></td>
          </tr>
          <tr>
            <td width="117" height="28" align="center"><a href="#" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="../Funcao/continua_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr>
            <td width="117" height="28" align="center"></td>
          </tr>
          <tr>
            <td width="117" height="27" align="center"><a href="javascript:;"  onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="../Funcao/continua2_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr>
            <td width="117" height="27" align="center"></td>
          </tr>
        </table>

<p style="margin: 0">&nbsp;</p>

<p style="margin: 0">&nbsp;</p>

<p style="margin: 0">&nbsp;</p>
  </form>

</body>

</html>
