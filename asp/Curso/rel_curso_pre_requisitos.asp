
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

processo=0

mega=request("mega")
curso=request("curso")

set rs_curso=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO ORDER BY CURS_CD_CURSO")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

set rscurso=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega &" AND CURS_CD_CURSO<>'" & curso & "' ORDER BY CURS_CD_CURSO")

ssql=""
ssql="SELECT * FROM " & Session("PREFIXO") & "CURSO_PRE_REQUISITO WHERE CURS_CD_CURSO='" & curso & "' ORDER BY CURS_CD_CURSO"

set rscursopre=db.execute(ssql)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="javascript" src="../js/troca_lista.js"></script>

<script>

function Confirma()
{
if(document.frm1.selMega.selectedIndex == 0)
{
alert("É obrigatória a seleção de um MEGA-PROCESSO!");
document.frm1.selMega.focus();
return;
}
if(document.frm1.selCurso.selectedIndex == 0)
{
alert("É obrigatória a seleção de um CURSO!");
document.frm1.selCurso.focus();
return;
}
else
{
carrega_txt(document.frm1.list2)
document.frm1.submit();
}
}

function carrega_txt(fbox) 
{
document.frm1.txtTrans.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtTrans.value = document.frm1.txtTrans.value + "," + fbox.options[i].value;
}
}

function envia1()
{
window.location="rel_curso_pre_requisitos.asp?mega="+document.frm1.selMega.value+"&curso="+document.frm1.selCurso.value
}

function envia2()
{
window.location="rel_curso_pre_requisitos.asp?mega=0&curso="+document.frm1.selCurso.value
}
</script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_rel_curso_pre_requisito.asp" name="frm1">
        <input type="hidden" name="curso" size="10" value="<%=curso%>"><input type="hidden" name="mega" size="10" value="<%=mega%>"><input type="hidden" name="txtTrans">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="confirma_f02.gif"></a></td>
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
          Curso x Pré-Requisitos</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="53">
          <tr>
            
      <td width="162" height="17"></td>
            
      <td width="136" height="19" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Curso
        :&nbsp;</b></font></td>
            
      <td width="531" height="19" valign="middle" align="left"> 
      <select size="1" name="selCurso" onchange="javascript:envia2()">
      <option value="0">== Selecione o Curso ==</option>
      <%DO UNTIL RS_curso.EOF=TRUE
        if trim(curso)=trim(RS_CURSO("CURS_CD_CURSO")) then
        %>
        <option selected value="<%=RS_CURSO("CURS_CD_CURSO")%>"><%=RS_CURSO("CURS_CD_CURSO")%>-<%=RS_CURSO("CURS_TX_NOME_CURSO")%></option>
        <%else%>
        <option value="<%=RS_CURSO("CURS_CD_CURSO")%>"><%=RS_CURSO("CURS_CD_CURSO")%>-<%=RS_CURSO("CURS_TX_NOME_CURSO")%></option>
        <%
        end if
        RS_CURSO.MOVENEXT
        LOOP
        %>        
      </select> 
      </td>
            
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="19" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo
        :</b></font></td>
            
      <td width="531" height="19" valign="middle" align="left"> 
      <font face="Verdana" size="2" color="#330099"><select size="1" name="selMega" onchange="javascript:envia1()">
      <option value="0">== Selecione o Mega-Processo ==</option>
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
        </font> 
      </td>
            
          </tr>
          <tr>
            
      <td width="162" height="8"></td>
            
      <td width="136" height="10" valign="top" align="left"></td>
            
      <td width="531" height="10" valign="middle" align="left"> 
      </td>
            
          </tr>
        </table>

  <table border="0" width="849" height="70">
        </table>

        <p style="margin: 0" align="center">&nbsp;<font face="Verdana" size="2" color="#330099"><b>Cursos</b></font></p>

        <table border="0" width="964" height="142">
          <tr>
            <td width="300" height="138" rowspan="5"></td>
            <td width="300" height="138" rowspan="5">
              <p style="margin: 0"><select size="7" name="list1" multiple>
              <%
              DO UNTIL rscurso.eof=true
              SET RSTEM=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "CURSO_PRE_REQUISITO WHERE CURS_CD_CURSO='" & CURSO & "' AND CURS_PRE_REQUISITO='" & rscurso("CURS_CD_CURSO") & "'")
              if RStem.eof=true then
              %>
				 <option value="<%=rscurso("CURS_CD_CURSO")%>"><%=rscurso("CURS_CD_CURSO")%>-<%=rscurso("CURS_TX_NOME_CURSO")%></option>
				 <%
				 end if
				 rscurso.movenext
				 loop
				 %>
              </select></td>
              
            <td width="117" height="28" align="center">
              <p style="margin: 0"></td>
            <td width="526" height="138" rowspan="5">
              <p style="margin: 0">
              <select size="7" name="list2" multiple>
           	<%
           	DO UNTIL rscursopre.eof=true
           	set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & rscursopre("CURS_PRE_REQUISITO") & "'")
            ON ERROR RESUME NEXT	
			VALOR = temp("CURS_TX_NOME_CURSO")
			IF ERR.NUMBER=0 THEN
           	%>
				 <option value="<%=rscursopre("CURS_PRE_REQUISITO")%>"><%=rscursopre("CURS_PRE_REQUISITO")%>-<%=VALOR%></option>
			<%
			END IF
			rscursopre.movenext
			loop
			%>
             </select>
              
              </td>
              
          </tr>
          <tr>
            <td width="117" height="28" align="center"><a href="#" onClick="move(document.frm1.list1,document.frm1.list2,1)"><img name="Image1611" border="0" src="continua_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr>
            <td width="117" height="28" align="center"></td>
          </tr>
          <tr>
            <td width="117" height="27" align="center"><a href="javascript:;"  onClick="move(document.frm1.list2,document.frm1.list1,1)"><img name="img0151111" border="0" src="continua2_F01.gif" width="24" height="24"></a></td>
          </tr>
          <tr>
            <td width="117" height="27" align="center"></td>
          </tr>
        </table>
  </form>
</body>

</html>
