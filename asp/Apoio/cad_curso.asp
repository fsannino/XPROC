<!--#include file="conn_consulta.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
mega=0

chave=request("chave")
atribb=request("attrib")

set fonte_mega=db.execute("SELECT DISTINCT SUMO_NR_CD_SEQUENCIA FROM APOIO_LOCAL_MODULO WHERE USMA_CD_USUARIO ='" & CHAVE & "' AND APLO_NR_ATRIBUICAO=" & atribb & " ORDER BY SUMO_NR_CD_SEQUENCIA")

do until fonte_mega.eof=true
	set q_mega=db.execute("SELECT * FROM SUB_MODULO WHERE SUMO_NR_CD_SEQUENCIA='" & fonte_mega("SUMO_NR_CD_SEQUENCIA") & "'")
	sequencia=sequencia & q_mega("mepr_cd_mega_processo_todos") & ","
	fonte_mega.movenext
loop

sequencia = replace(sequencia,"-",",")
sequencia = left(sequencia,len(sequencia)-1)

set Rusuario=db.execute("SELECT * FROM " & Session("Prefixo") & "USUARIO_MAPEAMENTO WHERE USMA_CD_USUARIO='" & chave & "'")

usuario=Rusuario("USMA_TX_NOME_USUARIO")

set rs_curso=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE MEPR_CD_MEGA_PROCESSO IN (" & SEQUENCIA & ") ORDER BY CURS_CD_CURSO")

do until rs_curso.eof=true
	cursos = cursos & "'" & rs_curso("CURS_CD_CURSO") & "',"
	rs_curso.movenext
loop

cursos = left(cursos,len(cursos)-1)

rs_curso.movefirst

ssql=""
ssql="SELECT DISTINCT CURS_CD_CURSO FROM " & Session("PREFIXO") & "APOIO_LOCAL_CURSO WHERE USMA_CD_USUARIO ='" & CHAVE & "' AND APLO_NR_ATRIBUICAO=" & atribb & " ORDER BY CURS_CD_CURSO"

set rscurso=db.execute(ssql)

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="javascript" src="troca_lista.js"></script>

<script>
function Confirma()
{
{
carrega_txt(document.frm1.list2)
document.frm1.submit();
}
}

function carrega_txt(fbox) 
{
document.frm1.txtCurso.value = "";
for(var i=0; i<fbox.options.length; i++) {
document.frm1.txtCurso.value = document.frm1.txtCurso.value + "," + fbox.options[i].value;
}
}

function envia1()
{
window.location.href="cad_curso.asp?chave=<%=chave%>&attrib=<%=atribb%>&mega="+document.frm1.selMega.value
}

function envia2()
{
window.location.href="rel_curso_cenario.asp?mega=0&curso="+document.frm1.selCurso.value
}

</script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_cad_curso.asp" name="frm1">
        <input type="hidden" name="txtCurso">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
                <p align="center">&nbsp;</div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center">&nbsp;</div>
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
          Multiplicador x Curso</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="50">
    <tr> 
      <td height="1"> <input name="txtchave" type="hidden" id="txtchave" value="<%=chave%>"></td>
      <td width="137" bgcolor="#000080"><b><font face="Verdana" size="2" color="#FFFFFF">Usuário&nbsp;&nbsp;</font></b></td>
      <td height="19" valign="middle" align="left"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=chave%> - <%=usuario%></font></td>
    </tr>
    <tr> 
      <td height="1"> <input name="txtapoio" type="hidden" id="txtapoio" value="<%=atribb%>"></td>
      <td width="137" bgcolor="#000080"><b><font face="Verdana" size="2" color="#FFFFFF">Atribuição&nbsp;&nbsp;</font></b></td>
      <td height="19" valign="middle" align="left"><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=atribb%> - MULTIPLICADOR</font></td>
    </tr>
    <tr> 
      <td width="165" height="1">
	<input name="txtmegas" type="hidden" id="txtmegas" value="<%=sequencia%>"></td>
      <td width="137" height="1" valign="middle" align="left"><input type="hidden" name="mega" size="10" value="<%=mega%>"></td>
      <td height="1" valign="middle" align="left" width="533"> </td>
    </tr>
    <tr> 
      <td width="165" height="1"></td>
      <td width="137" height="1" valign="middle" align="left"><input type="hidden" name="curso" size="10" value="<%=curso%>"></td>
      <td height="1" valign="middle" align="left" width="533"> </td>
    </tr>
  </table>
  <p style="margin: 0" align="center"><font face="Verdana" size="2" color="#330099"><b>Cursos 
    Dispon&iacute;veis </b></font></p>

        <table border="0" width="964" height="142">
          <tr>
            <td width="300" height="138" rowspan="5"></td>
            <td width="300" height="138" rowspan="5">
              <p style="margin: 0"><select size="7" name="list1" multiple>
               <%
			   DO UNTIL RS_CURSO.EOF=TRUE
			   set tem=db.execute("SELECT * FROM " & Session("PREFIXO") & "APOIO_LOCAL_CURSO WHERE USMA_CD_USUARIO ='" & CHAVE & "' AND APLO_NR_ATRIBUICAO=" & atribb & " AND CURS_CD_CURSO='" & RS_CURSO("CURS_CD_CURSO") & "' ORDER BY CURS_CD_CURSO")
			   if tem.eof=true then
			   %>
			   <option value="<%=RS_CURSO("CURS_CD_CURSO")%>"><%=RS_CURSO("CURS_CD_CURSO")%> - <%=RS_CURSO("CURS_TX_NOME_CURSO")%></option>
               <%
			   end if
			   RS_CURSO.MOVENEXT
			   LOOP
			   %>
			   </select></td>
            <td width="117" height="28" align="center">
              <p style="margin: 0"></td>
            <td width="526" height="138" rowspan="5">
              <p style="margin: 0"><select size="7" name="list2" multiple>
               <%
			   DO UNTIL RSCURSO.EOF=TRUE
			   set TEMP=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & RSCURSO("CURS_CD_CURSO") & "' ORDER BY CURS_CD_CURSO")
			   %>
			   <option value="<%=RSCURSO("CURS_CD_CURSO")%>"><%=RSCURSO("CURS_CD_CURSO")%> - <%=TEMP("CURS_TX_NOME_CURSO")%></option>
               <%
			   RSCURSO.MOVENEXT
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

  </form>

</body>

</html>