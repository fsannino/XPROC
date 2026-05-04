<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

processo=0

mega=request("mega")
curso=request("curso")

strSqlCurso= ""
strSqlCurso = strSqlCurso & "SELECT CURS_CD_CURSO, CURS_TX_NOME_CURSO FROM " & Session("PREFIXO") & "CURSO " 
strSqlCurso = strSqlCurso & "WHERE CURS_CD_CURSO = '" & trim(curso) & "'"
set rs_curso_Label = db.execute(strSqlCurso)

strSqlMega = ""
strSqlMega = strSqlMega & "SELECT MEPR_CD_MEGA_PROCESSO, MEPR_TX_DESC_MEGA_PROCESSO FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
strSqlMega = strSqlMega & "WHERE MEPR_CD_MEGA_PROCESSO=" & cint(mega)
set rsMega = db.execute(strSqlMega)

ssql= ""
ssql = ssql & "SELECT CURS_CD_CURSO, CURS_TX_NOME_CURSO FROM " & Session("PREFIXO") & "CURSO " 
ssql = ssql & "WHERE CURS_CD_CURSO <> '" & trim(curso) & "'"
ssql = ssql & " AND MEPR_CD_MEGA_PROCESSO=" & cint(mega)
ssql = ssql & " AND CURS_CD_CURSO NOT IN (SELECT CURS_CD_CURSO_CORRELATO FROM CURSO_CORRELATO WHERE CURS_CD_CURSO = '" & trim(curso) & "')"
ssql = ssql & " ORDER BY CURS_TX_NOME_CURSO"
set rsCursoNaoSelec = db.execute(ssql)

ssql_tem = ""
ssql_tem = ssql_tem & "SELECT CURS_CD_CURSO_CORRELATO "
ssql_tem = ssql_tem & "FROM CURSO_CORRELATO "
ssql_tem = ssql_tem & "WHERE CURS_CD_CURSO='" & curso & "'"
set rsCursoSelec = db.execute(ssql_tem)
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script language="javascript" src="../js/troca_lista.js"></script>

<script>
function Confirma()
{
carrega_txt(document.frm1.list2);
document.frm1.submit();
}

function carrega_txt(fbox) 
{
	document.frm1.txtTrans.value = "";
	for(var i=0; i<fbox.options.length; i++) 
	{
	document.frm1.txtTrans.value = document.frm1.txtTrans.value + "," + fbox.options[i].value;
	}
}

</script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_rel_curso_correlato.asp" name="frm1">
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
          Curso x Curso Correlato</font></div>
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="849" height="70">
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="19" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Curso
        :&nbsp;</b></font></td>
            
      <td width="531" height="19" valign="middle" align="left"> 
	  	<font face="Verdana" size="2" color="#330099"><%=rs_curso_Label("CURS_CD_CURSO") & " - " & rs_curso_Label("CURS_TX_NOME_CURSO")%></font>
	  </td>
       
	   	<%
	    rs_curso_Label.close
  		set rs_curso_Label = nothing 
  		%>  
		   
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="19" valign="middle" align="left"><font face="Verdana" size="2" color="#330099"><b>Mega-Processo
        :</b></font></td>
            
      <td width="531" height="19" valign="middle" align="left"> 
      	<font face="Verdana" size="2" color="#330099"><%=rsMega("MEPR_TX_DESC_MEGA_PROCESSO")%></font> 
      </td>
       <%
	   	rsMega.close
  		set rsMega = nothing
	   %>     
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"><input type="hidden" name="mega" size="10" value="<%=mega%>"></td>
            
      <td height="1" valign="middle" align="left" width="531"> 
      </td>
          </tr>
          <tr>
            
      <td width="162" height="1"></td>
            
      <td width="136" height="1" valign="middle" align="left"><input type="hidden" name="curso" size="10" value="<%=curso%>"></td>
            
      <td height="1" valign="middle" align="left" width="531"> 
      </td>
          </tr>
        </table>

<p style="margin: 0" align="center">&nbsp;<font face="Verdana" size="2" color="#330099"><b>Curso Correlato</b></font></p>

        <table border="0" width="964" height="142">
          <tr>
            <td width="300" height="138" rowspan="5"></td>
            
      <td width="300" height="138" rowspan="5"> <p style="margin: 0"> 
        <table width="75%" border="0">
          <tr>
            <td><font face="Verdana" size="2" color="#330099"><b>N&atilde;o selecionados</b></font></td>
          </tr>
          <tr>
            <td>
				<select size="7" name="list1" multiple>
					<%do until rsCursoNaoSelec.eof=true%>
						 <option value="<%=rsCursoNaoSelec("CURS_CD_CURSO")%>"><%=rsCursoNaoSelec("CURS_TX_NOME_CURSO")%></option>
					<%rsCursoNaoSelec.MOVENEXT               
					 loop%>
              	</select>
				  <%
				  rsCursoNaoSelec.close
				  set rsCursoNaoSelec = nothing    
				  %>
			 </td>
          </tr>
        </table></td>
            <td width="117" height="28" align="center">
              <p style="margin: 0"></td>
            
      <td width="526" height="138" rowspan="5"> <p style="margin: 0"> 
        <table width="75%" border="0">
          <tr>
            <td><font face="Verdana" size="2" color="#330099"><b>Selecionados</b></font></td>
          </tr>
          <tr>
            <td>
				<select size="7" name="list2" multiple>
					<%
					if not rsCursoSelec.eof then
						do until rsCursoSelec.eof = true			   		   	   
						set rsTemp = db.execute("SELECT DISTINCT CURS_TX_NOME_CURSO FROM CURSO WHERE CURS_CD_CURSO='" & rsCursoSelec("CURS_CD_CURSO_CORRELATO") & "'")
						if rsTemp.eof = false then
						   strDesc_Curso = rsTemp("CURS_TX_NOME_CURSO")
						   	%>							
							<option value="<%=rsCursoSelec("CURS_CD_CURSO_CORRELATO")%>"><%=strDesc_Curso%></option>
							<%
						end if
						rsCursoSelec.MOVENEXT               
						loop
						
					  	rsTemp.close
					  	set rsTemp = nothing    						
					end if
					
					rsCursoSelec.close
					set rsCursoSelec = nothing 					
				   %>
				 </select>
			</td>
          </tr>
        </table></td>
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