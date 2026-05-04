 
<!--#include file="../../asp/protege/protege.asp" -->
<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MEGA=REQUEST("MEGA")
CURSO=REQUEST("CURSO")
strStatus 	= REQUEST("STATUS")

ON ERROR RESUME NEXT

COMPL1 = " WHERE CURS_CD_CURSO <> '' "

IF REQUEST("CURSO")=0  THEN
	IF ERR.NUMBER=0 THEN
		IF LEN(MEGA)>0 THEN
			COMPL2=" AND MEPR_CD_MEGA_PROCESSO=" & MEGA
		END IF
	ELSE
		COMPL2=" AND CURS_CD_CURSO='" & REQUEST("CURSO") & "'"
	END IF
END IF

if strStatus = "0" then     '*** TODOS
	COMPL3 = ""
elseif strStatus = "1" then '*** ATIVOS	
	COMPL3 = " AND CURS_TX_STATUS_CURSO = '1'"	
elseif strStatus = "2" then	'*** INATIVOS	
	COMPL3 = " AND CURS_TX_STATUS_CURSO = '0'"	
end if

SSQL="SELECT * FROM " & Session("PREFIXO") & "CURSO" & COMPL1 & COMPL2 & COMPL3 & " ORDER BY MEPR_CD_MEGA_PROCESSO, CURS_CD_CURSO"

SET RS=DB.EXECUTE(SSQL)
%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>


</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="" name="frm1">
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
            <td width="26"></td>
          <td width="50"></td>
          <td width="26"></td>
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
        <div align="center">
          <p align="left" style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099" size="3">Relatório
          de Cursos x </font><font face="Verdana" color="#330099" size="3">Pré-Requisitos</font></div>
      </td>
    </tr>
  </table>
        <p style="margin-top: 0; margin-bottom: 0">
        <table border="0" width="100%">
          <tr>
            <td width="33%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Mega-Processo</font></b></td>
            <td width="33%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Curso</font></b></td>
            <td width="34%" bgcolor="#330099"><b><font face="Verdana" size="2" color="#FFFFFF">Curso
              Pré-Requisito</font></b></td>
          </tr>
          <%
          
          tem=0
          
          do until rs.eof=true
			
			set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO_PRE_REQUISITO WHERE CURS_CD_CURSO='" & rs("CURS_CD_CURSO") & "' ORDER BY CURS_CD_CURSO")			          
			
			mega_atual=""
			curso_atual=""
						
			do until rstemp.eof=true
			
          atual1=rs("mepr_cd_mega_processo")
          atual2=rs("curs_cd_curso")

          %>
          <tr>
				<%
             	SET RS1=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs("MEPR_CD_MEGA_PROCESSO"))
             	if atual1<>ant1 then
					NOME1=RS1("MEPR_TX_DESC_MEGA_PROCESSO")            
				else
					nome1=""
				end if
				if nome1="" then
					cor="white"
				else
					cor="#CCCCCC"	
				end if			
				%>
            <td width="33%" bgcolor="<%=cor%>"><font face="Verdana" size="2" color="#330099"><b><%=NOME1%></b></font></td>
            <%
				if atual2<>ant2 then
					NOME_CURSO=rs("CURS_CD_CURSO")& "-" & RS("CURS_TX_NOME_CURSO")            
				else
					nome_curso=""
				end if
				if nome_curso="" then
					cor="white"
				else
					cor="#FFFFDF"	
				end if			

            %>
            <td width="33%" bgcolor="<%=cor%>"><font face="Verdana" size="2" color="#330099"><%=NOME_CURSO%></font></td>
            <%
            	ssql1=	"SELECT * FROM " & Session("PREFIXO") & "CURSO WHERE CURS_CD_CURSO='" & RSTEMP("CURS_PRE_REQUISITO") & "'"
            	set rs5=db.execute(ssql1)
            	nome2= RSTEMP("CURS_PRE_REQUISITO") & "-" & RS5("CURS_TX_NOME_CURSO")            
            %>
            <td width="34%" bgcolor="#CCFFCC"><font face="Verdana" size="2" color="#330099"><%=nome2%></font></td>
          </tr>
          <%
          tem=tem+1
          
          ant1=rs("mepr_cd_mega_processo")
          ant2=rs("curs_cd_curso")
          
          rstemp.movenext
          
          atual1=rs("mepr_cd_mega_processo")
          atual2=rs("curs_cd_curso")

          loop

          rs.movenext
          loop
          %>
          
        </table>
<b>
<%if tem=0 then%>
<font face="Verdana" size="2" color="#800000">Nenhum Registro encontrado</font></b>
<%end if%>
  </form>

</body>

</html>
