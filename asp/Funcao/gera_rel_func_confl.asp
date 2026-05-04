<%
IF request("EXCEL")=1 THEN
	Response.Buffer = TRUE
	Response.ContentType = "application/vnd.ms-excel"
END IF

server.ScriptTimeout=99999999

dim int_Tot_Reg_Impresso

int_Tot_Reg_Impresso = 0

str_Uso = request("chkEmUso")
str_Desuso = request("chkEmDesuso")
if str_Uso = "" then
   str_Uso = 0
end if   
if str_Desuso = "" then
   str_Desuso = 0
end if   
if str_Uso = 1 and str_Desuso = 1 then
   str_usoDesuso =  " and (FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' or FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0')" 
else
   if str_Uso = 1 then
      str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '1' "
   else
      if str_Desuso = 1 then
         str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '0' "
	  else
     	 str_usoDesuso =  " and FUNCAO_NEGOCIO.FUNE_TX_INDICA_EM_USO = '2' "
	  end if	 
	end if        	  
end if


if request("selSubModulo") <> "" then
   str_Assunto=request("selSubModulo")
else
   str_Assunto= "0"
end if   

if request("selFuncao") <> "" then
   str_Funcao=request("selFuncao")
else
   str_Funcao = "0"
end if

compl5 = ""
if str_Assunto<>"0" then
	compl5=compl5 & " AND FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA =" & str_Assunto 
end if

if str_Funcao<>"0" then
	compl5=compl5+" AND FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO='" & str_Funcao & "'"
end if

mega=request("selMegaProcesso")
if mega="" then
   mega = request("selMega")
end if

compl = ""

if mega<>0 then
	compl=" and MEPR_CD_MEGA_PROCESSO=" & mega 
end if

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")
db.CursorLocation=3

str_Sql = "SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO where MEPR_CD_MEGA_PROCESSO > 0 " & compl & " ORDER BY MEPR_TX_DESC_MEGA_PROCESSO"
set rs_mega=db.execute(str_Sql)
'response.Write(str_Sql & "<p>")

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_func_confl.asp" name="frm1">

<%if request("excel")=0 then%>
		<input type="hidden" name="txtImp" size="20">
       <input type="hidden" name="txtFuncSelec" size="20">
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
          <td width="26">&nbsp;</td>
          <td width="195"></td>
            <td width="27"></td>  <td width="50"><a href="gera_rel_func_confl.asp?excel=1&selMega=<%=mega%>&chkEmUso=<%=str_uso%>&chkEmDesuso=<%=str_Desuso%>&selFuncao=<%=str_funcao%>&selSubModulo=<%=str_Assunto%>" target="_blank"><img border="0" src="../../imagens/exp_excel.gif"></a></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if%>        
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>
        &nbsp;
        <div align="center">
          <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" color="#330099">Relatório
          de Função(ões) R/3 Conflitante(s)</font></div>
      </td>
    </tr>
  </table>
  
          <p style="margin-top: 0; margin-bottom: 0">
          <table border="0" width="100%" bgcolor="#330099" cellspacing="0" cellpadding="2" bordercolor="#FFFFFF">
            <tr>
              <td width="27%"><b><font face="Verdana" size="1" color="#FFFFFF">Mega-Processo</font></b></td>
              <td width="36%"><b><font face="Verdana" size="1" color="#FFFFFF">Função
                R/3</font></b></td>
              <td width="37%"><b><font face="Verdana" size="1" color="#FFFFFF">Função
                R/3 Conflitante</font></b></td>
            </tr>
            <%
            im=0
            mega=rs_mega.RecordCount
            
            tem=0
			do until im = mega

			str_Sql_Func=""
			str_Sql_Func="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
			str_Sql_Func=str_Sql_Func+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
			str_Sql_Func=str_Sql_Func+"INNER JOIN FUNCAO_NEGOCIO ON "
			str_Sql_Func=str_Sql_Func+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
			str_Sql_Func=str_Sql_Func+"WHERE FUNCAO_NEGOCIO.MEPR_CD_MEGA_PROCESSO=" & RS_MEGA("MEPR_CD_MEGA_PROCESSO") & str_usoDesuso & compl5 			
			str_Sql_Func=str_Sql_Func+"ORDER BY FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO "
		
		   set temp=db.execute(str_Sql_Func)
	   
           valor=rs_mega("MEPR_TX_DESC_MEGA_PROCESSO")
               
			   it=0
			   itemp=temp.RecordCount
			   
			   do until it = itemp
                  set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_CONFL WHERE FUNE_CD_FUNCAO_NEGOCIO='" & TEMP("FUNE_CD_FUNCAO_NEGOCIO") & "' ORDER BY FUNE_CD_FUNCAO_NEGOCIO, FUNC_cD_FUNCAO_CONFL")
                  ANTERIOR=""
                  ATUAL=""
                  iss = 0
                  isrs = rs.RecordCount
                  
                  IF RS.EOF=FALSE THEN
                  
                  DO UNTIL iss = isrs
                     ATUAL=RS("FUNE_CD_FUNCAO_NEGOCIO") 
					 int_Tot_Reg_Impresso = int_Tot_Reg_Impresso + 1          
            %>
            <tr>
              <td width="27%" bgcolor="#FFFFFF"><font face="Verdana" size="1"><b><%=valor%></b></font></td>
              <%
              SET TEMP2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNE_CD_FUNCAO_NEGOCIO") & "'")
              IF ANTERIOR<>ATUAL THEN
	              VALOR2=RS("FUNE_CD_FUNCAO_NEGOCIO")& "-" & TEMP2("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
	           ELSE
	           	VALOR2=""
	           END IF
              %>
              <td width="36%" bgcolor="#FFFFFF"><font face="Verdana" size="1"><%=VALOR2%></font></td>
              <%
              SET TEMP2=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & RS("FUNC_CD_FUNCAO_CONFL") & "'")
              VALOR3=TEMP2("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
              %>
             <td width="37%" bgcolor="#FFFFFF"><font face="Verdana" size="1"><%=RS("FUNC_CD_FUNCAO_CONFL")%>-<%=VALOR3%></font></td>
            </tr>
            <%
			tem=tem+1
            ANTERIOR=RS("FUNE_CD_FUNCAO_NEGOCIO")
            
            iss=iss+1
            RS.MOVENEXT
            valor=" "
            LOOP
            
            END IF
            
            it=it+1
            temp.movenext
            Loop
            
            im=im+1
            rs_mega.movenext
            loop
            %>
          </table>
       <%if tem=0 then%>   
  <p style="margin-top: 0; margin-bottom: 0"><font color="#660000"><strong>Nenhum 
    Registro encontrado para a Sele&ccedil;&atilde;o</strong></font> </p>
	<%end if%>
  <table width="75%" border="0">
    <tr> 
      <td width="10%">&nbsp;</td>
      <td width="32%">&nbsp;</td>
      <td width="58%">&nbsp;</td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><div align="right"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total 
          de registros impressos:</font></strong></div></td>
      <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=int_Tot_Reg_Impresso%></font></strong></td>
    </tr>
  </table>

    </form>


</body>

</html>