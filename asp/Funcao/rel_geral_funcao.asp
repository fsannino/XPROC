<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

if request("selMegaProcesso") <> "" then
   str_mega=request("selMegaProcesso")
else
   str_mega=0
end if

'response.Write(request("selSubModulo"))

if request("selSubModulo")  <> "" then
   str_CdSubModulo = request("selSubModulo") 
else
   if request("txtSubModulo") <> "" then
      str_CdSubModulo = request("txtSubModulo")    
   else
      str_CdSubModulo = "0"
   end if
end if   
'response.Write("<p>" & str_SubModulo)

str_OPT = request("pOPT") 

str_SQL_MegaProc = ""
str_SQL_MegaProc = str_SQL_MegaProc & " SELECT DISTINCT "
str_SQL_MegaProc = str_SQL_MegaProc & " " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_CD_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " , " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " FROM " & Session("PREFIXO") & "MEGA_PROCESSO "
str_SQL_MegaProc = str_SQL_MegaProc & " order by " & Session("PREFIXO") & "MEGA_PROCESSO.MEPR_TX_DESC_MEGA_PROCESSO "
set rs=db.execute(str_SQL_MegaProc)

if str_mega<>0 then
	if str_CdSubModulo <> "" and str_CdSubModulo <> "0"  then
	   'str_Sql_SubModulo = " and SUMO_NR_SEQUENCIA = " &  str_CdSubModulo
	   str_Sql_SubModulo = " and FUNCAO_NEGOCIO_SUB_MODULO.SUMO_NR_CD_SEQUENCIA = " &  str_CdSubModulo
	else
	   str_Sql_SubModulo = " "
	end if   
	
	if str_Sql_SubModulo = " " then
		ssql=""
		ssql="SELECT DISTINCT FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
		ssql=ssql+"FROM FUNCAO_NEGOCIO "
		ssql=ssql+"WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega
		ssql=ssql+"ORDER BY FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	else
		ssql=""
		ssql="SELECT DISTINCT FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO, FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
		ssql=ssql+"FROM FUNCAO_NEGOCIO_SUB_MODULO "
		ssql=ssql+"INNER JOIN FUNCAO_NEGOCIO ON "
		ssql=ssql+"FUNCAO_NEGOCIO_SUB_MODULO.FUNE_CD_FUNCAO_NEGOCIO= FUNCAO_NEGOCIO.FUNE_CD_FUNCAO_NEGOCIO "
		ssql=ssql+"WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " " & str_Sql_SubModulo
		ssql=ssql+"ORDER BY FUNCAO_NEGOCIO.FUNE_TX_TITULO_FUNCAO_NEGOCIO "
	end if
	
	set rs1=db.execute(ssql)
	
else
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=0 ORDER BY FUNE_CD_FUNCAO_NEGOCIO")
	str_mega=0
end if

str_Sub_Modulo = ""
str_Sub_Modulo = str_Sub_Modulo & " SELECT DISTINCT "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_TX_DESC_SUB_MODULO, "
str_Sub_Modulo = str_Sub_Modulo & " SUMO_NR_CD_SEQUENCIA"
str_Sub_Modulo = str_Sub_Modulo & " FROM " & Session("PREFIXO") & "SUB_MODULO"
str_Sub_Modulo = str_Sub_Modulo + " WHERE MEPR_CD_MEGA_PROCESSO_TODOS LIKE '%" & Right("00" & str_mega,2) & "%'" 
str_Sub_Modulo = str_Sub_Modulo & " order by SUMO_TX_DESC_SUB_MODULO "
'response.write str_Sub_Modulo
set rs_SubModulo=db.execute(str_Sub_Modulo)

set rst_Area_Abrangencia=db.execute("SELECT * FROM " & Session("PREFIXO") & "ORGAO_AGLUTINADOR ORDER BY AGLU_SG_AGLUTINADO")

%>
<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
</head>

<script>
function manda()
{
//alert('_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value)
//+'&selSubModulo='+
//alert(document.frm1.selSubModulo.value)

//document.frm1.txtSubModulo.value = document.frm1.selSubModulo.value
//alert(document.frm1.txtSubModulo.value)
//window.location.href='rel_geral_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+'&txtSubModulo='+document.frm1.txtSubModulo.value
window.location.href='rel_geral_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&pOPT='+document.frm1.txtOPT.value+"&chkEmUso="+document.frm1.chkEmUso.value+"&chkEmDesuso="+document.frm1.chkEmDesuso.value

}
function manda1()
{
//alert(document.frm1.selSubModulo.value);
window.location.href='rel_geral_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selSubModulo='+document.frm1.selSubModulo.value+'&pOPT='+document.frm1.txtOPT.value+"&chkEmUso="+document.frm1.chkEmUso.checked+"&chkEmDesuso="+document.frm1.chkEmDesuso.checked
}

function Confirma()
{
   //if(document.frm1.selMegaProcesso.selectedIndex == 0)
      {
      //alert("É obrigatória a seleção de um MEGA-PROCESSO!");
      //document.frm1.selMegaProcesso.focus();
      //return;
      }
  		//else
        {
		document.frm1.txtDescModulo.value = document.frm1.selSubModulo.options[document.frm1.selSubModulo.selectedIndex].value
        document.frm1.submit();
        if(document.frm1.txtOPT.value == 1)
           {
           document.frm1.action="gera_rel_geral_funcao.asp";
           //document.frm1.target="corpo";
           document.frm1.submit();
           }
         if(document.frm1.txtOPT.value == 2)
           {
           document.frm1.action="gera_rel_geral_funcao_colunada.asp";
           //document.frm1.target="corpo";
           document.frm1.submit();
           }				
        }		
     }
    
</script>
<body topmargin="0" leftmargin="0">
<form method="POST" action="gera_rel_geral_funcao.asp" name="frm1">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
  <tr>
    <td width="20%" height="20">&nbsp;</td>
    <td width="44%" height="60">&nbsp;</td>
    <td width="36%" valign="top"> 
      <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/sinergia_total/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
            <div align="center"><a href="../../indexA.asp"><img src="../../imagens/home.gif" width="19" height="20" border="0"></a>&nbsp;</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor="#00FF99">
    <td colspan="3" height="20">
      <table width="625" border="0" align="center">
        <tr>
            <td width="26"><a href="javascript:Confirma()"><img border="0" src="../../imagens/confirma_f02.gif"></a></td>
          <td width="26"><b><font face="Verdana" size="2" color="#330099">Enviar</font></b></td>
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
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%"> 
        <div align="center"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#330099">Relatório
          Geral de Função</font></div>
      </td>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="25%">&nbsp;</td>
      <td width="55%">&nbsp;</td>
      <td width="20%">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="829" height="322">
    <tr> 
      <td width="66" height="37"> <% If str_mega <> 11 and str_mega <> 10 then %> <input type="hidden" name="selSubModulo22" value="0"> <% end if %> </td>
      <td width="115" height="37"> <div align="right"><b><font face="Verdana" color="#330099" size="2">Mega-Processo 
          : </font></b></div></td>
      <td height="37" width="634"> <select size="1" name="selMegaProcesso" onChange="javascript:manda()">
          <option value="0">== TODOS ==</option>
          <%do until rs.eof=true
         if trim(str_mega)=trim(rs("MEPR_CD_MEGA_PROCESSO")) then
                	%>
          <option selected value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%else%>
          <option value="<%=rs("MEPR_CD_MEGA_PROCESSO")%>"><%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
          <%
					end if
					rs.movenext
					loop
					%>
        </select> <% 'if InStrRev("11/10", Right("00" & str_mega, 2)) = 0 then %> <input type="hidden" name="txtDescModulo" value=""> <% 'end if %> </td>
    </tr>
    <% 
	   'if InStrRev("11/10", Right("00" & str_mega, 2)) <> 0 then
	%>
    <tr> 
      <td width="66" height="25">&nbsp;</td>
      <td width="115" height="25"> <div align="right"><font face="Verdana" size="2" color="#330099"><b>Assunto 
          : </b></font></div></td>
      <td width="634" height="25"> <select size="1" name="selSubModulo" onChange="javascript:manda1()">
          <option value="0">== Selecione o Assunto ==</option>
          <%do until rs_SubModulo.eof=true
		  if Trim(str_CdSubModulo)= Trim(rs_SubModulo("SUMO_NR_CD_SEQUENCIA")) then
		  %>
          <option selected value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <% else %>
          <option value="<%=rs_SubModulo("SUMO_NR_CD_SEQUENCIA")%>"><%=rs_SubModulo("SUMO_TX_DESC_SUB_MODULO")%></option>
          <%
		     end if
					rs_SubModulo.movenext
					loop
					%>
        </select> </td>
    </tr>
    <% 'end if %>
    <tr> 
      <td width="66" height="19"></td>
      <td><div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Uso 
          : </font></b></font></div></td>
      <td height="19"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        uso </font></b></font> <font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmUso" type="checkbox" value="1" checked>
        </b></font><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099">Em 
        desuso </font></b></font><font face="Verdana" size="2" color="#330099"><b> 
        <input name="chkEmDesuso" type="checkbox" value="1">
        </b></font></td>
    </tr>
    <tr> 
      <td width="66" height="15"></td>
      <td width="115" height="15"> <p align="right"> 
          <input type="checkbox" name="selG" value="1">
      </td>
      <td height="15" width="634"> <b><font face="Verdana" color="#330099" size="1">Listar 
        Somente Funções Genéricas</font></b> </td>
    </tr>
    <tr> 
      <td height="19"></td>
      <td height="15"> <p align="right"> 
          <input type="checkbox" name="chkCritica" value="1">
      </td>
      <td height="15"> <b><font face="Verdana" color="#330099" size="1">Listar 
        Somente Funções Cr&iacute;tica</font></b> </td>
    </tr>
    <tr> 
      <td height="19"></td>
      <td><div align="right"><strong><font color="#330099" size="2" face="Verdana, Arial, Helvetica, sans-serif">&Aacute;rea 
          de abrangência:</font></strong></div></td>
      <td height="41"><select name="selAreaAbrangencia">
          <option value="0">== Selecione a Área de Abrangência ==</option>
          <%do until rst_Area_Abrangencia.eof=true%>
          <option value="<%=rst_Area_Abrangencia("AGLU_CD_AGLUTINADO")%>"><%=rst_Area_Abrangencia("AGLU_SG_AGLUTINADO")%></option>
          <%
           			rst_Area_Abrangencia.movenext
           			loop
					rst_Area_Abrangencia.close
                    %>
        </select></td>
    </tr>
    <tr> 
      <td height="19"></td>
      <td> <div align="right"><b><font face="Verdana" color="#330099" size="2">Fun&ccedil;&atilde;o 
          R/3 : </font></b></div></td>
      <td height="41"> <select size="1" name="selFuncao">
          <option value="0">== Selecione a Fun&ccedil;&atilde;o R/3 ==</option>
          <%do until rs1.eof=true%>
          <option value="<%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rs1("FUNE_CD_FUNCAO_NEGOCIO")%> - <%=rs1("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
          <%
					rs1.movenext
					loop
					%>
        </select> </td>
    </tr>
    <tr> 
      <td width="66" height="19"></td>
      <td width="115" height="19"> <div align="right"> 
          <input type="hidden" name="txtOPT" value="<%=str_OPT%>">
        </div></td>
      <td height="19" width="634">&nbsp; </td>
    </tr>
  </table>
  </form>

<p>&nbsp;</p>

</body>

</html>
<%
rs.close
set rs = nothing
rs1.close
set rs1 = nothing
rs_SubModulo.close
set rs_SubModulo = nothing
db.close
set db = nothing

%>