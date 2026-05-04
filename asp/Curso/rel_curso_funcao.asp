<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

processo=0

mega=request("mega")
curso=request("curso")

set rs_curso=db.execute("SELECT * FROM " & Session("PREFIXO") & "CURSO ORDER BY CURS_CD_CURSO")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")

ssql=""
ssql="SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO, FUNE_TX_TITULO_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE MEPR_CD_MEGA_PROCESSO=" & mega & " ORDER BY FUNE_TX_TITULO_FUNCAO_NEGOCIO"

set rsfunc=db.execute(ssql)

ssql=""
ssql="SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "CURSO_FUNCAO WHERE CURS_CD_CURSO='" & curso & "'"

ssql_tem="SELECT DISTINCT FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "CURSO_FUNCAO "

ssql=ssql+"  ORDER BY FUNE_CD_FUNCAO_NEGOCIO"
'RESPONSE.Write(ssql)
'RESPONSE.End()
set rsfunccurso=db.execute(ssql)

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
window.location.href="rel_curso_funcao.asp?mega="+document.frm1.selMega.value+"&curso="+document.frm1.selCurso.value
}

function envia2()
{
window.location.href="rel_curso_funcao.asp?mega=0&curso="+document.frm1.selCurso.value

//onLoad="alert('teste');history.back(-1)"

}

</script>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="valida_rel_curso_funcao.asp" name="frm1">
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
          Curso x Fun&ccedil;&atilde;o R/3</font></div>
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

<p style="margin: 0" align="center">&nbsp;<font face="Verdana" size="2" color="#330099"><b>Funções
de Negócio</b></font></p>

        <table border="0" width="964" height="142">
          <tr>
            <td width="300" height="138" rowspan="5"></td>
            
      <td width="300" height="138" rowspan="5"> <p style="margin: 0"> 
        <table width="75%" border="0">
          <tr>
            <td><font face="Verdana" size="2" color="#330099"><b>N&atilde;o selecionadas</b></font></td>
          </tr>
          <tr>
            <td><select size="7" name="list1" multiple>
                <%do until rsfunc.eof=true
               set rstem=db.execute("SELECT * FROM CURSO_FUNCAO where CURS_CD_CURSO='" & curso & "' AND FUNE_CD_FUNCAO_NEGOCIO='" & rsfunc("FUNE_CD_FUNCAO_NEGOCIO") & "'")
               if rstem.eof=true then
               %>
                <option value="<%=rsfunc("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rsfunc("FUNE_CD_FUNCAO_NEGOCIO")%> 
                - <%=rsfunc("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></option>
                <%
				  end if
				  RSfunc.MOVENEXT               
                LOOP
               %>
              </select></td>
          </tr>
        </table></td>
            <td width="117" height="28" align="center">
              <p style="margin: 0"></td>
            
      <td width="526" height="138" rowspan="5"> <p style="margin: 0"> 
        <table width="75%" border="0">
          <tr>
            <td><font face="Verdana" size="2" color="#330099"><b>Selecionadas</b></font></td>
          </tr>
          <tr>
            <td><select size="7" name="list2" multiple>
                <%
				'response.Write("<p> eof " & rsfunccurso.eof)
				do until rsfunccurso.eof=true
               'str_Sql = " SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rsfunccurso("FUNE_CD_FUNCAO_NEGOCIO") & "'"
			   'response.Write("<p>" & str_Sql)
			   set rstemp=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & rsfunccurso("FUNE_CD_FUNCAO_NEGOCIO") & "'")
			   
               if rstemp.eof=false then
               		VALOR_FUNC2=rstemp("FUNE_TX_TITULO_FUNCAO_NEGOCIO")
               %>
                <option value="<%=rsfunccurso("FUNE_CD_FUNCAO_NEGOCIO")%>"><%=rsfunccurso("FUNE_CD_FUNCAO_NEGOCIO")%> 
                - <%=VALOR_func2%></option>
                <%
                end if
               RSfuncCURSO.MOVENEXT               
               LOOP
               %>
              </select></td>
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