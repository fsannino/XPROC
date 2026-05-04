 

<%
set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

str_funcao=request("selFuncao")
str_mega=request("selMegaProcesso")
str_proc=request("selProcesso")
str_sub=request("selSubProcesso")
str_ativ=request("selAtividade")

if len(str_mega)=0 then
	str_mega=0
end if

if len(str_proc)=0 then
	str_proc=0
end if

if len(str_sub)=0 then
	str_sub=0
end if

if len(str_ativ)=0 then
	str_ativ=0
end if

set rs_func=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUNCAO_NEGOCIO WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

set rs=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs_func("MEPR_CD_MEGA_PROCESSO"))

'set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TIPO_QUALIFICACAO ORDER BY TPQU_TX_DESC_TIPO_QUALIFICACAO")

set rs3=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TP_QUA WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

set rs4=db.execute("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_ORG_AGLU WHERE FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")

set rs_mega=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO")
set rs_proc=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " ORDER BY PROC_TX_DESC_PROCESSO")
set rs_sub=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc & " ORDER BY SUPR_TX_DESC_SUB_PROCESSO")
set rs_ativ=db.execute("SELECT DISTINCT ATCA_CD_ATIVIDADE_CARGA FROM " & Session("PREFIXO") & "RELACAO_FINAL WHERE MEPR_CD_MEGA_PROCESSO=" & str_mega & " AND PROC_CD_PROCESSO=" & str_proc & " AND SUPR_CD_SUB_PROCESSO=" & str_sub & " ORDER BY ATCA_CD_ATIVIDADE_CARGA")

SSQL=""
SSQL="SELECT DISTINCT MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO,ATCA_CD_ATIVIDADE_CARGA,FUNE_CD_FUNCAO_NEGOCIO FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO"

if str_mega<>0 then
	compl=compl+"MEPR_CD_MEGA_PROCESSO=" & STR_MEGA & " AND "
end if

if str_proc<>0 then
	compl=compl+"PROC_CD_PROCESSO=" & STR_PROC & " AND "
end if

if str_sub<>0 then
	compl=compl+"SUPR_CD_SUB_PROCESSO=" & STR_SUB & " AND "
end if

if str_ativ<>0 then
	compl=compl+"ATCA_CD_ATIVIDADE_CARGA=" & str_ativ & " AND "
end if

tamanho=len(compl)
tamanho=tamanho-5
compl=" WHERE " & LEFT(COMPL,TAMANHO) & " AND FUNE_CD_FUNCAO_NEGOCIO='"& str_funcao &"'"

SSQL=SSQL+COMPL+ " ORDER BY MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO,ATCA_CD_ATIVIDADE_CARGA"

ssql2=ssql

set rs10=db.execute(ssql2)
%>

<html>
<head>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>

<script>
function manda1()
{
window.location.href='exibe_dados_funcao.asp?selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncao='+document.frm1.selFuncao.value
}

function manda2()
{
window.location.href='exibe_dados_funcao.asp?selProcesso='+document.frm1.selProcesso.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncao='+document.frm1.selFuncao.value
}

function manda3()
{
window.location.href='exibe_dados_funcao.asp?selSubProcesso='+document.frm1.selSubProcesso.value+'&selProcesso='+document.frm1.selProcesso.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncao='+document.frm1.selFuncao.value
}

function manda4()
{
window.location.href='exibe_dados_funcao.asp?selAtividade='+document.frm1.selAtividade.value+'&selSubProcesso='+document.frm1.selSubProcesso.value+'&selProcesso='+document.frm1.selProcesso.value+'&selMegaProcesso='+document.frm1.selMegaProcesso.value+'&selFuncao='+document.frm1.selFuncao.value
}
</script>

</head>

<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
<form method="POST" action="../Funcao/valida_altera_funcao.asp" name="frm1">
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
            <div align="center"><a href="../../indexA.asp"><img src="../Funcao/home.gif" border="0"></a>&nbsp;</div>
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
            <td width="27"></td>  <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
        
  <table width="810" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="808">
        &nbsp;
        <div align="center"><font face="Verdana" color="#330099" size="3">Dados de Fun&ccedil;&atilde;o R/3 - <b><%=str_funcao%></b></font></div>
      </td>
    </tr>
    <tr>
      <td width="808">&nbsp;</td>
    </tr>
  </table>
  <table border="0" width="767" height="42">
          <tr>
            
      <td width="197" height="25"><input type="hidden" name="selFuncao" size="20" value="<%=str_funcao%>"></td>
            
      <td width="205" height="25" valign="top"><b><font face="Verdana" size="1" color="#330099">Mega-Processo</font></b></td>
            
      <td width="228" height="25" valign="top"> 
<font face="Verdana" size="1" color="#330099"> 
<%=rs("MEPR_TX_DESC_MEGA_PROCESSO")%></font></td>
            
      <td width="165" height="25" valign="top"> 
        <p align="left"><font face="Verdana" size="1" color="#330099">
        <%
        if rs_func("FUNE_TX_TP_FUN_NEG")="G" then%>
        <input type="hidden" name="selGenerica" value="1" checked>
        <%else%>
        <input type="hidden" name="selGenerica" value="0">
        <%end if%>
        <%if rs_func("FUNE_TX_TP_FUN_NEG")="G" then%> <b>Função Genérica </b>
        <%end if
        %></font></td>
          </tr>
          <tr>
            
      <td width="197" height="13"></td>
            
      <td width="205" height="13" valign="top"><b><font face="Verdana" size="1" color="#330099">Fun&ccedil;&atilde;o R/3</font></b></td>
            
      <td width="379" height="13" colspan="2" valign="top"> 
<font face="Verdana" size="1" color="#330099"> 
<%=rs_func("FUNE_TX_TITULO_FUNCAO_NEGOCIO")%></font></td>
          </tr>
          <tr>
            
      <td width="197" height="31"></td>
            
      <td width="205" height="31" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" height="34">
          <tr> 
            <td height="16"><b><font face="Verdana" size="1" color="#330099">Descrição da</font></b></td>
          </tr>
          <tr> 
            <td height="18"><b><font face="Verdana" size="1" color="#330099">Fun&ccedil;&atilde;o R/3</font></b></td>
          </tr>
        </table>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="margin-top: 0; margin-bottom: 0"><font size="1"><input type="hidden" name="Funcao" size="20" value="<%=str_funcao%>"><input type="hidden" name="txtQua" size="20"><input type="hidden" name="txtpub" size="20"></font></p>
            </td>
            
      <td width="379" height="31" valign="top" colspan="2"> 
        <p align="left" style="margin-top: 0; margin-bottom: 0">
          <font face="Verdana" size="1" color="#330099">
          <%=rs_func("FUNE_TX_DESC_FUNCAO_NEGOCIO")%>
          </font>
        <p align="left" style="margin-top: 0; margin-bottom: 0">
      </td>
          </tr>
          <tr>
            
      <td width="197" height="25" valign="top">
        <p style="margin-top: 0; margin-bottom: 0"></td>
            
      <td width="205" height="25" valign="top"> 
        <p style="margin-top: 0; margin-bottom: 0"> 
        <b><font face="Verdana" size="1" color="#330099">Qualificação 
          Não R/3</font></b>
        </p>
            </td>
            
            <%
				do until rs3.eof=true
				
				set temp_=db.execute("select * from " & Session("PREFIXO") & "TIPO_QUALIFICACAO where TPQU_CD_TIPO_QUALIFICACAO=" & rs3("TPQU_CD_TIPO_QUALIFICACAO"))            
				
				valor_QUA=valor_QUA & temp_("TPQU_TX_DESC_TIPO_QUALIFICACAO") & " / "
				rs3.movenext
            	loop
            	
            	tamanho=len(valor_Qua)
            	tamanho=tamanho-3
            	
            	on error resume next
            	valor=left(valor_QUA,tamanho)
            	
            %>

      <td width="379" height="25" valign="top" colspan="2"> 
       <p style="margin-top: 0; margin-bottom: 0"> 
       <font face="Verdana" size="1" color="#330099"><%=VALOR%></font>
      </td>
          </tr>
          <tr>
            
      <td width="197" height="25" valign="top">
      </td>
            
      <td width="205" height="25" valign="top"> 
        <font face="Verdana" size="1" color="#330099"><b>Abrangência de Impacto</b></font>
            </td>
            <%
				do until rs4.eof=true
				
				set temp_=db.execute("select * from " & Session("PREFIXO") & "ORGAO_AGLUTINADOR where AGLU_CD_AGLUTINADO=" & rs4("AGLU_CD_AGLUTINADO"))            
				
				valor_imp=valor_imp & temp_("AGLU_SG_AGLUTINADO") & " / "
				rs4.movenext
            	loop
            	
            	tamanho=len(valor_imp)
            	tamanho=tamanho-3
            	
            	on error resume next
            	valor_imp=left(valor_imp,tamanho)
            	
            %>
      <td width="379" height="25" valign="top" colspan="2"> 
       <font face="Verdana" size="1" color="#330099"><%=valor_imp%></font>
      </td>
          </tr>
        </table>
<p align="center" style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
<div align="center">
  <center>
  <table border="0" width="62%" cellspacing="0" height="81">
    <tr>
      <td width="35%" align="right" height="15">
        <p style="margin-top: 0; margin-bottom: 0"><b><font face="Verdana" size="1" color="#330099">Mega
        Processo :&nbsp;&nbsp; </font></b></td>
      <td width="65%" height="15">
        <select size="1" name="selMegaProcesso" onchange="javascript:manda1()">
        
        <option value="0">== Selecione o Mega-Processo ==</option>
        <%
        do until rs_mega.eof=true
        if trim(str_mega)=trim(rs_mega("mepr_cd_mega_processo")) then
        %>
        <option selected value="<%=rs_mega("mepr_cd_mega_processo")%>"><%=rs_mega("mepr_tx_desc_mega_processo")%></option>
        <%else%>
        <option value="<%=rs_mega("mepr_cd_mega_processo")%>"><%=rs_mega("mepr_tx_desc_mega_processo")%></option>
        <%
        end if
        rs_mega.movenext
        loop
        %>
        </select></td>
    </tr>
    <tr>
      <td width="35%" align="right" height="16"><b><font face="Verdana" size="1" color="#330099">Processo
        :&nbsp;&nbsp; </font></b></td>
      <td width="65%" height="16"><select size="1" name="selProcesso" onchange="javascript:manda2()">
        <option value="0">== Selecione o Processo ==</option>
        <%
        do until rs_proc.eof=true
        if trim(str_proc)=trim(rs_proc("proc_cd_processo")) then
        %>
        <option selected value="<%=rs_proc("proc_cd_processo")%>"><%=rs_proc("proc_tx_desc_processo")%></option>
        <%else%>
        <option value="<%=rs_proc("proc_cd_processo")%>"><%=rs_proc("proc_tx_desc_processo")%></option>
        <%
        end if
        rs_proc.movenext
        loop
        %>
        </select></td>
    </tr>
    <tr>
      <td width="35%" align="right" height="17"><b><font face="Verdana" size="1" color="#330099">Sub-Processo
        :&nbsp;&nbsp; </font></b></td>
      <td width="65%" height="17"><select size="1" name="selSubProcesso" onchange="javascript:manda3()">
        <option value="0">== Selecione o Sub-Processo ==</option>
        <%
        do until rs_sub.eof=true
        if trim(str_sub)=trim(rs_sub("supr_cd_sub_processo")) then
        %>
        <option selected value="<%=rs_sub("supr_cd_sub_processo")%>"><%=rs_sub("supr_tx_desc_sub_processo")%></option>
        <%else%>
        <option value="<%=rs_sub("supr_cd_sub_processo")%>"><%=rs_sub("supr_tx_desc_sub_processo")%></option>
        <%
        end if
        rs_sub.movenext
        loop
        %>
        </select></td>
    </tr>
    <tr>
      <td width="35%" align="right" height="17"><b><font face="Verdana" size="1" color="#330099">Atividade
        :&nbsp;&nbsp; </font></b></td>
      <td width="65%" height="17"><select size="1" name="selAtividade" onchange="javascript:manda4()">
        <option value="0">== Selecione a Atividade ==</option>
        <%
        do until rs_ativ.eof=true
        set temp=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & rs_ativ("atca_cd_atividade_carga"))
        valor_ativ=temp("ATCA_TX_DESC_ATIVIDADE")
        if trim(str_ativ)=trim(rs_ativ("atca_cd_atividade_carga")) then
        %>
        <option selected value="<%=rs_ativ("atca_cd_atividade_carga")%>"><%=valor_ativ%></option>
        <%else%>
        <option value="<%=rs_ativ("atca_cd_atividade_carga")%>"><%=valor_ativ%></option>
        <%
        end if
        rs_ativ.movenext
        loop
        %>
        </select></td>
    </tr>
  </table>
  </center>
</div>
<p align="center"><b><font face="Verdana" color="#330099" size="2">Transações
Relacionadas</font></b></p>
<div align="center">
  <center>
  <%DO UNTIL RS10.EOF=TRUE
  SET RS_TRANS=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "FUN_NEG_TRANSACAO WHERE MEPR_CD_MEGA_PROCESSO=" & rs10("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & rs10("PROC_CD_PROCESSO") & " AND SUPR_CD_SUB_PROCESSO=" & rs10("SUPR_CD_SUB_PROCESSO") & " AND ATCA_CD_ATIVIDADE_CARGA=" & rs10("ATCA_CD_ATIVIDADE_CARGA") & " AND FUNE_CD_FUNCAO_NEGOCIO='" & str_funcao & "'")
  tem=0
  IF RS_TRANS.EOF=FALSE THEN
  tem=1
  %>
  <table border="0" width="44%" height="61" cellspacing="0" cellpadding="0">
    <tr>
      <td width="35%" align="right" height="9">
        <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#330099">Mega
        Processo :&nbsp;&nbsp; </font></td>
        <%
			SET RS_TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS10("MEPR_CD_MEGA_PROCESSO"))        
        	VALOR_=RS_TEMP("MEPR_TX_DESC_MEGA_PROCESSO")
        %>
      <td width="67%" height="9"><font face="Verdana" size="1" color="#330099"><b><%=RS10("MEPR_CD_MEGA_PROCESSO")%> - <%=VALOR_%></b></font></td>
    </tr>
    <tr>
      <td width="35%" align="right" height="11"><font face="Verdana" size="1" color="#330099">Processo
        :&nbsp;&nbsp; </font></td>
        <%
			SET RS_TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS10("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & RS10("PROC_CD_PROCESSO"))        
        	VALOR_=RS_TEMP("PROC_TX_DESC_PROCESSO")
        %>

      <td width="67%" height="11"><font face="Verdana" size="1" color="#330099"><b><%=RS10("PROC_CD_PROCESSO")%> - <%=VALOR_%></b></font></td>
    </tr>
    <tr>
      <td width="35%" align="right" height="10"><font face="Verdana" size="1" color="#330099">Sub-Processo
        :&nbsp;&nbsp; </font></td>
        <%
			SET RS_TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & RS10("MEPR_CD_MEGA_PROCESSO") & " AND PROC_CD_PROCESSO=" & RS10("PROC_CD_PROCESSO") & " AND SUPR_CD_SUB_PROCESSO=" & RS10("SUPR_CD_SUB_PROCESSO"))        
        	VALOR_=RS_TEMP("SUPR_TX_DESC_SUB_PROCESSO")
        %>

      <td width="67%" height="10"><font face="Verdana" size="1" color="#330099"><b><%=RS10("SUPR_CD_SUB_PROCESSO")%> - <%=VALOR_%></b></font></td>
    </tr>
    <tr>
      <td width="35%" align="right" height="7"><font face="Verdana" size="1" color="#330099">Atividade
        :&nbsp;&nbsp; </font></td>
      <td width="67%" height="7">
       <%
			SET RS_TEMP=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & RS10("ATCA_CD_ATIVIDADE_CARGA"))        
        	VALOR_=RS_TEMP("ATCA_TX_DESC_ATIVIDADE")
        %>
        <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#330099"><b><%=RS10("ATCA_CD_ATIVIDADE_CARGA")%> - <%=VALOR_%></b></font></td>
    </tr>
  </table>
  </center>
</div>
<p align="center" style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <div align="center"> 
    <table border="0" width="73%" cellspacing="0" cellpadding="0">
   	 <%
   	 DO UNTIL RS_TRANS.EOF=TRUE
   	 IF COR="#DDDDDD" THEN
   	 	COR="WHITE"
   	 ELSE
   	 	COR="#DDDDDD"
   	 END IF
   	 %>	
  
	 <tr>
     <td width="30%" bgcolor="<%=COR%>">
        <p style="margin-top: 0; margin-bottom: 0" align="right"><font face="Verdana" size="1" color="#330099"><b><%=RS_TRANS("TRAN_CD_TRANSACAO")%> &nbsp;</b></font></td>
  
      <td width="70%" bgcolor="<%=COR%>">
      <%
      SET RS20=DB.EXECUTE("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & RS_TRANS("TRAN_CD_TRANSACAO") & "'")
      VALOR_TRANS=RS20("TRAN_TX_DESC_TRANSACAO")
      %>
          <p style="margin-top: 0; margin-bottom: 0"><font face="Verdana" size="1" color="#330099">- 
            <%=VALOR_TRANS%> 
			<%
			str_SQl = ""
			str_SQL = str_SQl & " Select " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.MCPT_NR_SITUACAO_ALTERACAO "
			str_SQL = str_SQl & " from " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO "
			str_SQL = str_SQl & " where " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.MEPR_CD_MEGA_PROCESSO = " & RS10("MEPR_CD_MEGA_PROCESSO")
			str_SQL = str_SQl & " AND " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.PROC_CD_PROCESSO = " & RS10("PROC_CD_PROCESSO")
			str_SQL = str_SQl & " AND " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.SUPR_CD_SUB_PROCESSO = " & RS10("SUPR_CD_SUB_PROCESSO")
			str_SQL = str_SQl & " AND " & Session("PREFIXO") & "MACRO_PERFIL_TRANSACAO.TRAN_CD_TRANSACAO = " & RS_TRANS("TRAN_CD_TRANSACAO")
			SET rsJaEditado=DB.EXECUTE(str_SQl)
			
			if rsJaEditado("MCPT_TX_SITUACAO_ATUA") = 0 then %>
			   <img src="../../imagens/func_tran_nao_marcada.gif" width="16" height="16"> 
			<% else %>
               <img src="../../imagens/func_tran_marcada.gif" width="16" height="16">
			<% end if 
			rsJaEditado.close
			set rsJaEditado = Nothing
			%>
			</font>
			
        </td>
    </tr>
    <%
    RS_TRANS.MOVENEXT
    LOOP
    %>
  </table>
  <P>
  <%
  END IF
  RS10.MOVENEXT
  LOOP
  %>
  </div>
  <%if tem=0 then%>
  <p align="center"><b><font face="Verdana" size="2" color="#800000">&nbsp;Não
  existem Transações para a Seleção</font></b></p>
  <%end if%>
  </form>

</body>

</html>