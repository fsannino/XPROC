<<<<<<< HEAD
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
dim int_MaxProcesso
dim int_MaxSubProcesso
dim int_MaxAtividade
dim str_MensagemMegaProc
dim str_MensagemProc
dim str_MensagemSubProc
dim str_Msg 

str_MensagemMegaProc = ""
str_MensagemProc = ""
str_MensagemSubProc = ""

int_MaxProcesso = 0
int_MaxSubProcesso = 0
int_MaxAtividade = 0

int_MegaProcesso= Request.Form("selMegaProcesso")
int_Processo = Request.Form("SelProcesso")
str_NovoProcesso = UCase(Request.Form("txtNovoProcesso"))
int_SubProcesso = Request.Form("SelSubProcesso")
str_NovoSubProcesso = UCase(Request.Form("txtNovoSubProcesso"))
int_Atividade = Request.Form("selAtividade")
str_NovaAtividade = UCase(Request.Form("txtNovaAtividade"))

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

'************GRAVA NOVO PROCESSO *************************
if str_NovoProcesso <> "" then
   if int_MegaProcesso <> 0 then
	str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " SELECT "
	str_SQL_Proc = str_SQL_Proc & " MAX(PROC_CD_PROCESSO) AS MAX_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " FROM " & Session("PREFIXO") & "PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & int_MegaProcesso
	Set rdsMaxProcesso = Conn_db.Execute(str_SQL_Proc)
	if rdsMaxProcesso.EOF then
	   int_MaxProcesso = 1	
	else
	   int_MaxProcesso = rdsMaxProcesso("MAX_PROCESSO") + 1	
	end if
	rdsMaxProcesso.Close
	set rdsMaxProcesso = Nothing
    str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " INSERT INTO " & Session("PREFIXO") & "PROCESSO ( "
    str_SQL_Proc = str_SQL_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_TX_DESC_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Proc = str_SQL_Proc & " ) Values( "
	str_SQL_Proc = str_SQL_Proc & int_MegaProcesso & "," & int_MaxProcesso & ","
	str_SQL_Proc = str_SQL_Proc & "'" & str_NovoProcesso & "', 'I', 'XXXX', GETDATE())" 
	Set rdsNovoProcesso = Conn_db.Execute(str_SQL_Proc)
   strChave = CStr(int_MegaProcesso) & CStr(int_MaxProcesso) ' &  CStr(strP) & CStr(strSP) & CStr(strEU)
   'call grava_log(strChave,"PROCESSO","I",0)	
	int_Processo = int_MaxProcesso
  else
    str_SemMegaProcesso = 0
	str_MensagemMegaProc = "Para cadastrar um novo Processo deve ser selecionado um Megaprocesso"
  end if	
end if
'********** GRAVA NOVO SUB PROCESSO ****************
if str_NovoSubProcesso <> "" then
 if int_MegaProcesso <> 0 then
  if int_Processo <> 0 or int_MaxProcesso <> 0  then
    if int_MaxProcesso <> 0 then
	   int_Processo = int_MaxProcesso
	end if
	str_SQL_Sub_Proc = ""
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " SELECT "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MAX(SUPR_CD_SUB_PROCESSO) AS MAXIMO_SUB "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " FROM " & Session("PREFIXO") & "SUB_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO, "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & int_MegaProcesso
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND PROC_CD_PROCESSO = " & int_Processo
	a = str_SQL_Sub_Proc
	Set rdsMaxSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
    'if IsNull(rdsMaxSubProcesso("MAXIMO_SUB")) then
    if rdsMaxSubProcesso.EOF then
	   int_MaxSubProcesso = 1	
    else
	   'A = rdsMaxSubProcesso("MAXIMO_SUB") 
   	   int_MaxSubProcesso = rdsMaxSubProcesso("MAXIMO_SUB") + 1	
	   'int_MaxSubProcesso = 10
    end if
	rdsMaxSubProcesso.Close
	set rdsMaxSubProcesso = Nothing
    str_SQL_Sub_Proc = ""
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " INSERT INTO " & Session("PREFIXO") & "SUB_PROCESSO ( "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_CD_SUB_PROCESSO "	
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_TX_DESC_SUB_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ) Values( "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & int_MegaProcesso & "," & int_Processo & "," & int_MaxSubProcesso & ","
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & "'" & str_NovoSubProcesso & "', 'I', 'XXXX', GETDATE())" 
	Set rdsNovoSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)

    strChave = CStr(int_MegaProcesso) & CStr(int_MaxProcesso) & CStr(int_MaxSubProcesso) '& CStr(strSP) & CStr(strEU)
    'call grava_log(strChave,"SUB_PROCESSO","I",0)	
			
	int_SubProcesso = int_MaxSubProcesso
   else
    str_SemProcesso = 0
    str_MensagemProc = "Para cadastrar um novo Sub Processo deve ser selecionado um Processo ou preencha um novo Processo"  
   end if	
 else
  str_SemMegaProcesso = 0
  str_MensagemMegaProc = "Para cadastrar um novo Sub Processo deve ser selecionado um Megaprocesso"  
 end if   
end if
'************GRAVA NOVA ATIVIDADE *************************
if str_NovaAtividade <> "" then
 if int_MegaProcesso <> 0 then
  if int_Processo <> 0 or int_MaxProcesso <> 0  then
    if int_MaxProcesso <> 0 then
	   int_Processo = int_MaxProcesso
	end if
   if int_SubProcesso <> 0 or int_MaxSubProcesso <> 0  then
      if int_MaxSubProcesso <> 0 then
	    int_SubProcesso = int_MaxSubProcesso
      end if
    str_SQL_Atividade = ""
    str_SQL_Atividade = str_SQL_Atividade & " SELECT "
    str_SQL_Atividade = str_SQL_Atividade & " MAX(ATIV_CD_ATIVIDADE) AS MAX_ATIVIDADE "
    str_SQL_Atividade = str_SQL_Atividade & " FROM " & Session("PREFIXO") & "ATIVIDADE "
    str_SQL_Atividade = str_SQL_Atividade & " GROUP BY MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, "
    str_SQL_Atividade = str_SQL_Atividade & " SUPR_CD_SUB_PROCESSO"
    str_SQL_Atividade = str_SQL_Atividade & " HAVING MEPR_CD_MEGA_PROCESSO = " & int_MegaProcesso
    str_SQL_Atividade = str_SQL_Atividade & " AND PROC_CD_PROCESSO = " & int_Processo
    str_SQL_Atividade = str_SQL_Atividade & " AND SUPR_CD_SUB_PROCESSO = " & int_SubProcesso
	Set rdsMaxAtividade = Conn_db.Execute(str_SQL_Atividade)
    if rdsMaxAtividade.EOF then
	   int_MaxAtividade = 1	
    else
	   int_MaxAtividade = rdsMaxAtividade("MAX_ATIVIDADE") + 1
    end if
	rdsMaxAtividade.Close
	set rdsMaxAtividade = Nothing
    str_SQL_Atividade = ""
	str_SQL_Atividade = str_SQL_Atividade & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE ( "
    str_SQL_Atividade = str_SQL_Atividade & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Atividade = str_SQL_Atividade & " ,PROC_CD_PROCESSO "
    str_SQL_Atividade = str_SQL_Atividade & " ,SUPR_CD_SUB_PROCESSO "
    str_SQL_Atividade = str_SQL_Atividade & " ,ATIV_CD_ATIVIDADE "		
    str_SQL_Atividade = str_SQL_Atividade & " ,ATIV_TX_DESC_ATIVIDADE "
    str_SQL_Atividade = str_SQL_Atividade & " ,ATUA_TX_OPERACAO "
	str_SQL_Atividade = str_SQL_Atividade & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Atividade = str_SQL_Atividade & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Atividade = str_SQL_Atividade & " ) Values( "
	str_SQL_Atividade = str_SQL_Atividade & int_MegaProcesso & "," & int_Processo & "," & int_SubProcesso & "," & int_MaxAtividade & ","
	str_SQL_Atividade = str_SQL_Atividade & "'" & str_NovaAtividade & "', 'I', 'XXXX', GETDATE())" 
	Set rdsNovaAtividade = Conn_db.Execute(str_SQL_Atividade)

    strChave = CStr(int_MegaProcesso) & CStr(int_Processo) & CStr(int_SubProcesso) & CStr(int_MaxAtividade) ' & CStr(strEU)
	'call grava_log(str_NovaAtividade,"" & Session("PREFIXO") & "ATIVIDADE","I",0)
	
	int_Atividade = int_MaxAtividade
   else
    str_SemSubProcesso = 0
    str_MensagemProc = "Para cadastrar uma nova Atividade deve ser selecionado um Sub Processo ou preencha um novo"  
   end if	
   else
    str_SemProcesso = 0
    str_MensagemProc = "Para cadastrar uma nova Atividade deve ser selecionado um Processo ou preencha um novo"  
   end if	
 else
  str_SemMegaProcesso = 0
  str_MensagemMegaProc = "Para cadastrar uma nova Atividade deve ser selecionado um Megaprocesso"  
 end if   
end if


%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function Confirma() 
{ 
	  document.frmResInc.submit();
	  }
//-->
</script>	  
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form name="frmResInc" method="post" action="form_relaciona_ativ_trans.asp?txtOpc=1">
  <table width="105%" border="0" cellpadding="0" cellspacing="0" height="353">
  <tr> 
    <td width="100%"> 

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
          <tr>
            <td width="20%" height="20">&nbsp;</td>
            <td width="44%" height="60">&nbsp;</td>
            <td width="36%" valign="top"> 
              <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
                    <div align="center">&nbsp;<a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
          </td>
        </tr>
      </table>
    </td>
        </tr>
        <tr bgcolor="#00FF99"> 
          <td colspan="3" height="36" bgcolor="#00FF99"> 
            <table width="625" border="0" align="center">
              <tr> 
                  <td width="26">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
                <td width="26">&nbsp;</td>
                <td width="195">
                  <table width="98%" border="0">
                    <tr>
                      <td width="19%"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
                        <td width="81%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099"><a href="javascript:Confirma()">Relaciona 
                          Transa&ccedil;&atilde;o</a> </font></b></font></td>
                    </tr>
                  </table>
                  <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
                  <td width="27">&nbsp;</td>
                  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
                <td width="28">&nbsp;</td>
                <td width="26">&nbsp;</td>
                <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="100%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="100%" height="295" valign="top"> 
      <table width="100%" border="0">
        <tr> 
          <td height="77"> 
            <table width="89%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="0%">&nbsp;</td>
                <td width="3%"> 
                  <%'=int_MegaProcesso%>
                </td>
                <td width="4%"> 
                  <%'=int_Processo%>
                </td>
                <td width="4%"> 
                  <%'=int_SubProcesso%>
                </td>
                <td width="4%"> 
                  <%'=int_Atividade%>
                </td>
                <td width="9%">&nbsp;</td>
                <td width="6%"> 
                  <%'=str_NovoProcesso%>
                </td>
                <td width="8%"> 
                  <%'=str_NovoSubProcesso%>
                </td>
                <td width="10%"> 
                  <%'=str_NovaAtividade%>
                </td>
                <td width="52%">
                  <%'=str_MensagemMegaProc%>
                </td>
              </tr>
            </table>
            <table width="89%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="0%">&nbsp;</td>
                <td width="3%">&nbsp; </td>
                <td width="4%">&nbsp; </td>
                <td width="4%">&nbsp; </td>
                <td width="4%">&nbsp; </td>
                <td width="9%">&nbsp;</td>
                <td width="6%">
                  <%'=A%>
                </td>
                <td width="8%">&nbsp; </td>
                <td width="10%">&nbsp; </td>
                <td width="52%">
                  <%'=str_MensagemProc%>
                </td>
              </tr>
            </table>
              <p> 
                <input type="hidden" name="txtMegaProcesso" value="<%=int_MegaProcesso%>">
                <input type="hidden" name="txtProcesso" value="<%=int_Processo%>">
                <input type="hidden" name="txtSubProcesso" value="<%=int_SubProcesso%>">
                <input type="hidden" name="txtAtividade" value="<%=int_Atividade%>">
              </p>
          </td>
        </tr>
        <tr> 
          <td> </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <%If int_MaxProcesso <> 0 then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Novo 
                    Processo:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=int_MaxProcesso%></font></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NovoProcesso%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
            </table>
            <%end if%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%If int_MaxSubProcesso <> 0 then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Novo 
                    Sub Processo:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=int_MaxSubProcesso%></font></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NovoSubProcesso%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
            </table>
            <%end if%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%If int_MaxAtividade <> 0 then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Nova 
                    Atividade:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=int_MaxAtividade%></font></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NovaAtividade%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
            </table>
            <%end if%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%str_Mensagem_Final = str_MensagemMegaProc & str_MensagemProc & str_MensagemSubProc
			If str_Mensagem_Final <> "" then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encontrado 
                    erro:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Mensagem_Final%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%"> 
                  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="form_inc_sub_processo.asp">Tela 
                    de Cadastro</a></font></div>
                </td>
                <td width="22%">&nbsp;</td>
              </tr>
            </table>
            <%end if%>
            <%str_Msg = "" 
			   if int_MegaProcesso = 0 then
                  str_Msg = "Para qualquer operação é necessário a seleção de um Mega Processo"
			   else
			      if str_NovoProcesso = "" and str_NovoSubProcesso = "" and str_NovaAtividade = "" then
			         if int_Processo = 0 and str_NovoProcesso = "" then
				        str_Msg = "Para qualquer operação é necessário a seleção de um Processo ou o cadastro de um novo"
                     else
                        if int_SubProcesso = 0 and str_NovoSubProcesso = "" then
				           str_Msg = "Para qualquer operação é necessário a seleção de um Sub Processo ou o cadastro de um novo"
                        else
                           if int_Atividade = 0 and str_NovaAtividade = "" then
		                      str_Msg = "Para qualquer operação é necessário a seleção de uma Atividade ou o cadastro de uma nova"
                           end if
						end if
				     end if 		   				     					 
                  else
				     if str_NovoSubProcesso <> "" then
				        if int_Processo = 0 and str_NovoProcesso = "" then
				           str_Msg = "Para cadastrar um novo Sub Processo deve ser selecionado um Processo ou preencha um novo"  				     
                        end if    						
				     else
			            if str_NovoAtividade <> "" then
				           if int_SubProcesso = 0 and str_NovoSubProcesso = "" then
				              str_Msg = "Para cadastrar uma nova Atividade deve ser selecionado um Sub Processo ou preencha um novo"  				     
                           else   						
				              if int_Processo = 0 and str_NovoProcesso = "" then
				                 str_Msg = "Para cadastrar uma nova Atividade deve ser selecionado um Processo ou preencha um novo"  				     						    	  
                              end if
					       end if	    						
				   	    end if
				     end if
				   end if
				 end if  
				 	  	
			'If int_MegaProcesso = 0 and int_Processo = 0 and int_SubProcesso = 0 and int_Atividade = 0 and str_NovoProcesso = "" and str_NovoSubProcesso = "" and str_NovaAtividade = "" then
			%>
            <% If str_Msg <> "" then %>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encontrado 
                    erro:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b></b></div>
                </td>
                <td width="45%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Msg%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%"> 
                  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="form_inc_sub_processo.asp">Tela 
                    de Cadastro</a></font></div>
                </td>
                <td width="22%">&nbsp;</td>
              </tr>
            </table>
            <% end if %>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
</body>
</html>
=======
<%@LANGUAGE="VBSCRIPT"%> 
 
<!--#include file="../asp/protege/protege.asp" -->
<%
dim int_MaxProcesso
dim int_MaxSubProcesso
dim int_MaxAtividade
dim str_MensagemMegaProc
dim str_MensagemProc
dim str_MensagemSubProc
dim str_Msg 

str_MensagemMegaProc = ""
str_MensagemProc = ""
str_MensagemSubProc = ""

int_MaxProcesso = 0
int_MaxSubProcesso = 0
int_MaxAtividade = 0

int_MegaProcesso= Request.Form("selMegaProcesso")
int_Processo = Request.Form("SelProcesso")
str_NovoProcesso = UCase(Request.Form("txtNovoProcesso"))
int_SubProcesso = Request.Form("SelSubProcesso")
str_NovoSubProcesso = UCase(Request.Form("txtNovoSubProcesso"))
int_Atividade = Request.Form("selAtividade")
str_NovaAtividade = UCase(Request.Form("txtNovaAtividade"))

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

'************GRAVA NOVO PROCESSO *************************
if str_NovoProcesso <> "" then
   if int_MegaProcesso <> 0 then
	str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " SELECT "
	str_SQL_Proc = str_SQL_Proc & " MAX(PROC_CD_PROCESSO) AS MAX_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " FROM " & Session("PREFIXO") & "PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO "
	str_SQL_Proc = str_SQL_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & int_MegaProcesso
	Set rdsMaxProcesso = Conn_db.Execute(str_SQL_Proc)
	if rdsMaxProcesso.EOF then
	   int_MaxProcesso = 1	
	else
	   int_MaxProcesso = rdsMaxProcesso("MAX_PROCESSO") + 1	
	end if
	rdsMaxProcesso.Close
	set rdsMaxProcesso = Nothing
    str_SQL_Proc = ""
	str_SQL_Proc = str_SQL_Proc & " INSERT INTO " & Session("PREFIXO") & "PROCESSO ( "
    str_SQL_Proc = str_SQL_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,PROC_TX_DESC_PROCESSO "
    str_SQL_Proc = str_SQL_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Proc = str_SQL_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Proc = str_SQL_Proc & " ) Values( "
	str_SQL_Proc = str_SQL_Proc & int_MegaProcesso & "," & int_MaxProcesso & ","
	str_SQL_Proc = str_SQL_Proc & "'" & str_NovoProcesso & "', 'I', 'XXXX', GETDATE())" 
	Set rdsNovoProcesso = Conn_db.Execute(str_SQL_Proc)
   strChave = CStr(int_MegaProcesso) & CStr(int_MaxProcesso) ' &  CStr(strP) & CStr(strSP) & CStr(strEU)
   'call grava_log(strChave,"PROCESSO","I",0)	
	int_Processo = int_MaxProcesso
  else
    str_SemMegaProcesso = 0
	str_MensagemMegaProc = "Para cadastrar um novo Processo deve ser selecionado um Megaprocesso"
  end if	
end if
'********** GRAVA NOVO SUB PROCESSO ****************
if str_NovoSubProcesso <> "" then
 if int_MegaProcesso <> 0 then
  if int_Processo <> 0 or int_MaxProcesso <> 0  then
    if int_MaxProcesso <> 0 then
	   int_Processo = int_MaxProcesso
	end if
	str_SQL_Sub_Proc = ""
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " SELECT "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MAX(SUPR_CD_SUB_PROCESSO) AS MAXIMO_SUB "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " FROM " & Session("PREFIXO") & "SUB_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " GROUP BY MEPR_CD_MEGA_PROCESSO, "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " HAVING MEPR_CD_MEGA_PROCESSO = " & int_MegaProcesso
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " AND PROC_CD_PROCESSO = " & int_Processo
	a = str_SQL_Sub_Proc
	Set rdsMaxSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)
    'if IsNull(rdsMaxSubProcesso("MAXIMO_SUB")) then
    if rdsMaxSubProcesso.EOF then
	   int_MaxSubProcesso = 1	
    else
	   'A = rdsMaxSubProcesso("MAXIMO_SUB") 
   	   int_MaxSubProcesso = rdsMaxSubProcesso("MAXIMO_SUB") + 1	
	   'int_MaxSubProcesso = 10
    end if
	rdsMaxSubProcesso.Close
	set rdsMaxSubProcesso = Nothing
    str_SQL_Sub_Proc = ""
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " INSERT INTO " & Session("PREFIXO") & "SUB_PROCESSO ( "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,PROC_CD_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_CD_SUB_PROCESSO "	
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,SUPR_TX_DESC_SUB_PROCESSO "
    str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_TX_OPERACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & " ) Values( "
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & int_MegaProcesso & "," & int_Processo & "," & int_MaxSubProcesso & ","
	str_SQL_Sub_Proc = str_SQL_Sub_Proc & "'" & str_NovoSubProcesso & "', 'I', 'XXXX', GETDATE())" 
	Set rdsNovoSubProcesso = Conn_db.Execute(str_SQL_Sub_Proc)

    strChave = CStr(int_MegaProcesso) & CStr(int_MaxProcesso) & CStr(int_MaxSubProcesso) '& CStr(strSP) & CStr(strEU)
    'call grava_log(strChave,"SUB_PROCESSO","I",0)	
			
	int_SubProcesso = int_MaxSubProcesso
   else
    str_SemProcesso = 0
    str_MensagemProc = "Para cadastrar um novo Sub Processo deve ser selecionado um Processo ou preencha um novo Processo"  
   end if	
 else
  str_SemMegaProcesso = 0
  str_MensagemMegaProc = "Para cadastrar um novo Sub Processo deve ser selecionado um Megaprocesso"  
 end if   
end if
'************GRAVA NOVA ATIVIDADE *************************
if str_NovaAtividade <> "" then
 if int_MegaProcesso <> 0 then
  if int_Processo <> 0 or int_MaxProcesso <> 0  then
    if int_MaxProcesso <> 0 then
	   int_Processo = int_MaxProcesso
	end if
   if int_SubProcesso <> 0 or int_MaxSubProcesso <> 0  then
      if int_MaxSubProcesso <> 0 then
	    int_SubProcesso = int_MaxSubProcesso
      end if
    str_SQL_Atividade = ""
    str_SQL_Atividade = str_SQL_Atividade & " SELECT "
    str_SQL_Atividade = str_SQL_Atividade & " MAX(ATIV_CD_ATIVIDADE) AS MAX_ATIVIDADE "
    str_SQL_Atividade = str_SQL_Atividade & " FROM " & Session("PREFIXO") & "ATIVIDADE "
    str_SQL_Atividade = str_SQL_Atividade & " GROUP BY MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, "
    str_SQL_Atividade = str_SQL_Atividade & " SUPR_CD_SUB_PROCESSO"
    str_SQL_Atividade = str_SQL_Atividade & " HAVING MEPR_CD_MEGA_PROCESSO = " & int_MegaProcesso
    str_SQL_Atividade = str_SQL_Atividade & " AND PROC_CD_PROCESSO = " & int_Processo
    str_SQL_Atividade = str_SQL_Atividade & " AND SUPR_CD_SUB_PROCESSO = " & int_SubProcesso
	Set rdsMaxAtividade = Conn_db.Execute(str_SQL_Atividade)
    if rdsMaxAtividade.EOF then
	   int_MaxAtividade = 1	
    else
	   int_MaxAtividade = rdsMaxAtividade("MAX_ATIVIDADE") + 1
    end if
	rdsMaxAtividade.Close
	set rdsMaxAtividade = Nothing
    str_SQL_Atividade = ""
	str_SQL_Atividade = str_SQL_Atividade & " INSERT INTO " & Session("PREFIXO") & "ATIVIDADE ( "
    str_SQL_Atividade = str_SQL_Atividade & " MEPR_CD_MEGA_PROCESSO "
    str_SQL_Atividade = str_SQL_Atividade & " ,PROC_CD_PROCESSO "
    str_SQL_Atividade = str_SQL_Atividade & " ,SUPR_CD_SUB_PROCESSO "
    str_SQL_Atividade = str_SQL_Atividade & " ,ATIV_CD_ATIVIDADE "		
    str_SQL_Atividade = str_SQL_Atividade & " ,ATIV_TX_DESC_ATIVIDADE "
    str_SQL_Atividade = str_SQL_Atividade & " ,ATUA_TX_OPERACAO "
	str_SQL_Atividade = str_SQL_Atividade & " ,ATUA_CD_NR_USUARIO "
	str_SQL_Atividade = str_SQL_Atividade & " ,ATUA_DT_ATUALIZACAO "
	str_SQL_Atividade = str_SQL_Atividade & " ) Values( "
	str_SQL_Atividade = str_SQL_Atividade & int_MegaProcesso & "," & int_Processo & "," & int_SubProcesso & "," & int_MaxAtividade & ","
	str_SQL_Atividade = str_SQL_Atividade & "'" & str_NovaAtividade & "', 'I', 'XXXX', GETDATE())" 
	Set rdsNovaAtividade = Conn_db.Execute(str_SQL_Atividade)

    strChave = CStr(int_MegaProcesso) & CStr(int_Processo) & CStr(int_SubProcesso) & CStr(int_MaxAtividade) ' & CStr(strEU)
	'call grava_log(str_NovaAtividade,"" & Session("PREFIXO") & "ATIVIDADE","I",0)
	
	int_Atividade = int_MaxAtividade
   else
    str_SemSubProcesso = 0
    str_MensagemProc = "Para cadastrar uma nova Atividade deve ser selecionado um Sub Processo ou preencha um novo"  
   end if	
   else
    str_SemProcesso = 0
    str_MensagemProc = "Para cadastrar uma nova Atividade deve ser selecionado um Processo ou preencha um novo"  
   end if	
 else
  str_SemMegaProcesso = 0
  str_MensagemMegaProc = "Para cadastrar uma nova Atividade deve ser selecionado um Megaprocesso"  
 end if   
end if


%>
<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
<!--
function Confirma() 
{ 
	  document.frmResInc.submit();
	  }
//-->
</script>	  
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form name="frmResInc" method="post" action="form_relaciona_ativ_trans.asp?txtOpc=1">
  <table width="105%" border="0" cellpadding="0" cellspacing="0" height="353">
  <tr> 
    <td width="100%"> 

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099">
          <tr>
            <td width="20%" height="20">&nbsp;</td>
            <td width="44%" height="60">&nbsp;</td>
            <td width="36%" valign="top"> 
              <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
        <tr> 
          <td bgcolor="#330099" width="39" valign="middle" align="center"> 
            <div align="center">
              <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.forward()"><img border="0" src="../imagens/avancar.gif"></a></div>
          </td>
          <td bgcolor="#330099" width="27" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../imagens/favoritos.gif"></a></div>
          </td>
        </tr>
        <tr> 
          <td bgcolor="#330099" height="12" width="39" valign="middle" align="center"> 
            <div align="center"><a href="javascript:print()"><img border="0" src="../imagens/imprimir.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="36" valign="middle" align="center"> 
            <div align="center"><a href="JavaScript:history.go()"><img border="0" src="../imagens/atualizar.gif"></a></div>
          </td>
          <td bgcolor="#330099" height="12" width="27" valign="middle" align="center"> 
                    <div align="center">&nbsp;<a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
          </td>
        </tr>
      </table>
    </td>
        </tr>
        <tr bgcolor="#00FF99"> 
          <td colspan="3" height="36" bgcolor="#00FF99"> 
            <table width="625" border="0" align="center">
              <tr> 
                  <td width="26">&nbsp;</td>  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
                <td width="26">&nbsp;</td>
                <td width="195">
                  <table width="98%" border="0">
                    <tr>
                      <td width="19%"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
                        <td width="81%"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b><font color="#330099"><a href="javascript:Confirma()">Relaciona 
                          Transa&ccedil;&atilde;o</a> </font></b></font></td>
                    </tr>
                  </table>
                  <font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
                  <td width="27">&nbsp;</td>
                  <td width="50"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
                <td width="28">&nbsp;</td>
                <td width="26">&nbsp;</td>
                <td width="159"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b></b></font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td width="100%">&nbsp;</td>
  </tr>
  <tr> 
    <td width="100%" height="295" valign="top"> 
      <table width="100%" border="0">
        <tr> 
          <td height="77"> 
            <table width="89%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="0%">&nbsp;</td>
                <td width="3%"> 
                  <%'=int_MegaProcesso%>
                </td>
                <td width="4%"> 
                  <%'=int_Processo%>
                </td>
                <td width="4%"> 
                  <%'=int_SubProcesso%>
                </td>
                <td width="4%"> 
                  <%'=int_Atividade%>
                </td>
                <td width="9%">&nbsp;</td>
                <td width="6%"> 
                  <%'=str_NovoProcesso%>
                </td>
                <td width="8%"> 
                  <%'=str_NovoSubProcesso%>
                </td>
                <td width="10%"> 
                  <%'=str_NovaAtividade%>
                </td>
                <td width="52%">
                  <%'=str_MensagemMegaProc%>
                </td>
              </tr>
            </table>
            <table width="89%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="0%">&nbsp;</td>
                <td width="3%">&nbsp; </td>
                <td width="4%">&nbsp; </td>
                <td width="4%">&nbsp; </td>
                <td width="4%">&nbsp; </td>
                <td width="9%">&nbsp;</td>
                <td width="6%">
                  <%'=A%>
                </td>
                <td width="8%">&nbsp; </td>
                <td width="10%">&nbsp; </td>
                <td width="52%">
                  <%'=str_MensagemProc%>
                </td>
              </tr>
            </table>
              <p> 
                <input type="hidden" name="txtMegaProcesso" value="<%=int_MegaProcesso%>">
                <input type="hidden" name="txtProcesso" value="<%=int_Processo%>">
                <input type="hidden" name="txtSubProcesso" value="<%=int_SubProcesso%>">
                <input type="hidden" name="txtAtividade" value="<%=int_Atividade%>">
              </p>
          </td>
        </tr>
        <tr> 
          <td> </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <%If int_MaxProcesso <> 0 then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Novo 
                    Processo:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=int_MaxProcesso%></font></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NovoProcesso%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
            </table>
            <%end if%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%If int_MaxSubProcesso <> 0 then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Novo 
                    Sub Processo:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=int_MaxSubProcesso%></font></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NovoSubProcesso%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
            </table>
            <%end if%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%If int_MaxAtividade <> 0 then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Nova 
                    Atividade:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=int_MaxAtividade%></font></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_NovaAtividade%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
            </table>
            <%end if%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%str_Mensagem_Final = str_MensagemMegaProc & str_MensagemProc & str_MensagemSubProc
			If str_Mensagem_Final <> "" then%>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encontrado 
                    erro:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b></b></div>
                </td>
                <td width="45%"> <b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Mensagem_Final%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%"> 
                  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="form_inc_sub_processo.asp">Tela 
                    de Cadastro</a></font></div>
                </td>
                <td width="22%">&nbsp;</td>
              </tr>
            </table>
            <%end if%>
            <%str_Msg = "" 
			   if int_MegaProcesso = 0 then
                  str_Msg = "Para qualquer operação é necessário a seleção de um Mega Processo"
			   else
			      if str_NovoProcesso = "" and str_NovoSubProcesso = "" and str_NovaAtividade = "" then
			         if int_Processo = 0 and str_NovoProcesso = "" then
				        str_Msg = "Para qualquer operação é necessário a seleção de um Processo ou o cadastro de um novo"
                     else
                        if int_SubProcesso = 0 and str_NovoSubProcesso = "" then
				           str_Msg = "Para qualquer operação é necessário a seleção de um Sub Processo ou o cadastro de um novo"
                        else
                           if int_Atividade = 0 and str_NovaAtividade = "" then
		                      str_Msg = "Para qualquer operação é necessário a seleção de uma Atividade ou o cadastro de uma nova"
                           end if
						end if
				     end if 		   				     					 
                  else
				     if str_NovoSubProcesso <> "" then
				        if int_Processo = 0 and str_NovoProcesso = "" then
				           str_Msg = "Para cadastrar um novo Sub Processo deve ser selecionado um Processo ou preencha um novo"  				     
                        end if    						
				     else
			            if str_NovoAtividade <> "" then
				           if int_SubProcesso = 0 and str_NovoSubProcesso = "" then
				              str_Msg = "Para cadastrar uma nova Atividade deve ser selecionado um Sub Processo ou preencha um novo"  				     
                           else   						
				              if int_Processo = 0 and str_NovoProcesso = "" then
				                 str_Msg = "Para cadastrar uma nova Atividade deve ser selecionado um Processo ou preencha um novo"  				     						    	  
                              end if
					       end if	    						
				   	    end if
				     end if
				   end if
				 end if  
				 	  	
			'If int_MegaProcesso = 0 and int_Processo = 0 and int_SubProcesso = 0 and int_Atividade = 0 and str_NovoProcesso = "" and str_NovoSubProcesso = "" and str_NovaAtividade = "" then
			%>
            <% If str_Msg <> "" then %>
            <table width="82%" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%"> 
                  <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encontrado 
                    erro:</font></div>
                </td>
                <td width="7%"> 
                  <div align="center"><b></b></div>
                </td>
                <td width="45%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=str_Msg%></font></b></td>
                <td width="22%"><b></b></td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%">&nbsp;</td>
                <td width="22%">&nbsp;</td>
              </tr>
              <tr> 
                <td width="5%">&nbsp;</td>
                <td width="21%">&nbsp;</td>
                <td width="7%">&nbsp;</td>
                <td width="45%"> 
                  <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="form_inc_sub_processo.asp">Tela 
                    de Cadastro</a></font></div>
                </td>
                <td width="22%">&nbsp;</td>
              </tr>
            </table>
            <% end if %>
          </td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
