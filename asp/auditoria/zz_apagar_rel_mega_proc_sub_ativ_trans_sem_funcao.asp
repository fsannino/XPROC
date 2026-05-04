<%
response.Buffer=false
Server.ScriptTimeOut=9990000

set db = Server.CreateObject("ADODB.Connection")
db.Open Session("Conn_String_Cogest_Gravacao")

MEGA=REQUEST("selMegaProcesso")
'PROC=REQUEST("selProcesso")
'SUBR=REQUEST("selSubProcesso")

'MODULO=REQUEST("selModulo")
'ATIVIDADE=REQUEST("selAtividade")
'TRANSACAO=REQUEST("selTransacao")

IF MEGA<>0 THEN
COMPL1="dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO=" & MEGA
END IF

'IF PROC<>0 THEN
'COMPL2="PROC_CD_PROCESSO=" & PROC
'END IF

'IF SUBR<>0 THEN
'COMPL3="SUPR_CD_SUB_PROCESSO=" & SUBR
'END IF

'IF MODULO<>0 THEN
'COMPL4="MODU_CD_MODULO=" & MODULO
'END IF

'IF ATIVIDADE<>0 THEN
'COMPL5="ATCA_CD_ATIVIDADE_CARGA=" & ATIVIDADE
'END IF

'IF TRANSACAO<>"0" THEN
'COMPL6="TRAN_CD_TRANSACAO='" & TRANSACAO & "'"
'END IF

IF COMPL1<>"" THEN
COMPLE = COMPL1
END IF

IF COMPL2<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE + " AND " + COMPL2
ELSE
COMPLE=COMPL2
END IF
END IF

IF COMPL3<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL3
ELSE
COMPLE=COMPL3
END IF
END IF

IF COMPL4<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL4
ELSE
COMPLE=COMPL4
END IF
END IF

IF COMPL5<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL5
ELSE
COMPLE=COMPL5
END IF
END IF

IF COMPL6<>"" THEN
IF LEN(COMPLE)>0 THEN
COMPLE = COMPLE +" AND " + COMPL6
ELSE
COMPLE=COMPL6
END IF
END IF

IF COMPLE<>"" THEN
CONECTA=" AND "
END IF

'SSQL="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL " & CONECTA & COMPLE & " ORDER BY RELA_NR_SEQUENCIA"

'SSQL="SELECT * FROM " & Session("PREFIXO") & "RELACAO_FINAL " & CONECTA & COMPLE & " ORDER BY MEPR_CD_MEGA_PROCESSO, PROC_CD_PROCESSO, SUPR_CD_SUB_PROCESSO, MODU_CD_MODULO, ATCA_CD_ATIVIDADE_CARGA"

str_Sql = ""
str_Sql = str_SQl & " SELECT"
str_Sql = str_SQl & " dbo.RELACAO_FINAL.MODU_CD_MODULO"
str_Sql = str_SQl & " , dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA"
str_Sql = str_SQl & " , dbo.RELACAO_FINAL.TRAN_CD_TRANSACAO"
str_Sql = str_SQl & " , dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO"
str_Sql = str_SQl & " , dbo.RELACAO_FINAL.PROC_CD_PROCESSO"
str_Sql = str_SQl & " , dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO"
str_Sql = str_SQl & " FROM dbo.RELACAO_FINAL LEFT OUTER JOIN"
str_Sql = str_SQl & " dbo.FUN_NEG_TRANSACAO ON dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = dbo.FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA AND "
str_Sql = str_SQl & " dbo.RELACAO_FINAL.TRAN_CD_TRANSACAO = dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO AND "
str_Sql = str_SQl & " dbo.RELACAO_FINAL.MODU_CD_MODULO = dbo.FUN_NEG_TRANSACAO.MODU_CD_MODULO"
str_Sql = str_SQl & " WHERE dbo.FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO IS NULL"
str_Sql = str_SQl & CONECTA & COMPLE & " ORDER BY dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.PROC_CD_PROCESSO, dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, dbo.RELACAO_FINAL.MODU_CD_MODULO, dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA"

str_Sql = " SELECT"
str_Sql = str_SQl & " dbo.RELACAO_FINAL.MODU_CD_MODULO, dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA, dbo.RELACAO_FINAL.TRAN_CD_TRANSACAO, "
str_Sql = str_SQl & " dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.PROC_CD_PROCESSO, dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO, "
str_Sql = str_SQl & " dbo.TRANSACAO_MEGA.MEPR_CD_MEGA_PROCESSO AS Mega_Dono"
str_Sql = str_SQl & " FROM dbo.RELACAO_FINAL INNER JOIN"
str_Sql = str_SQl & " dbo.TRANSACAO_MEGA ON dbo.RELACAO_FINAL.TRAN_CD_TRANSACAO = dbo.TRANSACAO_MEGA.TRAN_CD_TRANSACAO LEFT OUTER JOIN"
str_Sql = str_SQl & " dbo.FUN_NEG_TRANSACAO ON dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA = dbo.FUN_NEG_TRANSACAO.ATCA_CD_ATIVIDADE_CARGA AND "
str_Sql = str_SQl & " dbo.RELACAO_FINAL.TRAN_CD_TRANSACAO = dbo.FUN_NEG_TRANSACAO.TRAN_CD_TRANSACAO AND "
str_Sql = str_SQl & " dbo.RELACAO_FINAL.MODU_CD_MODULO = dbo.FUN_NEG_TRANSACAO.MODU_CD_MODULO"
str_Sql = str_SQl & " WHERE     (dbo.FUN_NEG_TRANSACAO.MEPR_CD_MEGA_PROCESSO IS NULL)"
str_Sql = str_SQl & CONECTA & COMPLE 
str_Sql = str_SQl & " ORDER BY dbo.TRANSACAO_MEGA.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.MEPR_CD_MEGA_PROCESSO, dbo.RELACAO_FINAL.PROC_CD_PROCESSO, dbo.RELACAO_FINAL.SUPR_CD_SUB_PROCESSO,"
str_Sql = str_SQl & " dbo.RELACAO_FINAL.MODU_CD_MODULO, dbo.RELACAO_FINAL.ATCA_CD_ATIVIDADE_CARGA"

'response.write str_SQl

set rs=db.execute(str_SQl)
%>

<html>
<head>
<script>
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
</script>
<title>SINERGIA # XPROC # Processos de Negócio</title>
<base target="_self">
</head>

<body topmargin="0" leftmargin="0">
<form method="POST" action="../gera_rel_modatca.asp" name="frm1">
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
            <div align="center"><a href="JavaScript:window.external.AddFavorite('http://S6000WS12.corp.petrobras.biz/xproc/index.htm','Sinergia  - X-Total')"><img border="0" src="../../imagens/favoritos.gif"></a></div>
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
          <td width="26"></td>
          <td width="50"></td>
          <td width="26">&nbsp;</td>
            <td width="94">&nbsp;</td>
          <td width="128">
            <p align="center"></td>
          <td width="50"></td>
          <td width="28"></td>
          <td width="26">&nbsp;</td>
          <td width="159"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
  <p style="margin-top: 0; margin-bottom: 0"><font face="Arial" size="2">&nbsp; 
    </font></p>
  <p style="margin-top: 0; margin-bottom: 0" align="center"></p>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%">
    <tr>
      <td width="65%"><div align="center"><b><font color="#330099" size="3" face="Verdana">Rela&ccedil;&atilde;o de Transa&ccedil;&otilde;es n&atilde;o associadas a Fun&ccedil;&atilde;o </font></b></div></td>
      <td width="35%"><font face="Verdana" color="#330099" size="3"> <img src="../../imagens/carregando01.gif" width="120" height="18" id="loader"></font></td>
    </tr>
  </table>
<%if rs.eof=true then%>  
<p style="margin-top: 0; margin-bottom: 0" align="left"><font face="Verdana" color="#800000" size="2"><b>Nenhum
Registro Encontrado</b></font></p>
<%else%>
<p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
<table border="0" width="100%">
  <tr>
    <td width="16%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Mega-Processo</font></b></td>
    <td width="16%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Mega-Processo-Dono</font></b></td>
    <td width="16%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Processo</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Sub-Processo</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Atividade</font></b></td>
    <td width="17%" bgcolor="#B5D6E8"><b><font face="Verdana" size="1">Transação</font></b></td>
  </tr>
<%

valor1=""
valor2=""
valor3=""
valor4=""
valor5=""
valor6=""

mega_dono_atual=RS("Mega_Dono")
mega_atual=RS("MEPR_CD_MEGA_PROCESSO")
proc_atual=RS("PROC_CD_PROCESSO")
sub_atual=RS("SUPR_CD_SUB_PROCESSO")
modulo_atual=RS("MODU_CD_MODULO")
atividade_atual=RS("ATCA_CD_ATIVIDADE_CARGA")

'response.Write(mega_atual & " - ")
'response.Write(proc_atual & " - ")
'response.Write(sub_atual & " - ")
'response.Write(modulo_atual & " - ")
'response.Write(atividade_atual & " - ")
'response.End()
do until rs.eof=true

	if mega_dono_ant<>mega_dono_atual then
		set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_dono_atual)
		valor0=rs1("MEPR_TX_DESC_MEGA_PROCESSO")
	else
		valor0=""
	end if

	if mega_ant<>mega_atual then
		set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual)
		valor1=rs1("MEPR_TX_DESC_MEGA_PROCESSO")
	else
		valor1=""
	end if

if proc_ant<>proc_atual then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual)
	valor2=rs1("PROC_TX_DESC_PROCESSO")
else
	valor2=""
end if

if sub_atual<>sub_ant then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual)
	valor3=rs1("SUPR_TX_DESC_SUB_PROCESSO")
else
	if proc_atual<>proc_ant then
		set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual)
		valor3=rs1("SUPR_TX_DESC_SUB_PROCESSO")
	else
		valor3=""
	end if
end if

if (modulo_atual<>modulo_ant) or (valor3<>"") then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & modulo_atual)
	valor4=rs1("MODU_TX_DESC_MODULO")
else
	valor4=""
end if
'response.Write("<p>" & Atividade_atual & "<p>")
if(atividade_atual<>atividade_ant) or (valor4<>"") then
	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & Atividade_atual)
	if not rs1.Eof then
		valor5=RS1("ATCA_TX_DESC_ATIVIDADE")
	else
		valor5= "não encontrado - " & Atividade_atual
	end if
else
	valor5=""
end if

set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "TRANSACAO WHERE TRAN_CD_TRANSACAO='" & RS("TRAN_CD_TRANSACAO") & "'")
valor6=RS("TRAN_CD_TRANSACAO") & "-" & rs1("TRAN_TX_DESC_TRANSACAO")

%>
  <tr>
    <%if valor5<>"" then
    set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_dono_atual)
	valor0=rs1("MEPR_TX_DESC_MEGA_PROCESSO")

    set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual)
	valor1=rs1("MEPR_TX_DESC_MEGA_PROCESSO")

	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual)
	valor2=rs1("PROC_TX_DESC_PROCESSO")

	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "SUB_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & mega_atual & " AND PROC_CD_PROCESSO=" & proc_atual & " AND SUPR_CD_SUB_PROCESSO=" & sub_atual)
	valor3=rs1("SUPR_TX_DESC_SUB_PROCESSO")

	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "MODULO_R3 WHERE MODU_CD_MODULO=" & modulo_atual)
	valor4=rs1("MODU_TX_DESC_MODULO")

	set rs1=db.execute("SELECT * FROM " & Session("PREFIXO") & "ATIVIDADE_CARGA WHERE ATCA_CD_ATIVIDADE_CARGA=" & Atividade_atual)
	if not rs1.Eof then
		valor5=RS1("ATCA_TX_DESC_ATIVIDADE")
	else
		valor5= "não encontrado - " & Atividade_atual
	end if
	%>
    <td width="17%" bgcolor="#99CCFF"><font face="Verdana" size="1"><%=valor1%></font></td>
    <td width="17%" bgcolor="#99CCFF"><font face="Verdana" size="1"><%=valor0%></font></td>	
    <td width="17%" bgcolor="#FFFF00"><font face="Verdana" size="1"><%=valor2%></font></td>
    <td width="17%" bgcolor="#C0C0C0"><font face="Verdana" size="1"><%=valor3%></font></td>
    <td width="17%" bgcolor="#FFFF00"><font face="Verdana" size="1"><%=valor5%></font></td>
    <%else%>
	<td width="17%"><font face="Verdana" size="1"><%=valor1%></font></td>
	<td width="17%"><font face="Verdana" size="1"><%=valor0%></font></td>
	<td width="17%"><font face="Verdana" size="1"><%=valor2%></font></td>
	<td width="17%"><font face="Verdana" size="1"><%=valor3%></font></td>
	<td width="17%"><font face="Verdana" size="1"><%=valor5%></font></td>
    <%end if%>
    
    <td width="17%" bgcolor="#AAFFDD"><font face="Verdana" size="1"><%=valor6%></font></td>
  </tr>
<%
	mega_dono_ant=RS("Mega_Dono")
	mega_ant=RS("MEPR_CD_MEGA_PROCESSO")
	proc_ant=RS("PROC_CD_PROCESSO")
	sub_ant=RS("SUPR_CD_SUB_PROCESSO")
	modulo_ant=RS("MODU_CD_MODULO")
	atividade_ant=RS("ATCA_CD_ATIVIDADE_CARGA")

	rs.movenext
	if not rs.Eof then
		mega_dono_atual=RS("Mega_Dono")
		mega_atual=RS("MEPR_CD_MEGA_PROCESSO")
		proc_atual=RS("PROC_CD_PROCESSO")
		sub_atual=RS("SUPR_CD_SUB_PROCESSO")
		modulo_atual=RS("MODU_CD_MODULO")
		atividade_atual=RS("ATCA_CD_ATIVIDADE_CARGA")
	end if
	
LOOP

end if%>
</table>
<p style="margin-top: 0; margin-bottom: 0" align="left">&nbsp;</p>
</form>
</body>
<script>
	MM_swapImage('loader','','../../imagens/carregando_limpa.gif',1);
</script>
</html>
