<%@LANGUAGE="VBSCRIPT"%> 
 
<%
str_Opc = Request("txtOpc")

if (Request("selUsuario") <> "") then 
    str_Usuario = Request("selUsuario")
else
    str_Usuario = "0"
end if

set conn_db = Server.CreateObject("ADODB.Connection")
conn_db.Open Session("Conn_String_Cogest_Gravacao")

str_SQL_Usuario = ""
str_SQL_Usuario = str_SQL_Usuario & " SELECT DISTINCT "
str_SQL_Usuario = str_SQL_Usuario & " " & Session("PREFIXO") & "USUARIO.USUA_CD_USUARIO "
str_SQL_Usuario = str_SQL_Usuario & " , " & Session("PREFIXO") & "USUARIO.USUA_TX_NOME_USUARIO "
str_SQL_Usuario = str_SQL_Usuario & " , " & Session("PREFIXO") & "USUARIO.USUA_TX_CATEGORIA "
str_SQL_Usuario = str_SQL_Usuario & " FROM " & Session("PREFIXO") & "USUARIO "
str_SQL_Usuario = str_SQL_Usuario & " order by " & Session("PREFIXO") & "USUARIO.USUA_TX_NOME_USUARIO "

set rs_destino=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO='" & str_Usuario & "'") 

ssql = ""
ssql="SELECT * FROM " & Session("PREFIXO") & "USUARIO ORDER BY USUA_TX_NOME_USUARIO"

SET RS=CONN_DB.EXECUTE(SSQL)

RS_MEGA_PROCESSO = "SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO"

%>

<html>
<head>
<STYLE type=text/css>
BODY {
	SCROLLBAR-HIGHLIGHT-COLOR: white; SCROLLBAR-SHADOW-COLOR: white; SCROLLBAR-ARROW-COLOR: yellow; SCROLLBAR-BASE-COLOR: #003399; scrollbar-3d-light-color: White}
</STYLE>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<LINK REL="SHORTCUT ICON" href="http://regina/imagens/Wrench.ico">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<script>
function Confirma() 
{ 
if (document.form1.Cd.value == "")
     { 
	 alert("O campo Chave do Usuário deve ser preenchido!");
     document.form1.Cd.focus();
     return;
     }
if (document.form1.txtNomeUsuario.value == "")
     { 
	 alert("O campo Nome do Usuário deve ser preenchido!");
     document.form1.txtNomeUsuario.focus();
     return;
     }
if ((!document.form1.rdbCategoria[0].checked)&&
   (!document.form1.rdbCategoria[1].checked)&&
   (!document.form1.rdbCategoria[2].checked)&&
   (!document.form1.rdbCategoria[3].checked)&&
   (!document.form1.rdbCategoria[4].checked)&&   
   (!document.form1.rdbCategoria[5].checked)&&  
   (!document.form1.rdbCategoria[6].checked)&&  
   (!document.form1.rdbCategoria[7].checked)){
     alert("A seleção da Categoria é obrigatória !");
     return;
     }	 
	 else
     {
	  document.form1.submit();
	 }
}
function Limpa(){
	document.form1.reset();
}

</script>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../imagens/continua_F02.gif','../imagens/continua2_F02.gif')">
<form name="form1" method="POST" action="grava_usuario.asp">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#330099" height="86">
    <tr> 
      <td width="150" height="66" colspan="2">&nbsp;</td>
      <td width="341" height="66" colspan="2">&nbsp;</td>
      <td width="276" valign="top" colspan="2" height="66"> 
        <table width="154" border="0" align="right" cellpadding="0" cellspacing="0" bgcolor="#0000CC">
          <tr> 
            <td bgcolor="#330099" width="39" valign="middle" align="center"> 
              <div align="center"> 
                <p align="center"><a href="JavaScript:history.back()"><img border="0" src="../imagens/voltar.gif"></a> 
              </div>
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
              <div align="center"><a href="../indexA.asp"><img src="../imagens/home.gif" width="19" height="20" border="0"></a></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr bgcolor="#00FF99"> 
      <td height="20" width="6%">&nbsp; </td>
      <td height="20" width="3%"><a href="javascript:Confirma()"><img src="../imagens/confirma_f02.gif" width="24" height="24" border="0"></a></td>
      <td height="20" width="21%"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Enviar</b></font></td>
      <td colspan="2" height="20"> 
        <div align="right"><a href="javascript:Limpa()"><img src="../imagens/limpa_F02.gif" width="24" height="24" border="0"></a></div>
      </td>
      <td height="20" width="39%"><font color="#330099" face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Limpa</b></font></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%">&nbsp;</td>
      <td width="14%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="24%">&nbsp;</td>
      <td width="62%"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#000099">Cadastro 
        de Usu&aacute;rio</font></td>
      <td width="14%">&nbsp;</td>
    </tr>
  </table>
  <table width="101%" border="0" cellspacing="0" cellpadding="4" height="197">
    <tr>
      <td width="2%" height="21">&nbsp;</td>
      <td width="24%" height="21">&nbsp;</td>
      <td width="59%" height="21">&nbsp;</td>
      <td width="15%" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="2%" height="21">&nbsp;</td>
      <td width="24%" height="21">&nbsp;</td>
      <td width="59%" height="21">&nbsp;</td>
      <td width="15%" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="2%" height="25"></td>
      <td width="24%" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Chave 
        do Usu&aacute;rio</b></font></td>
      <td width="59%" height="25"> 
        <input name="txtCdUsuario" type="text" id="txtCdUsuario" size="10" maxlength="4">
      </td>
      <td width="15%" height="25"></td>
    </tr>
    <tr> 
      <td width="2%" height="25">&nbsp;</td>
      <td width="24%" height="25"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Nome 
        do Usu&aacute;rio</b></font></td>
      <td width="59%" height="25"> 
        <input type="text" name="txtNomeUsuario" size="59" maxlength="80">
      </td>
      <td width="15%" height="25">&nbsp;</td>
    </tr>
    <tr> 
      <td width="2%" height="21">&nbsp;</td>
      <td width="24%" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Email</b></font></td>
      <td width="59%" height="21"><input name="txtEmail" type="text" id="txtEmail" size="59" maxlength="80"></td>
      <td width="15%" height="21">&nbsp;</td>
    </tr>
    <tr>
      <td height="21"></td>
      <td height="21" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Categoria</b></font></td>
      <td height="21"><select name="selCategaria" id="selCategaria">
        <option value="A">Cadastra decomposi&ccedil;&atilde;o</option>
        <option value="B">Administrador do sistema</option>
        <option value="C">Cadastra escopo</option>
        <option value="E">Cadastra curso</option>
        <option value="F">Perfil - acesso G1</option>
        <option value="G">Pefil - acesso G3</option>
        <option value="P">Cadastra fun&ccedil;&atilde;o conflitante</option>
      </select></td>
    </tr>
    <tr> 
      <td width="2%" height="21"></td>
      <td width="24%" height="21" valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#330099"><b>Categoria</b></font></td>
      <td width="59%" height="21"><table width="83%" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td width="52%">
            <div align="center"> <b>
              <select name="list1" size="8" multiple>
                <%
				RS_MEGA_PROCESSO = "SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO ORDER BY MEPR_TX_DESC_MEGA_PROCESSO"
				set rs_destino=conn_db.execute("SELECT * FROM " & Session("PREFIXO") & "ACESSO WHERE USUA_CD_USUARIO='" & str_Usuario & "'") 

                 Set RS1 = Conn_db.Execute(RS_MEGA_PROCESSO)
                 DO UNTIL RS1.EOF=TRUE
                 JATEM=0
                 
                 ON ERROR RESUME NEXT
                 RS_DESTINO.MOVEFIRST
                 
                 DO UNTIL RS_DESTINO.EOF=TRUE
                 		IF TRIM(RS1("MEPR_CD_MEGA_PROCESSO"))=TRIM(RS_DESTINO("MEPR_CD_MEGA_PROCESSO")) THEN
                 			JATEM=JATEM+1	
						END IF
						RS_DESTINO.MOVENEXT
                 LOOP
                 	IF JATEM=0 THEN
                 	%>
                <option value="<%=RS1("MEPR_CD_MEGA_PROCESSO")%>" ><%=RS1("MEPR_TX_DESC_MEGA_PROCESSO")%></option>
                <%
                 END IF
  					RS1.MoveNext
					LOOP
					%>
              </select>
          </b></div></td>
          <td width="5%" align="center">
            <table width="53%" cellpadding="0" cellspacing="0" border="0" align="center">
              <tr>
                <td><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image16','','../imagens/continua_F02.gif',1)" onClick="move(document.frm1.list1,document.frm1.list2)"><img name="Image16" border="0" src="../imagens/continua_F01.gif" width="24" height="24"></a></td>
              </tr>
              <tr>
                <td height="25">&nbsp;</td>
              </tr>
              <tr>
                <td height="25"><a href="javascript:;" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('img01511','','../imagens/continua2_F02.gif',1)" onClick="move(document.frm1.list2,document.frm1.list1)"><img name="img01511" border="0" src="../imagens/continua2_F01.gif" width="24" height="24"></a></td>
              </tr>
          </table></td>
          <td width="28%">
            <div align="center"> <font color="#000080">
              <select name="list2" size="8" multiple>
                <%
                RS_DESTINO.MOVEFIRST
                
                DO UNTIL RS_DESTINO.EOF=TRUE
                SSQL1="SELECT * FROM " & Session("PREFIXO") & "MEGA_PROCESSO WHERE MEPR_CD_MEGA_PROCESSO=" & rs_destino("MEPR_CD_MEGA_PROCESSO")
                SET RS_ATUAL=CONN_DB.EXECUTE(SSQL1)
                'RESPONSE.WRITE SSQL1
                ATUAL=RS_ATUAL("MEPR_TX_DESC_MEGA_PROCESSO")
                %>
                <option value=<%=RS_DESTINO("MEPR_CD_MEGA_PROCESSO")%>><%=ATUAL%></option>
                <%
                RS_DESTINO.MOVENEXT
                LOOP
                %>
              </select>
          </font></div></td>
        </tr>
        <tr>
          <td colspan="3">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="3">
            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#003366">Use a tecla Ctrl com o mouse para selecionar mais de uma op&ccedil;&atilde;o ou para desmarcar um item selecionado.</font></div></td>
        </tr>
        <tr>
          <td width="52%">
            <%'=str_SQL_Sub_Proc%>
          </td>
          <td width="5%" align="center">&nbsp;</td>
          <td width="28%">
            <input type="hidden" name="txtMegaSelecionado">
          </td>
        </tr>
      </table>
      </td>
    </tr>
    <tr> 
      <td width="2%" height="21">&nbsp;</td>
      <td width="24%" height="21">&nbsp;</td>
      <td width="59%" height="21"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;
        </font></td>
      <td width="15%" height="21">&nbsp;</td>
    </tr>
    <tr> 
      <td width="2%" height="21">&nbsp;</td>
      <td width="24%" height="21">&nbsp;</td>
      <td width="59%" height="21">&nbsp; </td>
      <td width="15%" height="21">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>
