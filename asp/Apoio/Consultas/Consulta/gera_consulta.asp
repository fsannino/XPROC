<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../conn_consulta.asp" -->
<%
Server.ScriptTimeOut=99999999

set db=server.createobject("ADODB.CONNECTION")
db.Open Session("Conn_String_Cogest_Gravacao")

set fso=server.createobject("Scripting.FileSystemObject")

atribui=request("atrib")
if atribui = 1 THEN
	TIT="Apoiadores Locais"
ELSE
	TIT="Multiplicadores"
END IF

org=request("org")

modo=request("modo")

if org=1 then
	ORGAO = "APOIO_LOCAL_ORGAO"
else
	ORGAO = "APOIO_LOCAL_MULT"
end if

orgao1=request("str01")
orgao2=request("str02")
orgao3=request("str03")
orgao4=request("str04")
orgao5=request("str05")

modulo=request("selModulo_")

if len(modulo)<1 then
	modulo=0
end if

if orgao1=0 and orgao2=0 and orgao3=0 and orgao4=0 and orgao5=0 and modulo<>0 then 

	SSQL=""
	SSQL="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO AS ATRIBUICAO, "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR AS LOTACAO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS MODULO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO "
	SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
	SSQL=SSQL+"dbo.ORGAO_MENOR ON dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND "
	SSQL=SSQL+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) "
	SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "	
	
	set fonte = db.execute(ssql)

else

	conta = 5

	if modo=1 then
		orgao2 = orgao2 & "00000000"
		orgao3 = orgao3 & "00000"
		orgao4 = orgao4 & "00"
	end if
	
	orgao_final=orgao5
	
	if modo=1 then
	
	if len(orgao5)<15 then
		conta=4
		orgao_final=orgao4
	end if

	if len(orgao4)<15 then
		conta=3
		orgao_final=orgao3
	end if

	if len(orgao3)<15 then
		conta=2
		orgao_final=orgao2
	end if

	if len(orgao2)<15 then
		conta=1
		orgao_final=orgao1
	end if
	
	else

	if len(orgao5)<15 then
		conta=4
		orgao_final=orgao4
	end if

	if len(orgao4)<13 then
		conta=3
		orgao_final=orgao3
	end if

	if len(orgao3)<10 then
		conta=2
		orgao_final=orgao2
	end if

	if len(orgao2)<7 then
		conta=1
		orgao_final=orgao1
	end if
	
	
	end if
	
'ssql=""
'ssql="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, "
'ssql=ssql+"dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, "
'ssql=ssql+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO "
'ssql=ssql+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
'ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
'ssql=ssql+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
'if modulo<>0 then		
'	ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND "
'else
'	ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND "
'end if
'if modo=1 then
'	ssql=ssql+"(dbo." & ORGAO & ".ORME_CD_ORG_MENOR = '" & orgao_final & "') "
'else
'	ssql=ssql+"(dbo." & ORGAO & ".ORME_CD_ORG_MENOR LIKE '" & orgao_final & "%') "
'end if
'ssql=ssql+"ORDER BY dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR "	
			
	SSQL=""
	SSQL="SELECT DISTINCT dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO AS CHAVE, dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO AS ATRIBUICAO, "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO AS NOME, dbo.USUARIO_MAPEAMENTO.USUA_TX_RAMAL AS RAMAL, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO, dbo.ORGAO_MENOR.ORME_SG_ORG_MENOR AS LOTACAO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO AS MODULO, "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO "
	SSQL=SSQL+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	SSQL=SSQL+"dbo.USUARIO_MAPEAMENTO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.USUARIO_MAPEAMENTO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	SSQL=SSQL+"dbo.SUB_MODULO ON dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = dbo.SUB_MODULO.SUMO_NR_CD_SEQUENCIA INNER JOIN "
	SSQL=SSQL+"dbo.ORGAO_MENOR ON dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR = dbo.ORGAO_MENOR.ORME_CD_ORG_MENOR INNER JOIN "
	SSQL=SSQL+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	
	SSQL=SSQL+"WHERE (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_MODULO.APLO_NR_ATRIBUICAO = " & atribui & ") AND (dbo.APOIO_LOCAL_ORGAO.APLO_NR_ATRIBUICAO = " & atribui & ")"
	
	if modo=1 then
		SSQL=SSQL+"AND (dbo." & ORGAO & ".ORME_CD_ORG_MENOR = '" & orgao_final & "') AND "
	else
		SSQL=SSQL+"AND (dbo." & ORGAO & ".ORME_CD_ORG_MENOR LIKE '" & orgao_final & "%') AND "
	end if
	
	if modulo<>0 then
		SSQL=SSQL+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) AND (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA IN (" & modulo & ")) "
	else
		SSQL=SSQL+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) "
	end if	
	SSQL=SSQL+"ORDER BY dbo.USUARIO_MAPEAMENTO.USMA_TX_NOME_USUARIO, dbo.SUB_MODULO.SUMO_TX_DESC_SUB_MODULO "	

	set fonte=db.execute(ssql)
		
end if

if fonte.eof=true then

	ssql=""
	ssql="SELECT DISTINCT dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR AS APOIO, dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO, dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO, "
	ssql=ssql+"dbo.APOIO_LOCAL_MULT.ORME_CD_ORG_MENOR, dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA "
	ssql=ssql+"FROM dbo.APOIO_LOCAL_MULT INNER JOIN "
	ssql=ssql+"dbo.APOIO_LOCAL_MODULO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_MODULO.USMA_CD_USUARIO INNER JOIN "
	ssql=ssql+"dbo.APOIO_LOCAL_ORGAO ON dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO = dbo.APOIO_LOCAL_ORGAO.USMA_CD_USUARIO "
	ssql=ssql+"WHERE (dbo.APOIO_LOCAL_MODULO.SUMO_NR_CD_SEQUENCIA = 000 ) AND (dbo.APOIO_LOCAL_MULT.APLO_NR_ATRIBUICAO = " & atribui & ") AND "
	ssql=ssql+"(dbo.APOIO_LOCAL_MULT.APLO_NR_SITUACAO = 1) "
	ssql=ssql+"ORDER BY dbo.APOIO_LOCAL_MULT.USMA_CD_USUARIO, dbo.APOIO_LOCAL_ORGAO.ORME_CD_ORG_MENOR"
	
	set fonte = db.execute(ssql)
	
	achou=0

else
	
	achou=1

end if

%>
<html>
<!--#include file="head.asp" -->
<title>....::::::: Sinergia</title>

<script language="javascript" src="../js/troca_lista2.js"></script>
<!-- #include file = "body.asp" -->
<form name="frm1" method="POST" action="gera_consulta_apoiador_excel.asp" target="blank">

<head>
<title></title>
</head>

<body link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">

  <table width="780" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="18" colspan="4" width="778"><img src="img/_0.gif" width="1" height="1"></td>
    </tr>
    <tr> 
      <td width="55" valign="top"><div align="right"></div></td>
      <td width="96">&nbsp;</td>
      <td width="498"><p><img src="004001.gif" alt=":: Apoiadores Locais" width="205" height="40"></p>
          <%
          chave_ant=0
          igual=0
                  
          do until fonte.eof=true
          
          IF COR="#DDDDDD" THEN
          		COR="WHITE"
          	ELSE
          		COR="#DDDDDD"
          	END IF
          
          
          chave_atual=fonte("chave")
                  
          if trim(chave_ant)<>trim(chave_atual) then
             	igual=igual+1
          end if
			ORG_APOIO=" "
			SET TEMP=DB.EXECUTE("SELECT * FROM ORGAO_MENOR WHERE ORME_CD_ORG_MENOR='" & fonte("apoio") & "'")
			on error resume next
			ORG_APOIO=UCASE(TEMP("ORME_SG_ORG_MENOR"))
			if err.number<>0 then
				SET TEMP=DB.EXECUTE("SELECT * FROM ORGAO_AGLUTINADOR WHERE AGLU_CD_AGLUTINADO='" & fonte("apoio") & "'")
				on error resume next
				ORG_APOIO=UCASE(TEMP("AGLU_SG_AGLUTINADO"))
				if err.number<>0 then
					ORG_APOIO=" "
				END IF
			END IF
          %>
<table width="467" height="70" border="0" cellpadding="0" cellspacing="0" background="img/000028.gif">
          <tr> 
            <td width="10" bgcolor="<%=COR%>"></td>
            <td width="70" rowspan="2" valign="bottom" bgcolor="<%=COR%>"> <div align="right"> 
                <p><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lota&ccedil;&atilde;o:&nbsp;<br>
                  M&oacute;dulo:&nbsp;&nbsp;</strong></font></p>
              </div></td>
            <td width="300" rowspan="2" valign="baseline" bgcolor="<%=COR%>"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ORG_APOIO%><br>
              <strong><%=FONTE("NOME")%></strong><br>
              <%=FONTE("LOTACAO")%><br>
              <%=FONTE("MODULO")%></font></td>
            <td width="87" rowspan="3" bgcolor="<%=COR%>">
            <%
	          TEM_FOTO=FALSE
				CAMINHO = SERVER.MAPPATH("..\FOTOS\"& fonte("CHAVE") & ".jpg")
			  	TEM_FOTO = FSO.FILEEXISTS(CAMINHO)
				       
		       IF TEM_FOTO=TRUE THEN
					FOTO=fonte("CHAVE")
				ELSE
					FOTO="SEM_FOTO"				      
				END IF
		      %>
            <img src="../fotos/<%=FOTO%>.jpg" width="62" height="62">
            
            </td>
          </tr>
          <tr> 
            <td width="10" bgcolor="<%=COR%>">&nbsp;</td>
          </tr>
          <tr> 
            <td width="10" bgcolor="<%=COR%>">&nbsp;</td>
            <td width="70" bgcolor="<%=COR%>"><div align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>chave 
                | Tel.:</strong></font></div></td>
            <td width="300" valign="middle" bgcolor="<%=COR%>"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=FONTE("CHAVE")%> | <%=FONTE("RAMAL")%></font></td>
          </tr>
        </table>
        <P>
          <%
          TEM=TEM+1
          chave_ANT=fonte("chave")
          FONTE.MOVENEXT
          LOOP
          %>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <%if tem<>0 then%>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif"><b>Total 
          de Registros Encontrados : <%=TEM%></b></font></p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif"><b>Total
        de Apoiadores Encontrados : <%=IGUAL%></b></font></p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p>
        <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"><b><font color="#666666" size="2" face="Georgia, Times New Roman, Times, serif"><img border="0" src="000027.gif" width="26" height="23" alt="Imprimir" onClick="javascript:print()"></font></b></p>
        <strong><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif"><br>
        <%else%>
        <b><font size="2" face="Georgia, Times New Roman, Times, serif" color="#800000">
        Nenhum Registro Encontrado</font></b>
        <%end if%>
        <img src="000024.gif" width="73" height="16" align="right" onClick="javascript:history.go(-1)"> </font></strong>
        <p>&nbsp;</p>
        <p><img src="000025.gif" width="467" height="1"></p>
        <p align="center"><font color="#666666" size="1" face="Verdana, Arial, Helvetica, sans-serif">© 
          2003 Sinergia | A Petrobras integrada rumo ao futuro</font></p>
        <p align="center">&nbsp;</p>
      </td>
      <td width="123">&nbsp;</td>
    </tr>
  </table>
</form>

</html>
