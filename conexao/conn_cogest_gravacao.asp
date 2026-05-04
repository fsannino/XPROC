<<<<<<< HEAD
<%
if Session("CdUsuario") = "XK45" then
   Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=p024720;pwd=;uid=sa;database=cogest"
else
   Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"
end if
if Session("CdUsuario") = "XD47" then
	Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=p024720;pwd=;uid=sa;database=cogest"
	'Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=127.0.0.1;pwd=;uid=sa;database=cogest"
end if
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=joao;pwd=cogestadm00;uid=cogestadm;database=cogest"
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=192.168.1.3;pwd=;uid=sa;database=cogest"
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=JOAO;pwd=cogestadm00;uid=cogestadm;database=cogest"
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=regina;pwd=cogestadm00;uid=cogestadm;database=cogest"

SUB GRAVA_LOG(COD,TABELA,OPE,CONECTA)

TX_CHAVE_TABELA=COD
TX_TABELA=TABELA
TX_OPERACAO=OPE

SSQL_LOG=""
SSQL_LOG="INSERT INTO " & Session("PREFIXO") & "LOG_GERAL(LOGE_TX_CHAVE_TABELA,LOGE_TX_TABELA,LOGE_DT_DATA_LOG,LOGE_TX_OPERACAO,LOGE_TX_USUARIO) "
SSQL_LOG=SSQL_LOG+"VALUES('" & TX_CHAVE_TABELA & "', "
SSQL_LOG=SSQL_LOG+"'" & TX_TABELA & "', "
SSQL_LOG=SSQL_LOG+"GETDATE(), "
SSQL_LOG=SSQL_LOG+"'" & TX_OPERACAO & "', "
SSQL_LOG=SSQL_LOG+"'" & Session("CdUsuario") & "')"

IF CONECTA=1 THEN
	DB.EXECUTE(SSQL_LOG)
ELSE
	CONN_DB.EXECUTE(SSQL_LOG)
END IF

END SUB

%>
=======
<%
if Session("CdUsuario") = "XK45" then
   Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=p024720;pwd=;uid=sa;database=cogest"
else
   Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=S6000DB11\I6000SQL01;pwd=cogest00;uid=cogest;database=cogest"
end if
if Session("CdUsuario") = "XD47" then
	Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=p024720;pwd=;uid=sa;database=cogest"
	'Session("Session("Conn_String_Cogest_Gravacao")2") = "Provider=SQLOLEDB.1;server=127.0.0.1;pwd=;uid=sa;database=cogest"
end if
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=joao;pwd=cogestadm00;uid=cogestadm;database=cogest"
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=192.168.1.3;pwd=;uid=sa;database=cogest"
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=JOAO;pwd=cogestadm00;uid=cogestadm;database=cogest"
'Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=regina;pwd=cogestadm00;uid=cogestadm;database=cogest"

SUB GRAVA_LOG(COD,TABELA,OPE,CONECTA)

TX_CHAVE_TABELA=COD
TX_TABELA=TABELA
TX_OPERACAO=OPE

SSQL_LOG=""
SSQL_LOG="INSERT INTO " & Session("PREFIXO") & "LOG_GERAL(LOGE_TX_CHAVE_TABELA,LOGE_TX_TABELA,LOGE_DT_DATA_LOG,LOGE_TX_OPERACAO,LOGE_TX_USUARIO) "
SSQL_LOG=SSQL_LOG+"VALUES('" & TX_CHAVE_TABELA & "', "
SSQL_LOG=SSQL_LOG+"'" & TX_TABELA & "', "
SSQL_LOG=SSQL_LOG+"GETDATE(), "
SSQL_LOG=SSQL_LOG+"'" & TX_OPERACAO & "', "
SSQL_LOG=SSQL_LOG+"'" & Session("CdUsuario") & "')"

IF CONECTA=1 THEN
	DB.EXECUTE(SSQL_LOG)
ELSE
	CONN_DB.EXECUTE(SSQL_LOG)
END IF

END SUB

%>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
<font color="#FFFFFF"><b> </b></font>