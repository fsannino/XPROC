<%
'*** PROD ***
strConnCoord = "Provider=SQLOLEDB;Server=s600136;Database=Busca_Sinergia;UID=Busca_SinergiaAdm;PWD=Busca_SinergiaAdm00;"
Session("Conn_String_Cogest_Gravacao") = "Provider=SQLOLEDB.1;server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest"   
Session("Conn_String_SGolive_Gravacao") = "Provider=SQLOLEDB.1;server=s5200db01\db01;pwd=sinergia;uid=ow_sinergia;database=sgolive"
MM_BibliOdbc_STRING = "Provider=SQLOLEDB.1;server=s5200db01\db01;pwd=sinergia;uid=ow_sinergia;database=biblioteca"
%>
