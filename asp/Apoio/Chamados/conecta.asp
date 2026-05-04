<%
Set Db = Server.CreateObject("AdoDb.Connection")

Db.Open "Provider=SQLOLEDB.1;server=S6000DB02\S6000DB08;pwd=cogest001;uid=cogest;database=cogest"
'Db.Open "Provider=SQLOLEDB.1;server=s5200db01\db01;pwd=sinergia;uid=ow_sinergia;database=sgolive"

Db.CursorLocation = 3
Session.LCID = 1046
%>