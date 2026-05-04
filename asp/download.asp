<<<<<<< HEAD
<%
arquivo=request("arquivo")
Response.buffer=true
set fso=server.createobject("Scripting.FileSystemObject")
set file = fso.getfile(arquivo)
set ts=fso.opentextfile(file)
response.ContentType="application/unknown"
response.AddHeader "Content-Disposition","Attachment; filename=" & arquivo
response.BinaryWrite ts.readall & ""
ts.close
set ts=nothing
set file=nothing
set fso=nothing
=======
<%
arquivo=request("arquivo")
Response.buffer=true
set fso=server.createobject("Scripting.FileSystemObject")
set file = fso.getfile(arquivo)
set ts=fso.opentextfile(file)
response.ContentType="application/unknown"
response.AddHeader "Content-Disposition","Attachment; filename=" & arquivo
response.BinaryWrite ts.readall & ""
ts.close
set ts=nothing
set file=nothing
set fso=nothing
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
%>