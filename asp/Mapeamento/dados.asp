<%
Set objUSR = Server.CreateObject("Seseg.Usuario")
%>
<p>Instância criada
<%
if not objUSR.GetUsuario then
    Response.write objUSR.CodErro
    Response.Write objUSR.MsgErro
else
%>
    <p>CHAVE=<%Response.Write objUSR.SEI_CHAVE%>
    <p>MATRICULA=<%Response.Write objUSR.SEI_MATRICULA%>
    <p>NOME=<%Response.Write objUSR.SEI_NOME%>
    <p>CARGO=<%Response.Write objUSR.SEI_CARGO%>
    <p>GERENTE=<%Response.Write objUSR.SEI_GERENTE%>
    <p>LOTACAO=<%Response.Write objUSR.SEI_LOTACAO%>
    <p>HOST=<%Response.Write objUSR.HOST%>
    <p>VERSAO=<%Response.Write objUSR.VERSAO%><%
end if
Set objUSR = nothing
%>