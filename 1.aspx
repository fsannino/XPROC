<<<<<<< HEAD
<%@ Page Language="VB" %>
<%@ import Namespace="System.Data.SQLClient" %>
<script runat="server">

    dim banco as new sqlconnection("server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest")
    
    Sub Page_Load()
    
        dim dados as new sqlcommand("SELECT * FROM MEGA_PROCESSO",banco)
    
        banco.open()
            datagrid1.datasource=dados.ExecuteReader()
            datagrid1.databind()
        banco.close()
    
    End Sub

</script>
<html>
<head>
</head>
<body>
    <form runat="server">
        <asp:DataGrid id="DataGrid1" runat="server" GridLines="None" BorderWidth="1px" BorderColor="Tan" BackColor="LightGoldenrodYellow" CellPadding="2" ForeColor="Black" Font-Names="Verdana" Font-Size="XX-Small">
            <FooterStyle backcolor="Tan"></FooterStyle>
            <HeaderStyle font-bold="True" backcolor="Tan"></HeaderStyle>
            <PagerStyle horizontalalign="Center" forecolor="DarkSlateBlue" backcolor="PaleGoldenrod"></PagerStyle>
            <SelectedItemStyle forecolor="GhostWhite" backcolor="DarkSlateBlue"></SelectedItemStyle>
            <AlternatingItemStyle backcolor="PaleGoldenrod"></AlternatingItemStyle>
        </asp:DataGrid>
        <!-- Insert content here -->
    </form>
</body>
</html>
=======
<%@ Page Language="VB" %>
<%@ import Namespace="System.Data.SQLClient" %>
<script runat="server">

    dim banco as new sqlconnection("server=S6000DB21;pwd=cogest00;uid=cogest;database=cogest")
    
    Sub Page_Load()
    
        dim dados as new sqlcommand("SELECT * FROM MEGA_PROCESSO",banco)
    
        banco.open()
            datagrid1.datasource=dados.ExecuteReader()
            datagrid1.databind()
        banco.close()
    
    End Sub

</script>
<html>
<head>
</head>
<body>
    <form runat="server">
        <asp:DataGrid id="DataGrid1" runat="server" GridLines="None" BorderWidth="1px" BorderColor="Tan" BackColor="LightGoldenrodYellow" CellPadding="2" ForeColor="Black" Font-Names="Verdana" Font-Size="XX-Small">
            <FooterStyle backcolor="Tan"></FooterStyle>
            <HeaderStyle font-bold="True" backcolor="Tan"></HeaderStyle>
            <PagerStyle horizontalalign="Center" forecolor="DarkSlateBlue" backcolor="PaleGoldenrod"></PagerStyle>
            <SelectedItemStyle forecolor="GhostWhite" backcolor="DarkSlateBlue"></SelectedItemStyle>
            <AlternatingItemStyle backcolor="PaleGoldenrod"></AlternatingItemStyle>
        </asp:DataGrid>
        <!-- Insert content here -->
    </form>
</body>
</html>
>>>>>>> 20204f36c6b9c077038ee81cbf1ea817475c484e
