<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Main.aspx.cs" Inherits="ProjectDownloadExcel.Main" %>


<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <link href="main.css" rel="stylesheet" type="text/css" />
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Panel ID="panel" runat="server">



        </asp:Panel>

        <div></div>
        <asp:TextBox ID="txtData" runat="server"></asp:TextBox>
        
        
    <asp:button ID="btnGerar" runat="server" text="Button" OnClick="btnGerar_Click" />
    </div>
    </form>
    <div id="areaCheckBox">
        

    </div>
</body>
</html>

