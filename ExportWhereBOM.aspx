<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ExportWhereBOM.aspx.cs" Inherits="ArasWebImportTool.ExportWhereBOM" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:FileUpload ID="FileUpload_Excel" runat="server" />
        </div>
        <div>
            <asp:Button ID="btnExport" runat="server" Text="Export" OnClick="btnExport_Click" />
            <asp:Button ID="btnDownloadTemplate" runat="server" Text="下載範本" OnClick="btnDownloadTemplate_Click" />
            <asp:Button ID="btnDownloadTemplate_EN" runat="server" Text="Download Template" OnClick="btnDownloadTemplate_EN_Click" />
        </div>
         <div style="border:1px solid;">
            <p>Log Message</p>
            <asp:Label ID="lblLog" runat="server" Text=""></asp:Label>
        </div>
    </form>
</body>
</html>
