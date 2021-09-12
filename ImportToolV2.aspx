<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ImportToolV2.aspx.cs" Inherits="ArasWebImportTool.ImportToolV2" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <h2>Excel BOM表匯入</h2>
    <form id="form1" runat="server">
        <div>
            請選擇檔案: <asp:FileUpload ID="FileUpload_Excel" runat="server" />
        </div>
        <div>
            匯入模式: <asp:DropDownList ID="DropDownList_Model" runat="server">
                <asp:ListItem Value="diff">差異匯入(新增、修改、取代)</asp:ListItem>
                <asp:ListItem Value="all">完整匯入(刪除重建)</asp:ListItem>
            </asp:DropDownList>
        </div>
        </br>
        <div>
            <asp:Button ID="btnImport" runat="server" Text="開始匯入Import" OnClick="btnImport_Click" />
            <asp:Button ID="btnDownloadTemplate" runat="server" Text="下載範本" OnClick="btnDownloadTemplate_Click" />
            <asp:Button ID="btnDownloadTemplate_EN" runat="server" Text="Download English Rev" OnClick="btnDownloadTemplate_EN_Click"  />
        </div>
        <div>
            <h3>BOM 匯入結果</h3>
            <asp:GridView ID="gvBOM" runat="server"></asp:GridView>
        </div>
        
        <div style="border:1px solid;">
            <p>Log Message</p>
            <asp:Label ID="lblLog" runat="server" Text=""></asp:Label>
        </div>
    </form>
</body>
</html>
