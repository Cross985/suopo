<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="WebApplication1._Default" %>

<%@ Register assembly="DevExpress.XtraReports.v8.1.Web" namespace="DevExpress.XtraReports.Web" tagprefix="dxxr" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>数据导入</title>
    <link href="style/default.css" rel="stylesheet" type="text/css" />
</head>
<body>
<form id="form1" runat="server">
    <div class="TABOFF" style="text-align: left; padding-left: 0px; margin-left: 0px;">
    <table border="0" cellspacing="0" cellpadding="0">
         <tbody>
         <tr>
         <td class="TABON" nowrap="nowrap">店铺导入</td>
         </tr>
         </tbody>
    </table>
    </div>
         <div class="PANEREPEAT">  请选择需要导入的店铺文档 </div>
    
        <asp:FileUpload ID="FileUpload1" runat="server" Width="305px" />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" Text="导入" />
    
        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" 
            Text="店铺列表" />
        
    </form>
</body>
</html>
