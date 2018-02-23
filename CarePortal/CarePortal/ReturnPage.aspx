<%@ Page Language="vb" AutoEventWireup="false" EnableViewState="false" EnableViewStateMac="false" CodeBehind="ReturnPage.aspx.vb" Inherits="CarePortal.ReturnPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Return Page</title>
</head>
<body>
    <form id="nfpForm" runat="server">
    <div>
        <input type ="hidden" name="gatewayFormResponse" value="test for response"/>
        <asp:Label runat="server" ID="lblReturn" Text="Expecting the response"></asp:Label> 
    </div>
    </form>
</body>
</html>
