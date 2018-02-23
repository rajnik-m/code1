<%@ Page Language="vb" AutoEventWireup="false" Codebehind="ShowErrors.aspx.vb" Inherits="CarePortal.ShowErrors" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Error Message Details</title>
    <link href="Styles.css" type="text/css" rel="stylesheet" />
</head>
<body>
    <asp:Label id= "BodyStart" runat ="server"></asp:Label>
    <form id="nfpform" runat="server">
        <div>
            <p>
            </p>
            <div class="ErrorDiv">
                <table class="ErrorTable" border="0" cellspacing="1" cellpadding="5">
                    <tr>
                        <td colspan="2"  class="ErrorHeading">
                            <asp:Label ID="lblError" runat="server"></asp:Label></td>
                    </tr>
                    <tr>
                        <td id="MessageRow" runat="server" class="ErrorLabel">
                            Error Message:</td>
                        <td>
                            <asp:Label ID="lblMessage" CssClass="ErrorItem" runat="server"></asp:Label></td>
                    </tr>
                    <tr id="SourceRow" runat="server">
                        <td class="ErrorLabel">
                            Source:</td>
                        <td>
                            <asp:Label ID="lblSource"  CssClass="ErrorItem" runat="server"></asp:Label></td>
                    </tr>
                    <tr id="LocationRow" runat="server">
                        <td class="ErrorLabel">
                            Location:</td>
                        <td>
                            <asp:Label ID="lblCallStack"  CssClass="ErrorItem" runat="server"></asp:Label></td>
                    </tr>
                    <tr id="HyperlinkRow" runat="server">
                    <td><asp:HyperLink ID="hyp" runat="server" CssClass="ErrorHyperlink" >Return to last page</asp:HyperLink></td></tr>
                </table>
            </div>
        </div>
    </form>
    <asp:Label id= "BodyEnd" runat ="server"></asp:Label>
</body>
</html>
