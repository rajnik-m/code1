<%@ Page Language="vb" AutoEventWireup="false" EnableViewState="True" CodeBehind="Default.aspx.vb"
    Inherits="CarePortal._Default" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Care Portal</title>
    <link href="Styles.css" type="text/css" rel="stylesheet" />
</head>
<body>
    <asp:Literal ID="BodyStart" runat="server"></asp:Literal>
    &nbsp;&nbsp;&nbsp;
    <form id="nfpform" runat="server" class="Form" autocomplete="<%$appSettings:AutoComplete%>">
    <script language="javascript" type="text/javascript" src="Scripts/ClientScript.js"></script>
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <script language="javascript" type="text/javascript">
      //<![CDATA[
        //The following script should not be added in ClientScript.js
        Sys.WebForms.PageRequestManager.getInstance().add_endRequest(EndRequestHandler);
        function EndRequestHandler(sender, args) {
            try {
                if ($get('LastControl').value.length > 0) {
                    var elem = $get($get('LastControl').value);
                    elem.focus();
                    if (elem.type == 'text' || elem.type == 'textarea') {
                        var Range = document.selection.createRange();
                        if (Range.offsetLeft != 0 && Range.offsetRight != 0)  //make sure the selection range/focus exists
                        {
                            var OriginalValue = elem.value;
                            var OriginalMaxLength = elem.maxLength;
                            elem.maxLength = 1000;
                            var specialchar = String.fromCharCode(1) + "scf";   //get a unique word
                            //Always append the specialchar manually to the RangeText as using Range.text = specialchar; can clear the value of the field
                            //when 1. the field is to be recreated (using AJAX) 2. the value is not to be updated 3. the value is selected(highlighted)
                            Range.text = Range.text + specialchar;
                            var pos = elem.value.indexOf(specialchar);
                            elem.value = elem.value.replace(specialchar, "");
                            elem.maxLength = OriginalMaxLength;
                            Range = elem.createTextRange();
                            if (pos <= 0) {
                                pos = OriginalValue.length * 2
                            }
                            //may add more code to handle carriage returns because the pos is incorrectly increased/decreased if the position is after carriage return
                            Range.move('character', pos);
                            Range.select();
                        }
                        else  //selection range/focus is not found as the control was recreated using AJAX
                        {
                            Range = elem.createTextRange();
                            //may add more code to handle carriage returns because the pos is incorrectly increased/decreased if the position is after carriage return
                            Range.move('character', elem.value.length);
                            Range.select();
                        }
                    }
                    else if (elem.type == 'checkbox' || elem.type == 'radio' || elem.type == 'select-one' || elem.type == 'select-multiple' || elem.type == 'submit') {
                        //a hack to to keep the focus in IE. This may not always work but it does most of the time
                        setTimeout("$get($get('LastControl').value).focus();", 0);
                    }
                }
            }
            catch (err)
                { }
        }
      //]]>
    </script>
    <table class="MainTable" id="MainTable" runat="server">
        <tr class="MainTableRow">
            <td id="SiteHeader" runat="server" colspan="3" align="center">
            </td>
        </tr>
        <tr class="MainTableRow">
            <td class="MenuCell" colspan="3" align="left">
                <asp:PlaceHolder ID="phcMenu" runat="server"></asp:PlaceHolder>
            </td>
        </tr>
        <tr class="MainTableRow">
            <td id="SiteLeftPanel" runat="server" align="center">
            </td>
            <td>
                <table class="InnerTable" id="CenterTable" runat="server">
                    <tr class="InnerTableRow">
                        <td id="HeadingData" runat="server" class="HeadingCell" colspan="3">
                        </td>
                    </tr>
                    <tr id="RowData" runat="server" class="InnerTableRow">
                        <td id="LeftData" runat="server" class="LeftCell">
                        </td>
                        <td id="CenterData" runat="server" class="CenterCell">
                        </td>
                        <td id="RightData" runat="server" class="RightCell">
                        </td>
                    </tr>
                    <tr class="InnerTableRow">
                        <td id="FootingData" runat="server" class="FootingCell" colspan="3">
                        </td>
                    </tr>
                </table>
            </td>
            <td id="SiteRightPanel" runat="server" align="center">
            </td>
        </tr>
        <tr class="MainTableRow">
            <td id="SiteFooter" runat="server" colspan="3" align="center">
            </td>
        </tr>
    </table>
    <asp:HiddenField ID="LastControl" runat="server" />
    </form>
    <asp:Literal ID="BodyEnd" runat="server"></asp:Literal>
</body>
</html>
