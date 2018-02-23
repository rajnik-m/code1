<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="DatePicker.aspx.vb" Inherits="CarePortal.DatePicker" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Date Picker</title>
    <link href="Styles.css" type="text/css" rel="stylesheet" />
</head>
  <body class = "DatePicker">
    <form id="nfpform" runat="server" class="styles">
    <div>
      <script type="text/javascript" language="javascript" src="Scripts/ClientScript.js"></script>
      <asp:Calendar id="calCalendar" runat="server" CssClass="DatePickerHeader" Width="100%" Height="100%">
          <TitleStyle CssClass="DatePickerTitle" />   
          <OtherMonthDayStyle CssClass="OtherMonthDayStyle"/>
          <TodayDayStyle CssClass="TodayDayStyle"/>
          <SelectedDayStyle CssClass="SelectedDayStyle"/>
      </asp:Calendar>
    </div>
    </form>
</body>
</html>
