<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="EventSelectionCalendar.ascx.vb" Inherits="CarePortal.EventSelectionCalendar" %>
<div>
    <table class="DataEntryTable" id="tblDataEntry" runat="server" cellspacing="1" cellpadding="3" border="0">    
    </table>
    <asp:Calendar ID="calEventCalendar" runat="server" CssClass="EventSelectionCalendar" 
           ShowGridLines="True">
        <OtherMonthDayStyle BackColor="#E0E0E0" />
        <TodayDayStyle BorderColor="Red" />
        <SelectedDayStyle Font-Bold="True" />
    </asp:Calendar>
</div>


