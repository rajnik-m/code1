<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="SurveyResponses.ascx.vb" Inherits="CarePortal.SurveyResponses" %>
<div>
<asp:UpdatePanel runat="server" id="updatPanel">
  <ContentTemplate>
    <table class="DataEntryTable" id="tblDataEntry" runat="server" cellspacing="1" cellpadding="3" border="0">
    </table>
  </ContentTemplate>
</asp:UpdatePanel>
</div>
