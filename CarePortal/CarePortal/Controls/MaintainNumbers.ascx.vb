Partial Public Class MaintainNumbers
  Inherits CareWebControl

  Private mvNumbers(5) As NumberInfo

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Overloads Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctMaintainNumbers, tblDataEntry, "", "DirectNumber,SwitchboardNumber,MobileNumber,EMailAddress,FaxNumber,WebAddress")
      mvNumbers(0) = New NumberInfo("DirectNumber", DataHelper.ControlValue(DataHelper.ControlValues.direct_device), 20)
      mvNumbers(1) = New NumberInfo("SwitchboardNumber", DataHelper.ControlValue(DataHelper.ControlValues.switchboard_device), 20)
      mvNumbers(2) = New NumberInfo("MobileNumber", DataHelper.ControlValue(DataHelper.ControlValues.mobile_device), 20)
      mvNumbers(3) = New NumberInfo("EMailAddress", DataHelper.ControlValue(DataHelper.ControlValues.email_device), 128)
      mvNumbers(4) = New NumberInfo("FaxNumber", DataHelper.ControlValue(DataHelper.ControlValues.fax_device), 20)
      mvNumbers(5) = New NumberInfo("WebAddress", DataHelper.ControlValue(DataHelper.ControlValues.web_device), 128)
      For Each vNumber As NumberInfo In mvNumbers
        If Me.FindControl(vNumber.Identifier) IsNot Nothing Then DirectCast(Me.FindControl(vNumber.Identifier), TextBox).MaxLength = vNumber.MaxLength
      Next
      Dim vTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactCommsNumbers, GetContactNumberFromParentGroup())
      If Not vTable Is Nothing Then
        For Each vRow As DataRow In vTable.Rows
          Dim vDeviceCode As String = vRow.Item("DeviceCode").ToString
          For Each vNumber As NumberInfo In mvNumbers
            If vDeviceCode = vNumber.DeviceCode Then
              Dim vExt As String = vRow.Item("Extension").ToString
              Dim vPhoneNumber As String = vRow.Item("PhoneNumber").ToString
              If vExt.Length > 0 AndAlso vPhoneNumber.EndsWith(vExt) Then
                Dim vPos As Integer = InStr(vPhoneNumber, " Ext ")
                If vPos > 0 Then vPhoneNumber = vPhoneNumber.Substring(0, vPos - 1)
              End If
              If Not IsPostBack Then SetTextBoxText(vNumber.Identifier, vPhoneNumber)
              vNumber.CommunicationNumber = IntegerValue(vRow.Item("CommunicationNumber").ToString)
              Exit For
            End If
          Next
        Next
      End If
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        SaveContactCommNumbers(mvNumbers, GetContactNumberFromParentGroup, UserAddressNumber)
        GoToSubmitPage()
      Catch vEx As ThreadAbortException
        Throw vEx
      Catch vException As Exception
        ProcessError(vException)
      End Try
    End If
  End Sub
End Class