Public Class UpdateDirectoryPreferences
  Inherits CareWebControl

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try      
      InitialiseControls(CareNetServices.WebControlTypes.wctDirectoryPreferences, tblDataEntry, "", "")

      If IsPostBack Then Exit Sub

      If UserContactNumber() > 0 Then
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ContactNumber") = UserContactNumber()
        vList("CarePortal") = "Y"
        If Not vList.Contains("WPD") Then vList.Add("WPD", "Y")
        If InitialParameters("AddressUsage") IsNot Nothing Then
          vList("AddressUsage") = InitialParameters("AddressUsage").ToString
          PopulateDDL("DirectoryAddress", "ADDRESS", "ADDRESS")
        End If

        If InitialParameters("CommunicationUsage1") IsNot Nothing Then
          vList("CommunicationUsage1") = InitialParameters("CommunicationUsage1").ToString
          PopulateDDL("CommunicationUsage1", "Device1", "Device2")
        End If

        If InitialParameters("CommunicationUsage2") IsNot Nothing Then
          vList("CommunicationUsage2") = InitialParameters("CommunicationUsage2").ToString
          PopulateDDL("CommunicationUsage2", "Device3", "Device4")
        End If

        If InitialParameters("CommunicationUsage3") IsNot Nothing Then
          vList("CommunicationUsage3") = InitialParameters("CommunicationUsage3").ToString
          PopulateDDL("CommunicationUsage3", "Device5", "Device6")
        End If

        If InitialParameters("CommunicationUsage4") IsNot Nothing Then
          vList("CommunicationUsage4") = InitialParameters("CommunicationUsage4").ToString
          PopulateDDL("CommunicationUsage4", "Device7", "Device8")
        End If

        Dim vDirectoryData As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtDirectoryUsage, vList)

        SetComboBoxItems("AddressUsage", vDirectoryData, "DirectoryAddress", "", "")
        SetComboBoxItems("CommunicationUsage1", vDirectoryData, "CommunicationUsage1", "Device1", "Device2")
        SetComboBoxItems("CommunicationUsage2", vDirectoryData, "CommunicationUsage2", "Device3", "Device4")
        SetComboBoxItems("CommunicationUsage3", vDirectoryData, "CommunicationUsage3", "Device5", "Device6")
        SetComboBoxItems("CommunicationUsage4", vDirectoryData, "CommunicationUsage4", "Device7", "Device8")
      End If

    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Private Sub SetComboBoxItems(ByVal pParameterName As String, ByVal pTable As DataTable, ByVal pDDL As String, ByVal pBuisiness As String, ByVal pPrivate As String)
    If InitialParameters.ContainsKey(pParameterName) AndAlso InitialParameters(pParameterName).ToString.Length > 0 Then
      Dim vParamValue As String = InitialParameters(pParameterName).ToString
      Dim vAddressDDL As DropDownList = TryCast(FindControlByName(Me, pDDL), DropDownList)

      If pTable IsNot Nothing Then
        For Each vDr As DataRow In pTable.Rows
          If vDr("CommunicationUsage").ToString = vParamValue Then
            If InitialParameters(pBuisiness) IsNot Nothing AndAlso vDr("Device").ToString = InitialParameters(pBuisiness).ToString Then
              vAddressDDL.SelectedValue = "O"
            ElseIf InitialParameters(pPrivate) IsNot Nothing AndAlso vDr("Device").ToString = InitialParameters(pPrivate).ToString Then
              vAddressDDL.SelectedValue = "C"
            Else
              vAddressDDL.SelectedValue = "N"
            End If

            Exit For
          ElseIf vDr("AddressUsage").ToString = vParamValue Then
            If vDr("AddressType").ToString = "O" Then
              vAddressDDL.SelectedValue = "O"
            ElseIf vDr("AddressType").ToString = "C" Then
              vAddressDDL.SelectedValue = "C"
            Else
              vAddressDDL.SelectedValue = "N"
            End If

            Exit For
          End If
        Next
      End If

    Else
      FindControlByName(Me, pDDL).Parent.Parent.Visible = False
      FindControlByName(Me, pDDL).Visible = False
    End If

  End Sub

  Private Sub PopulateDDL(ByRef pControlName As String, ByRef pBusinessDevice As String, ByRef pPrivateDevice As String)
    Dim vAddressDDL As DropDownList = TryCast(FindControlByName(Me, pControlName), DropDownList)
    If vAddressDDL IsNot Nothing Then
      If pPrivateDevice = "ADDRESS" Or InitialParameters(pPrivateDevice) IsNot Nothing Then vAddressDDL.Items.Add(New ListItem("Private", "C"))
      If pBusinessDevice = "ADDRESS" Or InitialParameters(pBusinessDevice) IsNot Nothing Then vAddressDDL.Items.Add(New ListItem("Business", "O"))
      vAddressDDL.Items.Add(New ListItem("None", "N"))
      vAddressDDL.SelectedValue = "N"
    End If
  End Sub

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    If IsValid() Then
      Try
        Dim vParams As New ParameterList(HttpContext.Current)

        If InitialParameters("AddressUsage") IsNot Nothing Then
          Dim vAddressDDL As DropDownList = TryCast(FindControlByName(Me, "DirectoryAddress"), DropDownList)
          vParams("AddressUsage") = InitialParameters("AddressUsage")
          If vAddressDDL.SelectedValue = "C" Then vParams("AddressType") = "C"
          If vAddressDDL.SelectedValue = "O" Then vParams("AddressType") = "O"
        End If

        For vI As Integer = 1 To 4
          If InitialParameters("CommunicationUsage" & vI.ToString) IsNot Nothing Then
            Dim vAddressDDL As DropDownList = TryCast(FindControlByName(Me, "CommunicationUsage" & vI.ToString), DropDownList)
            vParams("CommunicationUsage" & vI.ToString) = InitialParameters("CommunicationUsage" & vI.ToString)

            Dim vIDX As Integer = vI * 2
            If vAddressDDL.SelectedValue = "C" Then vParams("Device" & vI.ToString) = InitialParameters("Device" & vIDX.ToString)
            If vAddressDDL.SelectedValue = "O" Then vParams("Device" & vI.ToString) = InitialParameters("Device" & (vIDX - 1).ToString)
            ' work around if a param is not set, ensure a value (not null) is sent.
            If vParams("Device" & vI.ToString) Is Nothing Then vParams("Device" & vI.ToString) = ""
          End If
        Next

        vParams("ContactNumber") = UserContactNumber()

        Dim vResultList As ParameterList = DataHelper.UpdateDirectoryPreferences(vParams)
        GoToSubmitPage()
      Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
    End If
  End Sub

End Class
