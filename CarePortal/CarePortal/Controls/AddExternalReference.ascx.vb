Public Partial Class AddExternalReference
  Inherits CareWebControl
  Implements ICareChildWebControl

  Private mvDataSource As String = ""
  Private mvContactNumber As Integer
  Private mvExternalReference As String = ""


  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    Try
      mvUsesHiddenContactNumber = True
      mvNeedsParent = True
      mvHandlesExtReferences = True
      InitialiseControls(CareNetServices.WebControlTypes.wctAddExternalReference, tblDataEntry, "", "")
      AddHiddenField("OldExternalReference")
      AddHiddenField("OldDataSource")
    Catch vEX As ThreadAbortException
      Throw vEX
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      If Request.QueryString("DS") IsNot Nothing AndAlso Request.QueryString("DS").Length > 0 Then mvDataSource = Request.QueryString("DS")
      If Request.QueryString("ER") IsNot Nothing AndAlso Request.QueryString("ER").Length > 0 Then mvExternalReference = Request.QueryString("ER")
      If Request.QueryString("CN") IsNot Nothing AndAlso Request.QueryString("CN").Length > 0 Then mvContactNumber = IntegerValue(Request.QueryString("CN"))
      If mvDataSource.Length > 0 AndAlso mvExternalReference.Length > 0 AndAlso mvContactNumber > 0 Then
        'We have all the required data so just submit this page and go to the submit page
        SubmitChild(Nothing)
        GoToSubmitPage()
      End If
    Catch vEX As ThreadAbortException
      Throw vEX
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ClearControls()
    MyBase.ClearControls()
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild

    If mvDataSource.Length = 0 Then
      'DataSource may be hidden, if so use default value
      mvDataSource = GetDropDownValue("DataSource")
      If mvDataSource.Length = 0 Then mvDataSource = DefaultParameters("DataSource").ToString
    End If
    If mvExternalReference.Length = 0 Then mvExternalReference = GetTextBoxText("ExternalReference")
    If mvContactNumber = 0 AndAlso pList IsNot Nothing Then mvContactNumber = IntegerValue(pList("ContactNumber").ToString)
    Dim vAllowMaintenance As Boolean = DefaultParameters.Contains("AllowMaintenance") AndAlso BooleanValue(DefaultParameters("AllowMaintenance").ToString)
    Dim vAdd As Boolean = Not vAllowMaintenance
    Dim vDelete As Boolean
    Dim vUpdate As Boolean
    If vAllowMaintenance Then
      If GetHiddenText("OldExternalReference").Length = 0 OrElse GetHiddenContactNumber() <> mvContactNumber Then
        'Add new record
        vAdd = True
      ElseIf mvExternalReference.Length = 0 Then
        'Record is removed
        vDelete = True
      ElseIf GetHiddenText("OldExternalReference") = mvExternalReference AndAlso GetHiddenText("OldDataSource") = mvDataSource  Then
        'Nothing is changed. 
      Else
        'Data has been changed
        vUpdate = True
      End If
    Else
      vAdd = mvExternalReference.Length > 0
    End If
    If vAdd OrElse vUpdate OrElse vDelete Then
      Dim vList As New ParameterList(HttpContext.Current)

      If vDelete Then
        vList("DataSource") = GetHiddenText("OldDataSource")
        vList("ExternalReference") = GetHiddenText("OldExternalReference")
        vList("ContactNumber") = mvContactNumber
        DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctReference, vList)
      Else
        vList("DataSource") = mvDataSource
        vList("ExternalReference") = mvExternalReference
        vList("UseContactRestriction") = "N"  'Pass this parameter to check duplicate for ExternalReference and DataSource only
        'check for duplicate
        Dim vExtReferenceTable As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactExternalReferences, mvContactNumber, vList)
        If vExtReferenceTable Is Nothing Then
          vList.Remove("UseContactRestriction")
          vList("ContactNumber") = mvContactNumber
          If vAdd Then
            DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctReference, vList)
          Else
            vList("OldExternalReference") = GetHiddenText("OldExternalReference")
            vList("OldDataSource") = GetHiddenText("OldDataSource")
            DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctReference, vList)
          End If
        End If
      End If
    End If
  End Sub
  Public Overrides Sub ProcessExtReferenceSelection(ByVal pTable As DataTable)
    If pTable IsNot Nothing AndAlso pTable.Rows.Count > 0 Then
      Dim vExternalRef As String = ""
      Dim vDataSource As String = ""
      If DefaultParameters.Contains("DisplayFirstRecord") AndAlso BooleanValue(DefaultParameters("DisplayFirstRecord").ToString) Then
        If DefaultParameters.Contains("DataSource") AndAlso DefaultParameters("DataSource").ToString.Length > 0 Then
          'Only display the first record that has a Data Source same as the default parameter Data Source
          For Each vRow As DataRow In pTable.Rows
            If vRow("DataSource").ToString = DefaultParameters("DataSource").ToString Then
              vExternalRef = vRow("ExternalReference").ToString
              vDataSource = vRow("DataSource").ToString
              Exit For
            End If
          Next
        Else
          'Display the first record only
          vExternalRef = pTable.Rows(0)("ExternalReference").ToString
          vDataSource = pTable.Rows(0)("DataSource").ToString
        End If
      Else
        Dim vAmendedOn As Date = DateValue((pTable.Rows(0)("AmendedOn").ToString))
        For Each vRow As DataRow In pTable.Rows
          If DateValue(vRow("AmendedOn").ToString) >= vAmendedOn Then
            vExternalRef = vRow("ExternalReference").ToString
            vDataSource = vRow("DataSource").ToString
          End If
        Next
      End If
      If vExternalRef.Length > 0 Then
        SetTextBoxText("ExternalReference", vExternalRef)
        SetHiddenText("OldExternalReference", vExternalRef)
        SetDropDownText("DataSource", vDataSource)
        SetHiddenText("OldDataSource", vDataSource)
      End If
    End If
  End Sub
End Class