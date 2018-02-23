Imports System.Reflection

Partial Public Class AddAction
  Inherits CareWebControl
  Implements ICareChildWebControl

  Private mvActionNumber As Integer

  Public Sub New()
    'mvNeedsAuthentication = True
  End Sub

  Private Sub Page_Init(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Init
    mvNeedsParent = True
    Try
      InitialiseControls(CareNetServices.WebControlTypes.wctAddAction, tblDataEntry)
      AddHiddenField("OldActioner")
      AddHiddenField("OldTopic")
      AddHiddenField("OldSubTopic")
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Try
      If Request.QueryString("AC") IsNot Nothing Then mvActionNumber = IntegerValue(Request.QueryString("AC"))
      SetDefaults()
    Catch vEx As ThreadAbortException
      Throw vEx
    Catch vException As Exception
      ProcessError(vException)
    End Try
  End Sub

  Public Overrides Sub ClearControls()
    MyBase.ClearControls()
    mvActionNumber = 0
    SetDefaults()
  End Sub

  Private Sub SetDefaults()
    If Not IsPostBack Then
      If mvActionNumber > 0 Then
        'Need to display an existing Action
        Dim vList As New ParameterList(HttpContext.Current)
        vList("ActionNumber") = mvActionNumber
        vList("ContactNumber") = Session("CurrentContactNumber").ToString
        Dim vDR As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactActions, vList))
        If vDR IsNot Nothing Then
          ProcessActionSelection(vDR)
        Else
          mvActionNumber = 0
          If IntegerValue(Request.QueryString("AC")) > 0 Then
            Dim vIsReadOnly As PropertyInfo = GetType(System.Collections.Specialized.NameValueCollection).GetProperty("IsReadOnly", BindingFlags.Instance Or BindingFlags.NonPublic)
            ' Make collection editable
            vIsReadOnly.SetValue(Me.Request.QueryString, False, Nothing)
            ' Remove
            Me.Request.QueryString.Remove("AC")
            ProcessRedirect("default.aspx?" & Me.Request.QueryString.ToString)
          End If
        End If
      Else
        SetDropDownText("DocumentClass", InitialParameters("DocumentClass").ToString)
        SetDropDownText("ActionPriority", InitialParameters("ActionPriority").ToString)
        If Session.Contents.Item("UserContactNumber") IsNot Nothing AndAlso IntegerValue(Session("UserContactNumber").ToString) > 0 Then SetDropDownText("Actioner", Session("UserContactNumber").ToString)
      End If
    End If
  End Sub

  Public Sub SubmitChild(ByVal pList As ParameterList) Implements ICareChildWebControl.SubmitChild
    Dim vContactNumber As Integer = IntegerValue(pList("ContactNumber").ToString)
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vActionDesc As String = GetTextBoxText("ActionDesc")
    Dim vActionText As String = GetTextBoxText("ActionText")
    If vActionText.Length = 0 AndAlso vActionDesc.Length > 0 Then vActionText = vActionDesc
    If vActionDesc.Length = 0 AndAlso vActionText.Length > 0 Then vActionDesc = vActionText
    vList("ActionDesc") = vActionDesc
    vList("ActionText") = vActionText
    Dim vActionPriority As String = GetDropDownValue("ActionPriority")
    If vActionPriority.Length = 0 Then vActionPriority = InitialParameters("ActionPriority").ToString
    Dim vDocumentClass As String = GetDropDownValue("DocumentClass")
    If vDocumentClass.Length = 0 Then vDocumentClass = InitialParameters("DocumentClass").ToString
    vList("ActionPriority") = vActionPriority
    vList("DocumentClass") = vDocumentClass

    If vActionDesc.Trim.Length = 0 OrElse vActionText.Trim.Length = 0 Then
      'Do not save if these fields are empty
      Exit Sub
    End If

    Dim vDurationDays As String = GetTextBoxText("DurationDays")
    Dim vDurationHours As String = GetTextBoxText("DurationHours")
    Dim vDurationMinutes As String = GetTextBoxText("DurationMinutes")
    If vDurationDays.Length = 0 AndAlso vDurationHours.Length = 0 AndAlso vDurationMinutes.Length = 0 Then
      vDurationDays = InitialParameters("DurationDays").ToString
      vDurationHours = InitialParameters("DurationHours").ToString
      vDurationMinutes = InitialParameters("DurationMinutes").ToString
    End If
    If IntegerValue(vDurationDays) = 0 Then vDurationDays = ""
    If IntegerValue(vDurationHours) = 0 Then vDurationHours = ""
    If IntegerValue(vDurationMinutes) = 0 Then vDurationMinutes = ""
    vList("DurationDays") = vDurationDays
    vList("DurationHours") = vDurationHours
    vList("DurationMinutes") = vDurationMinutes
    AddOptionalTextBoxValue(vList, "ScheduledOn")
    AddOptionalTextBoxValue(vList, "Deadline")
    AddOptionalTextBoxValue(vList, "CompletedOn")

    Dim vReturnList As ParameterList
    If mvActionNumber > 0 Then
      'Update
      vList("ActionNumber") = mvActionNumber
      vReturnList = DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctAction, vList)
      'Links
      Dim vOldActioner As Integer = IntegerValue(GetHiddenText("OldActioner"))
      Dim vActioner As Integer = IntegerValue(GetDropDownValue("Actioner"))
      If vOldActioner <> vActioner Then
        'Delete old ActionLink
        Dim vLink As New ParameterList(HttpContext.Current)
        vLink("ContactNumber") = vOldActioner
        vLink("ActionNumber") = mvActionNumber
        vLink("ActionLinkType") = "A"
        DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctActionLink, vLink)
        'Create new ActionLink
        vLink("ContactNumber") = vActioner
        DataHelper.AddActionLink(vLink)
      End If
      'Topics
      Dim vOldTopic As String = GetHiddenText("OldTopic")
      Dim vOldSubTopic As String = GetHiddenText("OldSubTopic")
      Dim vTopic As String = GetDropDownValue("Topic")
      If vTopic.Length = 0 AndAlso InitialParameters.ContainsKey("Topic") Then vTopic = InitialParameters("Topic").ToString
      Dim vSubTopic As String = GetDropDownValue("SubTopic")
      If ((vOldTopic <> vTopic) Or (vOldSubTopic <> vSubTopic)) _
      AndAlso (vTopic.Length > 0 AndAlso vSubTopic.Length > 0) Then
        'Delete old Topic/SubTopic
        Dim vAnalysis As New ParameterList(HttpContext.Current)
        vAnalysis("ActionNumber") = mvActionNumber
        vAnalysis("Topic") = vOldTopic
        vAnalysis("SubTopic") = vOldSubTopic
        DataHelper.DeleteItem(CareNetServices.XMLMaintenanceControlTypes.xmctActionTopic, vAnalysis)
        'Add new Topic/SubTopic
        vAnalysis("Topic") = vTopic
        vAnalysis("SubTopic") = vSubTopic
        DataHelper.AddActionSubject(vAnalysis)
      End If
    Else
      'Create
      vReturnList = DataHelper.AddAction(vList)
      'Add ActionLink
      Dim vLink As New ParameterList(HttpContext.Current)
      vLink("ContactNumber") = vContactNumber
      vLink("ActionNumber") = vReturnList("ActionNumber")
      vLink("ActionLinkType") = "R"
      DataHelper.AddActionLink(vLink)
      Dim vActioner As String = GetDropDownValue("Actioner")
      If vActioner.Length > 0 Then
        vLink("ContactNumber") = vActioner
        vLink("ActionLinkType") = "A"
        DataHelper.AddActionLink(vLink)
      End If
      'Add Topic/SubTopic
      Dim vSubTopic As String = GetDropDownValue("SubTopic")
      If vSubTopic.Length > 0 Then
        Dim vTopic As String = GetDropDownValue("Topic")
        If vTopic.Length = 0 AndAlso InitialParameters.ContainsKey("Topic") Then vTopic = InitialParameters("Topic").ToString
        If vTopic.Length > 0 Then
          Dim vAnalysis As New ParameterList(HttpContext.Current)
          vAnalysis("ActionNumber") = vReturnList("ActionNumber")
          vAnalysis("Topic") = vTopic
          vAnalysis("SubTopic") = vSubTopic
          DataHelper.AddActionSubject(vAnalysis)
        End If
      End If
    End If
  End Sub

  Private Sub ProcessActionSelection(ByVal pDR As DataRow)
    SetTextBoxText("ActionDesc", pDR("ActionDesc").ToString)
    SetDropDownText("ActionPriority", pDR("ActionPriority").ToString)
    SetTextBoxText("ScheduledOn", pDR("ScheduledOn").ToString)
    SetTextBoxText("Deadline", pDR("Deadline").ToString)
    SetTextBoxText("DurationDays", pDR("DurationDays").ToString)
    SetTextBoxText("DurationHours", pDR("DurationHours").ToString)
    SetTextBoxText("DurationMinutes", pDR("DurationMinutes").ToString)
    SetDropDownText("DocumentClass", pDR("DocumentClass").ToString)
    SetDropDownText("Topic", pDR("Topic").ToString)
    SetDropDownText("Actioner", pDR("ContactNumber").ToString)
    SetTextBoxText("ActionText", pDR("ActionText").ToString)
    SetHiddenText("OldActioner", pDR("ContactNumber").ToString)
    SetHiddenText("OldTopic", pDR("Topic").ToString)
    SetHiddenText("OldSubTopic", pDR("SubTopic").ToString)
    'Setting the Topic may not have selected the SubTopics, so select them now
    Dim vSubTopicDDL As DropDownList = TryCast(FindControlByName(Me, "SubTopic"), DropDownList)
    If vSubTopicDDL IsNot Nothing Then
      Dim vList As New ParameterList(HttpContext.Current)
      vList("Topic") = pDR("Topic").ToString
      DataHelper.FillCombo(CareNetServices.XMLLookupDataTypes.xldtSubTopics, vSubTopicDDL, False, vList)
    End If
    SetDropDownText("SubTopic", pDR("SubTopic").ToString)
  End Sub

End Class