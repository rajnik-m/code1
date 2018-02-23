Public Class SurveyResponses
  Inherits CareWebControl

  Private mvResponsesTable As DataTable
  Private mvPageMap As Dictionary(Of Integer, Integer) 'Used to keep track of Paging.
  Private mvControlMap As Dictionary(Of String, Integer) 'Used to map controls with corresponding Row
  Private mvPreviousQuest As Integer
  Private mvContactSurveyNumber As Integer
  Private mvPreviousQuestionID As String
  Private mvMandatoryCheckBox As Boolean

  Public Sub New()
    mvNeedsAuthentication = True
  End Sub

  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim vSurveyVersionNumber As Integer = 0
    Dim vContactSurveyNumber As Integer = 0
    Dim vList As New ParameterList(HttpContext.Current)
    Dim vResult As New ParameterList()
    Dim vSurveyNumber As Integer
    Dim vPageNumber As Integer = 0
    Dim vMaxPageNumber As Integer = 0
    Try
      If Request.QueryString("SV") IsNot Nothing AndAlso Request.QueryString("SV").Length > 0 Then
        vSurveyVersionNumber = IntegerValue(Request.QueryString("SV"))
      Else
        If InitialParameters.ContainsKey("SurveyVersionNumber") Then vSurveyVersionNumber = IntegerValue(InitialParameters("SurveyVersionNumber").ToString)
      End If

      vList("SurveyVersionNumber") = vSurveyVersionNumber

      'Check for Survey Number
      Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSurveyVersions, vList))
      If vRow IsNot Nothing Then
        vSurveyNumber = IntegerValue(vRow.Item("SurveyNumber").ToString)
        vList("SurveyNumber") = vSurveyNumber
      End If

      'Set mvMandatoryCheckBox = False
      mvMandatoryCheckBox = False

      'Check for Contact Surveys
      vList("ContactNumber") = UserContactNumber()
      vList("UserID") = UserContactNumber.ToString
      Dim vContactSurveyTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtContactSurveys, vList)
      If vContactSurveyTable Is Nothing Then
        'Add a new ContactSurvey record
        vResult = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContactSurveys, vList)
      Else
        vResult.Add("ContactSurveyNumber", vContactSurveyTable.Rows(0).Item("ContactSurveyNumber"))
      End If
      vResult.Add("ContactNumber", UserContactNumber)
      mvContactSurveyNumber = IntegerValue(vResult("ContactSurveyNumber").ToString)
      If Not vResult.ContainsKey("Database") Then vResult.AddConectionData(HttpContext.Current)
      mvResponsesTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactSurveyResponses, vResult)
      If mvResponsesTable IsNot Nothing Then
        If Request.QueryString("SPN") IsNot Nothing AndAlso Request.QueryString("SPN").Length > 0 Then
          If mvResponsesTable IsNot Nothing Then
            'Call function to populate web page
            vPageNumber = IntegerValue(Request.QueryString("SPN"))
            vMaxPageNumber = PopulateControls(vPageNumber)
          End If
        Else
          vPageNumber = 0
          vMaxPageNumber = PopulateControls(vPageNumber)
        End If
        InitialiseControls(CareNetServices.WebControlTypes.wctSurveyResponses, tblDataEntry)

        'Set Previous and Next Button Visibility
        If vMaxPageNumber = 0 Then
          FindControlByName(Me, "Previous").Visible = False
          FindControlByName(Me, "Next").Visible = False
          FindControlByName(Me, "Complete").Visible = True
        Else
          If vPageNumber = 0 Then
            FindControlByName(Me, "Previous").Visible = False
            FindControlByName(Me, "Complete").Visible = False
          ElseIf vPageNumber > 0 AndAlso vPageNumber < vMaxPageNumber Then
            FindControlByName(Me, "Previous").Visible = True
            FindControlByName(Me, "Next").Visible = True
            FindControlByName(Me, "Complete").Visible = False
          ElseIf vPageNumber = vMaxPageNumber Then
            FindControlByName(Me, "Previous").Visible = True
            FindControlByName(Me, "Next").Visible = False
            FindControlByName(Me, "Complete").Visible = True
          End If
        End If
      End If
    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Sub

  ''' <summary>
  ''' Populates the controls depending on the answer types and other related values and returns the maximum number of pages in the survey
  ''' </summary>
  ''' <remarks></remarks>
  Private Function PopulateControls(ByVal pPageNumber As Integer) As Integer
    Try
      If mvResponsesTable IsNot Nothing Then
        Dim vPrevQuestNum As Integer
        Dim vQuestNum As Integer
        Dim vControlType As String = ""
        Dim vAnswers() As DataRow = Nothing
        Dim vAddAnswer As Boolean
        Dim vHasPattern As Boolean
        Dim vPageNumber As Integer = 0
        Dim vHTMLTable As New HtmlTable
        Dim vPreviousQuestion As Integer
        Dim vQuestionID As String = ""
        mvControlMap = New Dictionary(Of String, Integer)

        'Create a PageMap which keeps track of the Page on which the control will appear
        mvPageMap = New Dictionary(Of Integer, Integer)
        Dim vRow As Integer
        For vRow = 0 To mvResponsesTable.Rows.Count - 1
          If vRow = 0 Then vPreviousQuestion = IntegerValue(mvResponsesTable.Rows(vRow).Item("SurveyQuestionNumber").ToString)
          Dim vSurveyQuestionNumber As Integer = IntegerValue(mvResponsesTable.Rows(vRow).Item("SurveyQuestionNumber").ToString)
          If Not BooleanValue(mvResponsesTable.Rows(vRow).Item("NewPage").ToString) Then
            mvPageMap.Add(vRow, vPageNumber)
          Else
            If vPreviousQuestion <> vSurveyQuestionNumber Then
              vPageNumber = vPageNumber + 1
            End If
            mvPageMap.Add(vRow, vPageNumber)
          End If
          vPreviousQuestion = vSurveyQuestionNumber
        Next vRow

        For vIndex As Integer = 0 To mvResponsesTable.Rows.Count - 1
          'Check for Current Page and load the controls
          If mvPageMap(vIndex) = pPageNumber Then
            With mvResponsesTable.Rows(vIndex)
              vHasPattern = False
              vAddAnswer = True
              vQuestNum = IntegerValue(.Item("SurveyQuestionNumber").ToString)
              If vPrevQuestNum <> vQuestNum Then

                'Add Question
                Dim vHTMLRow As New HtmlTableRow

                'First Add Hyperlink for Question Text. 
                Dim vHTMLCell As New HtmlTableCell
                Dim vHyperLink As New HyperLink
                vHyperLink.Text = .Item("QuestionText").ToString
                vHyperLink.Attributes("Class") = "DataMessage"
                vHyperLink.Attributes("href") = "#SurveyQuestion" & vQuestNum.ToString
                vHyperLink.Attributes("name") = "SurveyQuestion" & vQuestNum.ToString
                vHyperLink.ID = "SurveyQuestion" & vQuestNum.ToString
                vHTMLCell.Controls.Add(vHyperLink)
                vHTMLCell.ColSpan = 2
                vQuestionID = vHyperLink.ID
                'Adding a Question on Page.
                vHTMLRow.Cells.Add(vHTMLCell)

                'Get all the answers for the current question
                vAnswers = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", vQuestNum))
                vPrevQuestNum = vQuestNum

                If vAnswers.Length = 1 AndAlso .Item("AnswerText").ToString.Length = 0 Then
                  'The question has only one Answer 
                  'Display the control at the end of the question
                  vHTMLRow.Cells.Add(vHTMLCell)
                  Dim vNewHTMLCell As New HtmlTableCell
                  SetHTMLCell(vNewHTMLCell, .Item("AnswerType").ToString, .Item("AnswerDataType").ToString, vIndex, vQuestionID)
                  vHTMLRow.Cells.Add(vNewHTMLCell)
                  tblDataEntry.Rows.Add(vHTMLRow)
                Else
                  tblDataEntry.Rows.Add(vHTMLRow)
                  'The question has multiple Answers 
                  'Loop through all the Answers for the perticular question and place the controls.
                  For vAnswer As Integer = 0 To vAnswers.Length - 1
                    vHTMLRow = New HtmlTableRow
                    vHTMLCell = New HtmlTableCell
                    vHTMLCell.InnerHtml = vAnswers(vAnswer).Item("AnswerText").ToString
                    vHTMLRow.Cells.Add(vHTMLCell)
                    'Add Control 
                    Dim vNewHTMLCell As New HtmlTableCell
                    SetHTMLCell(vNewHTMLCell, mvResponsesTable.Rows(vIndex + vAnswer).Item("AnswerType").ToString, mvResponsesTable.Rows(vIndex + vAnswer).Item("AnswerDataType").ToString, vIndex + vAnswer, vQuestionID)
                    vNewHTMLCell.Controls.GetType()
                    vHTMLRow.Cells.Add(vNewHTMLCell)
                    tblDataEntry.Rows.Add(vHTMLRow)
                  Next vAnswer
                End If
              End If
            End With
          End If
        Next vIndex

        Dim vMaxPages As Integer = mvPageMap(mvPageMap.Count - 1)
        Return vMaxPages
      End If

    Catch vEx As Exception
      ProcessError(vEx)
    End Try
  End Function

  ''' <summary>
  ''' This method generates the control depending on the AnswerType and AnswerDataType and places it on page
  ''' </summary>
  ''' <param name="pHTMLCell"></param>
  ''' <param name="pAnswerType"></param>
  ''' <param name="pAnswerDataType"></param>
  ''' <param name="pIndex"></param>
  ''' <param name="pQuestionId"></param>
  ''' <remarks></remarks>
  Private Sub SetHTMLCell(ByRef pHTMLCell As HtmlTableCell, ByVal pAnswerType As String, ByVal pAnswerDataType As String, ByVal pIndex As Integer, ByVal pQuestionId As String)
    Dim vControlType As String = ""
    Dim vAddAnswer As Boolean = True
    Dim vHasPattern As Boolean = False
    Dim vQuestNumber As Integer = IntegerValue(mvResponsesTable.Rows(pIndex).Item("SurveyQuestionNumber").ToString)
    Dim vAnsNumber As Integer = IntegerValue(mvResponsesTable.Rows(pIndex).Item("SurveyAnswerNumber").ToString)
    Dim vIsMandatory As Boolean = False

    Select Case pAnswerDataType
      Case "Y"
        Select Case pAnswerType.ToUpper
          Case "M"
            vControlType = "chk"
          Case "S"
            vControlType = "opt"
          Case Else
            vControlType = "cbo"
            vHasPattern = True
        End Select
      Case "I", "N", "A", "C"
        vControlType = "txt"
      Case "D", "T"
        vControlType = "dtp"
      Case "L"
        vControlType = "cbo"
      Case Else
        vAddAnswer = False
    End Select

    If vAddAnswer Then
      Select Case vControlType
        Case "opt"
          Dim vControl As New RadioButton
          vControl.ID = String.Format("SurveyAnswer{0}_{1}", vQuestNumber, vAnsNumber)
          mvControlMap.Add(vControl.ID, pIndex)
          vControl.GroupName = String.Format("SurveyAnswer{0}_", vQuestNumber)
          vControl.Checked = vAnsNumber = IntegerValue(mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString)
          vControl.AutoPostBack = True
          AddHandler vControl.CheckedChanged, AddressOf ControlValueChanged
          pHTMLCell.Controls.Add(vControl)
        Case "chk"
          Dim vControl As New CheckBox
          vControl.ID = String.Format("SurveyAnswer{0}_{1}", vQuestNumber, vAnsNumber)
          mvControlMap.Add(vControl.ID, pIndex)
          vControl.Checked = BooleanValue(mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString)
          If mvResponsesTable.Rows(pIndex).Item("NextQuestionNumber").ToString.Length > 0 Then
            vControl.AutoPostBack = True
          End If
          pHTMLCell.Controls.Add(vControl)
        Case "txt"
          Dim vControl As New TextBox
          Dim vMinValue As Integer = 0
          Dim vMaxValue As Integer = 0
          vControl.ID = String.Format("SurveyAnswer{0}_{1}", vQuestNumber, vAnsNumber)
          mvControlMap.Add(vControl.ID, pIndex)
          vControl.Text = mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString
          vControl.AutoPostBack = True

          AddHandler vControl.TextChanged, AddressOf ControlValueChanged

          pHTMLCell.Controls.Add(vControl)

          'Set Mandatory Field
          vIsMandatory = BooleanValue(mvResponsesTable.Rows(pIndex).Item("Mandatory").ToString)
          If mvResponsesTable.Rows(pIndex).Item("AnswerType").ToString.ToUpper = "S" AndAlso mvResponsesTable.Rows(pIndex).Item("AnswerDataType").ToString.ToUpper = "C" Then
            vControl.AutoPostBack = True
            'Do nothing
          Else
            If vIsMandatory Then
              vControl.CssClass = "DataEntryItemMandatory"
              vControl.AutoPostBack = True
              AddRequiredValidator(pHTMLCell, vControl.ID)
            End If
          End If

          'Check for Next Question
          If mvResponsesTable.Rows(pIndex).Item("NextQuestionNumber").ToString.Length > 0 Then
            vControl.AutoPostBack = True
            AddHandler vControl.TextChanged, AddressOf ControlValueChanged
          End If

          vControl.CssClass = "DataEntryItem"

          'Add Validations
          If pAnswerDataType = "A" OrElse pAnswerDataType = "C" Then
            'Add DatType validator for Min and Max value
            AddDataTypeValidator(pHTMLCell, vControl.ID, ValidationDataType.String)
            'Add Range validator for Min and Max value
            If mvResponsesTable.Rows(pIndex).Item("MaximumValue").ToString.Length > 0 OrElse mvResponsesTable.Rows(pIndex).Item("MinimumValue").ToString.Length > 0 Then
              vControl.AutoPostBack = True
              AddSurveyRangeValidator(pHTMLCell, vControl.ID, mvResponsesTable.Rows(pIndex).Item("MinimumValue").ToString, mvResponsesTable.Rows(pIndex).Item("MaximumValue").ToString, String.Format("Value must be between '{0}' and '{1}'", mvResponsesTable.Rows(pIndex).Item("MinimumValue").ToString, mvResponsesTable.Rows(pIndex).Item("MaximumValue").ToString), ValidationDataType.String)
            End If
            If pAnswerDataType = "A" Then
              AddSurveyRegExValidator(pHTMLCell, vControl.ID)
            End If
          ElseIf pAnswerDataType = "I" OrElse pAnswerDataType = "N" Then
            'Add DatType validator 
            AddDataTypeValidator(pHTMLCell, vControl.ID, ValidationDataType.Integer)
            'Add Range validator for Min and Max value
            vMinValue = IntegerValue(mvResponsesTable.Rows(pIndex).Item("MinimumValue").ToString)
            vMaxValue = IntegerValue(mvResponsesTable.Rows(pIndex).Item("MaximumValue").ToString)
            If vMinValue > 0 OrElse vMaxValue > 0 Then
              AddSurveyRangeValidator(pHTMLCell, vControl.ID, vMinValue.ToString, vMaxValue.ToString, String.Format("Value must be between {0} and {1}", vMinValue, vMaxValue), ValidationDataType.Integer)
            End If
          End If
        Case "dtp"
          If pAnswerDataType = "D" Then
            Dim vMinValue As String = ""
            Dim vMaxValue As String = ""
            Dim vControl As New TextBox
            vControl.ID = String.Format("SurveyAnswer{0}_{1}", vQuestNumber, vAnsNumber)
            mvControlMap.Add(vControl.ID, pIndex)
            pHTMLCell.Controls.Add(vControl)
            Dim vButton As New HtmlInputButton
            vButton.ID = "cmdFind" & vControl.ID
            vButton.Attributes("value") = "..."
            vButton.Attributes("class") = "Button"
            vButton.Attributes("style") = "width:2em;"
            vButton.CausesValidation = False
            AddDateTimePicker(vControl.ID)
            vControl.AutoPostBack = True
            vControl.Text = mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString
            pHTMLCell.Controls.Add(vButton)
            'Set Mandatory
            vIsMandatory = BooleanValue(mvResponsesTable.Rows(pIndex).Item("Mandatory").ToString)

            If vIsMandatory Then
              AddRequiredValidator(pHTMLCell, vControl.ID)
            End If
            'Add Range Validators
            vMinValue = mvResponsesTable.Rows(pIndex).Item("MinimumValue").ToString
            vMaxValue = mvResponsesTable.Rows(pIndex).Item("MaximumValue").ToString
            If vMinValue.Length > 0 OrElse vMaxValue.Length > 0 Then
              AddSurveyRangeValidator(pHTMLCell, vControl.ID, vMinValue, vMaxValue, String.Format("Value must be between {0} and {1}", vMinValue, vMaxValue), ValidationDataType.Date)
            End If
          ElseIf pAnswerDataType = "T" Then
            ' If the control is a Just a TimePicker then show two Dropdownlists respectively for Houre and for Minutes.
            Dim vHoursDropDown As New DropDownList
            Dim vMinutesDropDown As New DropDownList
            vIsMandatory = BooleanValue(mvResponsesTable.Rows(pIndex).Item("Mandatory").ToString)
            vMinutesDropDown.Width = 48
            vHoursDropDown.Width = 48
            'Populating Hours Dropdown
            vHoursDropDown.Items.Add("")
            For vHours As Integer = 0 To 23
              If vHours < 10 Then
                vHoursDropDown.Items.Add(String.Format("0{0}", vHours.ToString))
              Else
                vHoursDropDown.Items.Add(vHours.ToString)
              End If
            Next
            'Populating Minutes Dropdown
            vMinutesDropDown.Items.Add("")
            For vMins As Integer = 0 To 59
              If vMins < 10 Then
                vMinutesDropDown.Items.Add(String.Format("0{0}", vMins.ToString))
              Else
                vMinutesDropDown.Items.Add(vMins.ToString)
              End If
            Next
            If mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString.Length > 0 Then
              vHoursDropDown.Text = mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString.Split(":"c)(0).ToString
              vMinutesDropDown.Text = mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString.Split(":"c)(1).ToString
            Else
              vHoursDropDown.Text = ""
              vMinutesDropDown.Text = ""
            End If
            'Adding Labels for Hrs and Mins
            Dim vHrsLabel As New Label
            vHrsLabel.Text = "hrs "
            vHrsLabel.Font.Bold = True

            Dim vMinsLabel As New Label
            vMinsLabel.Text = " mins "
            vMinsLabel.Font.Bold = True

            vHoursDropDown.ID = String.Format("SurveyAnswer{0}_{1}", vQuestNumber, vAnsNumber)
            vMinutesDropDown.ID = String.Format("SurveyAnswer{0}_{1}_min", vQuestNumber, vAnsNumber)

            mvControlMap.Add(vHoursDropDown.ID, pIndex)
            pHTMLCell.Controls.Add(vHrsLabel)
            pHTMLCell.Controls.Add(vHoursDropDown)
            pHTMLCell.Controls.Add(vMinsLabel)
            pHTMLCell.Controls.Add(vMinutesDropDown)

            If vIsMandatory Then
              AddRequiredValidator(pHTMLCell, vHoursDropDown.ID)
            End If
          End If
        Case "cbo"
          Dim vControl As New DropDownList
          vControl.ID = String.Format("SurveyAnswer{0}_{1}", vQuestNumber, vAnsNumber)
          mvControlMap.Add(vControl.ID, pIndex)
          Dim vDataSource() As String = GetDataSourceForCombo(pAnswerDataType, pIndex, vHasPattern)
          vControl.DataSource = vDataSource
          vControl.DataBind()
          vControl.SelectedValue = mvResponsesTable.Rows(pIndex).Item("ResponseAnswerText").ToString
          If mvResponsesTable.Rows(pIndex).Item("NextQuestionNumber").ToString.Length > 0 Then
            vControl.AutoPostBack = True
            AddHandler vControl.SelectedIndexChanged, AddressOf ControlValueChanged
          End If
          pHTMLCell.Controls.Add(vControl)
          'Set Mandatory
          vIsMandatory = BooleanValue(mvResponsesTable.Rows(pIndex).Item("Mandatory").ToString)
          If vIsMandatory Then
            AddRequiredValidator(pHTMLCell, vControl.ID)
          End If
      End Select
    End If
    mvPreviousQuest = vQuestNumber
    mvPreviousQuestionID = pQuestionId
  End Sub

  Private Function GetDataSourceForCombo(ByVal pAnswerDataType As String, ByVal pIndex As Integer, ByVal pHasPatteren As Boolean) As String()
    Dim vOptionsArray() As String = {}
    Dim vListValues As String = Nothing
    If pAnswerDataType.ToUpper = "Y" Then
      If pHasPatteren Then vListValues = "|YES|NO"
    ElseIf pAnswerDataType.ToUpper = "L" Then
      vListValues = String.Format("|{0}", mvResponsesTable.Rows(pIndex).Item("ListValues").ToString.ToUpper.Replace(","c, "|"c))
    End If

    If Not String.IsNullOrEmpty(vListValues) Then
      vOptionsArray = vListValues.Split("|"c)
    End If
    Return vOptionsArray
  End Function

  Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      If Page.IsValid Then
        Dim vPageNumber As Integer
        Dim vParams As NameValueCollection
        Dim vMaxPage As Integer = 0
        Dim vIsSaved As Boolean = False

        If Request.QueryString("SPN") IsNot Nothing AndAlso Request.QueryString("SPN").Length > 0 Then
          vPageNumber = IntegerValue(Request.QueryString("SPN"))
        Else
          vPageNumber = 0
        End If
        If Not InWebPageDesigner() Then
          If CType(sender, Button).ID = "Next" Then
            If ProcessSave() Then
              vPageNumber = GetNextPageNumber(vPageNumber)
              vMaxPage = PopulateControls(vPageNumber)
              vIsSaved = True
            End If
          ElseIf CType(sender, Button).ID = "Previous" Then
            If ProcessSave() Then
              vPageNumber = GetPreviousPageNumber(vPageNumber)
              vMaxPage = PopulateControls(vPageNumber)
              vIsSaved = True
            End If
          ElseIf CType(sender, Button).ID = "Complete" Then
            If ProcessSave() Then
              'Set Completed On for Contact_Survey record
              Dim vCompletedList As New ParameterList(HttpContext.Current)
              vCompletedList("ContactSurveyNumber") = mvContactSurveyNumber.ToString
              vCompletedList("CompletedOn") = TodaysDate()
              vCompletedList("UserID") = UserContactNumber.ToString
              DataHelper.UpdateItem(CareNetServices.XMLMaintenanceControlTypes.xmctContactSurveys, vCompletedList)
              GoToSubmitPage()
              vIsSaved = True
            End If
          End If

          'submit page with next page number in query string
          If vIsSaved Then
            vParams = HttpUtility.ParseQueryString(Request.QueryString.ToString)
            vParams.Item("SPN") = vPageNumber.ToString
            ProcessRedirect(String.Format("Default.aspx?{0}", vParams.ToString))
          End If
        End If
      End If
    Catch vEx As Exception
      If Not TypeOf vEx Is ThreadAbortException Then
        ProcessError(vEx)
      End If
    End Try
  End Sub

  Protected Sub ControlValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    Try
      Page.Validate()
      If Page.IsValid Then
        Dim vControl As String = CType(sender, Control).ID
        Dim vRowNumber As Integer = mvControlMap(vControl)
        Dim vParams As NameValueCollection
        Dim vNextPageNumber As Integer = 0
        Dim vNextRowNumber As Integer = 0
        Dim vQueryString As String = ""
        'vIsNextQuestionValid is a boolean value which decides whether navigation to 'Next' associated question should be done or not.
        'in case of CheckBoxes if the checkbox is getting checked then it will be True otherwise it is False.So if the checkbox is not unchecked then it will not navigate in that case,
        Dim vIsNextQuestionValid As Boolean = True
        If TypeOf sender Is RadioButton Then
          If mvResponsesTable.Rows(vRowNumber).Item("AnswerType").ToString.ToUpper = "S" AndAlso mvResponsesTable.Rows(vRowNumber).Item("AnswerDataType").ToString.ToUpper = "Y" Then
            Dim vAnswers() As DataRow = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(vRowNumber).Item("SurveyQuestionNumber").ToString))
            For Each vAnswerRow As DataRow In vAnswers
              If vAnswerRow.Item("AnswerDataType").ToString.ToUpper = "C" Then
                Dim vControlId As String = String.Format("SurveyAnswer{0}_{1}", vAnswerRow.Item("SurveyQuestionNumber").ToString, vAnswerRow.Item("SurveyAnswerNumber").ToString)
                Dim vTempControl As Control = FindControlFromHTMLTable(tblDataEntry, vControlId)
                If Not vTempControl Is Nothing Then
                  Dim vTextBox As TextBox = CType(vTempControl, TextBox)
                  vTextBox.Text = ""
                End If
              End If
            Next
          End If
        ElseIf TypeOf sender Is CheckBox Then
          Dim vCheckBox As New CheckBox
          vCheckBox = CType(sender, CheckBox)
          vIsNextQuestionValid = vCheckBox.Checked
        ElseIf TypeOf sender Is TextBox Then
          If mvResponsesTable.Rows(vRowNumber).Item("AnswerType").ToString.ToUpper = "S" AndAlso mvResponsesTable.Rows(vRowNumber).Item("AnswerDataType").ToString.ToUpper = "C" Then
            Dim vAnswers() As DataRow = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(vRowNumber).Item("SurveyQuestionNumber").ToString))
            If CType(sender, TextBox).Text.Length > 0 Then
              For Each vAnswerRow As DataRow In vAnswers
                If vAnswerRow.Item("AnswerDataType").ToString.ToUpper = "Y" Then
                  Dim vControlId As String = String.Format("SurveyAnswer{0}_{1}", vAnswerRow.Item("SurveyQuestionNumber").ToString, vAnswerRow.Item("SurveyAnswerNumber").ToString)
                  Dim vTempControl As Control = FindControlFromHTMLTable(tblDataEntry, vControlId)
                  If Not vTempControl Is Nothing Then
                    Dim vRadioButton As RadioButton = CType(vTempControl, RadioButton)
                    vRadioButton.Checked = False
                  End If
                End If
              Next
            End If
          End If
        End If
        If Not InWebPageDesigner() Then
          If vIsNextQuestionValid Then
            If mvResponsesTable.Rows(vRowNumber).Item("NextQuestionNumber").ToString.Length > 0 Then
              If mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(vRowNumber).Item("NextQuestionNumber").ToString)).Length > 0 Then
                vNextRowNumber = mvResponsesTable.Rows.IndexOf(mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(vRowNumber).Item("NextQuestionNumber").ToString))(0))
                vNextPageNumber = mvPageMap(vNextRowNumber)
                If ProcessSave() Then
                  'submit page with next page number in query string
                  vParams = HttpUtility.ParseQueryString(Request.QueryString.ToString)
                  vParams.Item("SPN") = vNextPageNumber.ToString
                  vQueryString = vParams.ToString & String.Format("#SurveyQuestion{0}", mvResponsesTable.Rows(vRowNumber).Item("NextQuestionNumber").ToString)
                  ProcessRedirect("Default.aspx?" & vQueryString)
                End If
              End If
            End If
          End If
        End If
      End If
    Catch vEx As Exception
      If Not TypeOf vEx Is ThreadAbortException Then
        ProcessError(vEx)
      End If
    End Try
  End Sub

  ''' <summary>
  ''' This method saves control value if it is changed.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ProcessSave() As Boolean
    Page.Validate()
    If Page.IsValid Then
      Dim vList As New ParameterList
      Dim vRow As Integer = 0
      Dim vOptionGroup As String = ""
      If AreCheckBoxesValid() AndAlso ValidateTime() Then
        For Each vHtmlRow As HtmlTableRow In tblDataEntry.Rows
          For Each vCell As HtmlTableCell In vHtmlRow.Cells
            For Each vControl As Control In vCell.Controls
              If Not TypeOf vControl Is LiteralControl AndAlso Not TypeOf vControl Is Label Then
                If mvControlMap.ContainsKey(vControl.ID) Then
                  vRow = mvControlMap(vControl.ID)
                  vList = New ParameterList(HttpContext.Current)
                  vList("ContactSurveyNumber") = mvResponsesTable.Rows(vRow).Item("ContactSurveyNumber")
                  vList("SurveyQuestionNumber") = mvResponsesTable.Rows(vRow).Item("SurveyQuestionNumber")
                  vList("SurveyAnswerNumber") = mvResponsesTable.Rows(vRow).Item("SurveyAnswerNumber")
                  vList("UserID") = UserContactNumber.ToString

                  If TypeOf vControl Is RadioButton Then
                    Dim vRadioButton As RadioButton = CType(vControl, RadioButton)
                    If vRadioButton.Checked Then
                      vList("ResponseAnswerText") = mvResponsesTable.Rows(vRow).Item("SurveyAnswerNumber")
                    Else
                      vList("ResponseAnswerText") = ""
                    End If
                    vOptionGroup = vRadioButton.GroupName
                  ElseIf TypeOf vControl Is TextBox Then
                    Dim vTextBox As TextBox = CType(vControl, TextBox)
                    If mvResponsesTable.Rows(vRow).Item("ResponseAnswerText").ToString <> vTextBox.Text Then
                      vList("ResponseAnswerText") = vTextBox.Text
                    End If
                  ElseIf TypeOf vControl Is DropDownList Then
                    If mvResponsesTable.Rows(vRow).Item("AnswerDataType").ToString.ToUpper = "T" Then
                      Dim vTime As String = ""

                      For Each vTimeControl As Control In vCell.Controls
                        If Not TypeOf vTimeControl Is Label Then
                          vTime = vTime & CType(vTimeControl, DropDownList).SelectedValue & ":"
                        End If
                      Next
                      If vTime = "::" Then
                        vList("ResponseAnswerText") = ""
                        DataHelper.UpdateContactSurveyResponse(vList)
                        Exit For
                      ElseIf vTime.StartsWith(":") Then
                        vTime = "00" & vTime
                      End If

                      vTime = vTime.Remove(vTime.LastIndexOf(":"))
                      If vTime.EndsWith(":") Then
                        vTime = vTime & "00"
                      End If

                      vList("ResponseAnswerText") = vTime
                      DataHelper.UpdateContactSurveyResponse(vList)
                      Exit For
                    Else
                      Dim vComboBox As DropDownList = CType(vControl, DropDownList)
                      If mvResponsesTable.Rows(vRow).Item("ResponseAnswerText").ToString <> vComboBox.SelectedValue Then
                        vList("ResponseAnswerText") = vComboBox.SelectedValue
                      End If
                    End If
                  ElseIf TypeOf vControl Is CheckBox Then
                    Dim vCheckBox As CheckBox = CType(vControl, CheckBox)
                    If BooleanValue(mvResponsesTable.Rows(vRow).Item("Mandatory").ToString) Then
                      If CheckCheckBoxValidationRequired(vRow) Then
                        Dim vErrorLabel As New Label
                        vErrorLabel.ForeColor = Drawing.Color.Red
                        vErrorLabel.Font.Bold = True
                        vErrorLabel.Text = "Atleast one of the values must be selected"
                        vCell.Controls.Add(vErrorLabel)
                        Exit Function
                      End If
                    End If
                    If BooleanValue(mvResponsesTable.Rows(vRow).Item("ResponseAnswerText").ToString) <> vCheckBox.Checked Then
                      vList("ResponseAnswerText") = IIf(vCheckBox.Checked, "Y", "N").ToString 'CBoolYN(vCheckBox.Checked)
                    End If
                  End If
                  If vList.Count > 0 AndAlso vList.Contains("ResponseAnswerText") Then
                    DataHelper.UpdateContactSurveyResponse(vList)
                  End If
                End If
              End If
            Next
          Next
        Next
        Return True
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Protected Sub AddDataTypeValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pDataType As ValidationDataType)
    Dim vCV As New CompareValidator
    With vCV
      .CssClass = "DataValidator"
      .ID = "afv" & pID
      .ControlToValidate = pID
      .Display = ValidatorDisplay.Dynamic
      .ErrorMessage = "Invalid Value"
      .Operator = ValidationCompareOperator.DataTypeCheck
      .Type = pDataType
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vCV)
  End Sub

  Private Sub AddSurveyRangeValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String, ByVal pMinimumValue As String, ByVal pMaximumValue As String, Optional ByVal pErrorMessage As String = "", Optional ByVal pDataType As ValidationDataType = ValidationDataType.String)
    Dim vRNV As New RangeValidator
    With vRNV
      .Type = pDataType
      .CssClass = "DataValidator"
      .ID = "rnv" & pID
      .ControlToValidate = pID
      .MinimumValue = pMinimumValue
      .MaximumValue = pMaximumValue
      .Display = ValidatorDisplay.Dynamic
      If pErrorMessage.Length = 0 Then
        .ErrorMessage = "Invalid Value"
      Else
        .ErrorMessage = pErrorMessage
      End If
      .SetFocusOnError = True
    End With
    pHTMLCell.Controls.Add(vRNV)
  End Sub

  Private Sub AddSurveyRegExValidator(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String)
    Dim vREV As New RegularExpressionValidator
    With vREV
      .CssClass = "DataValidator"
      .ID = "rev" & pID
      .ControlToValidate = pID
      .Display = ValidatorDisplay.Dynamic
      .SetFocusOnError = True
      .ErrorMessage = "Invalid Value"
      .ValidationExpression = "^[a-zA-Z]"
    End With
    pHTMLCell.Controls.Add(vREV)
  End Sub

  ''' <summary>
  ''' This method finds the control in HTMLTable and returns it.
  ''' </summary>
  ''' <param name="pHTMLTable"></param>
  ''' <param name="pControlID"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function FindControlFromHTMLTable(ByVal pHTMLTable As HtmlTable, ByVal pControlID As String) As Control
    Dim vResultControl As Control = Nothing
    For Each vHtmlRow As HtmlTableRow In tblDataEntry.Rows
      For Each vCell As HtmlTableCell In vHtmlRow.Cells
        For Each vControl As Control In vCell.Controls
          If Not TypeOf vControl Is LiteralControl Then
            If vControl.ID = pControlID Then vResultControl = vControl
          End If
        Next
      Next
    Next
    Return vResultControl
  End Function

  ''' <summary>
  ''' This is method checks if the Mandatory checkboxes and checked.
  ''' </summary>
  ''' <param name="pRow"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function CheckCheckBoxValidationRequired(ByVal pRow As Integer) As Boolean
    Dim vRequired As Boolean = False
    'Check only if the control is a check-box
    If mvResponsesTable.Rows(pRow).Item("AnswerDataType").ToString.ToUpper = "Y" AndAlso mvResponsesTable.Rows(pRow).Item("AnswerType").ToString.ToUpper = "M" Then
      Dim vQuestionNumber As Integer = IntegerValue(mvResponsesTable.Rows(pRow).Item("SurveyQuestionNumber").ToString)
      Dim vAnswers() As DataRow = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", vQuestionNumber))
      Dim vControlId As String = ""
      If vAnswers.Length > 0 Then
        For Each vRow As DataRow In vAnswers
          vControlId = String.Format("SurveyAnswer{0}_{1}", vRow.Item("SurveyQuestionNumber").ToString, vRow.Item("SurveyAnswerNumber").ToString)
          Dim vTempControl As Control = FindControlFromHTMLTable(tblDataEntry, vControlId)
          If Not vTempControl Is Nothing Then
            Dim vCheckBox As CheckBox = CType(vTempControl, CheckBox)
            If vCheckBox.Checked Then
              vRequired = False
              Exit For
            Else
              vRequired = True
            End If
          End If
        Next
      End If
      'Check only if the control is a Group of Radio Buttons
    ElseIf mvResponsesTable.Rows(pRow).Item("AnswerType").ToString.ToUpper = "S" AndAlso (mvResponsesTable.Rows(pRow).Item("AnswerDataType").ToString.ToUpper = "Y" OrElse mvResponsesTable.Rows(pRow).Item("AnswerDataType").ToString.ToUpper = "C") Then
      Dim vQuestionNumber As Integer = IntegerValue(mvResponsesTable.Rows(pRow).Item("SurveyQuestionNumber").ToString)
      Dim vAnswers() As DataRow = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", vQuestionNumber))
      Dim vControlId As String = ""
      If vAnswers.Length > 0 Then
        For Each vRow As DataRow In vAnswers
          vControlId = String.Format("SurveyAnswer{0}_{1}", vRow.Item("SurveyQuestionNumber").ToString, vRow.Item("SurveyAnswerNumber").ToString)
          Dim vTempControl As Control = FindControlFromHTMLTable(tblDataEntry, vControlId)
          If Not vTempControl Is Nothing Then
            If vRow.Item("AnswerDataType").ToString.ToUpper = "Y" Then
              Dim vRadioButton As RadioButton = CType(vTempControl, RadioButton)
              If vRadioButton.Checked Then
                vRequired = False
                Exit For
              Else
                vRequired = True
              End If
            ElseIf vRow.Item("AnswerDataType").ToString.ToUpper = "C" Then
              Dim vTextBox As TextBox = CType(vTempControl, TextBox)
              If vTextBox.Text.Length > 0 Then
                vRequired = False
                Exit For
              Else
                vRequired = True
              End If
            End If
          End If
        Next
      End If
    End If
    Return vRequired
  End Function

  ''' <summary>
  ''' This method validates Mandatory Checkboxes
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function AreCheckBoxesValid() As Boolean
    Dim vValid As Boolean = True
    Dim vControlId As String = ""
    For Each vRow As DataRow In mvResponsesTable.Rows
      If vRow.Item("AnswerDataType").ToString.ToUpper = "Y" AndAlso vRow.Item("AnswerType").ToString.ToUpper = "M" Then
        If BooleanValue(vRow.Item("Mandatory").ToString) Then
          If CheckCheckBoxValidationRequired(mvResponsesTable.Rows.IndexOf(vRow)) Then
            vValid = False
            Dim vErrorLabel As New Label
            vErrorLabel.ForeColor = Drawing.Color.Red
            vErrorLabel.Font.Bold = True
            vErrorLabel.Text = "Atleast one of the values must be selected"
            vControlId = String.Format("SurveyAnswer{0}_{1}", vRow.Item("SurveyQuestionNumber").ToString, vRow.Item("SurveyAnswerNumber").ToString)
            Dim vControl As Control = FindControlFromHTMLTable(tblDataEntry, vControlId)
            If TypeOf vControl.Parent Is HtmlTableCell Then
              vControl.Parent.Controls.Add(vErrorLabel)
            End If
            Exit Function
          End If
        End If
      End If
      If vRow.Item("AnswerType").ToString.ToUpper = "S" AndAlso (vRow.Item("AnswerDataType").ToString.ToUpper = "Y" OrElse vRow.Item("AnswerDataType").ToString.ToUpper = "C") Then
        If BooleanValue(vRow.Item("Mandatory").ToString) Then
          If CheckCheckBoxValidationRequired(mvResponsesTable.Rows.IndexOf(vRow)) Then
            vValid = False
            Dim vErrorLabel As New Label
            vErrorLabel.ForeColor = Drawing.Color.Red
            vErrorLabel.Font.Bold = True
            vErrorLabel.Text = "Atleast one of the values must be selected"
            vControlId = String.Format("SurveyAnswer{0}_{1}", vRow.Item("SurveyQuestionNumber").ToString, vRow.Item("SurveyAnswerNumber").ToString)
            Dim vControl As Control = FindControlFromHTMLTable(tblDataEntry, vControlId)
            If TypeOf vControl.Parent Is HtmlTableCell Then
              vControl.Parent.Controls.Add(vErrorLabel)
            End If
            Exit Function
          End If
        End If

      End If
    Next
    Return vValid
  End Function

  ''' <summary>
  ''' This method validates timepicker.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ValidateTime() As Boolean
    Dim vMaxTime As Integer = 0
    Dim vMinTime As Integer = 0
    Dim vHrs As Integer = 0
    Dim vMins As Integer = 0
    Dim vResponseTime As Integer = 0
    Dim vResponse As String = ""
    Dim vControlID As String = ""
    Dim vTempControl As New Control
    Dim vErrorLabel As New Label
    Dim vValid As Boolean = True
    For Each vRow As DataRow In mvResponsesTable.Rows
      If vRow.Item("AnswerDataType").ToString.ToUpper = "T" Then
        If vRow.Item("MaximumValue").ToString.Length > 0 Then
          vHrs = IntegerValue(vRow.Item("MaximumValue").ToString.Split(":"c)(0).ToString)
          vMins = IntegerValue(vRow.Item("MaximumValue").ToString.Split(":"c)(1).ToString)
          vMaxTime = (vHrs * 60) + vMins
        End If
        If vRow.Item("MinimumValue").ToString.Length > 0 Then
          vHrs = IntegerValue(vRow.Item("MinimumValue").ToString.Split(":"c)(0).ToString)
          vMins = IntegerValue(vRow.Item("MinimumValue").ToString.Split(":"c)(1).ToString)
          vMinTime = (vHrs * 60) + vMins
        End If

        If vMaxTime > 0 OrElse vMinTime > 0 Then
          vControlID = String.Format("SurveyAnswer{0}_{1}", vRow.Item("SurveyQuestionNumber").ToString, vRow.Item("SurveyAnswerNumber").ToString)
          vTempControl = FindControlFromHTMLTable(tblDataEntry, vControlID)
          If vTempControl IsNot Nothing Then
            For Each vCellControl As Control In vTempControl.Parent.Controls
              If TypeOf vCellControl Is DropDownList Then
                Dim vDropDown As DropDownList
                vDropDown = CType(vCellControl, DropDownList)
                vResponse = vResponse & vDropDown.SelectedValue & ":"
              End If
            Next
            If Not vResponse = "::" Then
              If vResponse.StartsWith(":") Then
                vResponse = "00" & vResponse
              End If
              vResponse = vResponse.Remove(vResponse.LastIndexOf(":"))
              If vResponse.EndsWith(":") Then
                vResponse = vResponse & ":"
              End If
              vResponseTime = (IntegerValue(vResponse.Split(":"c)(0).ToString) * 60) + (IntegerValue(vResponse.Split(":"c)(1).ToString))
              If vMaxTime > 0 AndAlso vResponseTime > vMaxTime Then
                vErrorLabel.ForeColor = Drawing.Color.Red
                vErrorLabel.Font.Bold = True
                vErrorLabel.Text = String.Format("Value should be less than or equal to {0}", vRow.Item("MaximumValue").ToString)
                vTempControl.Parent.Controls.Add(vErrorLabel)
                vValid = False
                Exit For
              Else
                vValid = True
              End If
              If vMinTime > 0 AndAlso vResponseTime < vMinTime Then
                vErrorLabel.ForeColor = Drawing.Color.Red
                vErrorLabel.Font.Bold = True
                vErrorLabel.Text = String.Format("Value should be greater than or equal to {0}", vRow.Item("MinimumValue").ToString)
                vTempControl.Parent.Controls.Add(vErrorLabel)
                vValid = False
                Exit For
              Else
                vValid = True
              End If
            End If
          End If
        End If
      End If
    Next
    Return vValid
  End Function

  ''' <summary>
  ''' this method gets Next Page Number
  ''' </summary>
  ''' <param name="pPageNumber"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetNextPageNumber(ByVal pPageNumber As Integer) As Integer
    Dim vRowNumber As Integer = 0
    Dim vQuestionNumber As Integer = 0
    Dim vPreviousQuestionNumber As Integer = 0
    Dim vSingleQuestionOnPage As Boolean = True
    Dim vNextRowNumber As Integer = 0
    Dim vNextPageNumber As Integer = pPageNumber + 1
    'If there are more than one questions on the current page then go forward in sequesnce else if there is only one question
    'on the current page then check if there is a next question associated with it and navigate to it.
    If Not SingleQuestionOnPage() Then
      Return pPageNumber + 1
    Else
      For Each vHtmlRow As HtmlTableRow In tblDataEntry.Rows
        For Each vCell As HtmlTableCell In vHtmlRow.Cells
          For Each vControl As Control In vCell.Controls
            If Not TypeOf vControl Is LiteralControl AndAlso Not TypeOf vControl Is Label Then
              If mvControlMap.ContainsKey(vControl.ID) Then
                If TypeOf vControl Is RadioButton AndAlso CType(vControl, RadioButton).Checked AndAlso mvResponsesTable.Rows(mvControlMap(vControl.ID)).Item("NextQuestionNumber").ToString.Length > 0 Then
                  If mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(mvControlMap(vControl.ID)).Item("NextQuestionNumber").ToString)).Length > 0 Then
                    vNextRowNumber = mvResponsesTable.Rows.IndexOf(mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(mvControlMap(vControl.ID)).Item("NextQuestionNumber").ToString))(0))
                    vNextPageNumber = mvPageMap(vNextRowNumber)
                    Return vNextPageNumber
                  End If
                End If
              End If
            End If
          Next
        Next
      Next
      Return vNextPageNumber
    End If
  End Function

  ''' <summary>
  ''' This method returns a boolean value depending on number of questions on page. If there is only 1 question on page it will return 'True' else 'False'
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SingleQuestionOnPage() As Boolean
    Dim vCurrentPageNumber As Integer
    Dim vRowNumber As Integer = 0
    Dim vPreviousQuestionNumber As Integer = 0
    Dim vQuestionNumber As Integer = 0
    Dim vSingleQuestionOnPage As Boolean = True

    If Request.QueryString("SPN") IsNot Nothing AndAlso Request.QueryString("SPN").Length > 0 Then
      vCurrentPageNumber = IntegerValue(Request.QueryString("SPN"))
    Else
      vCurrentPageNumber = 0
    End If
    For Each vHtmlRow As HtmlTableRow In tblDataEntry.Rows
      For Each vCell As HtmlTableCell In vHtmlRow.Cells
        For Each vControl As Control In vCell.Controls
          If Not TypeOf vControl Is LiteralControl AndAlso Not TypeOf vControl Is Label Then
            If mvControlMap.ContainsKey(vControl.ID) Then
              vRowNumber = mvControlMap(vControl.ID)
              vQuestionNumber = IntegerValue(mvResponsesTable.Rows(vRowNumber).Item("SurveyQuestionNumber").ToString)
              If vPreviousQuestionNumber <> 0 AndAlso vPreviousQuestionNumber <> vQuestionNumber Then
                vSingleQuestionOnPage = False
                Exit For
              Else
                vPreviousQuestionNumber = vQuestionNumber
              End If
            End If
          End If
        Next
        If Not vSingleQuestionOnPage Then
          Exit For
        End If
      Next
      If Not vSingleQuestionOnPage Then
        Exit For
      End If
    Next
    Return vSingleQuestionOnPage
  End Function

  ''' <summary>
  ''' This function returns Previous Page number.
  ''' </summary>
  ''' <param name="pPageNumber"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetPreviousPageNumber(ByVal pPageNumber As Integer) As Integer
    Dim vRowNumber As Integer = 0
    Dim vQuestionNumber As Integer = 0
    Dim vPreviousRowNumber As Integer = 0
    Dim vFound As Boolean = False
    'If there are more than one questions on the current page then go back in sequesnce else if there is only one question
    'on the current page then check if the current qustion is set as next-question for any other question and navigate to it.
    If SingleQuestionOnPage() Then
      For Each vHtmlRow As HtmlTableRow In tblDataEntry.Rows
        For Each vCell As HtmlTableCell In vHtmlRow.Cells
          For Each vControl As Control In vCell.Controls
            If Not TypeOf vControl Is LiteralControl AndAlso Not TypeOf vControl Is Label Then
              If mvControlMap.ContainsKey(vControl.ID) Then
                vRowNumber = mvControlMap(vControl.ID)
                vQuestionNumber = IntegerValue(mvResponsesTable.Rows(vRowNumber).Item("SurveyQuestionNumber").ToString)
              Else
                If TypeOf vControl Is HyperLink AndAlso vControl.ID.StartsWith("SurveyQuestion") Then
                  vQuestionNumber = GetQuestionNumberFromControl(vControl.ID)
                End If
              End If
              For Each vRow As DataRow In mvResponsesTable.Rows
                If IntegerValue(vRow.Item("NextQuestionNumber").ToString) = vQuestionNumber Then
                  vPreviousRowNumber = mvResponsesTable.Rows.IndexOf(vRow)
                  vFound = True
                  Exit For
                End If
              Next
              If Not vFound Then
                Return pPageNumber - 1
              End If
            End If
          Next
          If vFound Then
            Exit For
          End If
        Next
        If vFound Then
          Exit For
        End If
      Next
      Return mvPageMap(vPreviousRowNumber)
    Else
      Return pPageNumber - 1
    End If
  End Function

  ''' <summary>
  ''' This function returs question number depending on ControlId.
  ''' </summary>
  ''' <param name="pControlId"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetQuestionNumberFromControl(ByVal pControlId As String) As Integer
    Dim vControlId As String = pControlId.Remove(0, "SurveyQuestion".Length)
    Return IntegerValue(vControlId)
  End Function

End Class