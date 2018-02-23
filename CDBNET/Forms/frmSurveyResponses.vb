Public Class frmSurveyResponses

  Private mvResponsesTable As DataTable
  Private mvResponseMap As Dictionary(Of String, Integer) 'Used to keep track of the response and its corresponding field on the form
  Private mvFirstLoad As Boolean
  Private Const CONTROL_WIDTH As Integer = 2000

  Public Sub New(ByVal pContactNumber As Integer, ByVal pContactSurveyNumber As Integer)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    mvResponseMap = New Dictionary(Of String, Integer)
    mvFirstLoad = True
    InitialiseControls(pContactNumber, pContactSurveyNumber)
  End Sub

  Private Sub InitialiseControls(ByVal pContactNumber As Integer, ByVal pContactSurveyNumber As Integer)
    Try
      Dim vList As New ParameterList(True)
      vList("ContactSurveyNumber") = pContactSurveyNumber.ToString
      mvResponsesTable = DataHelper.GetTableFromDataSet(DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactSurveyResponses, pContactNumber, vList))
      PopulateControls()
      mvFirstLoad = False
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub
  ''' <summary>
  ''' Populates the controls depending on the answer types and other related values.
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub PopulateControls()
    Try

      If mvResponsesTable IsNot Nothing Then
        Dim vTop As Integer = EditPanelInfo.StartY
        Dim vOffSet As Integer = EditPanelInfo.OffsetY
        Dim vPanelItems As New PanelItems("SurveyResponse")
        Dim vPrevQuestNum As Integer
        Dim vQuestNum As Integer
        Dim vPanelItem As PanelItem = Nothing
        Dim vControlType As PanelItem.ControlTypes
        Dim vParamName As String = String.Empty
        Dim vAnswers() As DataRow = Nothing
        Dim vAddAnswer As Boolean
        Dim vLabel As TransparentLabel = Nothing
        Dim vStartPos As Integer
        Dim vControlWidth As Integer = CInt(CONTROL_WIDTH / AppValues.TwipsConversionX)
        Dim vCaptionWidth As Integer
        Dim vList As New ParameterList
        Dim vHasPattern As Boolean
        Dim vQuestionNumber As String = ""
        Dim vOldQuestion As Integer = 0
        Dim vValid As Boolean

        'The table will contain a row for each available answer. Loop through the rows and create the controls
        For vIndex As Integer = 0 To mvResponsesTable.Rows.Count - 1
          With mvResponsesTable.Rows(vIndex)
            vHasPattern = False
            vAddAnswer = True
            vQuestNum = IntegerValue(.Item("SurveyQuestionNumber"))
            If vPrevQuestNum <> vQuestNum Then
              vStartPos = EditPanelInfo.StartX
              vCaptionWidth = vControlWidth
              'Display the next question. 
              If vPrevQuestNum <> 0 Then vTop += vOffSet
              vPanelItems.Add(New PanelItem("SurveyQuestion" & vQuestNum.ToString, PanelItem.ControlTypes.ctLabelOnly, New Rectangle(EditPanelInfo.StartX, vTop, epl.Width, EditPanelInfo.DefaultHeight), .Item("QuestionText").ToString, epl.Width))

              'Get all the answers for the current question
              vAnswers = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", vQuestNum))
              vPrevQuestNum = vQuestNum

              If vAnswers.Length = 1 AndAlso .Item("AnswerText").ToString.Length = 0 Then
                'The question has only one Answer and the AnswerText is not set 
                'Display the control at the end of the question
                vLabel = New TransparentLabel
                vLabel.Width = vControlWidth
                vLabel.AutoSize = True
                vLabel.Text = .Item("QuestionText").ToString
                vStartPos = vLabel.PreferredWidth + vStartPos
                vCaptionWidth = vStartPos
              Else
                'calculate max width for the caption
                For Each vAnswer As DataRow In vAnswers
                  vLabel = New TransparentLabel
                  vLabel.Width = vControlWidth
                  vLabel.AutoSize = True
                  vLabel.Text = vAnswer.Item("AnswerText").ToString

                  If vLabel.PreferredWidth + EditPanelInfo.StartX > vCaptionWidth Then vCaptionWidth = vLabel.PreferredWidth + EditPanelInfo.StartX
                Next
              End If
            End If
            'Set a default param name. Using the quest and answer num to build the param name
            vParamName = String.Format("SurveyAnswer{0}-{1}", vQuestNum, .Item("SurveyAnswerNumber"))
            If vAnswers.Length > 1 OrElse .Item("AnswerText").ToString.Length <> 0 Then
              'Display each answer on a new line
              vTop += vOffSet 'increase offset to display answers below the question
            End If

            'get type of control to be displayed based on the answer data type
            Select Case .Item("AnswerDataType").ToString
              Case "Y"
                Select Case .Item("AnswerType").ToString.ToUpper
                  Case "M"
                    vControlType = PanelItem.ControlTypes.ctCheckBox
                  Case "S"
                    vControlType = PanelItem.ControlTypes.ctOptionButton
                  Case Else
                    vControlType = PanelItem.ControlTypes.ctComboBox
                    vHasPattern = True
                End Select
              Case "I", "N", "A", "C"
                vControlType = PanelItem.ControlTypes.ctRichTextBox 'PanelItem.ControlTypes.ctTextBox
              Case "D", "T"
                vControlType = PanelItem.ControlTypes.ctDTP
              Case "L"
                vControlType = PanelItem.ControlTypes.ctComboBox
              Case Else
                vAddAnswer = False
            End Select

            If vAddAnswer Then
              If vControlType = PanelItem.ControlTypes.ctOptionButton Then
                'Need to add an underscore for option buttons as it strips away trailing numbers
                vParamName = vParamName & "_"
                vPanelItem = New PanelItem(vParamName, .Item("SurveyAnswerNumber").ToString, New Rectangle(EditPanelInfo.StartX, vTop, vControlWidth, EditPanelInfo.DefaultHeight), .Item("AnswerText").ToString, 0)
                vPanelItem.SetAttributeData("none", .Item("SurveyQuestionNumber").ToString)
              Else
                Dim vNewWidth As Integer = (Me.Width - (EditPanelInfo.StartX * 4) - vControlWidth)
                If vControlType = PanelItem.ControlTypes.ctRichTextBox Then
                  vPanelItem = New PanelItem(vParamName, vControlType, New Rectangle(EditPanelInfo.StartX, vTop, vNewWidth, (EditPanelInfo.DefaultHeight * 3)), .Item("AnswerText").ToString, vCaptionWidth)
                  vTop += (vOffSet * 2)
                Else
                  vPanelItem = New PanelItem(vParamName, vControlType, New Rectangle(EditPanelInfo.StartX, vTop, vNewWidth, EditPanelInfo.DefaultHeight), .Item("AnswerText").ToString, vCaptionWidth)
                End If


              End If

              Select Case .Item("AnswerDataType").ToString
                Case "L"
                  'replace "," with "|" so the list works off the existing code that handles patterns
                  'Use upper case as SelectComboBoxItem expects the lookupcode to be in upper case
                  vPanelItem.Pattern = .Item("ListValues").ToString.ToUpper.Replace(", ", ",").Replace(","c, "|"c)
                Case "T"
                  vPanelItem.FieldType = PanelItem.FieldTypes.cftTime
                Case "Y"
                  If vHasPattern Then vPanelItem.Pattern = "YES|NO"
              End Select

              vPanelItem.Mandatory = BooleanValue(.Item("Mandatory").ToString)
              If .Item("AnswerDataType").ToString = "C" AndAlso .Item("AnswerType").ToString = "S" Then
                vPanelItem.Mandatory = False
              End If

              vPanelItems.Add(vPanelItem)
                'Add response if any
              If .Item("ResponseAnswerText").ToString.Length > 0 Then vList(vParamName) = .Item("ResponseAnswerText").ToString
                'Keep track of the param name used for the response so that it can be retrieved later to compare/update the values
                mvResponseMap.Add(vParamName, vIndex)
              End If
          End With
        Next
        epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optSurveyResponse, vPanelItems))
        ' Check if the question for which answers are Radio Buttons, is mandatory and Set a default value.
        For Each vPanelItem In epl.PanelInfo.PanelItems
          If vPanelItem.ControlType = PanelItem.ControlTypes.ctOptionButton Then
            If Not vList.Contains(vPanelItem.ParameterName) AndAlso BooleanValue(mvResponsesTable.Rows(mvResponseMap(vPanelItem.ParameterName)).Item("Mandatory").ToString) Then
              Dim vQuestNumber As Integer = IntegerValue(mvResponsesTable.Rows(mvResponseMap(vPanelItem.ParameterName)).Item("SurveyQuestionNumber").ToString)
              vAnswers = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", vQuestNumber))
              For Each vAnswer As DataRow In vAnswers
                If epl.FindRadioButton(vPanelItem.ParameterName & String.Format("_{0}", mvResponsesTable.Rows(mvResponseMap(vPanelItem.ParameterName)).Item("SurveyAnswerNumber"))).Checked Then
                  vValid = True
                End If
              Next
              If Not vValid Then
                If Not vList.Contains(vPanelItem.ParameterName) Then
                  Dim vAdd As Boolean = True
                  For Each vItem As DictionaryEntry In vList
                    If vItem.Key.ToString.Contains("SurveyAnswer" & vQuestNumber.ToString & "-") Then
                      vAdd = False
                    End If
                  Next
                  If vAdd Then vList(vPanelItem.ParameterName) = mvResponsesTable.Rows(mvResponseMap(vPanelItem.ParameterName)).Item("SurveyAnswerNumber").ToString
                End If
              End If
            End If
          End If
        Next
        If vList.Count > 0 Then epl.Populate(vList)
      End If
    Catch ex As Exception
      DataHelper.HandleException(ex)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      Dim vFormHasErrors As Boolean = False
      Dim vList As New ParameterList(True)
      Dim vAnswer As ParameterList = Nothing
      Dim vValid As Boolean = epl.AddValuesToList(vList, True, EditPanel.AddNullValueTypes.anvtAll)
      If vValid AndAlso mvResponsesTable IsNot Nothing Then
        For Each vParam As String In vList.Keys
          Try
            If mvResponseMap.ContainsKey(vParam) Then
              If (mvResponsesTable.Rows(mvResponseMap(vParam)).Item("ResponseAnswerText").ToString <> vList(vParam)) OrElse
                (mvResponsesTable.Rows(mvResponseMap(vParam)).Item("AnswerDataType").ToString = "Y" AndAlso mvResponsesTable.Rows(mvResponseMap(vParam)).Item("AnswerType").ToString.ToUpper = "S") Then 'Update only if the response has changed
                If mvResponsesTable.Rows(mvResponseMap(vParam)).Item("AnswerDataType").ToString = "Y" AndAlso mvResponsesTable.Rows(mvResponseMap(vParam)).Item("AnswerType").ToString.ToUpper = "S" Then
                  'Reset other values for this question
                  Dim vRows() As DataRow = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(mvResponseMap(vParam)).Item("SurveyQuestionNumber")))
                  For Each vRow As DataRow In vRows
                    vAnswer = New ParameterList(True)
                    vAnswer("ContactSurveyNumber") = vRow.Item("ContactSurveyNumber").ToString
                    vAnswer("SurveyQuestionNumber") = vRow.Item("SurveyQuestionNumber").ToString
                    vAnswer("SurveyAnswerNumber") = vRow.Item("SurveyAnswerNumber").ToString
                    vAnswer("ResponseAnswerText") = IIf(vRow.Item("SurveyAnswerNumber").ToString = mvResponsesTable.Rows(mvResponseMap(vParam)).Item("SurveyAnswerNumber").ToString, vList(vParam), String.Empty).ToString
                    DataHelper.UpdateContactSurveyResponse(vAnswer)
                  Next
                ElseIf mvResponsesTable.Rows(mvResponseMap(vParam)).Item("AnswerDataType").ToString = "C" AndAlso mvResponsesTable.Rows(mvResponseMap(vParam)).Item("AnswerType").ToString.ToUpper = "S" Then
                  'Reset other values for this question
                  Dim vRows() As DataRow = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", mvResponsesTable.Rows(mvResponseMap(vParam)).Item("SurveyQuestionNumber")))
                  For Each vRow As DataRow In vRows
                    If Not vList.Contains(String.Format("SurveyAnswer{0}-{1}_", vRow.Item("SurveyQuestionNumber"), vRow.Item("SurveyAnswerNumber"))) Then
                      vAnswer = New ParameterList(True)
                      vAnswer("ContactSurveyNumber") = vRow.Item("ContactSurveyNumber").ToString
                      vAnswer("SurveyQuestionNumber") = vRow.Item("SurveyQuestionNumber").ToString
                      vAnswer("SurveyAnswerNumber") = vRow.Item("SurveyAnswerNumber").ToString
                      vAnswer("ResponseAnswerText") = IIf(vRow.Item("SurveyAnswerNumber").ToString = mvResponsesTable.Rows(mvResponseMap(vParam)).Item("SurveyAnswerNumber").ToString, vList(vParam), String.Empty).ToString
                      DataHelper.UpdateContactSurveyResponse(vAnswer)
                    End If
                  Next
                Else
                  vAnswer = New ParameterList(True)
                  vAnswer("ContactSurveyNumber") = mvResponsesTable.Rows(mvResponseMap(vParam)).Item("ContactSurveyNumber").ToString
                  vAnswer("SurveyQuestionNumber") = mvResponsesTable.Rows(mvResponseMap(vParam)).Item("SurveyQuestionNumber").ToString
                  vAnswer("SurveyAnswerNumber") = mvResponsesTable.Rows(mvResponseMap(vParam)).Item("SurveyAnswerNumber").ToString
                  vAnswer("ResponseAnswerText") = vList(vParam)
                  DataHelper.UpdateContactSurveyResponse(vAnswer)
                End If
              End If
            End If
          Catch vEx As CareException
            vFormHasErrors = True
            Select Case vEx.ErrorNumber
              Case CareException.ErrorNumbers.enSurveyResponseIsNotaValidCharacter,
                CareException.ErrorNumbers.enSurveyResponseIsNotaValidDate,
                CareException.ErrorNumbers.enSurveyResponseIsNotaValidInteger,
                CareException.ErrorNumbers.enSurveyResponseIsNotaValidNumber,
                CareException.ErrorNumbers.enSurveyResponseIsNotaValidTime,
                CareException.ErrorNumbers.enSurveyResponseNotYorN,
                CareException.ErrorNumbers.enSurveyResponseValueGreaterThanMaximum,
                CareException.ErrorNumbers.enSurveyResponseValueLessThanMinimum,
                CareException.ErrorNumbers.enResponseNotInListOfValidResponses
                Me.epl.SetErrorField(vParam, vEx.Message)
              Case CareException.ErrorNumbers.enMaximumValueGreaterThanMinimum,
                CareException.ErrorNumbers.enSurveyAnswerDataTypeInvalid,
                CareException.ErrorNumbers.enSurveyAnswerListEmpty,
                CareException.ErrorNumbers.enSurveyAnswerListNotAppropriate,
                CareException.ErrorNumbers.enSurveyAnswerMaximumNotAppropriate,
                CareException.ErrorNumbers.enSurveyAnswerMinimumNotAppropriate,
                CareException.ErrorNumbers.enSurveyAnswerRangeNotAppropriate,
                CareException.ErrorNumbers.enSurveyQuestionIsMandatory
                Me.epl.SetErrorField(vParam, vEx.Message)
              Case CareException.ErrorNumbers.enContactSurveyNumberInvalid,
              CareException.ErrorNumbers.enSurveyAnswerNumberInvalid
                ShowErrorMessage(vEx.Message)
            End Select
          End Try
        Next
        If Not vFormHasErrors Then
          Me.DialogResult = System.Windows.Forms.DialogResult.OK
          Me.Close()
        End If
      End If
    Catch vEx As CareException
      DataHelper.HandleException(vEx)
    End Try
  End Sub

  Private Sub epl_ValueChanged(ByVal sender As System.Object, ByVal pParameterName As System.String, ByVal pValue As System.String) Handles epl.ValueChanged
    If Not mvFirstLoad Then
      If mvResponseMap.ContainsKey(pParameterName) Then
        Dim vParamName As String
        Dim vControl As Control = Nothing
        Dim vQuestionNumber As Integer
        With mvResponsesTable.Rows(mvResponseMap(pParameterName))
          ' Check for NextQuestion, If present, set focus to NextQuestion
          If .Item("NextQuestionNumber").ToString.Length > 0 AndAlso mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", .Item("NextQuestionNumber").ToString)).Length > 0 Then
            Dim vRow As DataRow = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}'", .Item("NextQuestionNumber").ToString))(0)
            If vRow IsNot Nothing AndAlso vRow.Item("SurveyAnswerNumber").ToString.Length > 0 Then
              vParamName = String.Format("SurveyAnswer{0}-{1}", .Item("NextQuestionNumber").ToString, vRow.Item("SurveyAnswerNumber"))
              'Handling Option Buttons as special case since EditPanel adds '_' for option buttons.
              If vRow.Item("AnswerDataType").ToString = "Y" AndAlso vRow.Item("AnswerType").ToString.ToUpper = "S" Then
                vParamName = String.Format(String.Format("SurveyAnswer{0}-{1}", .Item("NextQuestionNumber").ToString, vRow.Item("SurveyAnswerNumber")) & "__{0}", vRow.Item("SurveyAnswerNumber"))
              End If
            Else
              vParamName = String.Format("SurveyQuestion{0}_Label", .Item("NextQuestionNumber").ToString)
            End If
            vControl = epl.FindPanelControl(vParamName, False)
            If vControl IsNot Nothing Then
              epl.ScrollControlIntoView(vControl)
              vControl.Focus()
            End If
            epl.Refresh()
          End If

          'Check for TextBox in Question of type "S" 
          If .Item("AnswerType").ToString = "S" AndAlso .Item("AnswerDataType").ToString = "C" Then
            vQuestionNumber = IntegerValue(.Item("SurveyQuestionNumber"))
            Dim vChecked As Boolean = False
            Dim vRows() As DataRow = Nothing
            If mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}' and AnswerDataType = 'C'", vQuestionNumber)).Length > 0 Then
              vRows = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}' and AnswerDataType = 'Y'", vQuestionNumber))
              For Each vRow As DataRow In vRows
                If pValue.Length > 0 Then
                  epl.FindRadioButton(String.Format("SurveyAnswer{0}-{1}__{2}", vRow.Item("SurveyQuestionNumber").ToString, vRow("SurveyAnswerNumber").ToString, vRow("SurveyAnswerNumber").ToString)).Checked = False
                  vChecked = True
                Else
                  If vRow.Item("ResponseAnswerText").ToString.Length > 0 Then
                    epl.FindRadioButton(String.Format("SurveyAnswer{0}-{1}__{2}", vRow.Item("SurveyQuestionNumber").ToString, vRow("SurveyAnswerNumber").ToString, vRow("SurveyAnswerNumber").ToString)).Checked = True
                    vChecked = True
                  End If
                End If
              Next
              If Not vChecked AndAlso BooleanValue(.Item("Mandatory").ToString) Then
                epl.FindRadioButton(String.Format("SurveyAnswer{0}-{1}__{2}", vRows(0).Item("SurveyQuestionNumber").ToString, vRows(0)("SurveyAnswerNumber").ToString, vRows(0)("SurveyAnswerNumber").ToString)).Checked = True
              End If
            End If
          End If
          'Handle RadioButton click for questions which have AnswerDataType 'C'
          If .Item("AnswerType").ToString = "S" AndAlso .Item("AnswerDataType").ToString = "Y" Then
            vQuestionNumber = IntegerValue(.Item("SurveyQuestionNumber"))
            Dim vRows() As DataRow = Nothing
            vRows = mvResponsesTable.Select(String.Format("SurveyQuestionNumber = '{0}' and AnswerDataType = 'C'", vQuestionNumber))
            If vRows.Length > 0 Then
              For Each vRow As DataRow In vRows
                epl.SetValue(String.Format("SurveyAnswer{0}-{1}", vRow.Item("SurveyQuestionNumber"), vRow.Item("SurveyAnswerNumber")), "")
              Next
            End If
          End If
        End With
      End If
    End If
  End Sub

  Private Sub epl_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles epl.Scroll
    epl.Refresh()
  End Sub

  Private Sub epl_validateAllItems(ByVal sender As Object, ByVal pList As ParameterList, ByRef pValid As Boolean) Handles epl.ValidateAllItems
    Dim vPanelItem As PanelItem
    Dim vGroup As Boolean = False
    Dim vQuestionNumber As String = ""
    Dim vOldQuestion As Integer = 0
    Dim vValid As Boolean = False
    Dim vReturn As Boolean = True
    Dim vMandatory As Boolean = False
    For Each vPanelItem In epl.PanelInfo.PanelItems
      If vPanelItem.ControlType = PanelItem.ControlTypes.ctCheckBox Then
        With mvResponsesTable.Rows(mvResponseMap(vPanelItem.ParameterName))
          vMandatory = BooleanValue(.Item("Mandatory").ToString)

          If vOldQuestion <> 0 AndAlso vOldQuestion <> IntegerValue(.Item("SurveyQuestionNumber")) Then
            If Not vValid Then
              epl.SetErrorField(vQuestionNumber, InformationMessages.ImQuestionAnswerMandatory)
              vReturn = False
            End If
            vValid = False
          End If
          If vMandatory Then
            If vPanelItem.ControlType = PanelItem.ControlTypes.ctCheckBox AndAlso epl.FindCheckBox(vPanelItem.ParameterName).Checked Then
              vValid = True
            End If
            vQuestionNumber = vPanelItem.ParameterName
            vOldQuestion = IntegerValue(.Item("SurveyQuestionNumber"))
          End If
        End With
      End If
    Next
    If Not vValid AndAlso vMandatory Then
      epl.SetErrorField(vQuestionNumber, InformationMessages.ImQuestionAnswerMandatory)
      vReturn = False
    End If
    pValid = vReturn
  End Sub
End Class