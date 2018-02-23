Imports System.Linq
Public Class frmExamResults

  Private mvExamMarkType As String
  Private mvExamCompMarkType As String = "M"
  Private mvComponentBasedEntry As Boolean = False
  Private mvTabFromMark As Boolean = False
  Friend WithEvents erp As System.Windows.Forms.ErrorProvider

  Private Const GRADE = "G"
  Private Const MARK = "M"
  Private Const RESULT = "P"
  Private mvUnitMarkType As New Dictionary(Of String, String)
  Private mvMultipleUnits As Boolean
  Private Const COLUMN_BOOKING_ID = "BookingId"
  Private Const COLUMN_CHECK = "Check"
  Private Const COLUMN_GRADE_TYPE = "GradeType"

  Public Sub New()
    InitializeComponent()
    InitialiseControls()
    ' TODO Set Form Title from C:\dev32\CDBNETCL\Settings\ControlText.strings
    Me.Text = "Exam Result Entry"

    TabFindReplaceHost.ItemSize = New Size(0, 1)
    TabFindReplaceHost.SizeMode = TabSizeMode.Fixed

    AddHandler txtMarkFind.KeyPress, AddressOf NumericKeyPressHandler
    AddHandler txtMarkReplace.KeyPress, AddressOf NumericKeyPressHandler

    'Create a dummy PanelItem for the Raw Mark tag to force it to behave like a numeric field with 3 decimals
    Dim vDummyPanel As New PanelItem("RawMark", PanelItem.ControlTypes.ctTextBox, txtMarkFind.DisplayRectangle, lblFindWhat.Text, 0, PanelItem.FieldTypes.cftNumeric)
    vDummyPanel.SetAttributeData("exam_booking_units", "raw_mark")
    txtMarkFind.Tag = vDummyPanel
    txtMarkReplace.Tag = vDummyPanel

    AddHandler txtMarkFind.Validating, AddressOf NumericReformatHandler
    AddHandler txtMarkReplace.Validating, AddressOf NumericReformatHandler

  End Sub

  Private Sub InitialiseControls()
    Try
      SetControlTheme()
      SettingsName = "frmExamResults"
      splMain.FixedPanel = FixedPanel.Panel1
      dgr.Clear()
      dgrComponents.Clear()
      ShowHideComponentGrid(False)

      If AppValues.ConfigurationValue(AppValues.ConfigurationValues.ex_allow_multiple_results, "N") = "Y" Then mvMultipleUnits = True

      epl.Init(New EditPanelInfo(CDBNETCL.CareNetServices.FunctionParameterTypes.fptExamResultEntry))
      Me.erp = New System.Windows.Forms.ErrorProvider(Me)
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub epl_GetCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByVal pList As CDBNETCL.ParameterList) Handles epl.GetCodeRestrictions
    Select Case pParameterName
      Case "ExamUnitDescription"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList.IntegerValue("AllUnits") = 1
        pList.Item("AllowMarkEntry") = "Y"
        pList.Item("ExcludeQuestionUnits") = "Y"
        pList.Item("ResultEntry") = "Y"
        If FindControl(epl, "ExamSessionCode", False) IsNot Nothing And FindControl(epl, "ExamUnitDescription", False) IsNot Nothing Then
          Dim vSessionLookUp As CDBNETCL.TextLookupBox = CType(FindControl(epl, "ExamSessionCode", False), TextLookupBox)
          pList("ExamSessionId") = vSessionLookUp.GetDataRowItem("ExamSessionId")
        End If
        If FindControl(epl, "ExamCentreCode", False) IsNot Nothing Then
          Dim vCentreLookUp As CDBNETCL.TextLookupBox = CType(FindControl(epl, "ExamCentreCode", False), TextLookupBox)
          pList("ExamCentreId") = vCentreLookUp.GetDataRowItem("ExamCentreId")
          If pList("ExamCentreId").Length > 0 Then
            pList("CentreUnitDescription") = "Y"
            pList("AllowSelection") = "N"
          End If
        End If

      Case "ExamCentreCode"
        pList.Item("ResultEntry") = "Y"
    End Select
  End Sub

  Private Sub epl_GetInitialCodeRestrictions(ByVal sender As Object, ByVal pParameterName As String, ByRef pList As ParameterList) Handles epl.GetInitialCodeRestrictions
    Select Case pParameterName
      Case "ExamSessionCode"
        If pList Is Nothing Then pList = New ParameterList(True)
        pList.Item("NonSessionBased") = "Y"
    End Select
  End Sub


  Private Sub epl_ValueChanged(ByVal sender As System.Object, ByVal pParameterName As System.String, ByVal pValue As System.String) Handles epl.ValueChanged
    Dim vEPL As EditPanel = DirectCast(sender, EditPanel)

    Select Case pParameterName
      Case "ExamSessionCode"
        Dim vTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("ExamSessionCode")
        Dim vSessionId As Integer = vTextLookupBox.GetDataRowInteger("ExamSessionId")
        Dim vUnitTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("ExamUnitDescription")
        vUnitTextLookupBox.FillComboWithRestriction(vSessionId.ToString)
        If vSessionId = 0 Then
          vEPL.PanelInfo.PanelItems("ExamCentreCode").Mandatory = False
        End If
        'vEPL.SetValue("ExamUnitCode", "")
        vEPL.SetValue("ExamUnitDescription", "")
        vUnitTextLookupBox = vEPL.FindTextLookupBox("ExamCentreCode")
        vUnitTextLookupBox.FillComboWithRestriction(vSessionId.ToString)
        vEPL.SetValue("ExamCentreCode", "")
      Case "ExamCentreCode"
        Dim vList As New ParameterList(True)
        Dim vTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox(pParameterName)
        Dim vCentreId As Integer = vTextLookupBox.GetDataRowInteger("ExamCentreId")
        Dim vSessionCode As TextLookupBox = vEPL.FindTextLookupBox("ExamSessionCode")
        Dim vSessionId As Integer = vSessionCode.GetDataRowInteger("ExamSessionId")
        If vCentreId > 0 Then vList("ExamCentreId") = vCentreId.ToString()
        vList("ExamSessionId") = vSessionId.ToString()
        vList("CentreCode") = "Y"
        'Dim vUnitTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("ExamUnitCode")
        Dim vUnitTextLookupBox As TextLookupBox = vEPL.FindTextLookupBox("ExamUnitDescription")
        vUnitTextLookupBox.FillComboWithRestriction(vCentreId.ToString, "", False, vList)
        vEPL.SetValue("ExamUnitDescription", "")
    End Select
  End Sub


  Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click
    Try
      Dim pList As New ParameterList(True, True)
      If epl.AddValuesToList(pList, True) Then
        dgr.Clear()
        dgr.Refresh()
        pList.IntegerValue("ExamSessionId") = epl.FindTextLookupBox("ExamSessionCode").GetDataRowInteger("ExamSessionId")
        pList.IntegerValue("ExamCentreId") = epl.FindTextLookupBox("ExamCentreCode").GetDataRowInteger("ExamCentreId")
        pList.IntegerValue("ExamUnitId") = epl.FindTextLookupBox("ExamUnitDescription").GetDataRowInteger("ExamUnitId")
        pList.IntegerValue("ExamUnitLinkId") = epl.FindTextLookupBox("ExamUnitDescription").GetDataRowInteger("ExamUnitLinkId")
        Dim vDataSet As DataSet = ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamStudentResults, pList)

        If vDataSet IsNot Nothing AndAlso vDataSet.Tables.Contains("DataRow") Then
          ' Stop the ContactName appearing as a hyper link in the Dialog based grid by renaming it
          If vDataSet.Tables("DataRow").Columns.Contains("ContactName") Then
            vDataSet.Tables("DataRow").Columns("ContactName").ColumnName = "CandidateName"
            Dim vRowArray() As DataRow = vDataSet.Tables("Column").Select("Name = 'ContactName'")
            vRowArray(0).Item("Name") = "CandidateName"
          End If
        End If

        'Initialise Change Reason
        InitTextLookupBox(txtChangeReason, "exam_grade_change_reasons", "exam_grade_change_reason")

        If mvMultipleUnits Then
          SetupForMultipleResultEntry(vDataSet)
        Else
          'Me.TabFindReplacePage.Visible = False
          'Me.TabPage3.Visible = False
          dgr.Populate(vDataSet)
          dgr.SetCellsEditable()
          If dgr.RowCount > 0 Then
            dgr.SetCellsReadOnly(-1, dgr.GetColumn("Name"), True)
            dgr.SetColumnReadOnly(dgr.GetColumn("Name"), True)
            mvComponentBasedEntry = (dgr.GetValue(0, "ExamUnitChildLink").ToString.Length > 0)
            mvExamMarkType = dgr.GetValue(0, "ExamMarkType").ToString
            dgr.SetColumnVisible("RawMark", (mvComponentBasedEntry Or (mvExamMarkType = "M")))
            dgr.SetColumnVisible("OriginalGrade", (Not mvComponentBasedEntry And (mvExamMarkType = "G")))
            dgr.SetColumnVisible("OriginalResult", (Not mvComponentBasedEntry And (mvExamMarkType = "P")))
            If Not mvComponentBasedEntry Then
              dgr.SetColumnWritable("RawMark")
            Else
              dgr.SetCellsReadOnly(-1, dgr.GetColumn("RawMark"), True, True)
              dgr.SetColumnReadOnly(dgr.GetColumn("RawMark"))
            End If

            dgr.SetColumnWritable("OriginalResult")
            dgr.SetColumnWritable("OriginalGrade")
            dgr.SetColumnReadOnly(dgr.GetColumn("ContactNumber"), True, True)
            dgr.SetColumnReadOnly(dgr.GetColumn("CandidateName"), True, True)

            If Not mvComponentBasedEntry Then
              ShowHideComponentGrid(False)
              Dim vDataTable As DataTable = Nothing
              Dim vParams As New ParameterList(True)
              Select Case mvExamMarkType
                Case "G"
                  vDataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamGrades, vParams)
                  PopulateGridCombo(dgr, "OriginalGrade", "ExamGrade", "ExamGrade", vDataTable)
                  txtMark.Visible = False
                  txtGrade.Visible = True
                  txtResult.Visible = False
                  InitTextLookupBox(txtGrade, "exam_grades", "exam_grade")
                  lblMark.Text = "Grade"
                Case "M"
                  txtMark.Visible = True
                  txtGrade.Visible = False
                  txtResult.Visible = False
                  lblMark.Text = "Marks"
                Case "P"
                  vDataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamResultTypes, vParams)
                  PopulateGridCombo(dgr, "OriginalResult", "LookUpCode", "LookUpDesc", vDataTable)
                  txtMark.Visible = False
                  txtGrade.Visible = False
                  txtResult.Visible = True
                  InitTextLookupBox(txtResult, "exam_booking_units", "original_result")
                  lblMark.Text = "P/F"
              End Select
            End If
            dgr.SelectRow(0)
          End If

          LoadComponentGrid()
        End If
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function PopulateGridCombo(ByVal Grid As CDBNETCL.DisplayGrid, ByVal comboColumn As String, ByVal codeColumn As String, ByVal DescriptionColumn As String, ByVal SourceData As DataTable) As Boolean
    Try
      Dim vItemsCodes() As String = {""}
      Dim vItemsDescriptions() As String = {""}

      If (comboColumn <> "") AndAlso (codeColumn <> "") AndAlso (DescriptionColumn <> "") AndAlso (SourceData IsNot Nothing) AndAlso (SourceData.Rows.Count > 0) Then
        Dim vIndex As Integer = 2

        For Each vRow As DataRow In SourceData.Rows
          Array.Resize(vItemsCodes, vIndex)
          Array.Resize(vItemsDescriptions, vIndex)
          vItemsCodes.SetValue(vRow(codeColumn).ToString, vIndex - 1)
          vItemsDescriptions.SetValue(vRow(DescriptionColumn).ToString, vIndex - 1)
          vIndex += 1
        Next
      End If

      Grid.SetComboBoxColumn(comboColumn, vItemsDescriptions, vItemsCodes)
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Function

  Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
    If mvComponentBasedEntry Then
      SaveComponentGrid()
    Else
      Dim vSaved As Boolean = False
      vSaved = SaveGrid()
      If vSaved Then
        cmdSelect_Click(sender, e) 'To refresh datagrid dgr with current db values
      End If
    End If
  End Sub

  Private Function SaveGrid() As Boolean
    Try
      If dgr.RowCount > 0 Then
        Dim vNewExamResult As String = ""
        Dim vOldExamResult As String = ""
        Dim vReturnList As ParameterList = New ParameterList()
        Dim vResultColumnName As String = ""
        Dim vTotalColumnName As String = ""
        Dim vChangeReason As String = ""

        If GetChangeReason(vChangeReason) Then

          If mvMultipleUnits Then
            For vRow As Integer = 0 To dgr.RowCount - 1
              For vCol As Integer = 0 To dgr.ColumnCount - 1

                If dgr.ColumnName(vCol) <> "ContactNumber" AndAlso
                  dgr.ColumnName(vCol) <> "CandidateName" AndAlso
                  dgr.ColumnName(vCol) <> "ExamMarkType" AndAlso
                  dgr.ColumnName(vCol) <> "ExamBookingUnitId" AndAlso
                  Not dgr.ColumnName(vCol).Contains(COLUMN_CHECK) AndAlso
                  Not dgr.ColumnName(vCol).Contains(COLUMN_BOOKING_ID) AndAlso
                  Not dgr.ColumnName(vCol).Contains(COLUMN_GRADE_TYPE) Then
                  Dim vList As ParameterList = New ParameterList(True)
                  Dim vColumnName As String = dgr.ColumnName(vCol)

                  vNewExamResult = dgr.GetValue(vRow, vColumnName)
                  vOldExamResult = dgr.GetValue(vRow, vColumnName & COLUMN_CHECK)

                  If mvMultipleUnits Then
                    GetMarkColumns(mvUnitMarkType.Item(vColumnName), vResultColumnName, vTotalColumnName)
                  Else
                    GetMarkColumns(dgr.GetValue(vRow, "ExamMarkType").Trim, vResultColumnName, vTotalColumnName)
                  End If


                  If (vNewExamResult <> vOldExamResult) Then
                    vList(vResultColumnName) = vNewExamResult
                    vList(vTotalColumnName) = vNewExamResult
                    If vChangeReason.Length > 0 Then vList("ExamGradeChangeReason") = vChangeReason

                    If mvMultipleUnits Then
                      vList("ExamBookingUnitId") = dgr.GetValue(vRow, vColumnName & COLUMN_BOOKING_ID)
                    Else
                      vList("ExamBookingUnitId") = dgr.GetValue(vRow, "ExamBookingUnitId")
                    End If

                    If vList("ExamBookingUnitId").ToString.Length = 0 Then
                      Throw New CareException(String.Format(InformationMessages.ImExamBookingNotFound, dgr.GetValue(vRow, "ContactNumber"), vColumnName))
                    End If
                    vReturnList = ExamsDataHelper.UpdateItem(ExamsAccess.XMLExamMaintenanceTypes.ExamBookingUnit, vList)
                  End If
                End If
              Next
            Next
          Else
            GetMarkColumns(mvExamMarkType, vResultColumnName, vTotalColumnName)

            For row As Integer = 0 To dgr.RowCount - 1
              Dim vList As ParameterList = New ParameterList(True)
              vNewExamResult = dgr.GetValue(row, vResultColumnName)
              vOldExamResult = dgr.GetValue(row, vResultColumnName & COLUMN_CHECK)
              If (vNewExamResult <> vOldExamResult) Then
                Dim vSameValue As Boolean = False
                Dim vColumnNumber As Integer = dgr.GetColumn(vResultColumnName)
                Dim vDataType As DBField.FieldTypes = dgr.GetDataType(vColumnNumber)
                If dgr.GetDataType(vColumnNumber) = DBField.FieldTypes.cftNumeric Then
                  'Could be format different but numerically the same so check numerically e.g. 10 <> 10.000
                  Dim vNewAdjusted As Decimal
                  Dim vOldAdjusted As Decimal
                  If Decimal.TryParse(vNewExamResult, vNewAdjusted) And Decimal.TryParse(vOldExamResult, vOldAdjusted) Then
                    If vNewAdjusted = vOldAdjusted Then
                      vSameValue = True
                    End If
                  End If
                End If
                If Not vSameValue Then
                  vList(vResultColumnName) = vNewExamResult
                  vList(vTotalColumnName) = vNewExamResult
                  If vChangeReason.Length > 0 Then vList("ExamGradeChangeReason") = vChangeReason
                  vList("ExamBookingUnitId") = dgr.GetValue(row, "ExamBookingUnitId")
                  vReturnList = ExamsDataHelper.UpdateItem(ExamsAccess.XMLExamMaintenanceTypes.ExamBookingUnit, vList)
                End If
              End If
            Next
          End If
          Return True
        Else
          Return False
        End If
      End If
    Catch vException As CareException
      Select Case vException.ErrorNumber
        Case CareException.ErrorNumbers.enParameterInvalidValue
          ShowErrorMessage(vException.Message)
        Case Else
          DataHelper.HandleException(vException)
      End Select

    End Try
  End Function

  Private Function quickEntryGridUpdate() As Boolean
    If mvComponentBasedEntry Then
      ' Put mark in grid, in currently selected row
      If dgrComponents.ActiveRow > -1 Then

        Dim vColNo As Integer = dgrComponents.GetColumn("RawMark")
        Dim vNewMark As Double
        Dim vCurrentMark As Double = CDbl(dgrComponents.GetRawValue(dgrComponents.ActiveRow, vColNo))
        If Double.TryParse(txtMark.Text, vNewMark) Then
          Dim vValueChanged As Boolean = Not (vCurrentMark = vNewMark)
          If vValueChanged Then
            dgrComponents.SetValue(dgrComponents.ActiveRow, vColNo, vNewMark)
            UpdateComponentMarkSum()
          End If

          If dgrComponents.ActiveRow < dgrComponents.RowCount - 1 Then
            ' Move focus in grid to next row
            dgrComponents.SelectRow(dgrComponents.ActiveRow + 1, True)
          ElseIf dgrComponents.ActiveRow = dgrComponents.RowCount - 1 Then
            ' Move main grid focus to next Contact
            If SaveComponentGridCheck() Then
              SaveComponentGrid()
              If dgr.ActiveRow < dgr.RowCount - 1 Then dgr.SelectRow(dgr.ActiveRow + 1)
              If dgrComponents.ActiveRow < 0 AndAlso dgrComponents.RowCount > 0 Then dgrComponents.SelectRow(0, True)
            End If
          End If

          If dgrComponents.ActiveRow > -1 Then
            txtMark.Text = dgrComponents.GetValue(dgrComponents.ActiveRow, "RawMark")
          Else
            txtMark.Text = ""
          End If

          If txtMark.Text <> String.Empty Then txtMark.SelectAll()
          quickEntryGridUpdate = True
        End If
      End If

    Else
      Dim vCurrentRow As Integer = dgr.FindRow("ContactNumber", txtContactNumber.Text)
      If vCurrentRow > -1 Then
        Select Case mvExamMarkType
          Case "M"
            Dim vNewMark As Double
            If Double.TryParse(txtMark.Text, vNewMark) Then
              Dim vColNo As Integer = dgr.GetColumn("RawMark")
              dgr.SetValue(vCurrentRow, vColNo, vNewMark)
              quickEntryGridUpdate = Not String.IsNullOrEmpty(txtMark.Text)
            End If
          Case "G"
            dgr.SetValue(vCurrentRow, "OriginalGrade", txtGrade.Text)
            quickEntryGridUpdate = Not String.IsNullOrEmpty(txtGrade.Text)
          Case "P"
            'dgr.SetValue(vCurrentRow, "OriginalResult", "Edit")
            dgr.SetValue(vCurrentRow, "OriginalResult", txtResult.Text)
            quickEntryGridUpdate = Not String.IsNullOrEmpty(txtResult.Text)
        End Select

        dgr.SelectRow("ContactNumber", txtContactNumber.Text)
      Else
        If Not String.IsNullOrEmpty(txtContactNumber.Text) Then ShowInformationMessage(InformationMessages.ImInvalidContactNumber)
        quickEntryGridUpdate = False
      End If
    End If
  End Function

  Private Sub dgr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgr.KeyDown
    If e.KeyCode = 46 Then
      dgr.SetValue(dgr.ActiveRow, dgr.ActiveColumn, "")
      UpdateQuickEntry(dgr.ActiveRow)
      e.SuppressKeyPress = True
    End If
  End Sub

  Private Sub dgr_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgr.KeyUp
    UpdateQuickEntry(dgr.ActiveRow)
  End Sub

  Private Sub dgr_RowSelected(ByVal sender As System.Object, ByVal pRow As System.Int32, ByVal pDataRow As System.Int32) Handles dgr.RowSelected
    UpdateQuickEntry(pRow)

    If dgr.ActiveRow > -1 Then
      ' Load components for selected contact
      LoadComponentGrid()
    End If
  End Sub

  Private Sub InitTextLookupBox(ByRef pTxtLookup As TextLookupBox, ByVal pValidationTable As String, ByVal pValidationAttribute As String, Optional ByVal pVal As String = "")
    Dim vParamList As New ParameterList(True)
    vParamList("TableName") = pValidationTable
    vParamList("FieldName") = pValidationAttribute
    vParamList("FieldType") = "C"  ' Character FieldType
    Dim vParams As ParameterList = DataHelper.GetMaintenanceData(vParamList)
    vParams("AttributeName") = pValidationAttribute
    vParams("ValidationAttribute") = pValidationAttribute
    vParams("ValidationTable") = pValidationTable
    pTxtLookup.BackColor = Me.BackColor
    Dim vPanelItem As PanelItem = New PanelItem(pTxtLookup, pValidationAttribute)
    vPanelItem.InitFromMaintenanceData(vParams)
    pTxtLookup.Tag = vPanelItem
    pTxtLookup.Init(vPanelItem, False, False)
    AddHandler pTxtLookup.Validating, AddressOf LookupValidatingHandler
  End Sub

  Private Function ValidateControl(ByVal pControl As System.Windows.Forms.Control, ByVal pPanelItem As PanelItem, ByVal pValue As String) As Boolean
    Dim vValid As Boolean = True

    If pPanelItem.ValidationError Then
      vValid = False
    Else
      erp.SetError(pControl, "")                                'Clear any errors
    End If

    If TypeOf pControl Is TextLookupBox AndAlso DirectCast(pControl, TextLookupBox).IsValid = False Then
      erp.SetError(pControl, GetInformationMessage(InformationMessages.ImInvalidValue))
      vValid = False
    End If

    Return vValid
  End Function

  Private Sub LookupValidatingHandler(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
    Try
      Dim vTextLookupBox As TextLookupBox = DirectCast(sender, TextLookupBox)
      If ValidateControl(vTextLookupBox, DirectCast(vTextLookupBox.Tag, PanelItem), vTextLookupBox.Text) Then
        If quickEntryGridUpdate() Then
          txtContactNumber.Text = ""
          txtMark.Text = ""
          txtGrade.Text = ""
          txtResult.Text = ""
          txtContactNumber.Focus()
        End If
      Else
        e.Cancel = False
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub


  Private Sub txtContactNumber_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtContactNumber.Validating
    erp.SetError(txtContactNumber, "")
    Dim vContactNumberRow As Integer = dgr.FindRow("ContactNumber", txtContactNumber.Text)
    If Not (String.IsNullOrEmpty(txtContactNumber.Text) Or vContactNumberRow > -1) Then
      erp.SetError(txtContactNumber, InformationMessages.ImInvalidContactNumber)
      e.Cancel = True
    ElseIf Not mvComponentBasedEntry AndAlso vContactNumberRow > -1 Then
      Dim vOrigCol As String = ""
      Dim vTotalCol As String = ""
      GetMarkColumns(mvExamMarkType, vOrigCol, vTotalCol)

      'If Multiple Unit result entry config is set then select the row if a valid contact number is entered 
      'in the contact number text box
      If mvMultipleUnits Then dgr.SetActiveCell(vContactNumberRow, "0")
      If vOrigCol.Length > 0 Then
        Dim vOldValue As String = dgr.GetValue(vContactNumberRow, vOrigCol)

        If txtMark.Visible Then
          txtMark.Text = vOldValue
        ElseIf txtGrade.Visible Then
          txtGrade.Text = vOldValue
        ElseIf txtResult.Visible Then
          txtResult.Text = vOldValue
        End If
      End If

    End If
  End Sub

  Private Sub txtMark_Validating(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtMark.Validating
    Dim mark As Double = 0

    If Not String.IsNullOrEmpty(txtMark.Text) Then
      e.Cancel = True

      If Double.TryParse(txtMark.Text, mark) AndAlso mark >= 0 Then
        If quickEntryGridUpdate() AndAlso Not mvComponentBasedEntry Then
          txtMark.Text = ""
          txtContactNumber.Text = ""
          txtContactNumber.Focus()
          e.Cancel = False
        End If
      End If
    End If

    If (Not mvComponentBasedEntry) And e.Cancel Then
      erp.SetError(txtMark, InformationMessages.ImInvalidValue)
    Else
      erp.SetError(txtMark, String.Empty)
    End If

    mvTabFromMark = False
  End Sub

  Private Function LoadComponentGrid() As Boolean
    dgrComponents.Clear()
    ShowHideComponentGrid(mvComponentBasedEntry)

    If mvComponentBasedEntry And dgr.RowCount > 0 Then
      Dim pList As New ParameterList(True, True)
      If epl.AddValuesToList(pList, True) Then
        pList.IntegerValue("ExamSessionId") = epl.FindTextLookupBox("ExamSessionCode").GetDataRowInteger("ExamSessionId")
        pList.IntegerValue("ExamCentreId") = epl.FindTextLookupBox("ExamCentreCode").GetDataRowInteger("ExamCentreId")
        pList.IntegerValue("ExamUnitId") = epl.FindTextLookupBox("ExamUnitDescription").GetDataRowInteger("ExamUnitId")
        Dim vDgrRow As Integer = dgr.CurrentRow
        If vDgrRow < 0 Then vDgrRow = 0
        pList.IntegerValue("ContactNumber") = CInt(dgr.GetValue(vDgrRow, "ContactNumber"))
        Dim vContactComponentDataSet As DataSet = ExamsDataHelper.GetExamData(ExamsAccess.XMLExamDataSelectionTypes.ExamStudentComponentResults, pList)
        dgrComponents.Populate(vContactComponentDataSet)
        dgrComponents.SetCellsEditable()
        If dgrComponents.RowCount > 0 Then
          dgrComponents.SetColumnVisible("RawMark", (mvExamCompMarkType = "M"))
          dgrComponents.SetColumnWritable("RawMark")

          dgrComponents.SetCellsReadOnly(-1, dgrComponents.GetColumn("ExamUnitCode"), True, True)
          dgrComponents.SetCellsReadOnly(-1, dgrComponents.GetColumn("ExamUnitDescription"), True, True)

          Dim vDataTable As DataTable = Nothing
          Dim vParams As New ParameterList(True)
          txtMark.Visible = True
          txtGrade.Visible = False
          txtResult.Visible = False
          lblMark.Text = "Marks"
          dgrComponents.SelectRow(0)
          txtMark.Focus()
          txtMark.SelectAll()
        End If
      End If
    End If
  End Function

  Private Sub dgr_RowChanging(ByVal sender As System.Object, ByRef pCancel As System.Boolean) Handles dgr.RowChanging
    ' If entering by component, check to see if values have changed before allowing row change
    If mvComponentBasedEntry Then
      If SaveComponentGridCheck() Then
        SaveComponentGrid()
      End If
    End If
  End Sub

  Private Function SaveComponentGrid() As Boolean
    SaveComponentGrid = False

    ' If entering by component, check to see if values have changed before allowin row change
    If mvComponentBasedEntry Then
      Dim vRawResultColumnName As String = "RawMark"
      Dim vNewExamResult As String = ""
      Dim vOldExamResult As String = ""
      Dim vList As ParameterList = New ParameterList(True)
      Dim vReturnList As ParameterList = New ParameterList()
      Dim vUpdateHasOccured As Boolean = False

      For vRow As Integer = 0 To dgrComponents.RowCount - 1
        vNewExamResult = dgrComponents.GetValue(vRow, vRawResultColumnName)
        vOldExamResult = dgrComponents.GetValue(vRow, vRawResultColumnName + COLUMN_CHECK)

        If (vNewExamResult <> vOldExamResult) Then
          vUpdateHasOccured = True
          vList(vRawResultColumnName) = vNewExamResult
          vList("ExamBookingUnitId") = dgrComponents.GetValue(vRow, "ExamBookingUnitId")
          vReturnList = ExamsDataHelper.UpdateItem(ExamsAccess.XMLExamMaintenanceTypes.ExamBookingUnit, vList)
          dgrComponents.SetValue(vRow, vRawResultColumnName & COLUMN_CHECK, vNewExamResult)
        End If
      Next

      If vUpdateHasOccured Then
        ' Update marks in parent component
        Dim vActiveRow As Integer = dgr.FindRow("ContactNumber", dgrComponents.GetValue(dgrComponents.ActiveRow, "ContactNumber"))
        vList("ExamBookingUnitId") = dgr.GetValue(vActiveRow, "ExamBookingUnitId")
        vList(vRawResultColumnName) = dgr.GetValue(vActiveRow, vRawResultColumnName) 'Raw Mark
        vReturnList = ExamsDataHelper.UpdateItem(ExamsAccess.XMLExamMaintenanceTypes.ExamBookingUnit, vList)
      End If
    End If
  End Function

  Private Sub GetMarkColumns(ByVal pMarkType As String, ByRef pOriginalColumn As String, ByRef pTotalColumn As String)
    pOriginalColumn = ""
    pTotalColumn = ""

    Select Case pMarkType
      Case "M"
        pOriginalColumn = "RawMark"
        pTotalColumn = "TotalMark"
      Case "G"
        pOriginalColumn = "OriginalGrade"
        pTotalColumn = "TotalGrade"
      Case "P"
        pOriginalColumn = "OriginalResult"
        pTotalColumn = "TotalResult"
    End Select
  End Sub


  Private Sub dgrComponents_RowSelected(ByVal sender As Object, ByVal pRow As Integer, ByVal pDataRow As Integer) Handles dgrComponents.RowSelected
    UpdateComponentQuickEntry(pRow)
  End Sub

  Private Sub UpdateQuickEntry(ByVal pRow As Integer, Optional ByVal pNewValue As String = Nothing)
    If (pRow > -1) Then
      txtContactNumber.Text = dgr.GetValue(pRow, "ContactNumber")

      If (Not mvComponentBasedEntry) Then
        Select Case mvExamMarkType
          Case "M"
            If Not pNewValue = Nothing Then
              txtMark.Text = pNewValue
            Else
              txtMark.Text = dgr.GetValue(pRow, "RawMark")
            End If
          Case "G"
            If Not pNewValue = Nothing Then
              txtGrade.Text = pNewValue
            Else
              txtGrade.Text = dgr.GetValue(pRow, "OriginalGrade")
            End If
          Case "P"
            If Not pNewValue = Nothing Then
              txtResult.Text = pNewValue
            Else
              txtResult.Text = dgr.GetValue(pRow, "OriginalResult")
            End If
        End Select
      End If
    Else
      txtContactNumber.Text = ""
      txtMark.Text = ""
      txtGrade.Text = ""
      txtResult.Text = ""
    End If
  End Sub

  Private Sub UpdateComponentQuickEntry(ByVal pRow As Integer)
    If pRow > -1 Then txtMark.Text = dgrComponents.GetValue(pRow, "RawMark")
    dgrComponents.SetActiveCell(pRow, dgrComponents.ColumnName(2))
  End Sub

  Private Function SaveComponentGridCheck() As Boolean
    SaveComponentGridCheck = False

    If mvComponentBasedEntry Then
      Dim vResultColumnName As String = "RawMark"
      Dim vNewExamResult As String = ""
      Dim vOldExamResult As String = ""

      For vRow As Integer = 0 To dgrComponents.RowCount - 1
        vNewExamResult = dgrComponents.GetValue(vRow, vResultColumnName)
        vOldExamResult = dgrComponents.GetValue(vRow, vResultColumnName & COLUMN_CHECK)

        If (vNewExamResult <> vOldExamResult) Then
          ' TODO - Hardcoded text needs putting in resource file
          If ShowQuestion("Answers have been updated, save?", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            SaveComponentGridCheck = True
          End If
          Exit For
        End If
      Next
    End If
  End Function

  Private Function GetRawMarkSum() As Double
    GetRawMarkSum = 0

    If mvComponentBasedEntry Then
      Dim vColNo = dgrComponents.GetColumn("RawMark")
      For i As Integer = 0 To dgrComponents.RowCount - 1
        Dim vRawMark As Double = CDbl(dgrComponents.GetRawValue(i, vColNo))
        GetRawMarkSum += vRawMark
      Next
    End If
  End Function

  Private Sub UpdateComponentMarkSum()
    Dim vSumOfMarks As Double = GetRawMarkSum()
    If vSumOfMarks < 0 Then vSumOfMarks = 0
    Dim vColNo As Integer = dgr.GetColumn("RawMark")
    dgr.SetValue(dgr.ActiveRow, vColNo, vSumOfMarks)
  End Sub

  Private Sub dgrComponents_ValueChanged(ByVal sender As System.Object, ByVal pRow As System.Int32, ByVal pCol As System.Int32, ByVal pValue As System.String, ByVal pOldValue As System.String) Handles dgrComponents.ValueChanged
    If pCol = dgrComponents.GetColumn("RawMark") And pOldValue <> pValue Then
      UpdateComponentQuickEntry(dgrComponents.ActiveRow)
      UpdateComponentMarkSum()
    End If
  End Sub

  Private Sub cmdOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Dim vValid As Boolean = True
    If mvComponentBasedEntry Then
      If SaveComponentGridCheck() Then SaveComponentGrid()
    Else
      vValid = SaveGrid()
    End If

    If vValid Then Me.Close()
  End Sub

  Private Sub dgrComponents_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgrComponents.KeyDown
    dgrComponents.SetActiveCell(dgrComponents.ActiveRow, dgrComponents.ColumnName(dgrComponents.ActiveColumn))
    If e.KeyCode = 46 Then
      dgrComponents.SetValue(dgrComponents.ActiveRow, dgrComponents.ActiveColumn, "")
      UpdateComponentQuickEntry(dgrComponents.ActiveRow)
      UpdateComponentMarkSum()
      e.SuppressKeyPress = True
    End If
  End Sub

  Private Sub dgrComponents_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles dgrComponents.KeyUp
    UpdateComponentQuickEntry(dgrComponents.ActiveRow)
  End Sub

  Private Sub dgr_ValueChanged(ByVal sender As Object, ByVal pRow As Integer, ByVal pCol As Integer, ByVal pValue As String, ByVal pOldValue As String) Handles dgr.ValueChanged
    UpdateQuickEntry(dgr.ActiveRow, pValue)
  End Sub

  Private Sub dgr_MouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgr.MouseUp
    UpdateQuickEntry(dgr.ActiveRow)
  End Sub

  Private Sub ShowHideComponentGrid(ByVal pShow As Boolean)
    If pShow Then
      dgrComponents.Visible = True
      dgrComponents.Height = 100
      splComponentSeparator.Visible = True
    Else
      dgrComponents.Visible = False
      splComponentSeparator.Visible = False
      'splBottom.Panel2Collapsed = True
    End If
    splBottom.Panel2Collapsed = False
  End Sub

  Private Sub cmdCancel_Click(sender As Object, e As System.EventArgs) Handles cmdCancel.Click
    If mvComponentBasedEntry AndAlso SaveComponentGridCheck() Then SaveComponentGrid()
    Me.Close()
  End Sub

  Private Sub cmdFindNext_Click(sender As Object, e As EventArgs) Handles cmdFindNext.Click
    If rdoResult.Checked Then
      FindText(txtResultFind.Text)
    ElseIf rdoGrade.Checked Then
      FindText(txtGradeFind.Text)
    ElseIf rdoMark.Checked Then
      Dim vMark As Double
      If String.IsNullOrWhiteSpace(txtMarkFind.Text) OrElse Double.TryParse(txtMarkFind.Text, vMark) Then
        FindText(txtMarkFind.Text)
      End If
      'FindText(txtMarkFind.Text)
    Else
      'Do nothing
    End If
  End Sub


  Private Sub RefreshSearchSection(ByVal pExamMarkType As String)
    Dim vDataTable As DataTable = Nothing
    Dim vParams As New ParameterList(True)
    Select Case pExamMarkType
      Case GRADE
        DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamGrades, vParams)
        'PopulateGridCombo(dgr, "OriginalGrade", "ExamGrade", "ExamGrade", vDataTable)
        'txtGradeFind.Visible = True
        'txtGradeReplace.Visible = True
        'txtMarkFind.Visible = False
        'txtMarkReplace.Visible = False
        'txtResultFind.Visible = False
        'txtResultReplace.Visible = False
        TabFindReplaceHost.SelectedTab = TabFindReplaceGrade
        InitTextLookupBox(txtGradeFind, "exam_grades", "exam_grade")
        InitTextLookupBox(txtGradeReplace, "exam_grades", "exam_grade")
        'lblMark.Text = "Grade"
      Case MARK
        'txtGradeFind.Visible = False
        'txtGradeReplace.Visible = False
        'txtMarkFind.Visible = True
        'txtMarkReplace.Visible = True
        'txtResultFind.Visible = False
        'txtResultReplace.Visible = False
        'lblMark.Text = "Marks"
        TabFindReplaceHost.SelectedTab = TabFindReplaceMark
      Case RESULT
        DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamResultTypes, vParams)
        'PopulateGridCombo(dgr, "OriginalResult", "LookUpCode", "LookUpDesc", vDataTable)
        'txtGradeFind.Visible = False
        'txtGradeReplace.Visible = False
        'txtMarkFind.Visible = False
        'txtMarkReplace.Visible = False
        'txtResultFind.Visible = True
        'txtResultReplace.Visible = True
        TabFindReplaceHost.SelectedTab = TabFindReplaceResult
        InitTextLookupBox(txtResultFind, "exam_booking_units", "original_result")
        InitTextLookupBox(txtResultReplace, "exam_booking_units", "original_result")
        'lblMark.Text = "P/F"
    End Select
  End Sub

  Private Sub ReplaceText(ByVal pSearchText As String, ByVal pReplaceText As String)
    dgr.ReplaceText(pSearchText, pReplaceText)
  End Sub

  Private Sub FindText(ByVal pSearchText As String)
    dgr.FindText(pSearchText)
  End Sub

  Private Sub cmdReplace_Click(sender As Object, e As EventArgs) Handles cmdReplace.Click
    If rdoResult.Checked Then
      ReplaceText(txtResultFind.Text, txtResultReplace.Text)
    ElseIf rdoGrade.Checked Then
      ReplaceText(txtGradeFind.Text, txtGradeReplace.Text)
    ElseIf rdoMark.Checked Then
      ReplaceText(txtMarkFind.Text, txtMarkReplace.Text)
    Else
      'do nothing
    End If
  End Sub

  Private Sub cmdReplaceAll_Click(sender As Object, e As EventArgs) Handles cmdReplaceAll.Click
    If rdoResult.Checked Then
      dgr.FindAndReplaceAll(txtResultFind.Text, txtResultReplace.Text)
    ElseIf rdoGrade.Checked Then
      dgr.FindAndReplaceAll(txtGradeFind.Text, txtGradeReplace.Text)
    ElseIf rdoMark.Checked Then
      dgr.FindAndReplaceAll(txtMarkFind.Text, txtMarkReplace.Text)
    Else

    End If
  End Sub

  Private Sub rdoMark_CheckedChanged(sender As Object, e As EventArgs) Handles rdoMark.CheckedChanged
    RefreshSearchSection(MARK)
  End Sub

  Private Sub rdoGrade_CheckedChanged(sender As Object, e As EventArgs) Handles rdoGrade.CheckedChanged
    RefreshSearchSection(GRADE)
  End Sub

  Private Sub rdoResult_CheckedChanged(sender As Object, e As EventArgs) Handles rdoResult.CheckedChanged
    RefreshSearchSection(RESULT)
  End Sub

  Private Sub SetupForMultipleResultEntry(ByVal pDataSet As DataSet)
    Dim vMultiResultsDS As New DataSet
    Dim vMultiResultsDT As New DataTable("DataRow")

    If pDataSet.Tables.Contains("DataRow") Then
      Dim vDRView As New DataView(pDataSet.Tables("DataRow"))
      Dim vExamUnits As DataTable = vDRView.ToTable(True, "ExamUnitCode")

      vMultiResultsDT.Columns.Add("ContactNumber")
      vMultiResultsDT.Columns.Add("CandidateName")
      vMultiResultsDT.Columns.Add("ExamMarkType")
      vMultiResultsDT.Columns.Add("ExamBookingUnitId")

      'Create all the required columns for new consolidated table to show multiple 
      'units on the same row for a contact (if contact has multiple bookings)
      For vRow As Integer = 0 To vExamUnits.Rows.Count - 1
        vMultiResultsDT.Columns.Add(vExamUnits.Rows(vRow).Item(0).ToString)
        vMultiResultsDT.Columns.Add(vExamUnits.Rows(vRow).Item(0).ToString + COLUMN_CHECK)
        vMultiResultsDT.Columns.Add(vExamUnits.Rows(vRow).Item(0).ToString + COLUMN_BOOKING_ID)
        vMultiResultsDT.Columns.Add(vExamUnits.Rows(vRow).Item(0).ToString + COLUMN_GRADE_TYPE)
      Next

      For Each vRow As DataRow In pDataSet.Tables("DataRow").Rows
        If Not mvUnitMarkType.ContainsKey(vRow.Item("ExamUnitCode").ToString) Then mvUnitMarkType.Add(vRow.Item("ExamUnitCode").ToString, vRow.Item("ExamMarkType").ToString)
      Next

      'Record exists for the conatct, so update it with new exam unit details
      If pDataSet IsNot Nothing AndAlso pDataSet.Tables.Contains("DataRow") Then
        For Each vDR As DataRow In pDataSet.Tables("DataRow").Rows
          Dim vExisting As Boolean = False
          If vMultiResultsDT.Rows.Count > 0 Then
            For Each vRow As DataRow In vMultiResultsDT.Rows
              If CInt(vRow.Item("ContactNumber")) = CInt(vDR.Item("ContactNumber")) Then
                Select Case vDR.Item("ExamMarkType").ToString.Trim
                  Case GRADE
                    vRow.Item(vDR.Item("ExamUnitCode").ToString) = vDR.Item("OriginalGrade")
                    vRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_CHECK) = vDR.Item("OriginalGradeCheck")
                  Case MARK
                    vRow.Item(vDR.Item("ExamUnitCode").ToString) = vDR.Item("RawMark")
                    vRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_CHECK) = vDR.Item("RawMarkCheck")
                  Case RESULT
                    vRow.Item(vDR.Item("ExamUnitCode").ToString) = vDR.Item("OriginalResult")
                    vRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_CHECK) = vDR.Item("OriginalResultCheck")
                End Select
                vRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_BOOKING_ID) = vDR.Item("ExamBookingUnitId")
                vRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_GRADE_TYPE) = vDR.Item("ExamMarkType")
                vExisting = True
                Exit For
              End If
            Next
          End If

          'No record exists for the contact so, create a new row 
          If Not vExisting Then
            Dim vDataRow As DataRow = vMultiResultsDT.NewRow
            vDataRow.Item("ContactNumber") = vDR.Item("ContactNumber")
            vDataRow.Item("CandidateName") = vDR.Item("CandidateName")

            'Add exam mark type to a new column as there is a posiblity that the contact has bookings for 
            'different types of exam units (Mark, grade or Pass/fail)
            vDataRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_GRADE_TYPE) = vDR.Item("ExamMarkType")

            'Add exam booking id to a new column as there is a posiblity that the contact has more than one booking(s) for 
            'and well will need the booking number to update the record for the new result value
            vDataRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_BOOKING_ID) = vDR.Item("ExamBookingUnitId")
            Select Case vDR.Item("ExamMarkType").ToString.Trim
              Case GRADE
                vDataRow.Item(vDR.Item("ExamUnitCode").ToString) = vDR.Item("OriginalGrade")
                vDataRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_CHECK) = vDR.Item("OriginalGradeCheck")
              Case MARK
                vDataRow.Item(vDR.Item("ExamUnitCode").ToString) = vDR.Item("RawMark")
                vDataRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_CHECK) = vDR.Item("RawMarkCheck")
              Case RESULT
                vDataRow.Item(vDR.Item("ExamUnitCode").ToString) = vDR.Item("OriginalResult")
                vDataRow.Item(vDR.Item("ExamUnitCode").ToString + COLUMN_CHECK) = vDR.Item("OriginalResultCheck")
            End Select
            vMultiResultsDT.Rows.Add(vDataRow)
          End If
        Next
      End If

      Dim vMarkColumns As List(Of String) = mvUnitMarkType.Where(Function(vItem) vItem.Value = MARK).Select(Function(vItem) vItem.Key).ToList()
      If vMultiResultsDT IsNot Nothing AndAlso vMultiResultsDT.Rows.Count > 0 Then
        Dim vColumnDT As New DataTable("Column")
        vColumnDT.Columns.Add("Name")
        vColumnDT.Columns.Add("Value")
        vColumnDT.Columns.Add("Visible")
        vColumnDT.Columns.Add("DataType")
        vColumnDT.Columns.Add("Heading")
        For Each vCol As DataColumn In vMultiResultsDT.Columns
          Dim vDataRow As DataRow = vColumnDT.NewRow
          vDataRow.Item("Name") = vCol.ColumnName
          vDataRow.Item("Heading") = vCol.ColumnName
          vDataRow.Item("Visible") = "Y"
          vDataRow.Item("DataType") = "Char"
          If vMarkColumns.Contains(vCol.ColumnName) Then
            vDataRow.Item("DataType") = "Numeric"
          End If
          vColumnDT.Rows.Add(vDataRow)
        Next
        If vColumnDT IsNot Nothing Then vMultiResultsDS.Tables.Add(vColumnDT)
      End If


      vMultiResultsDS.Tables.Add(vMultiResultsDT)
      dgr.Populate(vMultiResultsDS)
      dgr.SetColumnReadOnly(dgr.GetColumn("ContactNumber"), True, True)
      dgr.SetColumnReadOnly(dgr.GetColumn("CandidateName"), True, True)
      dgr.SetColumnVisible(dgr.GetColumn("ExamMarkType"), False)
      dgr.SetColumnVisible(dgr.GetColumn("ExamBookingUnitId"), False)
      dgr.SetCellsEditable()
      If vMarkColumns IsNot Nothing AndAlso vMarkColumns.Count > 0 Then
        Dim vNumberCellType As New FarPoint.Win.Spread.CellType.NumberCellType
        vNumberCellType.DecimalPlaces = AppValues.GetDecimalPlaces("RawMark")
        vNumberCellType.Separator = NumberGroupSeparator
        vNumberCellType.ShowSeparator = AppValues.ShowNumberGroupSeparator
        vMarkColumns.ForEach(Sub(vColName) dgr.SetColumnCellType(vColName, vNumberCellType))
      End If
      HideGridColumns()
      SetGridColumns()
      DisableInvalidCells()
      SetUpTabControl()
    End If
  End Sub

  ''' <summary>
  ''' Set Grid Columns
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub SetGridColumns()

    For Each vColumnName As String In mvUnitMarkType.Keys
      Select Case mvUnitMarkType(vColumnName)
        Case GRADE
          PopulateGridCombo(dgr, vColumnName, "ExamGrade", "ExamGrade", DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamGrades, New ParameterList(True)))
        Case RESULT
          PopulateGridCombo(dgr, vColumnName, "LookUpCode", "LookUpDesc", DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtExamResultTypes, New ParameterList(True)))
        Case Else
          'do nothing
      End Select
    Next
  End Sub

  Private Sub SetUpTabControl()
    'Me.txtContactNumber.Enabled = False
    Me.txtGrade.Enabled = False
    Me.txtMark.Enabled = False
    Me.txtResult.Enabled = False
    Me.lblMark.Enabled = False
    Me.lblInfo.Visible = True
  End Sub

  Private Sub HideGridColumns()
    For vCol As Integer = 0 To dgr.ColumnCount - 1
      If dgr.ColumnName(vCol).Contains(COLUMN_CHECK) OrElse dgr.ColumnName(vCol).Contains(COLUMN_BOOKING_ID) OrElse dgr.ColumnName(vCol).Contains(COLUMN_GRADE_TYPE) Then
        dgr.SetColumnVisible(vCol, False)
      End If
    Next
  End Sub

  ''' <summary>
  ''' This method will go through each column and disable all the cells where contacts do not have any bookings
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub DisableInvalidCells()
    For vRow As Integer = 0 To dgr.RowCount - 1
      For vColumn As Integer = 0 To dgr.ColumnCount - 1
        If dgr.ColumnName(vColumn).Contains(COLUMN_BOOKING_ID) AndAlso dgr.GetValue(vRow, vColumn).Length = 0 Then
          dgr.SetCellsReadOnly(vRow, dgr.GetColumn(dgr.ColumnName(vColumn).Replace(COLUMN_BOOKING_ID, "")), True, True)
        End If
      Next
    Next
  End Sub

  Private Function GetChangeReason(ByRef pChangeReason As String) As Boolean
    If AppValues.ControlValue(AppValues.ControlTables.exam_controls, AppValues.ControlValues.record_grade_change_history).ToUpper = "M" Then
      If txtChangeReason.Text.Length = 0 Then
        ShowErrorMessage(InformationMessages.ImChangeReasonNotSelected, "Change Reason")
        Return False
      Else
        pChangeReason = txtChangeReason.Text
        Return True
      End If
    Else
      Return True
    End If

  End Function

  Private Sub txtMarkFind_KeyPress(sender As Object, e As KeyPressEventArgs)
    Dim allowedChars As String = "."
    If Char.IsDigit(e.KeyChar) = False And Char.IsControl(e.KeyChar) = False And allowedChars.IndexOf(e.KeyChar) = -1 Then
      e.Handled = True
    End If
  End Sub

End Class


