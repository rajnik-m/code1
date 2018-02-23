Imports FarPoint.Win.Spread

Public Class ActivityDataSheet
  Inherits System.Windows.Forms.UserControl

  Private Enum ActivityEntryColumns
    aecActivity
    aecActivityValue
    aecOldValidFrom
    aecOldValidTo
    aecOldSource
    aecMandatory
    aecQtyRequired
    aecGroup
    aecMultipleValues
    aecActivityDesc
    aecActivityValueDesc
    aecActivityQuantity
    aecActivityDate
    aecActivityNotes
    aecActivityValidFrom
    aecActivityValidTo
    aecActivitySource
  End Enum

  Private mvContactInfo As ContactInfo
  Private mvSource As String
  Private mvDefaultValidFrom As Date
  Private mvDefaultValidTo As Date
  Private mvSheet As SheetView
  Private mvInitialised As Boolean

  Public Sub Init(ByVal pContactInfo As ContactInfo, ByVal pGroupCode As String, ByVal pTable As DataTable, ByVal pSource As String)
    mvContactInfo = pContactInfo
    mvSource = pSource
    mvSheet = vas.Sheets(0)
    mvSheet.Columns.Count = ActivityEntryColumns.aecActivitySource + 1
    mvSheet.OperationMode = OperationMode.Normal
    mvSheet.RowHeaderVisible = False
    mvSheet.GrayAreaBackColor = DisplayTheme.GridBackAreaColor

    Dim vInputMap As InputMap
    vInputMap = vas.GetInputMap(InputMapMode.WhenAncestorOfFocused)
    vInputMap.Put(New Keystroke(Keys.Tab, Keys.None), SpreadActions.None)
    vInputMap = vas.GetInputMap(InputMapMode.WhenAncestorOfFocused)
    vInputMap.Put(New Keystroke(Keys.Tab, Keys.Shift), SpreadActions.None)

    Dim vDateCellType As New CellType.DateTimeCellType
    vDateCellType.DateTimeFormat = CellType.DateTimeFormat.ShortDate
    vDateCellType.DropDownButton = True
    vDateCellType.MinimumDate = New Date(1900, 1, 1)
    vDateCellType.MaximumDate = New Date(2500, 12, 31)

    Dim vNumberCellType As New CellType.NumberCellType
    vNumberCellType.DecimalPlaces = 2
    vNumberCellType.MaximumValue = 99999999

    Dim vCheckBoxType As New CellType.CheckBoxCellType
    Dim vComboBoxType As CellType.ComboBoxCellType
    Dim vSourceTextType As New CellType.TextCellType
    vSourceTextType.CharacterCasing = CharacterCasing.Upper
    vSourceTextType.MaxLength = 10
    Dim vNotesTextType As New CellType.TextCellType
    vNotesTextType.Multiline = True
    vNotesTextType.MaxLength = 1024
    Dim vStaticTextType As New CellType.TextCellType
    vStaticTextType.ReadOnly = True

    mvDefaultValidFrom = Date.Now
    mvDefaultValidTo = Date.Now.AddYears(100)

    Dim vTable As DataTable
    If pTable Is Nothing Then
      Dim vList As New ParameterList(True)
      vList("UsageCode") = "B"
      vList("ActivityGroup") = pGroupCode
      vList(pContactInfo.ContactGroupParameterName) = pContactInfo.ContactGroup
      vTable = DataHelper.GetLookupData(CareServices.XMLLookupDataTypes.xldtActivityDataSheet, vList)
    Else
      vTable = pTable
    End If
    'ActivityCode,ActivityDesc,ActivityValueCode,ActivityValueDesc,Mandatory,QuantityRequired,MultipleValues,ContactGroup,SourceCode
    Dim vRow As Integer = 0
    Dim vActivity As String
    Dim vValue As String
    Dim vValueDesc As String
    Dim vLastActivity As String = ""
    Dim vValueCount As Integer
    Dim vLastValue As String = ""
    Dim vLastValueDesc As String = ""

    With mvSheet
      .ColumnHeaderVisible = True
      .Columns(ActivityEntryColumns.aecActivity).Visible = False
      .Columns(ActivityEntryColumns.aecActivityValue).Visible = False
      .Columns(ActivityEntryColumns.aecOldValidFrom).Visible = False
      .Columns(ActivityEntryColumns.aecOldValidTo).Visible = False
      .Columns(ActivityEntryColumns.aecOldSource).Visible = False
      .Columns(ActivityEntryColumns.aecMandatory).Visible = False
      .Columns(ActivityEntryColumns.aecQtyRequired).Visible = False
      .Columns(ActivityEntryColumns.aecGroup).Visible = False
      .Columns(ActivityEntryColumns.aecMultipleValues).Visible = False

      .Columns(ActivityEntryColumns.aecActivityDesc).CellType = vStaticTextType
      .Columns(ActivityEntryColumns.aecActivitySource).CellType = vSourceTextType
      .Columns(ActivityEntryColumns.aecActivityNotes).CellType = vNotesTextType
      .Columns(ActivityEntryColumns.aecActivityQuantity).CellType = vNumberCellType
      .Columns(ActivityEntryColumns.aecActivityDate).CellType = vDateCellType
      .Columns(ActivityEntryColumns.aecActivityValidFrom).CellType = vDateCellType
      .Columns(ActivityEntryColumns.aecActivityValidTo).CellType = vDateCellType

      .ColumnHeader.Columns(ActivityEntryColumns.aecActivityDesc).Label = "Category"
      .ColumnHeader.Columns(ActivityEntryColumns.aecActivityValueDesc).Label = "Value"
      .ColumnHeader.Columns(ActivityEntryColumns.aecActivityQuantity).Label = "Quantity"
      .ColumnHeader.Columns(ActivityEntryColumns.aecActivityDate).Label = "Date"
      .ColumnHeader.Columns(ActivityEntryColumns.aecActivityNotes).Label = "Notes"
      .ColumnHeader.Columns(ActivityEntryColumns.aecActivityValidFrom).Label = "ValidFrom"
      .ColumnHeader.Columns(ActivityEntryColumns.aecActivityValidTo).Label = "ValidTo"
      .ColumnHeader.Columns(ActivityEntryColumns.aecActivitySource).Label = "Source"

      If vTable IsNot Nothing Then
        For Each vDataRow As DataRow In vTable.Rows
          If mvSource.Length = 0 Then mvSource = vDataRow("SourceCode").ToString
          vActivity = vDataRow("ActivityCode").ToString
          vValue = vDataRow("ActivityValueCode").ToString
          vValueDesc = vDataRow("ActivityValueDesc").ToString
          If vActivity = vLastActivity Then            'Same as last so it goes in the combo
            vValueCount = vValueCount + 1
            vRow -= 1
            If vValueCount = 1 Then
              vComboBoxType = New CellType.ComboBoxCellType
              .Cells(vRow, ActivityEntryColumns.aecActivityValueDesc).CellType = vComboBoxType
              Dim vItems() As String = {"", vLastValueDesc, vValueDesc}
              Dim vItemData() As String = {"", vLastValue, vValue}
              vComboBoxType.Items = vItems
              vComboBoxType.ItemData = vItemData
              vComboBoxType.EditorValue = CellType.EditorValue.ItemData
              .SetValue(vRow, ActivityEntryColumns.aecActivityValueDesc, "")
              .SetValue(vRow, ActivityEntryColumns.aecActivityValue, "")
            Else
              vComboBoxType = CType(.Cells(vRow, ActivityEntryColumns.aecActivityValueDesc).CellType, CellType.ComboBoxCellType)
              Dim vItems() As String = vComboBoxType.Items
              Dim vItemsData() As String = vComboBoxType.ItemData
              Dim vItemCount As Integer = vItems.GetLength(0)
              ReDim Preserve vItems(vItemCount)
              ReDim Preserve vItemsData(vItemCount)
              vItems(vItemCount) = vValueDesc
              vItemsData(vItemCount) = vValue
              vComboBoxType.Items = vItems
              vComboBoxType.ItemData = vItemsData
            End If
            vRow += 1
          Else
            .RowCount = vRow + 1
            .SetValue(vRow, ActivityEntryColumns.aecActivity, vDataRow("ActivityCode").ToString)
            .SetValue(vRow, ActivityEntryColumns.aecActivityValue, vDataRow("ActivityValueCode").ToString)
            .SetValue(vRow, ActivityEntryColumns.aecMandatory, vDataRow("Mandatory").ToString)
            .SetValue(vRow, ActivityEntryColumns.aecQtyRequired, vDataRow("QuantityRequired").ToString)
            .SetValue(vRow, ActivityEntryColumns.aecGroup, vDataRow("ContactGroup").ToString)
            .SetValue(vRow, ActivityEntryColumns.aecMultipleValues, vDataRow("MultipleValues").ToString)
            .SetValue(vRow, ActivityEntryColumns.aecActivityDesc, vDataRow("ActivityDesc").ToString)
            .SetValue(vRow, ActivityEntryColumns.aecActivityValueDesc, vDataRow("ActivityValueDesc").ToString)

            .SetValue(vRow, ActivityEntryColumns.aecActivityValidFrom, mvDefaultValidFrom)
            .SetValue(vRow, ActivityEntryColumns.aecActivityValidTo, mvDefaultValidTo)
            .SetValue(vRow, ActivityEntryColumns.aecActivitySource, mvSource)

            If vValueCount = 0 And vRow > 0 Then
              .Cells(vRow - 1, ActivityEntryColumns.aecActivityValueDesc).CellType = vCheckBoxType
              .SetValue(vRow - 1, ActivityEntryColumns.aecActivityValueDesc, "")
              .SetValue(vRow - 1, ActivityEntryColumns.aecMultipleValues, "N")
            End If
            vValueCount = 0
            vRow = vRow + 1
          End If
          vLastValue = vValue
          vLastValueDesc = vValueDesc
          vLastActivity = vActivity
          'vRow = vRow + 1
        Next
        If vValueCount = 0 And vRow >= 1 Then
          .Cells(vRow - 1, ActivityEntryColumns.aecActivityValueDesc).CellType = vCheckBoxType
          .SetValue(vRow - 1, ActivityEntryColumns.aecActivityValueDesc, "")
          .SetValue(vRow - 1, ActivityEntryColumns.aecMultipleValues, "N")
        End If
      End If
    End With
    SetPreferredWidths()
    PopulateActivities()
    mvInitialised = True
    SetHeight()
  End Sub

  Public Sub SetHeight()
    If mvInitialised Then
      Dim vRequiredHeight As Integer = CInt((mvSheet.Rows(0).Height * mvSheet.Rows.Count) + mvSheet.ColumnHeader.Rows(0).Height + SystemInformation.HorizontalScrollBarHeight + (SystemInformation.Border3DSize.Height * 2) + Me.DockPadding.Top + Me.DockPadding.Bottom)
      Me.Height = Math.Min(vRequiredHeight, Me.Parent.DisplayRectangle.Height)
    End If
  End Sub

  Public ReadOnly Property Source() As String
    Get
      Return mvSource
    End Get
  End Property

  Private Sub SetPreferredWidths()
    Dim vDefaultWidth As Single
    Dim vPreferredWidth As Single
    For Each vCol As Column In mvSheet.Columns
      Select Case vCol.Index
        Case ActivityEntryColumns.aecActivityDesc, ActivityEntryColumns.aecActivityValueDesc
          vDefaultWidth = 200
        Case ActivityEntryColumns.aecActivityQuantity
          vDefaultWidth = 60
        Case ActivityEntryColumns.aecActivityNotes
          vDefaultWidth = 300
        Case ActivityEntryColumns.aecActivityValidFrom, ActivityEntryColumns.aecActivityValidTo, ActivityEntryColumns.aecActivityDate
          vDefaultWidth = 100
      End Select
      vPreferredWidth = vCol.GetPreferredWidth
      vCol.Width = Math.Max(vPreferredWidth, vDefaultWidth)
    Next
  End Sub

  Private Sub PopulateActivities()
    Dim vList As New ParameterList(True)
    Dim vActivities As New ArrayListEx

    With mvSheet
      For vRow As Integer = 0 To .Rows.Count - 1
        vActivities.Add(.GetText(vRow, ActivityEntryColumns.aecActivity))
      Next
      vList("Activities") = vActivities.CSStringList
      vList("SystemColumns") = "N"
      'vList("IgnoreUnknownParameters") = "Y"        'TODO Add Activities as a valid XML parameter name
      Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactCategories, mvContactInfo.ContactNumber, vList)
      Dim vCell As Cell
      Dim vActivity As String
      Dim vSetValues As Boolean
      If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Rows.Count > 0 Then
        For Each vDataRow As DataRow In vDataSet.Tables("DataRow").Rows
          vActivity = vDataRow("ActivityCode").ToString
          'Find the first location for the activity in the database
          For vRow As Integer = 0 To .Rows.Count - 1
            vSetValues = False
            If .GetText(vRow, ActivityEntryColumns.aecActivity) = vActivity Then
              vCell = .Cells(vRow, ActivityEntryColumns.aecActivityValueDesc)
              If TypeOf vCell.CellType Is CellType.CheckBoxCellType Then
                vCell.Text = "True"
                vSetValues = True
                vCell.Locked = True
              Else                      'Must be a combo box
                If vCell.Text.Length = 0 Then
                  If .GetText(vRow, ActivityEntryColumns.aecMultipleValues) = "Y" Then
                    .AddRows(vRow + 1, 1)
                    .CopyRange(vRow, 0, vRow + 1, 0, 1, .Columns.Count, False)
                    .SetValue(vRow + 1, ActivityEntryColumns.aecActivityValue, "")
                  End If
                  .SetValue(vRow, ActivityEntryColumns.aecActivityValue, vDataRow("ActivityValueCode"))
                  .SetValue(vRow, ActivityEntryColumns.aecActivityValueDesc, vDataRow("ActivityValueCode"))
                  'vCell.Locked = True
                  vSetValues = True
                End If
              End If
              If vSetValues Then
                .SetValue(vRow, ActivityEntryColumns.aecActivitySource, vDataRow("SourceCode"))
                .SetValue(vRow, ActivityEntryColumns.aecOldSource, vDataRow("SourceCode"))
                .SetValue(vRow, ActivityEntryColumns.aecActivityValidFrom, LimitDateValue(vDataRow("ValidFrom").ToString))
                .SetValue(vRow, ActivityEntryColumns.aecOldValidFrom, vDataRow("ValidFrom"))
                .SetValue(vRow, ActivityEntryColumns.aecActivityValidTo, LimitDateValue(vDataRow("ValidTo").ToString))
                .SetValue(vRow, ActivityEntryColumns.aecOldValidTo, vDataRow("ValidTo"))
                If vDataRow("Quantity").ToString.Length > 0 Then
                  .SetValue(vRow, ActivityEntryColumns.aecActivityQuantity, vDataRow("Quantity"))
                End If
                .SetValue(vRow, ActivityEntryColumns.aecActivityDate, vDataRow("ActivityDate"))
                .SetValue(vRow, ActivityEntryColumns.aecActivityNotes, vDataRow("Notes"))
                Exit For
              End If
            End If
          Next
        Next
      End If
    End With
  End Sub

  Private Function LimitDateValue(ByVal pValue As String) As String
    If pValue.Length > 0 Then
      Dim vDate As Date = CDate(pValue)
      If vDate.Year > 2500 Then
        vDate = New Date(2500, 12, 31)
      ElseIf vDate.Year < 1900 Then
        vDate = New Date(1900, 1, 1)
      End If
      Return vDate.ToShortDateString
    Else
      Return ""
    End If
  End Function

  Public Function ValidateActivities() As Boolean
    Dim vRow As Integer
    Dim vValid As Boolean
    Dim vFromDate As String
    Dim vToDate As String
    Dim vInsert As Boolean
    Dim vActivity As String
    Dim vSameAct As Boolean
    Dim vValidSources As New CollectionList(Of String)
    Dim vCell As Cell
    Dim vSource As String

    Dim vList As New ParameterList(True)
    With mvSheet
      vValid = True
      For vRow = 0 To .Rows.Count - 1
        vInsert = False
        vCell = .Cells(vRow, ActivityEntryColumns.aecActivityValueDesc)
        If TypeOf vCell.CellType Is CellType.CheckBoxCellType Then
          If vCell.Text = "True" Then vInsert = True
        Else                     'Must be a combo box
          If vCell.Text.Length > 0 Then
            vInsert = True
          End If
        End If
        If .Cells(vRow, ActivityEntryColumns.aecMandatory).Text = "Y" And vInsert = False Then
          vSameAct = False
          If vRow > 0 Then
            vActivity = .Cells(vRow, ActivityEntryColumns.aecActivity).Text
            If vActivity = .Cells(vRow - 1, ActivityEntryColumns.aecActivity).Text Then vSameAct = True 'Same activity
          End If
          If Not vSameAct Then
            .SetActiveCell(vRow, ActivityEntryColumns.aecActivityValueDesc)
            ShowWarningMessage(InformationMessages.imRequiredEntry, .Cells(vRow, ActivityEntryColumns.aecActivityDesc).Text)
            Return False
          End If
        End If
        If vValid Then
          vFromDate = .Cells(vRow, ActivityEntryColumns.aecActivityValidFrom).Text
          vToDate = .Cells(vRow, ActivityEntryColumns.aecActivityValidTo).Text
          If IsDate(vFromDate) And IsDate(vToDate) Then
            If CDate(vFromDate) > CDate(vToDate) Then
              .SetActiveCell(vRow, ActivityEntryColumns.aecActivityValidFrom)
              ShowWarningMessage(InformationMessages.imValidFromGTValidTo)
              Return False
            End If
          End If
        End If
        If vValid Then
          If .Cells(vRow, ActivityEntryColumns.aecQtyRequired).Text = "Y" And vInsert = True Then
            Dim vValue As String = .Cells(vRow, ActivityEntryColumns.aecActivityQuantity).Text
            If vValue.Length = 0 OrElse CSng(vValue) <= 0 Then
              .SetActiveCell(vRow, ActivityEntryColumns.aecActivityQuantity)
              ShowWarningMessage(InformationMessages.imQuantityRequired)
              Return False
            End If
          End If
        End If
        If vValid And vInsert Then
          vSource = .Cells(vRow, ActivityEntryColumns.aecActivitySource).Text
          If Not vValidSources.ContainsKey(vSource) Then
            vList("Source") = vSource
            Dim vResult As ParameterList = DataHelper.GetLookupItem(CareServices.XMLLookupDataTypes.xldtSources, vList)
            If vResult.ContainsKey("Source") Then
              vValidSources.Add(vSource, vSource)
            Else
              .SetActiveCell(vRow, ActivityEntryColumns.aecActivitySource)
              ShowWarningMessage(InformationMessages.imInvalidSource)
              Return False
            End If
          End If
        End If
      Next
    End With
    Return True
  End Function
  Public Sub SaveActivities(ByVal pSource As String)
    Dim vRow As Integer
    Dim vCell As Cell
    Dim vList As New ParameterList(True)
    Dim vInsert As Boolean
    Dim vInsertCount As Integer
    Dim vResult As ParameterList

    vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
    With mvSheet
      For vRow = 0 To .Rows.Count - 1
        vInsert = False
        vCell = .Cells(vRow, ActivityEntryColumns.aecActivityValueDesc)
        If TypeOf vCell.CellType Is CellType.CheckBoxCellType Then
          If vCell.Text = "True" Then
            vList("ActivityValue") = .GetText(vRow, ActivityEntryColumns.aecActivityValue)
            vInsert = True
          End If
        Else                              'Must be a combo box
          If vCell.Text.Length > 0 Then
            vList("ActivityValue") = .GetValue(vRow, ActivityEntryColumns.aecActivityValueDesc).ToString
            vInsert = True
          End If
        End If
        If vInsert Then
          'Check if there was an existing record
          If .GetValue(vRow, ActivityEntryColumns.aecOldSource) IsNot Nothing Then
            AddActivityItems(vRow, vList, pSource, True)
            vResult = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctActivities, vList)
            vInsertCount += 1
            vInsert = False
          End If
        End If
        If vInsert Then
          AddActivityItems(vRow, vList, pSource)
          vResult = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctActivities, vList)
          vInsertCount += 1
        End If
      Next
    End With
    If vInsertCount > 0 Then FormHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactCategories, mvContactInfo.ContactNumber)
  End Sub

  Private Sub AddActivityItems(ByVal pRow As Integer, ByVal pList As ParameterList, ByVal pSource As String, Optional ByVal pForUpdate As Boolean = False)
    With mvSheet
      pList("Activity") = .GetText(pRow, ActivityEntryColumns.aecActivity)
      pList("ValidFrom") = .GetText(pRow, ActivityEntryColumns.aecActivityValidFrom)
      pList("ValidTo") = .GetText(pRow, ActivityEntryColumns.aecActivityValidTo)
      pList("Quantity") = .GetText(pRow, ActivityEntryColumns.aecActivityQuantity)
      pList("ActivityDate") = .GetText(pRow, ActivityEntryColumns.aecActivityDate)
      pList("Notes") = .GetText(pRow, ActivityEntryColumns.aecActivityNotes)
      If .GetText(pRow, ActivityEntryColumns.aecActivitySource).Length > 0 Then
        pList("Source") = .GetText(pRow, ActivityEntryColumns.aecActivitySource)
      Else
        pList("Source") = pSource
      End If
      If pForUpdate Then
        pList.IntegerValue("OldContactNumber") = mvContactInfo.ContactNumber
        pList("OldActivity") = .GetText(pRow, ActivityEntryColumns.aecActivity)
        pList("OldActivityValue") = .GetText(pRow, ActivityEntryColumns.aecActivityValue)
        pList("OldSource") = .GetText(pRow, ActivityEntryColumns.aecOldSource)
        pList("OldValidFrom") = .GetText(pRow, ActivityEntryColumns.aecOldValidFrom)
        pList("OldValidTo") = .GetText(pRow, ActivityEntryColumns.aecOldValidTo)
      End If
    End With
  End Sub

  Private Sub vas_ComboCloseUp(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vas.ComboCloseUp
    HandleMultipleValues(True)
  End Sub
  Private Sub vas_EditModeOff(ByVal sender As Object, ByVal e As System.EventArgs) Handles vas.EditModeOff
    HandleMultipleValues(False)
  End Sub

  Private Sub HandleMultipleValues(ByVal pComboCloseUp As Boolean)
    Static vHandling As Boolean

    With mvSheet
      If .ActiveColumnIndex = ActivityEntryColumns.aecActivityValueDesc And Not vHandling Then
        'Exit from edit mode on a combo box - check for multiple values allowed
        Dim vRow As Integer = .ActiveRowIndex
        Dim vLen As Integer = .ActiveCell.Text.Length
        If .GetText(vRow, ActivityEntryColumns.aecMultipleValues) = "Y" Then
          vHandling = True
          Try
            'Check if at end of activity value section (Next is different
            If vRow = .Rows.Count - 1 OrElse .GetText(vRow, ActivityEntryColumns.aecActivity) <> .GetText(vRow + 1, ActivityEntryColumns.aecActivity) Then
              If vLen > 0 Then
                If ComboValueExists() Then
                  DoBeep()
                  .ActiveCell.Text = ""
                Else
                  vRow += 1
                  .AddRows(vRow, 1)
                  .CopyRange(vRow - 1, 0, vRow, 0, 1, .Columns.Count, False)
                  .SetValue(vRow, ActivityEntryColumns.aecActivityValueDesc, "")
                  SetHeight()
                End If
              End If
            Else                  'Row in the middle of a group of multiple values
              If vLen > 0 Then
                If ComboValueExists() Then
                  DoBeep()
                  If pComboCloseUp Then
                    '.ActiveCell.Text = ""
                  Else
                    .RemoveRows(vRow, 1)
                  End If
                End If
              Else
                If Not pComboCloseUp Then .RemoveRows(vRow, 1)
              End If
            End If
          Finally
            vHandling = False
          End Try
        End If
      End If
    End With
  End Sub
  Private Function ComboValueExists() As Boolean
    With mvSheet
      Dim vRow As Integer = .ActiveRowIndex
      Dim vActivity As String = .GetValue(vRow, ActivityEntryColumns.aecActivity).ToString    'Get selected activity 
      Dim vActivityValueDesc As String = .ActiveCell.Text                        'Get selected activity value desc
      'Go back to the earliest row with this activity in it
      While vRow > 0
        vRow -= 1
        If .Cells(vRow, ActivityEntryColumns.aecActivity).Text <> vActivity Then
          vRow += 1
          Exit While
        End If
      End While
      Do While vRow < .Rows.Count AndAlso .Cells(vRow, ActivityEntryColumns.aecActivity).Text = vActivity
        If vRow <> .ActiveRowIndex AndAlso .Cells(vRow, ActivityEntryColumns.aecActivityValueDesc).Text = vActivityValueDesc Then
          Return True
        End If
        vRow += 1
      Loop
    End With
    Return False
  End Function
End Class
