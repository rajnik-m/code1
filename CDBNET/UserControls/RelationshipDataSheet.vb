Imports FarPoint.Win.Spread

Public Class RelationshipDataSheet
  Inherits System.Windows.Forms.UserControl

  Private Enum RelationshipEntryColumns
    recRelationship
    recMandatory
    recGroup
    recMultipleValues
    recPostPoint
    recExists
    recContactType
    recSelectionType
    recComplementary
    recRelationshipDesc
    recContactFind
    recContactNumber
    recContactName
    recValidFrom
    recValidTo
    recNotes
  End Enum

  Private mvContactInfo As ContactInfo
  Private mvDefaultValidFrom As Date
  Private mvDefaultValidTo As Date
  Private mvSheet As SheetView
  Private mvInitialised As Boolean

  Public Sub Init(ByVal pContactInfo As ContactInfo, ByVal pGroupCode As String, ByVal pTable As DataTable)
    mvContactInfo = pContactInfo

    mvSheet = vas.Sheets(0)
    mvSheet.OperationMode = OperationMode.Normal
    mvSheet.RowHeaderVisible = False
    mvSheet.GrayAreaBackColor = DisplayTheme.GridBackAreaColor
    mvSheet.Columns.Count = RelationshipEntryColumns.recNotes + 1

    Dim vInputMap As InputMap
    vInputMap = vas.GetInputMap(InputMapMode.WhenAncestorOfFocused)
    vInputMap.Put(New Keystroke(Keys.Tab, Keys.None), SpreadActions.None)
    vInputMap = vas.GetInputMap(InputMapMode.WhenAncestorOfFocused)
    vInputMap.Put(New Keystroke(Keys.Tab, Keys.Shift), SpreadActions.None)

    Dim vDateCellType As New CellType.DateTimeCellType
    vDateCellType.DateTimeFormat = CellType.DateTimeFormat.ShortDate
    vDateCellType.DropDownButton = True
    vDateCellType.MinimumDate = New Date(1900, 1, 1)
    vDateCellType.MaximumDate = New Date(2500, 1, 1)

    Dim vNotesTextType As New CellType.TextCellType
    vNotesTextType.Multiline = True
    vNotesTextType.MaxLength = 1024

    Dim vStaticTextType As New CellType.TextCellType
    vStaticTextType.ReadOnly = True

    Dim vFindCellType As New CellType.ButtonCellType
    vFindCellType.Text = "?"

    mvDefaultValidFrom = Date.Today
    mvDefaultValidTo = Nothing

    Dim vTable As DataTable
    If pTable Is Nothing Then
      Dim vParams As New ParameterList(True)
      vParams("UsageCode") = "B"
      vParams("RelationshipGroup") = pGroupCode
      vParams(pContactInfo.ContactGroupParameterName) = pContactInfo.ContactGroup
      vTable = DataHelper.GetLookupData(CareServices.XMLLookupDataTypes.xldtRelationshipDataSheet, vParams)
    Else
      vTable = pTable
    End If
    'RelationshipCode,RelationshipDesc,Mandatory,ToContactGroup,MultipleValues,ContactSelectionType,ComplimentaryRelationship,PostPoint"
    With mvSheet
      .ColumnHeaderVisible = True
      .Columns(RelationshipEntryColumns.recRelationship).Visible = False
      .Columns(RelationshipEntryColumns.recMandatory).Visible = False
      .Columns(RelationshipEntryColumns.recGroup).Visible = False
      .Columns(RelationshipEntryColumns.recMultipleValues).Visible = False
      .Columns(RelationshipEntryColumns.recSelectionType).Visible = False
      .Columns(RelationshipEntryColumns.recComplementary).Visible = False
      .Columns(RelationshipEntryColumns.recPostPoint).Visible = False

      .Columns(RelationshipEntryColumns.recExists).Visible = False
      .Columns(RelationshipEntryColumns.recContactType).Visible = False
      .Columns(RelationshipEntryColumns.recContactNumber).Visible = False

      .Columns(RelationshipEntryColumns.recContactName).CellType = vStaticTextType
      .Columns(RelationshipEntryColumns.recContactFind).CellType = vFindCellType

      .Columns(RelationshipEntryColumns.recValidFrom).CellType = vDateCellType
      .Columns(RelationshipEntryColumns.recValidTo).CellType = vDateCellType
      .Columns(RelationshipEntryColumns.recNotes).CellType = vNotesTextType

      .ColumnHeader.Columns(RelationshipEntryColumns.recRelationshipDesc).Label = "Relationship"
      .ColumnHeader.Columns(RelationshipEntryColumns.recContactName).Label = "With"
      .ColumnHeader.Columns(RelationshipEntryColumns.recValidFrom).Label = "ValidFrom"
      .ColumnHeader.Columns(RelationshipEntryColumns.recValidTo).Label = "ValidTo"
      .ColumnHeader.Columns(RelationshipEntryColumns.recNotes).Label = "Notes"

      Dim vRow As Integer = 0
      For Each vDataRow As DataRow In vTable.Rows
        .RowCount = vRow + 1
        'RelationshipCode,RelationshipDesc,Mandatory,ToContactGroup,MultipleValues,ContactSelectionType,ComplimentaryRelationship,PostPoint"
        .SetValue(vRow, RelationshipEntryColumns.recRelationship, vDataRow("RelationshipCode").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recRelationshipDesc, vDataRow("RelationshipDesc").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recMandatory, vDataRow("Mandatory").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recGroup, vDataRow("ToContactGroup").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recMultipleValues, vDataRow("MultipleValues").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recSelectionType, vDataRow("ContactSelectionType").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recComplementary, vDataRow("ComplimentaryRelationship").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recPostPoint, vDataRow("PostPoint").ToString)
        .SetValue(vRow, RelationshipEntryColumns.recValidFrom, mvDefaultValidFrom)
        If mvDefaultValidTo <> Nothing Then .SetValue(vRow, RelationshipEntryColumns.recValidTo, mvDefaultValidTo)
        vRow += 1
      Next
    End With
    SetPreferredWidths()
    PopulateRelationships()
    mvInitialised = True
    SetHeight()
  End Sub

  Public Sub SetHeight()
    If mvInitialised Then
      Dim vRequiredHeight As Integer = CInt((mvSheet.Rows(0).Height * mvSheet.Rows.Count) + mvSheet.ColumnHeader.Rows(0).Height + SystemInformation.HorizontalScrollBarHeight + (SystemInformation.Border3DSize.Height * 2) + Me.DockPadding.Top + Me.DockPadding.Bottom)
      Me.Height = Math.Min(vRequiredHeight, Me.Parent.DisplayRectangle.Height)
    End If
  End Sub

  Private Sub SetPreferredWidths()
    Dim vDefaultWidth As Single
    Dim vPreferredWidth As Single
    For Each vCol As Column In mvSheet.Columns
      Select Case vCol.Index
        Case RelationshipEntryColumns.recRelationshipDesc, RelationshipEntryColumns.recContactName
          vDefaultWidth = 200
        Case RelationshipEntryColumns.recNotes
          vDefaultWidth = 300
        Case RelationshipEntryColumns.recValidFrom, RelationshipEntryColumns.recValidTo
          vDefaultWidth = 100
        Case RelationshipEntryColumns.recContactFind
          vCol.Width = 20
          Continue For
      End Select
      vPreferredWidth = vCol.GetPreferredWidth
      vCol.Width = Math.Max(vPreferredWidth, vDefaultWidth)
    Next
  End Sub

  Private Sub PopulateRelationships()
    Dim vList As New ParameterList(True)
    Dim vRelationships As New ArrayListEx
    Dim vRelationship As String
    Dim vRow As Integer

    With mvSheet
      Dim vStaticTextType As New CellType.TextCellType
      vStaticTextType.ReadOnly = True
      For vRow = 0 To .Rows.Count - 1
        vRelationships.Add(.GetText(vRow, RelationshipEntryColumns.recRelationship))
      Next
      vList("Relationships") = vRelationships.CSStringList
      vList("IgnoreUnknownParameters") = "Y"        'TODO Add Relationships as a valid XML parameter name
      Dim vDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, mvContactInfo.ContactNumber, vList)
      If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Rows.Count > 0 Then
        For Each vDataRow As DataRow In vDataSet.Tables("DataRow").Rows
          'Now go and read any existing relationships
          If vDataRow("Type2").ToString = "O" Then    'It's an org type
            vRelationship = vDataRow("RelationshipCode").ToString
            For vRow = 0 To .Rows.Count - 1
              If vRelationship = .GetText(vRow, RelationshipEntryColumns.recRelationship) Then
                If .GetText(vRow, RelationshipEntryColumns.recContactNumber).Length = 0 Then      'Not already set so add it here
                  SetContact(vRow, CInt(vDataRow("ContactNumber")), vDataRow("ContactName").ToString, vDataRow("Type2").ToString)
                  .SetValue(vRow, RelationshipEntryColumns.recValidFrom, vDataRow.Item("ValidFrom"))
                  .SetValue(vRow, RelationshipEntryColumns.recValidTo, vDataRow.Item("ValidTo"))
                  .SetText(vRow, RelationshipEntryColumns.recNotes, vDataRow.Item("Notes").ToString)
                  .SetText(vRow, RelationshipEntryColumns.recExists, "Y")
                  .Cells(vRow, RelationshipEntryColumns.recContactFind).CellType = vStaticTextType
                  Exit For
                End If
              End If
            Next
          End If
        Next
        'Next do the contact type relationships
        'The complication here is they need to go in under the correct organisation
        Dim vLastOrg As Integer
        Dim vFoundPosition As Boolean

        For Each vDataRow As DataRow In vDataSet.Tables("DataRow").Rows
          If vDataRow("Type2").ToString = "C" Then    'It's a contact type
            vRelationship = vDataRow("RelationshipCode").ToString
            vLastOrg = 0
            For vRow = 0 To .Rows.Count - 1
              If .GetText(vRow, RelationshipEntryColumns.recContactType) = "O" Then
                vLastOrg = IntegerValue(.GetText(vRow, RelationshipEntryColumns.recContactNumber))
              End If
              If vRelationship = .GetText(vRow, RelationshipEntryColumns.recRelationship) Then
                If IntegerValue(.GetText(vRow, RelationshipEntryColumns.recContactNumber)) = 0 Then      'Not already set so add it here
                  vFoundPosition = False
                  If .GetText(vRow, RelationshipEntryColumns.recSelectionType) = "P" And vLastOrg > 0 Then
                    Dim vPostionsDataSet As DataSet = DataHelper.GetContactData(CareServices.XMLContactDataSelectionTypes.xcdtContactPositions, CInt(vDataRow("ContactNumber")))
                    If vDataSet.Tables.Contains("DataRow") AndAlso vDataSet.Tables("DataRow").Rows.Count > 0 Then
                      For Each vPositionRow As DataRow In vDataSet.Tables("DataRow").Rows
                        If CInt(vPositionRow("ContactNumber")) = vLastOrg Then
                          vFoundPosition = True
                          Exit For
                        End If
                      Next
                    End If
                  Else
                    vFoundPosition = True
                  End If
                  If vFoundPosition Then
                    SetContact(vRow, CInt(vDataRow("ContactNumber")), vDataRow("ContactName").ToString, vDataRow("Type2").ToString)
                    .SetValue(vRow, RelationshipEntryColumns.recValidFrom, vDataRow.Item("ValidFrom"))
                    .SetValue(vRow, RelationshipEntryColumns.recValidTo, vDataRow.Item("ValidTo"))
                    .SetText(vRow, RelationshipEntryColumns.recNotes, vDataRow.Item("Notes").ToString)
                    .SetText(vRow, RelationshipEntryColumns.recExists, "Y")
                    .Cells(vRow, RelationshipEntryColumns.recContactFind).CellType = vStaticTextType
                    Exit For
                  End If
                End If
              End If
            Next
          End If
        Next
        For vRow = 0 To .Rows.Count - 1
          If .Cells(vRow, RelationshipEntryColumns.recContactFind).CellType IsNot vStaticTextType Then
            .SetActiveCell(vRow, RelationshipEntryColumns.recContactFind)
            Exit For
          End If
        Next
      End If
    End With
  End Sub

  Public Sub SetContact(ByVal pRow As Integer, ByVal pNumber As Integer, ByVal pName As String, ByVal pType As String)
    Dim vRelationship As String
    Dim vInsert As Boolean
    Dim vRow As Integer
    Dim vGroupCode As String
    Dim vTempRow As Integer
    Dim vSelType As String
    Dim vRowCount As Integer
    Dim vContactInfo As ContactInfo
    Dim vInValid As Boolean

    'This routine assumes .Row is set to the row that needs to have the contact set
    With mvSheet
      vRow = pRow
      vRelationship = .GetText(vRow, RelationshipEntryColumns.recRelationship)
      If pNumber = mvContactInfo.ContactNumber Then
        ShowWarningMessage(InformationMessages.imCannotLinkToSelf)
        vas.Focus()
      ElseIf CheckLinkedAlready(vRelationship, pNumber) Then
        ShowWarningMessage(InformationMessages.imLinkExists)
        vas.Focus()
      Else
        vGroupCode = .GetText(vRow, RelationshipEntryColumns.recGroup)
        If vGroupCode.Length > 0 Then
          vContactInfo = New ContactInfo(pNumber)
          If vGroupCode <> vContactInfo.ContactGroup Then
            vInValid = True
            DoBeep()
          End If
        End If
        If Not vInValid Then
          vInsert = .GetText(vRow, RelationshipEntryColumns.recContactNumber).Length = 0     'Set up for multi value insert only if not editing          
          .SetText(vRow, RelationshipEntryColumns.recContactNumber, pNumber.ToString)
          .SetText(vRow, RelationshipEntryColumns.recContactName, pName)
          .SetText(vRow, RelationshipEntryColumns.recContactType, pType)

          If .GetText(vRow, RelationshipEntryColumns.recMultipleValues) = "Y" And vInsert Then
            'If this multiple is an org type and there are following records which are dependant then we need to copy them
            vRowCount = 1
            If pType = "O" Then
              vTempRow = vRow
              Do
                vTempRow = vTempRow + 1
                If vTempRow < .Rows.Count Then
                  vSelType = .GetText(vTempRow, RelationshipEntryColumns.recSelectionType)
                  If vSelType = "P" Then vRowCount = vRowCount + 1
                Else
                  vSelType = ""
                End If
              Loop While vSelType = "P"
            End If
            .AddRows(vRow + vRowCount, vRowCount)
            .CopyRange(vRow, 0, vRow + vRowCount, 0, vRowCount, .Columns.Count, False)      'Make a copy of the existing row
            vRow += vRowCount
            For vIndex As Integer = 1 To vRowCount
              .SetText(vRow, RelationshipEntryColumns.recMandatory, "N")
              .SetText(vRow, RelationshipEntryColumns.recContactType, "")
              .SetText(vRow, RelationshipEntryColumns.recContactName, "")
              .SetText(vRow, RelationshipEntryColumns.recContactNumber, "")
              .SetText(vRow, RelationshipEntryColumns.recNotes, "")
              vRow += 1
            Next
            .SetActiveCell(vRow, RelationshipEntryColumns.recContactFind)
          End If
        End If
      End If
    End With
  End Sub

  Private Function CheckLinkedAlready(ByVal pRelationship As String, ByVal pNumber As Integer) As Boolean
    Dim vRow As Integer

    With mvSheet
      For vRow = 0 To .Rows.Count - 1
        If .GetText(vRow, RelationshipEntryColumns.recRelationship) = pRelationship Then
          If IntegerValue(.GetText(vRow, RelationshipEntryColumns.recContactNumber)) = pNumber Then
            Return True
          End If
        End If
      Next
    End With
  End Function

  Public Function ValidateRelationships() As Boolean
    Dim vRow As Integer
    Dim vRelationship As String
    Dim vRelationshipDesc As String
    Dim vValidFrom As String
    Dim vValidTo As String
    Dim vLinkedContact As Integer

    With mvSheet
      For vRow = 0 To .Rows.Count - 1
        vRelationship = .GetText(vRow, RelationshipEntryColumns.recRelationship)
        vRelationshipDesc = .GetText(vRow, RelationshipEntryColumns.recRelationshipDesc)
        vValidFrom = .GetText(vRow, RelationshipEntryColumns.recValidFrom)
        vValidTo = .GetText(vRow, RelationshipEntryColumns.recValidTo)
        vLinkedContact = IntegerValue(.GetText(vRow, RelationshipEntryColumns.recContactNumber))
        If vRelationship.Length > 0 And vLinkedContact > 0 Then
          If IsDate(vValidFrom) And IsDate(vValidTo) Then
            If CDate(vValidFrom) > CDate(vValidTo) Then
              ShowWarningMessage(InformationMessages.imValidFromGTValidTo)
              .SetActiveCell(vRow, RelationshipEntryColumns.recContactNumber)
              Return False
            End If
          End If
        Else
          If .GetText(vRow, RelationshipEntryColumns.recMandatory) = "Y" Then
            ShowWarningMessage(InformationMessages.imRequiredEntry, vRelationshipDesc)
            .SetActiveCell(vRow, RelationshipEntryColumns.recContactFind)
            Return False
          End If
        End If
      Next
    End With
    Return True
  End Function

  Public Sub SaveRelationships()
    Dim vRow As Integer
    Dim vRelationship As String
    Dim vRelationshipDesc As String
    Dim vValidFrom As String
    Dim vValidTo As String
    Dim vLinkedContact As Integer
    Dim vNotes As String
    Dim vParams As ParameterList
    Dim vList As ParameterList
    Dim vToCount As Integer
    Dim vFromCount As Integer

    With mvSheet
      For vRow = 0 To .Rows.Count - 1
        vRelationship = .GetText(vRow, RelationshipEntryColumns.recRelationship)
        vRelationshipDesc = .GetText(vRow, RelationshipEntryColumns.recRelationshipDesc)
        vValidFrom = .GetText(vRow, RelationshipEntryColumns.recValidFrom)
        vValidTo = .GetText(vRow, RelationshipEntryColumns.recValidTo)
        vLinkedContact = IntegerValue(.GetText(vRow, RelationshipEntryColumns.recContactNumber))
        vNotes = .GetText(vRow, RelationshipEntryColumns.recNotes)

        If vRelationship.Length > 0 And vLinkedContact > 0 Then
          vParams = New ParameterList(True)
          vParams.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
          vParams.IntegerValue("ContactNumber2") = vLinkedContact
          vParams("OldRelationship") = vRelationship
          vParams("Relationship") = vRelationship
          vParams("ValidFrom") = vValidFrom
          vParams("ValidTo") = vValidTo
          vParams("Notes") = vNotes
          If .GetText(vRow, RelationshipEntryColumns.recExists) = "Y" Then
            vList = DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctLink, vParams)
          Else
            vList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctLink, vParams)
            vRelationship = .GetText(vRow, RelationshipEntryColumns.recComplementary)
            If vRelationship.Length > 0 Then
              vParams.IntegerValue("ContactNumber") = vLinkedContact
              vParams.IntegerValue("ContactNumber2") = mvContactInfo.ContactNumber
              vParams("Relationship") = vRelationship
              vList = DataHelper.AddItem(CareServices.XMLMaintenanceControlTypes.xmctLink, vParams)
              FormHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, vLinkedContact)
              vFromCount += 1
            End If
          End If
          FormHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom, vLinkedContact)
          vToCount += 1
        End If
      Next
      If vToCount > 0 Then FormHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksTo, mvContactInfo.ContactNumber)
      If vFromCount > 0 Then FormHelper.RefreshData(CareServices.XMLContactDataSelectionTypes.xcdtContactLinksFrom, mvContactInfo.ContactNumber)
    End With
  End Sub

  Private Sub vas_ButtonClicked(ByVal sender As Object, ByVal e As FarPoint.Win.Spread.EditorNotifyEventArgs) Handles vas.ButtonClicked
    Dim vGroupCode As String
    Dim vSelectionType As String
    Dim vOrgNumber As Integer
    Dim vRow As Integer
    Dim vFinderType As CareServices.XMLDataFinderTypes

    Dim vList As New ParameterList
    With mvSheet
      vRow = e.Row
      vGroupCode = .GetText(vRow, RelationshipEntryColumns.recGroup)
      If vGroupCode.Length = 0 Then vGroupCode = "CON"
      If DataHelper.ContactAndOrganisationGroups(vGroupCode).Type = EntityGroup.EntityGroupTypes.egtContactGroup Then
        vFinderType = CareServices.XMLDataFinderTypes.xdftContacts
        If vGroupCode.Length > 0 Then
          vList("ContactGroup") = vGroupCode
          vList("LockContactGroup") = "Y"
        End If
        vSelectionType = .GetText(vRow, RelationshipEntryColumns.recSelectionType)
        Select Case vSelectionType
          Case "P"                         'Position related
            Do While vRow > 0
              vRow -= 1
              If .GetText(vRow, RelationshipEntryColumns.recContactType) = "O" Then
                vOrgNumber = IntegerValue(.GetText(vRow, RelationshipEntryColumns.recContactNumber))
                If vOrgNumber > 0 Then
                  vList.IntegerValue("OrganisationNumber") = vOrgNumber
                  Exit Do
                End If
              End If
            Loop
          Case "O"                          'My organisation related
            vList.IntegerValue("OrganisationNumber") = DataHelper.UserInfo.OrganisationNumber
          Case "R"                          'Post Point recipient
            vList("PostPoint") = .GetText(vRow, RelationshipEntryColumns.recPostPoint)
          Case Else
            'regular finder
        End Select
      Else
        vFinderType = CareServices.XMLDataFinderTypes.xdftOrganisations
        If vGroupCode.Length > 0 Then
          vList("OrganisationGroup") = vGroupCode
          vList("LockContactGroup") = "Y"
        End If
      End If

    End With
    Dim vResult As Integer = FormHelper.ShowFinder(vFinderType, vList, Me.ParentForm)
    If vResult > 0 Then
      Dim vContactInfo As New ContactInfo(vResult)
      SetContact(e.Row, vContactInfo.ContactNumber, vContactInfo.ContactName, vContactInfo.ContactTypeCode)
      SetHeight()
    End If
  End Sub
End Class
