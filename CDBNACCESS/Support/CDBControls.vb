Imports System.Linq
Namespace Access

  Public Class CDBControls
    Inherits CollectionList(Of CDBControl)

    Private Const RADIO_BUTTON_GAP As Integer = 50
    Private Const DEFAULT_HEIGHT As Integer = 300
    Private Const DEFAULT_BUTTON_HEIGHT As Integer = 400
    Private Const DEFAULT_LEFT As Integer = 100
    Private Const DEFAULT_TOP As Integer = 100
    Private Const DEFAULT_FINDER_TOP As Integer = 400
    Private Const DEFAULT_GAP As Integer = 50

    'Protected mvMaxHeight As Long
    Private mvPageType As String
    Private mvCustomisedControls As Boolean
    Private mvMultiplePages As Boolean = False

    Public Sub New()
      MyBase.New(1)
    End Sub

    'Public Overloads Function Add(ByVal pCDBControl As CDBControl) As CDBControl
    '  'If pCDBControl.ControlTop + pCDBControl.ControlHeight > mvMaxHeight Then mvMaxHeight = pCDBControl.ControlTop + pCDBControl.ControlHeight
    '  If Exists(pCDBControl.ParameterName) Then MyBase.Add(pCDBControl.AttributeName, pCDBControl) Else MyBase.Add(pCDBControl.ParameterName, pCDBControl)
    '  Return pCDBControl
    'End Function

    'Public Function AddFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet) As CDBControl
    '  Return AddFromRecordSet(pEnv, pRecordSet, 0)
    'End Function

    Public Overloads Function Add(ByVal pCDBControl As CDBControl) As CDBControl
      MyBase.Add(MyBase.Count.ToString, pCDBControl)
      Return pCDBControl
    End Function

    Private Function AddFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pErrorOnDuplicateSequenceNumber As Boolean, ByVal pMultiplePages As Boolean, pQBE As Boolean) As CDBControl
      Dim vCDBControl As New CDBControl(pEnv)
      vCDBControl.InitFromRecordSet(pRecordSet)
      'If pUseSequenceAsIndex = False AndAlso vCDBControl.ParameterName.Length = 0 Then
      '  Dim vBaseName As String
      '  If vCDBControl.ControlType = "tab" Then
      '    vBaseName = "TAB"
      '  Else
      '    vBaseName = vCDBControl.AttributeName
      '  End If
      '  'If pStartNumberingAt > 0 Then vBaseName = vBaseName & pStartNumberingAt
      '  Dim vName As String
      '  If Exists(vBaseName) Then
      '    Dim vCount As Integer
      '    'If pStartNumberingAt > 0 Then
      '    ' vBaseName = Substring(vBaseName, 0, vBaseName.Length - pStartNumberingAt.ToString.Length)
      '    'vCount = pStartNumberingAt
      '    'Else
      '    vCount = 1                              'Start with 2
      '    'End If
      '    Do
      '      vCount = vCount + 1
      '      vName = vBaseName & vCount
      '    Loop While MyBase.ContainsKey(vName)
      '  Else
      '    vName = vBaseName
      '  End If
      '  vCDBControl.ParameterName = vName
      'End If
      'If vCDBControl.ControlTop + vCDBControl.ControlHeight > mvMaxHeight Then mvMaxHeight = vCDBControl.ControlTop + vCDBControl.ControlHeight
      'MyBase.Add(vCDBControl.ParameterName, vCDBControl)
      Dim vIndexKey As String = vCDBControl.SequenceNumber.ToString
      If pMultiplePages Then vIndexKey = vCDBControl.FpPageType & vIndexKey
      If pQBE AndAlso vCDBControl.FpApplication.Length = 0 Then
        If MyBase.ContainsKey(vIndexKey) Then
          Return Nothing                          'A customised version of this exists
        ElseIf MyBase.ContainsKey(vCDBControl.FpPageType & "1") Then
          If MyBase.Item(vCDBControl.FpPageType & "1").FpApplication.Length > 0 Then
            Return Nothing                          'A customised version of the page exists
          End If
        End If
      ElseIf MyBase.ContainsKey(vIndexKey) Then
        If pErrorOnDuplicateSequenceNumber Then
          RaiseError(DataAccessErrors.daeDatabaseUpgrade)
        Else
          'Handle incorrect sequence numbers in the dbcontrols list
          Dim vIndex As Integer = vCDBControl.SequenceNumber + 1
          While MyBase.ContainsKey(vIndex.ToString)
            vIndex += 1
          End While
          vIndexKey = vIndex.ToString
        End If
      End If
      MyBase.Add(vIndexKey, vCDBControl)
      Return vCDBControl
    End Function

    Public Sub GetPageControls(ByVal pEnv As CDBEnvironment, ByVal pPageType As String, ByVal pPageList As StringList, ByVal pErrorOnDuplicateSequenceNumber As Boolean)
      Dim vWhereFields As New CDBFields
      If pPageList Is Nothing Then
        vWhereFields.Add("fp_application")
      Else
        vWhereFields.Add("fp_application", pPageList.ItemList, CDBField.FieldWhereOperators.fwoIn Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("fp_application#2", "", CDBField.FieldWhereOperators.fwoEqual Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      GetPageControls(pEnv, pPageType, vWhereFields, pErrorOnDuplicateSequenceNumber)
    End Sub

    Public Sub GetPageControls(ByVal pEnv As CDBEnvironment, ByVal pPageType As String, ByVal pApplication As Integer, ByVal pErrorOnDuplicateSequenceNumber As Boolean)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("fp_application", pApplication, CDBField.FieldWhereOperators.fwoNullOrEqual)
      GetPageControls(pEnv, pPageType, vWhereFields, pErrorOnDuplicateSequenceNumber)
    End Sub

    Public Sub GetPageControls(ByVal pEnv As CDBEnvironment, ByVal pPageType As String, pWhereFields As CDBFields, ByVal pErrorOnDuplicateSequenceNumber As Boolean)
      Dim vFWO As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoEqual
      Dim vQBE As Boolean
      If pPageType.Contains(",") Then
        Dim vStringList As New StringList(pPageType)
        vFWO = CDBField.FieldWhereOperators.fwoIn
        pPageType = vStringList.InList
        mvMultiplePages = True
      ElseIf pPageType.Contains("*") Then
        vFWO = CDBField.FieldWhereOperators.fwoLike
        mvMultiplePages = True
        vQBE = True         'This should only be in the case of QBE
      End If
      pWhereFields.Add("fp_page_type", pPageType, vFWO)
      mvPageType = pPageType
      Dim vAnsiJoins As New AnsiJoins
      vAnsiJoins.Add("maintenance_attributes ma", "fc.table_name", "ma.table_name", "fc.attribute_name", "ma.attribute_name")
      Dim vCDBControl As CDBControl = New CDBControl(pEnv)
      vCDBControl.Init()
      Dim vOrderBy As String
      If pEnv.Connection.NullsSortAtEnd Then
        vOrderBy = "fp_application"
      Else
        vOrderBy = "fp_application DESC"
      End If
      If mvMultiplePages Then vOrderBy &= ", fp_page_type"
      vOrderBy &= ", fc.sequence_number"
      Dim vSQL As New SQLStatement(pEnv.Connection, vCDBControl.GetRecordSetFields, "fp_controls fc", pWhereFields, vOrderBy, vAnsiJoins)
      Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet
      mvCustomisedControls = False
      While vRecordSet.Fetch
        If vRecordSet.Fields("fp_application").LongValue > 0 Then
          mvCustomisedControls = True
        Else
          If mvCustomisedControls Then
            'Normally if we have read some customised controls and then get to the default ones we want to ignore them
            'But not necessarily when dealing with QBE pages
            If Not vQBE Then Exit While
          End If
        End If
        AddFromRecordSet(pEnv, vRecordSet, pErrorOnDuplicateSequenceNumber, mvMultiplePages, vQBE)
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public ReadOnly Property CustomisedControls() As Boolean
      Get
        Return mvCustomisedControls
      End Get
    End Property

    Private Function ControlBySequenceNumber(ByVal pSequenceNumber As Integer) As CDBControl
      For Each vControl As CDBControl In Me
        If vControl.SequenceNumber = pSequenceNumber Then Return vControl
      Next
      Return Nothing
    End Function

    Public Sub RevertPageControls(ByVal pApplicationNumber As Integer)
      If Me.Count > 0 Then
        Me.Item(1).RevertPageControls(pApplicationNumber, mvPageType)
      End If
    End Sub

    Public Enum UpdateControlTypes
      None
      AdjustPositons
      AdjustTops
      AdjustTopsForFinders
      AdjustTopsForTrader
    End Enum

    Public Sub UpdatePageControls(ByVal pEnv As CDBEnvironment, ByVal pControlList As Xml.XmlNodeList, ByVal pApplication As Integer, ByVal pUpdateType As UpdateControlTypes)
      For Each vControl As CDBControl In Me
        If vControl.FpApplication.Length = 0 Then vControl.SetApplication(pApplication)
      Next

      Dim vSequenceChanged As Boolean
      Dim vWhereFields As New CDBFields
      Dim vDataSet As DataSet
      vWhereFields.Add("fp_application", CDBField.FieldTypes.cftCharacter, pApplication)
      vWhereFields.Add("fp_page_type", Me(1).FpPageType)
      Dim vSQL As New SQLStatement(pEnv.Connection, "attribute_name,sequence_number,control_top", "fp_controls", vWhereFields)
      vDataSet = vSQL.GetDataSet()
      If vDataSet.Tables(0).Rows.Count = 0 Then
        vWhereFields.Remove("fp_application")
        vWhereFields.Add("fp_application", CDBField.FieldTypes.cftCharacter, "")
        vDataSet = vSQL.GetDataSet()
      End If

      Dim vLastControl As CDBControl = Nothing
      For Each vItem As Xml.XmlNode In pControlList
        Dim vAttributes As Xml.XmlAttributeCollection = vItem.Attributes
        Dim vSequenceNumber As Integer = IntegerValue(vAttributes("OldSequenceNumber").InnerText)
        Dim vControlType As String = Mid(vAttributes("ControlType").InnerText, 1, 3)
        Dim vControl As CDBControl = Nothing
        If MyBase.ContainsKey(vSequenceNumber.ToString) Then vControl = Item(vSequenceNumber.ToString)
        If vControl IsNot Nothing Then
          Dim vCaption As String = ""
          Dim vVisible As Boolean = True
          Dim vMandatory As Boolean = False
          Dim vReadonly As Boolean = False
          Dim vWidth As Integer = 1
          Dim vHeight As Integer = 1
          Dim vTop As Integer = 1
          Dim vLeft As Integer = 1
          Dim vNextToPreviousItem As Boolean = False
          Dim vHelpText As String = ""
          Dim vNewSequenceNumber As Integer = vControl.SequenceNumber
          Dim vDefaultValue As String = String.Empty
          Dim vContactGroup As String = String.Empty
          Dim vCaptionWidth As Integer = vControl.CaptionWidth
          For Each vAttribute As Xml.XmlAttribute In vAttributes
            Select Case vAttribute.LocalName
              Case "ControlCaption"
                vCaption = vAttribute.Value
              Case "Visible"
                vVisible = vAttribute.Value = "True"
              Case "MandatoryItem"
                vMandatory = vAttribute.Value = "True"
              Case "ControlWidth"
                vWidth = CInt(vAttribute.Value) * 100
              Case "ControlHeight"
                If vControlType = "cmd" Then
                  If vAttributes("OriginalControlHeight") IsNot Nothing AndAlso vAttribute.Value <> vAttributes("OriginalControlHeight").Value Then
                    vHeight = CInt(vAttribute.Value) * DEFAULT_BUTTON_HEIGHT
                  Else
                    vHeight = CInt(vAttributes("OriginalControlHeightValue").Value)
                  End If
                Else
                  vHeight = CInt(vAttribute.Value) * DEFAULT_HEIGHT
                End If
              Case "SequenceNumber"
                vNewSequenceNumber = IntegerValue(vAttribute.Value)
              Case "ReadonlyItem"
                vReadonly = vAttribute.Value = "True"
              Case "DefaultValue"
                vDefaultValue = vAttribute.Value
              Case "ControlTop"
                vTop = IntegerValue((CDbl(vAttribute.Value) * 100).ToString)
              Case "ControlLeft"
                vLeft = IntegerValue((CDbl(vAttribute.Value) * 100).ToString)
              Case "NextToPreviousItem"
                vNextToPreviousItem = (vAttribute.Value = "True" OrElse vAttribute.Value.Equals("Y", StringComparison.InvariantCultureIgnoreCase)) AndAlso vLastControl IsNot Nothing AndAlso vLastControl.Visible = True
              Case "HelpText"
                'BR17672 truncate to 80 chars
                If vAttribute.Value.Length > 80 Then
                  vHelpText = vAttribute.Value.Substring(0, 80)
                Else
                  vHelpText = vAttribute.Value
                End If
              Case "ContactGroup"
                vContactGroup = vAttribute.Value
              Case "CaptionWidth"
                vCaptionWidth = CInt(vAttribute.Value) * 100
            End Select
          Next
          If vNewSequenceNumber <> vControl.SequenceNumber AndAlso vControl.Existing Then vSequenceChanged = True
          vControl.Update(vCaption, vVisible, vWidth, vHeight, vMandatory, vNewSequenceNumber, vReadonly, vDefaultValue, vTop, vLeft, vNextToPreviousItem, vHelpText, vContactGroup, vCaptionWidth)
          vLastControl = vControl
        End If
      Next
      If pUpdateType <> UpdateControlTypes.None AndAlso pUpdateType <> UpdateControlTypes.AdjustTopsForTrader Then
        'Reorder the controls into the new sequence so we can adjust the tops by stepping through the controls in the correct sequence
        Dim vNewList As New CollectionList(Of CDBControl)
        If Me.Count > 0 Then
          Do
            Dim vControlToAdd As CDBControl = Me.Item(1)
            For Each vItem As CDBControl In Me
              If vItem.SequenceNumber < vControlToAdd.SequenceNumber Then vControlToAdd = vItem
            Next
            vNewList.Add(vControlToAdd.SequenceNumber.ToString, vControlToAdd)
            Me.Remove(vControlToAdd)
          Loop While Me.Count > 0
        End If

        'Need to do this separately as pControlList is not always in the correct order
        Dim vTop As Integer = DEFAULT_TOP
        Dim vLastTop As Integer = DEFAULT_TOP
        vLastControl = Nothing
        Dim vControlsOnRow As Integer
        Dim vLeft As Integer
        Dim vStandardLeft As Integer
        Dim vMaxControlHeightInRow As Integer
        Dim vNextLeftTop As Integer
        For Each vControl As CDBControl In vNewList
          Select Case pUpdateType
            Case UpdateControlTypes.AdjustPositons
              If vControl.ControlType.StartsWith("opt") OrElse vControl.ControlType.StartsWith("chk") Then
                vLeft = DEFAULT_LEFT
              ElseIf vControl.ControlType.StartsWith("tab") Then
                vTop = DEFAULT_TOP
              Else
                vLeft = vControl.ControlLeft
              End If
              If vLastControl IsNot Nothing Then
                If vStandardLeft = 0 Then vStandardLeft = vLastControl.ControlLeft + vControl.CaptionWidth
                If vControl.NextToPreviousItem Then
                  If vControl.Visible Then
                    If vControl.ControlHeight > vLastControl.ControlHeight AndAlso vMaxControlHeightInRow < vControl.ControlHeight Then
                      vMaxControlHeightInRow = vControl.ControlHeight
                    Else
                      If vMaxControlHeightInRow = 0 Then vMaxControlHeightInRow = vLastControl.ControlHeight
                    End If
                    'This Control is the same type and either a checkbox or the same attribute name
                    If vControlsOnRow < 2 OrElse Not (vLastControl.ControlType.Substring(0, 3) = vControl.ControlType.Substring(0, 3) AndAlso _
                      (vControl.ControlType.StartsWith("chk") OrElse vLastControl.AttributeName = vControl.AttributeName)) Then
                      vTop = vLastControl.ControlTop
                      If vLastControl.ControlType.Substring(0, 3) = vControl.ControlType.Substring(0, 3) AndAlso _
                      (vControl.ControlType.StartsWith("chk") OrElse vControl.ControlType.StartsWith("opt")) Then
                        vLeft = (vLastControl.ControlLeft + vLastControl.ControlWidth + DEFAULT_GAP)
                      Else
                        If vControl.ControlType = "btn" AndAlso vLastControl.ControlType = "btn" Then 'Don't add CaptionWidth
                          vLeft = (vLastControl.ControlLeft + vLastControl.ControlWidth + DEFAULT_GAP)
                        Else
                          vLeft = (vLastControl.ControlLeft + vLastControl.ControlWidth + vLastControl.CaptionWidth + DEFAULT_GAP)  'Always add CaptionWidth for all other controls
                        End If
                      End If
                      vControlsOnRow += 1
                    Else
                      vLeft = DEFAULT_LEFT
                      vControlsOnRow = 0
                    End If
                  Else
                    'vLeft = DEFAULT_LEFT
                  End If
                ElseIf (vLeft > vStandardLeft) Then
                  If vControl.ControlType = "btn" Then
                    vLeft = vStandardLeft
                  Else
                    If vControl.CaptionWidth < (vStandardLeft - DEFAULT_LEFT) Then vControl.CaptionWidth = vStandardLeft - DEFAULT_LEFT
                    vLeft = DEFAULT_LEFT
                  End If
                Else
                  vControlsOnRow = 0
                End If
                If vLastControl.ControlType.StartsWith("opt") AndAlso (vLastControl.AttributeName <> vControl.AttributeName) Then
                  If vControl.ControlType.StartsWith("opt") Then
                    vTop += RADIO_BUTTON_GAP
                  End If
                  vTop += RADIO_BUTTON_GAP
                End If
              End If
              vControl.SetControlLocation(vLeft, vTop)
              If vControl.Visible Then
                If vMaxControlHeightInRow > 0 AndAlso vControl.NextToPreviousItem Then
                  vTop += (vMaxControlHeightInRow + DEFAULT_GAP)
                Else
                  vTop += (vControl.ControlHeight + DEFAULT_GAP)
                  vMaxControlHeightInRow = 0
                End If
                vLastControl = vControl
              End If
            Case UpdateControlTypes.AdjustTops, UpdateControlTypes.AdjustTopsForFinders
              vLeft = vControl.ControlLeft
              If vControl.ControlType.StartsWith("tab") Then
                If vControl.Visible Then
                  vNextLeftTop = DEFAULT_TOP
                  vTop = DEFAULT_TOP
                  vLastTop = DEFAULT_TOP
                  vControl.SetControlLocation(vLeft, vTop)
                  If pUpdateType = UpdateControlTypes.AdjustTopsForFinders Then
                    vTop = DEFAULT_FINDER_TOP
                    vLastControl = vControl
                  Else
                    vLastControl = Nothing
                  End If
                End If
              Else
                If vLastControl Is Nothing Then
                  vTop = DEFAULT_TOP
                Else
                  If (vControl.Visible OrElse (vControl.ControlTop.Equals(DEFAULT_TOP) AndAlso vLastControl.ControlTop > vControl.ControlTop)) AndAlso _
                     ((vLeft + vControl.CaptionWidth) > (vLastControl.ControlLeft + vLastControl.CaptionWidth + vLastControl.ControlWidth)) AndAlso _
                     (vLastControl.ValidationTable.Length = 0 _
                    OrElse vLastControl.AttributeName = "action_status" OrElse vLastControl.AttributeName = "action_priority" _
                    OrElse vLastControl.AttributeName = "workstream_group_outcome") Then
                    If vControl.ControlTop.Equals(DEFAULT_TOP) AndAlso vLastControl.ControlTop > vControl.ControlTop Then
                      'This control is on a new column to the right of the last control and at the top of the page
                      vTop = DEFAULT_TOP
                      vLastTop = DEFAULT_TOP
                    Else
                      'Set this control at the same vertical position as the last one if it is visible and would fit to the right of the last control
                      vTop = vLastTop
                    End If
                  Else
                    Dim vGap As Integer = DEFAULT_GAP
                    If vLastTop = DEFAULT_TOP AndAlso pUpdateType = UpdateControlTypes.AdjustTopsForFinders Then vGap = 0
                    If vControl.ControlLeft = DEFAULT_LEFT AndAlso vNextLeftTop > vLastTop + (vLastControl.ControlHeight) + vGap Then
                      vTop = vNextLeftTop
                    Else
                      vTop = vLastTop + (vLastControl.ControlHeight) + vGap
                    End If
                  End If
                End If
                vControl.SetControlLocation(vLeft, vTop)
                If vControl.Visible AndAlso vTop + vControl.ControlHeight + DEFAULT_GAP > vNextLeftTop Then vNextLeftTop = vTop + vControl.ControlHeight + DEFAULT_GAP
                If vControl.Visible AndAlso vControl.ControlType <> "btn" Then
                  vLastControl = vControl 'Remember the last visible control
                  vLastTop = vControl.ControlTop
                End If
              End If
          End Select
          Me.Add(vControl.SequenceNumber.ToString, vControl)
        Next
      End If
      If Me.Count > 0 AndAlso vSequenceChanged Then
        pEnv.Connection.StartTransaction()
        Item(1).UpdatePageSequences(Me.Count)
        For Each vControl As CDBControl In Me
          vControl.ModifyOriginalSequence(Me.Count)
        Next
      End If

      For Each vControl As CDBControl In Me
        If pUpdateType = UpdateControlTypes.AdjustTopsForTrader AndAlso vControl.OldSequenceNumber <> 0 Then
          For vCtr As Integer = 0 To vDataSet.Tables(0).Rows.Count - 1
            If vDataSet.Tables(0).Rows(vCtr).Item("sequence_number").ToString = vControl.SequenceNumber.ToString Then
              If vControl.ControlTop = IntegerValue(vDataSet.Tables(0).Rows(vCtr).Item("control_top").ToString) Then
                If vControl.SequenceNumber > vControl.OldSequenceNumber Then
                  vControl.ControlTop = FindTopPosition(vCtr, vDataSet, False)
                Else
                  vControl.ControlTop = FindTopPosition(vCtr, vDataSet, True)
                End If
              Else
                vControl.ControlTop = IntegerValue(vDataSet.Tables(0).Rows(vCtr).Item("control_top").ToString)
              End If
            End If
          Next
        End If
        vControl.Save()
      Next
      pEnv.Connection.CommitTransaction()
    End Sub

    Public Function FindTopPosition(ByVal pRowNum As Integer, ByVal pDataSet As DataSet, ByVal pSearchUp As Boolean) As Integer
      Dim vTopPosition As Integer = IntegerValue(pDataSet.Tables(0).Rows(pRowNum).Item("control_top").ToString)
      If pSearchUp Then
        For vCtr As Integer = pRowNum To 0 Step -1
          If IntegerValue(pDataSet.Tables(0).Rows(vCtr).Item("control_top").ToString) <> vTopPosition Then
            Return IntegerValue(pDataSet.Tables(0).Rows(vCtr).Item("control_top").ToString)
          End If
        Next
      Else
        For vCtr As Integer = pRowNum To pDataSet.Tables(0).Rows.Count - 1 Step 1
          If IntegerValue(pDataSet.Tables(0).Rows(vCtr).Item("control_top").ToString) <> vTopPosition Then
            Return IntegerValue(pDataSet.Tables(0).Rows(vCtr).Item("control_top").ToString)
          End If
        Next
      End If
      Return vTopPosition
    End Function
    Public Sub SetParameterNameCasing()
      For Each vControl As CDBControl In Me
        vControl.SetParameterNameCasing()
      Next
    End Sub

    Public Sub GetCustomFormControls(ByVal pEnv As CDBEnvironment, ByVal pCustomForm As Integer)
      Dim vCustomFormControls As New CustomFormControls
      vCustomFormControls.Init(pEnv, pCustomForm)
      For Each vCustomFormControl As CustomFormControl In vCustomFormControls
        Dim vCDBControl As New CDBControl(pEnv)
        vCDBControl.InitFromCustomFormControl(vCustomFormControl)
        Me.Add(vCDBControl)
      Next
    End Sub

    Public Function TableNames() As CDBParameters
      Dim vControl As CDBControl
      Dim vTables As New CDBParameters

      For Each vControl In Me
        If Len(vControl.TableName) > 0 Then
          If Not vTables.Exists(vControl.TableName) Then vTables.Add(vControl.TableName)
        End If
      Next
      TableNames = vTables
    End Function

    Public Function WhereFields(ByVal pTableName As String, ByVal pParams As CDBParameters) As CDBFields
      Dim vWhereFields As New CDBFields
      Dim vControl As CDBControl

      For Each vControl In Me
        If vControl.TableName = pTableName AndAlso vControl.PrimaryKey = "Y" Then
          If pParams.Exists(vControl.ParameterName) Then
            vWhereFields.Add(vControl.AttributeName, vControl.Type, pParams(vControl.ParameterName).Value)
          End If
        End If
      Next
      WhereFields = vWhereFields
    End Function

    Public Function Fields(ByVal pTableName As String, ByVal pParams As CDBParameters) As CDBFields
      Dim vFields As New CDBFields
      Dim vControl As CDBControl

      For Each vControl In Me
        If vControl.TableName = pTableName Then
          If pParams.Exists(vControl.ParameterName) Then
            vFields.Add(vControl.AttributeName, vControl.Type, pParams(vControl.ParameterName).Value)
          End If
        End If
      Next
      Fields = vFields
    End Function

    Public Function DataTable(ByVal pEnv As CDBEnvironment, ByVal pForMaintenance As Boolean, ByVal pSetParameterNames As Boolean) As CDBDataTable

      Dim vControl As CDBControl
      Dim vNames As New List(Of String)
      Dim vBaseName As String
      For Each vControl In Me
        If vControl.ParameterName.Length = 0 Then
          If vControl.ControlType = "tab" Then
            vBaseName = "TAB"
          Else
            vBaseName = ProperName(vControl.AttributeName)
          End If
          Dim vCount As Integer = 1
          Dim vName As String = vBaseName
          While vNames.Contains(vName)
            vCount += 1
            vName = vBaseName & vCount.ToString
          End While
          vControl.ParameterName = vName
          vNames.Add(vName)
        End If
      Next
      If Count > 0 Then
        vControl = DirectCast(Me(1), CDBControl)
      Else
        vControl = New CDBControl(pEnv)
        vControl.Init()
      End If
      Dim vTable As New CDBDataTable
      vControl.AddColumnsForDataTable(vTable, pForMaintenance, mvMultiplePages)
      For Each vControl In Me
        vControl.AddRowForDataTable(vTable)
      Next
      Return vTable
    End Function

    Public Overloads ReadOnly Property Item(ByVal pIndexKey As Object) As CDBControl
      Get
        Return DirectCast(Me(pIndexKey.ToString), CDBControl)
      End Get
    End Property

    Public Overloads Sub Remove(ByVal pIndexKey As Integer)
      MyBase.Remove(pIndexKey)
    End Sub
    Public Function Exists(ByVal pIndexKey As Object) As Boolean
      Return MyBase.ContainsKey(pIndexKey.ToString)
    End Function

    Public Shared Sub UpgradeCustomizedPage(pEnv As CDBEnvironment, pPageTypeCode As String, pPageNumber As Integer)

      Dim vLastSeqNumWhere As New CDBFields()
      vLastSeqNumWhere.Add("fp_application", pPageNumber)
      vLastSeqNumWhere.Add("fp_page_type", pPageTypeCode)
      Dim vLastSeqNumSQL As New SQLStatement(pEnv.Connection, "MAX(sequence_number)", "fp_controls", vLastSeqNumWhere)
      Dim vSQLRtn As String = pEnv.Connection.GetValue(vLastSeqNumSQL.SQL)
      Dim pLastSequenceNumber As Integer = 1
      Integer.TryParse(vSQLRtn, pLastSequenceNumber)

      Dim vCustomizedControlsWhere As New CDBFields()
      vCustomizedControlsWhere.Add("customizedControls.fp_application", pPageNumber)
      vCustomizedControlsWhere.AddJoin("customizedControls.fp_page_type", "systemControls.fp_page_type")
      vCustomizedControlsWhere.AddJoin("customizedControls.table_name", "systemControls.table_name")
      vCustomizedControlsWhere.AddJoin("customizedControls.attribute_name", "systemControls.attribute_name")
      vCustomizedControlsWhere.AddJoin("customizedControls.control_type", "systemControls.control_type")
      vCustomizedControlsWhere.AddJoin("customizedControls.parameter_row_num", "systemControls.parameter_row_num")
      Dim vCustomizedControlsSQL As New SQLStatement(pEnv.Connection, "*", "VFPControlsUniqueParam customizedControls", vCustomizedControlsWhere)

      Dim vMissingControlsWhere As New CDBFields()
      vMissingControlsWhere.Add("fp_application")
      vMissingControlsWhere.Add("fp_page_type", pPageTypeCode)
      vMissingControlsWhere.Add("Exclude", vCustomizedControlsSQL.SQL, CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoExist)
      Dim vMissingControlsSQL As New SQLStatement(pEnv.Connection, "fp_page_type, table_name, attribute_name, control_type, fp_application, sequence_number", "VFPControlsUniqueParam systemControls", vMissingControlsWhere, "systemControls.sequence_number")

      Dim vMissingControlsData As New CDBDataTable
      vMissingControlsData.Timeout = 300 ''This was taking 40 seconds to run on a QA Db.  As it's DatabaseUpgrade, I'm setting a high value to ensure it always executes
      vMissingControlsData.FillFromSQL(pEnv, vMissingControlsSQL)

      For Each vControlRow As CDBDataRow In vMissingControlsData.Rows
        Dim vControlWhere As New CDBFields()
        For Each vControlColumn As CDBDataColumn In vMissingControlsData.Columns
          vControlWhere.Add(vControlColumn.AttributeName, vControlRow.Item(vControlColumn.Name).ToString())
        Next
        Dim vMissingControl As New CDBControl(pEnv)
        vMissingControl.InitWithPrimaryKey(vControlWhere)
        If vMissingControl.Existing Then
          pLastSequenceNumber = pLastSequenceNumber + 1
          vMissingControl.SaveFPControlAs(pPageNumber.ToString(), pPageTypeCode, pLastSequenceNumber, True)
        End If
      Next
    End Sub

  End Class
End Namespace
