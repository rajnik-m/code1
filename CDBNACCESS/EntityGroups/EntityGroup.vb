Namespace Access

  Public MustInherit Class EntityGroup
    Inherits CARERecord

#Region "Non-AutoGenerated Code"

    Public Enum EntityGroupTypes
      egtUnknown = 0
      egtContact = 1                        'Keep same as contact class
      egtOrganisation = 2                   'Keep same as contact class
      egtEvent
    End Enum

    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub
    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public MustOverride ReadOnly Property EntityGroupCode() As String
    Public MustOverride ReadOnly Property EntityGroupDesc() As String
    Public MustOverride ReadOnly Property Client() As String
    Public MustOverride ReadOnly Property Name() As String
    Public MustOverride ReadOnly Property SequenceNo() As Integer
    Public MustOverride ReadOnly Property RgbValue() As Integer
    Public MustOverride ReadOnly Property TabPrefix() As String
    Public MustOverride ReadOnly Property HiddenAttributes() As String
    Public MustOverride ReadOnly Property NamedAttributes() As String
    Public MustOverride ReadOnly Property AmendedBy() As String
    Public MustOverride ReadOnly Property AmendedOn() As String
    Public MustOverride ReadOnly Property EntityGroupType() As EntityGroupTypes
    Public MustOverride ReadOnly Property DefaultGroup() As Boolean
    Public MustOverride ReadOnly Property NameFormat As String
    Public MustOverride ReadOnly Property LastUsedId As Integer

    Public MustOverride Property UnknownAddress() As String
    Public MustOverride Property UnknownTown() As String
    Public MustOverride ReadOnly Property AllAddressesUnknown() As Boolean
    Public MustOverride ReadOnly Property PrimaryRelationship() As String
    Public MustOverride ReadOnly Property UseEventPricingMatrix() As Boolean
    Public MustOverride ReadOnly Property PositionActivityPrompt() As Boolean
    Public MustOverride ReadOnly Property PositionRelationshipPrompt() As Boolean
    Public MustOverride ReadOnly Property ViewOrganisationInContactCard() As Boolean
    Public MustOverride ReadOnly Property OrganisationNumber() As Integer

    Public Sub GetColorPreference()
      'Dim vReader As INIReader = New INIReader
      'mvClassFields.Item("rgb_value").IntegerValue = vReader.ReadInteger("TABS", EntityGroupDesc & "TabColor", RgbValue)
    End Sub

    Public Function GetAttributeTable(ByVal pNoTabs As Boolean) As CDBDataTable
      Dim vTable As New CDBDataTable
      vTable.AddColumnsFromList("Attribute,AttributeName")
      Dim vParams As New CDBParameters
      vParams.InitKeysFromUniqueList(HiddenAttributes)                    'First add the hidden attributes into the list
      Dim vNamedAttributes As String() = NamedAttributes.Split(",".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
      Dim vAttrName As String
      For Each vItem As String In vNamedAttributes
        Dim vPos As Integer
        vPos = (vItem.IndexOf("=", 0) + 1)
        If vPos > 0 Then
          vAttrName = vItem.Substring(0, vPos - 1)
          If vAttrName.Length > 0 AndAlso Not vParams.Exists(vAttrName) Then
            vParams.Add(vAttrName, vItem.Substring(vPos))
          End If
        End If
      Next
      vParams.Add("GroupName", Name)
      vParams.Add("GroupDescription", EntityGroupDesc)
      vParams.Add("UnknownAddress", UnknownAddress)
      vParams.Add("UnknownTown", UnknownTown)
      vParams.Add("AllAddressesUnknown", IIf(AllAddressesUnknown, "Y", "N").ToString)
      vParams.Add("RGBValue", CDBField.FieldTypes.cftLong, RgbValue.ToString)
      vParams.Add("PrimaryRelationship", PrimaryRelationship)
      vParams.Add("UseEventPricingMatrix", IIf(UseEventPricingMatrix, "Y", "N").ToString)
      vParams.Add("PositionActivityPrompt", IIf(PositionActivityPrompt, "Y", "N").ToString)
      vParams.Add("PositionRelationshipPrompt", IIf(PositionRelationshipPrompt, "Y", "N").ToString)
      vParams.Add("ViewOrganisationInContactCard", BooleanString(ViewOrganisationInContactCard))
      vParams.Add("OrganisationNumber", CDBField.FieldTypes.cftLong, OrganisationNumber.ToString)
      If pNoTabs = False Then
        'vDefaultGroup = mvEnv.EntityGroups.DefaultGroup(EntityGroupType)
        'If vDefaultGroup.Tabs.Count = 0 Then vDefaultGroup.LoadTabSet(vDefaultGroup, vColl)
        'If Tabs.Count = 0 Then LoadTabSet(vDefaultGroup, vColl)
        'For Each vEntityTab In Tabs
        '  vParams.Add("Tab" & vEntityTab.TabType, cftCharacter, vEntityTab.Heading)
        'Next
      End If
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("report_code", EntityGroupCode & "*", CDBField.FieldWhereOperators.fwoLikeOrEqual)
      vWhereFields.Add("client", mvEnv.ClientCode, CDBField.FieldWhereOperators.fwoOpenBracket)
      vWhereFields.Add("client#1", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      Dim vRecordSet As CDBRecordSet = New SQLStatement(mvEnv.Connection, "report_name,report_number", "reports", vWhereFields).GetRecordSet
      While vRecordSet.Fetch()
        vParams.Add("Report_" & vRecordSet.Fields("report_number").Value, vRecordSet.Fields("report_name").Value)
      End While
      vRecordSet.CloseRecordSet()
      For Each vParam As CDBParameter In vParams
        vTable.AddRowFromItems(vParam.Name, vParam.Value)
      Next
      Return vTable
    End Function

    Public ReadOnly Property CanUseUnknownAddress() As Boolean
      Get
        If UnknownAddress.Length > 0 And UnknownTown.Length > 0 Then
          Return True
        End If
      End Get
    End Property

    Public Sub CheckUnknownAddress()
      If UnknownAddress.Length = 0 Then UnknownAddress = "Unknown Address"
      If UnknownTown.Length = 0 Then UnknownTown = "Unknown Town"
    End Sub

    Public Function GetDefaultName() As String

      Dim vResult As String = ""
      Try
        If NameFormat.Length = 0 Then
          vResult = ""
        Else
          Dim vPreFix As String = NameFormat.Substring(0, NameFormat.IndexOf("["))
          Dim vSuffix As String = NameFormat.Substring(NameFormat.LastIndexOf("]") + 1, NameFormat.Length - (NameFormat.LastIndexOf("]") + 1))
          Dim vYear As String = NameFormat.Substring(NameFormat.IndexOf("[") + 1, NameFormat.IndexOf("]") - (NameFormat.IndexOf("[") + 1))
          Dim vNumber As String = NameFormat.Substring(NameFormat.LastIndexOf("[") + 1, NameFormat.LastIndexOf("]") - (NameFormat.LastIndexOf("[") + 1))
          Dim vNewNumber As Integer = GetNextId(NameFormat, LastUsedId.ToString, vYear, vNumber)
          If vNewNumber = 0 Then Return ""
          Dim vEntityGroups As EntityGroups = New EntityGroups(mvEnv)
          If vNewNumber < vEntityGroups(EntityGroupCode).GetNextId(NameFormat, vEntityGroups(EntityGroupCode).LastUsedId.ToString, vYear, vNumber) Then
            vEntityGroups(EntityGroupCode).GetDefaultName()
          Else
            Dim vUpdateParam As New CDBParameters
            vUpdateParam.Add("LastUsedId", vNewNumber)
            Update(vUpdateParam)
            Save()
            vResult = vPreFix + vNewNumber.ToString + vSuffix
          End If
        End If
        Return vResult
      Catch ex As ArgumentOutOfRangeException
        'Invalid format return empty string as the name cannot be generated
        Return vResult
      End Try
    End Function

    Public Function GetNextId(ByVal pNameFormat As String, ByVal pLastNumber As String, ByVal pYear As String, ByVal pNumber As String) As Integer
      Dim vReturn As Integer
      Dim vFormatNumber As String = "00000000000000"
      Try

        Dim vLastYear As String = ""
        If CInt(pLastNumber) > 0 Then
          vLastYear = pLastNumber.Substring(0, pYear.Length)
        Else
          If pYear.Length < 2 OrElse pYear.Length > 4 OrElse pYear.Length = 3 Then Return 0
        End If

        Dim vCurrentYear As String = DateTime.Today.ToString(pYear.ToLower())
        Dim vAppNumber As Integer = 0
        If pLastNumber.Length > 0 Then vAppNumber = CInt(pLastNumber)

        If vLastYear.Length > 0 AndAlso CInt(vLastYear) = CInt(vCurrentYear) Then
          vAppNumber = vAppNumber + 1
          vReturn = CInt(FormatNumber(vAppNumber, , TriState.True, , ))
        Else
          vAppNumber = 1
          vReturn = CInt(vCurrentYear + Format(vAppNumber, vFormatNumber.Substring(0, pNumber.Length)))

        End If
        Return vReturn
      Catch ex As ArgumentOutOfRangeException
        Throw ex
      End Try
    End Function

#End Region

  End Class
End Namespace
