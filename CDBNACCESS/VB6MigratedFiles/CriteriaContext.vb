

Namespace Access
  Public Class CriteriaContext

    Public Enum SaveTypes
      stUpdateCounted = 1
      stUpdateBrackets = 2
      stUpdateAll = 4
      stInsert = 8
    End Enum

    Private mvEnv As CDBEnvironment
    Private mvConn As CDBConnection
    Private mvProcessed As Boolean
    Private mvID As Integer
    'Criteria Set Details
    Private mvCriteriaSet As Integer
    Private mvSequenceNumber As Integer
    Private mvSearchArea As String
    Private mvCorO As String
    Private mvIorE As String
    Private mvMainValue As String
    Private mvMainValueHeading As String
    Private mvSubValue As String
    Private mvSubValueHeading As String
    Private mvMainDataType As String
    Private mvSubDataType As String
    Private mvPeriod As String
    Private mvCounted As Integer
    Private mvAndOr As String
    Private mvLeftParenthesis As String
    Private mvRightParenthesis As String
    'Selection Control
    Private mvValidationTable As String
    Private mvMainAttribute As String
    Private mvMainValidationAttribute As String
    Private mvSubAttribute As String
    Private mvSubValidationAttribute As String
    Private mvToAttribute As String
    Private mvFromAttribute As String
    Private mvIndexed As Boolean
    Private mvSpecial As Boolean
    Private mvTableName As String
    Private mvAddressLink As Boolean
    Private mvGeographic As Boolean
    Private mvNullsAllowed As Boolean
    'Maintenance Attributes
    Private mvPattern As String
    Friend Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pConn As CDBConnection, ByRef pCC As CriteriaContext)
      mvEnv = pEnv
      mvConn = pConn
      mvProcessed = pCC.Processed
      mvID = pCC.ID
      'Criteria Set Details
      mvCriteriaSet = pCC.CriteriaSet
      mvSequenceNumber = pCC.SequenceNumber
      mvSearchArea = pCC.SearchArea
      mvCorO = pCC.CorO
      mvIorE = pCC.IorE
      mvMainValue = pCC.MainValue
      mvSubValue = pCC.SubValue
      mvPeriod = pCC.Period
      mvCounted = pCC.Counted
      mvAndOr = pCC.AndOr
      mvLeftParenthesis = pCC.LeftParenthesis
      mvRightParenthesis = pCC.RightParenthesis
      'Selection Control
      mvValidationTable = pCC.ValidationTable
      mvMainAttribute = pCC.MainAttribute
      mvMainDataType = pCC.MainDataType
      mvMainValueHeading = pCC.MainValueHeading
      mvMainValidationAttribute = pCC.MainValidationAttribute
      mvSubAttribute = pCC.SubAttribute
      mvSubDataType = pCC.SubDataType
      mvSubValueHeading = pCC.SubValueHeading
      mvSubValidationAttribute = pCC.SubValidationAttribute
      mvToAttribute = pCC.ToAttribute
      mvFromAttribute = pCC.FromAttribute
      mvIndexed = pCC.Indexed
      mvSpecial = pCC.Special
      mvTableName = pCC.TableName
      mvAddressLink = pCC.AddressLink
      mvGeographic = pCC.Geographic
      mvNullsAllowed = pCC.NullsAllowed
    End Sub
    Friend Property CriteriaSet() As Integer
      Get
        CriteriaSet = mvCriteriaSet
      End Get
      Set(ByVal Value As Integer)
        mvCriteriaSet = Value
      End Set
    End Property
    Friend ReadOnly Property AddressLink() As Boolean
      Get
        AddressLink = mvAddressLink
      End Get
    End Property
    Friend ReadOnly Property ContactOrOrgValue() As String
      Get
        ContactOrOrgValue = mvCorO
      End Get
    End Property
    Friend Property Contacts() As Boolean
      Get
        Contacts = mvCorO = "C"
      End Get
      Set(ByVal Value As Boolean)
        If Value Then
          mvCorO = "C"
        Else
          mvCorO = "O"
        End If
      End Set
    End Property
    Friend ReadOnly Property Geographic() As Boolean
      Get
        Geographic = mvGeographic
      End Get
    End Property
    Friend ReadOnly Property ID() As Integer
      Get
        ID = mvID
      End Get
    End Property
    Friend Property Include() As Boolean
      Get
        Include = mvIorE = "I"
      End Get
      Set(ByVal Value As Boolean)
        If Value Then
          mvIorE = "I"
        Else
          mvIorE = "E"
        End If
      End Set
    End Property
    Friend ReadOnly Property Indexed() As Boolean
      Get
        Indexed = mvIndexed
      End Get
    End Property
    Friend ReadOnly Property FromAttribute() As String
      Get
        FromAttribute = mvFromAttribute
      End Get
    End Property
    Friend ReadOnly Property Pattern() As String
      Get
        Pattern = mvPattern
      End Get
    End Property

    Friend ReadOnly Property ToAttribute() As String
      Get
        ToAttribute = mvToAttribute
      End Get
    End Property
    Friend ReadOnly Property MainAttribute() As String
      Get
        MainAttribute = mvMainAttribute
      End Get
    End Property
    Friend ReadOnly Property MainDataType() As String
      Get
        MainDataType = mvMainDataType
      End Get
    End Property
    Friend ReadOnly Property MainValue() As String
      Get
        MainValue = mvMainValue
      End Get
    End Property
    Friend ReadOnly Property MainValueHeading() As String
      Get
        MainValueHeading = mvMainValueHeading
      End Get
    End Property
    Friend ReadOnly Property MainValidationAttribute() As String
      Get
        MainValidationAttribute = mvMainValidationAttribute
      End Get
    End Property
    Friend ReadOnly Property NullsAllowed() As Boolean
      Get
        NullsAllowed = mvNullsAllowed
      End Get
    End Property
    Friend ReadOnly Property Period() As String
      Get
        Period = mvPeriod
      End Get
    End Property
    Friend Property Processed() As Boolean
      Get
        Processed = mvProcessed
      End Get
      Set(ByVal Value As Boolean)
        mvProcessed = Value
      End Set
    End Property
    Friend ReadOnly Property SearchArea() As String
      Get
        SearchArea = mvSearchArea
      End Get
    End Property
    Friend ReadOnly Property SequenceNumber() As Integer
      Get
        SequenceNumber = mvSequenceNumber
      End Get
    End Property
    Friend ReadOnly Property Special() As Boolean
      Get
        Special = mvSpecial
      End Get
    End Property
    Friend ReadOnly Property SubAttribute() As String
      Get
        SubAttribute = mvSubAttribute
      End Get
    End Property
    Friend ReadOnly Property SubDataType() As String
      Get
        SubDataType = mvSubDataType
      End Get
    End Property
    Friend ReadOnly Property SubValidationAttribute() As String
      Get
        SubValidationAttribute = mvSubValidationAttribute
      End Get
    End Property
    Friend ReadOnly Property SubValue() As String
      Get
        SubValue = mvSubValue
      End Get
    End Property
    Friend ReadOnly Property SubValueHeading() As String
      Get
        SubValueHeading = mvSubValueHeading
      End Get
    End Property
    Friend ReadOnly Property TableName() As String
      Get
        TableName = mvTableName
      End Get
    End Property
    Friend ReadOnly Property ValidationTable() As String
      Get
        ValidationTable = mvValidationTable
      End Get
    End Property
    Friend Property Counted() As Integer
      Get
        Counted = mvCounted
      End Get
      Set(ByVal Value As Integer)
        mvCounted = Value
      End Set
    End Property
    Friend Property AndOr() As String
      Get
        AndOr = mvAndOr
      End Get
      Set(ByVal Value As String)
        mvAndOr = Value
      End Set
    End Property
    Friend ReadOnly Property CorO() As String
      Get
        CorO = mvCorO
      End Get
    End Property
    Friend ReadOnly Property IorE() As String
      Get
        IorE = mvIorE
      End Get
    End Property
    Friend Property LeftParenthesis() As String
      Get
        LeftParenthesis = mvLeftParenthesis
      End Get
      Set(ByVal Value As String)
        mvLeftParenthesis = Value
      End Set
    End Property
    Friend Property RightParenthesis() As String
      Get
        RightParenthesis = mvRightParenthesis
      End Get
      Set(ByVal Value As String)
        mvRightParenthesis = Value
      End Set
    End Property
    Friend Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pConn As CDBConnection, ByVal pRecordSet As CDBRecordSet, ByRef pID As Integer)
      mvEnv = pEnv
      mvConn = pConn
      With pRecordSet.Fields
        mvCriteriaSet = .Item("criteria_set").IntegerValue
        mvSearchArea = .Item("search_area").Value
        mvCorO = .Item("c_o").Value
        mvIorE = .Item("i_e").Value
        mvMainValue = .Item("main_value").Value
        mvMainValueHeading = .Item("main_value_heading").Value
        mvSubValue = .Item("subsidiary_value").Value
        mvSubValueHeading = .Item("subsidiary_value_heading").Value
        mvMainDataType = .Item("main_data_type").Value
        mvSubDataType = .Item("subsidiary_data_type").Value
        mvPeriod = .Item("period").Value
        mvCounted = .Item("counted").IntegerValue
        mvAndOr = .Item("and_or").Value
        mvLeftParenthesis = .Item("left_parenthesis").Value
        mvRightParenthesis = .Item("right_parenthesis").Value

        mvValidationTable = .Item("validation_table").Value
        mvMainAttribute = .Item("main_attribute").Value
        mvMainValidationAttribute = .Item("main_validation_attribute").Value
        mvSubAttribute = .Item("subsidiary_attribute").Value
        mvSubValidationAttribute = .Item(pConn.DBAttrName("subsidiary_validation_attribute")).Value
        mvToAttribute = .Item("to_attribute").Value
        mvFromAttribute = .Item("from_attribute").Value
        mvIndexed = .Item("indexed").Bool
        mvSpecial = .Item("special").Bool
        mvTableName = .Item("table_name").Value
        mvAddressLink = .Item("address_link").Bool
        mvGeographic = .Item("geographic").Bool
        mvSequenceNumber = .Item("sequence_number").IntegerValue
        mvNullsAllowed = .Item("nulls_allowed").Bool
        If .Exists("pattern") Then mvPattern = .Item("pattern").Value
      End With
      mvProcessed = False
      mvID = pID
    End Sub
    Friend Sub Save(ByVal pType As SaveTypes)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields

      With vUpdateFields
        If pType = SaveTypes.stInsert Or pType = SaveTypes.stUpdateAll Or pType = SaveTypes.stUpdateCounted Then .Add("counted", CDBField.FieldTypes.cftLong, mvCounted)
        If pType = SaveTypes.stInsert Or pType = SaveTypes.stUpdateAll Or pType = SaveTypes.stUpdateBrackets Then
          .Add("and_or", CDBField.FieldTypes.cftCharacter, mvAndOr)
          .Add("left_parenthesis", CDBField.FieldTypes.cftCharacter, mvLeftParenthesis)
          .Add("right_parenthesis", CDBField.FieldTypes.cftCharacter, mvRightParenthesis)
        End If
        .Add("amended_by", CDBField.FieldTypes.cftCharacter, mvEnv.User.Logname)
        .Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
      End With
      If pType = SaveTypes.stInsert Then
        With vUpdateFields
          .Add("criteria_set", CDBField.FieldTypes.cftLong, mvCriteriaSet)
          .Add("sequence_number", CDBField.FieldTypes.cftInteger, mvSequenceNumber)
          .Add("search_area", CDBField.FieldTypes.cftCharacter, mvSearchArea)
          .Add("i_e", CDBField.FieldTypes.cftCharacter, mvIorE)
          .Add("c_o", CDBField.FieldTypes.cftCharacter, mvCorO)
          .Add("main_value", CDBField.FieldTypes.cftUnicode, mvMainValue)
          .Add("subsidiary_value", CDBField.FieldTypes.cftUnicode, mvSubValue)
          .Add("period", CDBField.FieldTypes.cftCharacter, mvPeriod)
        End With
        mvEnv.Connection.InsertRecord("criteria_set_details", vUpdateFields)
      Else
        With vWhereFields
          .Add("criteria_set", CDBField.FieldTypes.cftLong, mvCriteriaSet)
          .Add("sequence_number", CDBField.FieldTypes.cftInteger, mvSequenceNumber)
        End With
        mvEnv.Connection.UpdateRecords("criteria_set_details", vUpdateFields, vWhereFields)
      End If
    End Sub
  End Class
End Namespace
