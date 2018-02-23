Namespace Access
  Public Class UpgradeAttribute

    Public Enum NullOptions
      noNullsInvalid = 0
      noNullsAllowed
    End Enum

    Public Key As String
    Private mvDataType As String
    Private mvStructureModified As Boolean
    Private mvParameter1 As Integer
    Private mvParameter2 As Integer
    Private mvParameter3 As Integer
    Private mvDefaultValue As String
    Private mvBeforeAttribute As String
    Private mvRangeCheck As String
    Private mvEmpressRange As String
    Private mvRangeCheckOnly As Boolean
    Private mvToBeCreated As Boolean
    Private mvToBeDeleted As Boolean
    Private mvNullable As NullOptions
    Private mvParameterCount As Integer

    Public Property Nullable() As NullOptions
      Get
        Return mvNullable
      End Get
      Set(ByVal value As NullOptions)
        mvNullable = value
      End Set
    End Property

    Public Property ToBeDeleted() As Boolean
      Get
        Return mvToBeDeleted
      End Get
      Set(ByVal value As Boolean)
        mvToBeDeleted = value
      End Set
    End Property

    Public Property ToBeCreated() As Boolean
      Get
        Return mvToBeCreated
      End Get
      Set(ByVal value As Boolean)
        mvToBeCreated = value
      End Set
    End Property

    Public Property RangeCheckOnly() As Boolean
      Get
        Return mvRangeCheckOnly
      End Get
      Set(ByVal value As Boolean)
        mvRangeCheckOnly = value
      End Set
    End Property

    Public Property RangeCheck() As String
      Get
        Return mvRangeCheck
      End Get
      Set(ByVal value As String)
        mvRangeCheck = value
      End Set
    End Property

    Public Property EmpressRange() As String
      Get
        Return mvEmpressRange
      End Get
      Set(ByVal value As String)
        mvEmpressRange = value
      End Set
    End Property

    Public Property BeforeAttribute() As String
      Get
        Return mvBeforeAttribute
      End Get
      Set(ByVal value As String)
        mvBeforeAttribute = value
      End Set
    End Property

    Public Property DefaultValue() As String
      Get
        Return mvDefaultValue
      End Get
      Set(ByVal value As String)
        mvDefaultValue = value
      End Set
    End Property

    Public Property Parameter3() As Integer
      Get
        Return mvParameter3
      End Get
      Set(ByVal value As Integer)
        mvParameter3 = value
      End Set
    End Property

    Public Property Parameter2() As Integer
      Get
        Return mvParameter2
      End Get
      Set(ByVal value As Integer)
        mvParameter2 = value
      End Set
    End Property

    Public Property Parameter1() As Integer
      Get
        Return mvParameter1
      End Get
      Set(ByVal value As Integer)
        mvParameter1 = value
      End Set
    End Property

    Public Property ParameterCount() As Integer
      Get
        Return mvParameterCount
      End Get
      Set(ByVal value As Integer)
        mvParameterCount = value
      End Set
    End Property

    Public Property StructureModified() As Boolean
      Get
        Return mvStructureModified
      End Get
      Set(ByVal value As Boolean)
        mvStructureModified = value
      End Set
    End Property

    Public Property DataType() As String
      Get
        Return mvDataType
      End Get
      Set(ByVal value As String)
        mvDataType = value
      End Set
    End Property

    Public Function Rangechanged(ByVal pRange As String) As Boolean
      Dim vRange As String
      Dim vPos As Integer
      Dim vFrom As String
      Dim vTo As String

      'Try and turn the range into what empress would store internally
      vPos = pRange.IndexOf("to")
      If vPos > 0 Then
        vFrom = Left(pRange, vPos - 1).Trim
        vTo = Mid(pRange, vPos + 2).Trim
        vRange = "range '" + vFrom + "' i '" + vTo + "' i"

      Else
        vRange = "smatch '["
        vRange = vRange + pRange.Replace("|", "") + "]'"
      End If
      If vRange <> mvEmpressRange Then Rangechanged = True
    End Function

    Public Function Alter(ByVal pConn As CDBConnection, ByVal pDataType As String, ByVal pDTP1 As String, ByVal pDTP2 As String, ByVal pDTP3 As String, ByVal pNullsAllowed As String, ByVal pDefaultValue As String, ByRef pDoAlter As Boolean, Optional ByVal pRemoveUnicode As Boolean = False) As Boolean
      Dim vTestDataType As String
      Dim vParmCount As Integer
      Dim vTestNulls As NullOptions
      Dim vError As Boolean

      pDoAlter = True
      vTestDataType = pDataType
      DBSetup.GetNativeDataType(pConn, LCase(Key), vTestDataType, vParmCount)
      If vTestDataType = "char" OrElse vTestDataType = "nlschar" Then
        vTestDataType = vTestDataType.Replace("char", "character")
      ElseIf Left(vTestDataType, 8) = "varchar2" Then
        If Mid(vTestDataType, 9, 1) = "(" Then pDTP1 = Math.Round(Val(Mid(vTestDataType, 10))).ToString
        vTestDataType = "varchar2"
      ElseIf Not vTestDataType.Equals("varchar(max)", StringComparison.InvariantCultureIgnoreCase) AndAlso
             Left(vTestDataType, 7) = "varchar" Then
        If Mid(vTestDataType, 8, 1) = "(" Then pDTP1 = Math.Round(Val(Mid(vTestDataType, 9))).ToString
        vTestDataType = "varchar"
      ElseIf vTestDataType = "decimal" And DataType = "numeric" And pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer Then
        vTestDataType = "numeric"
      End If

      If pNullsAllowed = "NULLS" Then
        vTestNulls = NullOptions.noNullsAllowed
      Else
        vTestNulls = NullOptions.noNullsInvalid
      End If

      If DataType.ToUpper = vTestDataType.ToUpper Then 'UCase$ used because Oracle returns all uppercase values
        Select Case vTestDataType
          Case "date", "time", "datetime", "int", "nlstext", "text", "integer", "longinteger", "smallint", "clob", "varchar(max)"
            If Nullable = vTestNulls Then pDoAlter = False
          Case "char", "character", "varchar", "varchar2", "nlschar", "nlscharacter", "nvarchar"
            If Parameter1 = IntegerValue(pDTP1) And Nullable = vTestNulls Then pDoAlter = pRemoveUnicode
          Case "decimal", "number", "numeric"
            If Parameter1 = IntegerValue(pDTP1) And Parameter2 = IntegerValue(pDTP2) And Nullable = vTestNulls Then pDoAlter = False
        End Select
      ElseIf DataType = "character" And vTestDataType = "nlscharacter" Then
        'Change empress attribute from char to nlschar
      ElseIf DataType = "nlscharacter" And vTestDataType = "nlstext" Then
        'Change empress attribute from nlschar to text
      ElseIf DataType = "integer" And vTestDataType = "longinteger" Then
        'Change empress attribute from integer to longinteger
      ElseIf DataType = "smallint" And vTestDataType = "integer" Then
        'Change attribute from smallint to integer
      ElseIf DataType = "smallint" And vTestDataType = "int" Then
        'Change sql server attribute from smallint to int
      ElseIf DataType = "nlstext" And vTestDataType = "nlscharacter" Then
        'Change empress attribute from nlstext to nlschar
        'Could truncate data so only allow in specific cases
        If Key <> "mailmerge_header_desc" Then vError = True
      ElseIf DataType = "varchar" And vTestDataType = "text" Then
        'Change sql server attribute from varchar to text
      ElseIf DataType = "varchar2" And vTestDataType = "long" Then
        'Change oracle attribute from varchar2 to text
      ElseIf DataType = "LONG" And vTestDataType = "varchar2" Then
        'Change oracle attribute from long to varchar2
        'Could truncate data so only allow in specific cases
        Select Case Key
          Case "main_value", "subsidiary_value", "period"
            'allow
          Case Else
            vError = True
        End Select
      ElseIf DataType = "text" And vTestDataType = "varchar" Then
        'Change sql server attribute from text to varchar
        'Could truncate data so only allow in specific cases
        Select Case Key
          Case "main_value", "subsidiary_value", "period", "account_name", "left_parenthesis", "right_parenthesis", _
               "question_text", "answer_text", "help_text", "response_answer_text"
            'allow
          Case Else
            vError = True
        End Select
      ElseIf DataType = "varchar" And vTestDataType = "nvarchar" Then
        'Change sql server attribute from varchar to nvarchar
      ElseIf DataType = "text" And vTestDataType = "nvarchar" Then
        'Change sql server attribute from text to nvarchar
      ElseIf DataType = "nvarchar" And vTestDataType = "varchar" Then
        'This will do nothing
        pDoAlter = pRemoveUnicode
      ElseIf (DataType = "int" AndAlso vTestDataType = "decimal") OrElse (DataType = "integer" AndAlso vTestDataType = "number") Then
        Select Case Key
          Case "cpd_points", "cpd_points_2",
            "raw_mark", "original_mark", "moderated_mark", "total_mark",
            "current_mark", "previous_mark",
            "previous_raw_mark", "previous_original_mark", "previous_moderated_mark", "previous_mark",
            "exemption_mark"
            'Change from Integer to Decimal. Allowed in SQL Server. For Oracle, ConvertOracleTables will be used in DatabaseUpgrade class
            If pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsOracle Then pDoAlter = False
          Case Else
            vError = True
        End Select
      ElseIf (DataType = "int" AndAlso vTestDataType = "varchar") OrElse (DataType = "integer" AndAlso vTestDataType = "varchar2") Then
        'Changing from int to varchar - could truncate data so only allow in specific circumstances
        If Key = "bacs_user_number" OrElse Key = "authorised_transaction_number" Then
          pDoAlter = (pConn.RDBMSType = CDBConnection.RDBMSTypes.rdbmsSqlServer)  'Allowed in SQL Server. For Oracle, ConvertOracleTables will be used in DatabaseUpgrade class as this cannnot be changed if the attriute contains data
        Else
          vError = True
        End If
      Else
        vError = True
      End If

      If pDoAlter And Not vError Then
        DataType = vTestDataType
        Parameter1 = IntegerValue(pDTP1)
        Parameter2 = IntegerValue(pDTP2)
        Parameter3 = IntegerValue(pDTP3)
        Nullable = vTestNulls
        DefaultValue = pDefaultValue
        StructureModified = True
        ParameterCount = vParmCount
      End If
      Alter = vError
    End Function

  End Class
End Namespace
