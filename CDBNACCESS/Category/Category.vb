Namespace Access

  Public Class Category
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum CategoryFields
      AllFields = 0
      CategoryId
      Activity
      ActivityValue
      Quantity
      Source
      ValidFrom
      ValidTo
      ActivityDate
      Notes
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("category_id", CDBField.FieldTypes.cftInteger)
        .Add("activity")
        .Add("activity_value")
        .Add("quantity", CDBField.FieldTypes.cftNumeric)
        .Add("source")
        .Add("valid_from", CDBField.FieldTypes.cftDate)
        .Add("valid_to", CDBField.FieldTypes.cftDate)
        .Add("activity_date", CDBField.FieldTypes.cftDate)
        .Add("notes", CDBField.FieldTypes.cftMemo)

        .Item(CategoryFields.CategoryId).PrimaryKey = True
        .Item(CategoryFields.CategoryId).PrefixRequired = True

        .Item(CategoryFields.Activity).PrefixRequired = True
        .Item(CategoryFields.ActivityValue).PrefixRequired = True
        .Item(CategoryFields.Source).PrefixRequired = True
        .SetControlNumberField(CategoryFields.CategoryId, "CTI")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "cat"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "categories"
      End Get
    End Property

    Public Overrides Sub AddDeleteCheckItems()
      AddCascadeDeleteItem("category_links", "category_id")
    End Sub
'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property CategoryId() As Integer
      Get
        Return mvClassFields(CategoryFields.CategoryId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property Activity() As String
      Get
        Return mvClassFields(CategoryFields.Activity).Value
      End Get
    End Property
    Public ReadOnly Property ActivityValue() As String
      Get
        Return mvClassFields(CategoryFields.ActivityValue).Value
      End Get
    End Property
    Public ReadOnly Property Quantity() As Double
      Get
        Return mvClassFields(CategoryFields.Quantity).DoubleValue
      End Get
    End Property
    Public ReadOnly Property Source() As String
      Get
        Return mvClassFields(CategoryFields.Source).Value
      End Get
    End Property
    Public ReadOnly Property ValidFrom() As String
      Get
        Return mvClassFields(CategoryFields.ValidFrom).Value
      End Get
    End Property
    Public ReadOnly Property ValidTo() As String
      Get
        Return mvClassFields(CategoryFields.ValidTo).Value
      End Get
    End Property
    Public ReadOnly Property ActivityDate() As String
      Get
        Return mvClassFields(CategoryFields.ActivityDate).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(CategoryFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(CategoryFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(CategoryFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non Auto Generated"

    ''' <summary>
    ''' What has happened to a new Contact/Organisation Catergory
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum CategoryUpdated
      Insert = 0
      Update = 1
      Duplicate = 2
    End Enum

    Public Enum ActivityEntryStyles     'Make these bit wise values
      aesNormal = 0
      aesAllowMultipleSource = 1
      aesCheckDateRange = 2
      aesIgnoreExisting = 4
      aesPositionActivity = 8
      aesCreateJournal = 16
      aesSmartClient = 32
      aesCarePortal = 64
      aesForceAmendmentHistory = 128   'BR17834
    End Enum

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pParams"></param>
    ''' <param name="pStyle"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetExistingCategories(ByVal pParams As CDBParameters, pStyle As Category.ActivityEntryStyles) As DataTable
      Return GetExistingCategories(pParams, Nothing, pStyle)
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="pParams"></param>
    ''' <param name="pWhereClause"></param>
    ''' <param name="pStyle"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetExistingCategories(ByVal pParams As CDBParameters, ByVal pWhereClause As CDBFields, pStyle As Category.ActivityEntryStyles) As DataTable
      Dim vAnsiJoins As New AnsiJoins

      vAnsiJoins.Add("category_links cl", "cl.category_id", "cat.category_id")
      Dim vFields As String = "cat.category_id,activity,activity_value,valid_from,valid_to,source"

      If pWhereClause Is Nothing Then
        Dim vWhereFields As New CDBFields
        If pParams.HasValue("ExamUnitLinkId") Then
          vWhereFields.Add("exam_unit_link_id", pParams("ExamUnitLinkId").IntegerValue)
        End If
        If pParams.HasValue("ExamCentreId") Then
          vWhereFields.Add("exam_centre_id", pParams("ExamCentreId").IntegerValue)
        End If
        If pParams.HasValue("ExamCentreUnitId") Then
          vWhereFields.Add("exam_centre_unit_id", pParams("ExamCentreUnitId").IntegerValue)
        End If
        If pParams.HasValue("WorkstreamId") Then
          vWhereFields.Add("workstream_id", pParams("WorkstreamId").IntegerValue)
        End If
        If pParams.HasValue("Activity") Then
          vWhereFields.Add("activity", pParams("Activity").Value)
        End If
        If pParams.HasValue("ActivityValue") Then
          vWhereFields.Add("activity_value", pParams("ActivityValue").Value)
        End If
        Dim vValidFrom As String
        If pParams.HasValue("ValidFrom") AndAlso pParams("ValidFrom").Value.Length > 0 Then
          vValidFrom = pParams("ValidFrom").Value
          If (pStyle And ActivityEntryStyles.aesCheckDateRange) > 0 Then
            vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, vValidFrom, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
          Else
            vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, vValidFrom)
          End If
        Else
          vValidFrom = TodaysDate()
        End If

        Dim vValidTo As String
        If pParams.HasValue("ValidTo") AndAlso pParams("ValidTo").Value.Length > 0 Then
          vValidTo = pParams("ValidTo").Value
          If (pStyle And ActivityEntryStyles.aesCheckDateRange) > 0 Then
            If (pStyle And ActivityEntryStyles.aesSmartClient) > 0 Then vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, vValidTo, CDBField.FieldWhereOperators.fwoLessThanEqual)
          End If
        Else
          vValidTo = TodaysDate()
        End If

        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "categories cat", vWhereFields, "", vAnsiJoins)
        Dim vDT As DataTable = mvEnv.Connection.GetDataTable(vSQL)
        Return vDT
      Else
        Dim vSQL As New SQLStatement(mvEnv.Connection, vFields, "categories cat", pWhereClause, "", vAnsiJoins)
        Dim vDT As DataTable = mvEnv.Connection.GetDataTable(vSQL)
        Return vDT
      End If

    End Function

    ''' <summary>
    ''' Check if the record already exists in the database with activity, activityvalue, source, validfrom, validto   
    ''' </summary>
    ''' <param name="pParams"></param>
    ''' <param name="pActivity"></param>
    ''' <param name="pActivityValue"></param>
    ''' <param name="pSource"></param>
    ''' <param name="pValidFrom"></param>
    ''' <param name="pValidTo"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function IsValid(ByVal pParams As CDBParameters, ByVal pActivity As String, ByVal pActivityValue As String, Optional ByVal pSource As String = "", Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "") As Boolean
      Dim vFields As New CDBFields
      With vFields
        .Add(mvClassFields.Item(CategoryFields.Activity).Name, pActivity)
        .Add(mvClassFields.Item(CategoryFields.ActivityValue).Name, pActivityValue)
        If pSource.Length > 0 AndAlso Not mvEnv.GetConfigOption("activity_exclude_source_check", False) Then .Add(mvClassFields.Item(CategoryFields.Source).Name, pSource)
        If pValidTo.Length > 0 Then
          .Add(mvClassFields.Item(CategoryFields.ValidFrom).Name, CDBField.FieldTypes.cftDate, pValidTo, CDBField.FieldWhereOperators.fwoLessThanEqual)
        End If
        If pValidFrom.Length > 0 Then
          .Add(mvClassFields.Item(CategoryFields.ValidTo).Name, CDBField.FieldTypes.cftDate, pValidFrom, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        End If
        Dim vExistingValidFrom As String = ""
        Dim vExistingValidTo As String = ""
        Dim vExistingSource As String = ""
        With mvClassFields
          If Not (.Item(CategoryFields.Activity).ValueChanged OrElse .Item(CategoryFields.ActivityValue).ValueChanged OrElse .Item(CategoryFields.Source).ValueChanged) _
            OrElse mvEnv.GetConfigOption("activity_exclude_source_check", False) AndAlso .Item(CategoryFields.Source).SetValue <> pSource Then
            vExistingValidFrom = .Item(CategoryFields.ValidFrom).SetValue
            vExistingValidTo = .Item(CategoryFields.ValidTo).SetValue
            If mvEnv.GetConfigOption("activity_exclude_source_check", False) AndAlso .Item(CategoryFields.Source).SetValue <> pSource Then vExistingSource = .Item(CategoryFields.Source).SetValue
          End If
        End With
        If vExistingValidFrom.Length > 0 OrElse vExistingValidTo.Length > 0 Then
          Dim vWhereOperator As CDBField.FieldWhereOperators = CDBField.FieldWhereOperators.fwoNotEqual
          If (vExistingValidFrom.Length > 0 AndAlso vExistingValidTo.Length > 0) OrElse (vExistingValidFrom.Length > 0 OrElse vExistingValidTo.Length > 0 AndAlso vExistingSource.Length > 0) Then vWhereOperator = CDBField.FieldWhereOperators.fwoNOT Or CDBField.FieldWhereOperators.fwoOpenBracket
          If vExistingValidFrom.Length > 0 Then .Add(mvClassFields.Item(CategoryFields.ValidFrom).Name & "#2", CDBField.FieldTypes.cftDate, vExistingValidFrom, vWhereOperator)
          If vExistingValidTo.Length > 0 Then
            If vExistingValidFrom.Length > 0 Then
              If vExistingSource.Length > 0 Then .Add(mvClassFields.Item(CategoryFields.Source).Name, CDBField.FieldTypes.cftCharacter, vExistingSource)
              vWhereOperator = CDBField.FieldWhereOperators.fwoCloseBracket
            ElseIf vExistingSource.Length > 0 Then
              .Add(mvClassFields.Item(CategoryFields.Source).Name, CDBField.FieldTypes.cftCharacter, vExistingSource, vWhereOperator)
              vWhereOperator = CDBField.FieldWhereOperators.fwoCloseBracket
            End If
            .Add(mvClassFields.Item(CategoryFields.ValidTo).Name & "#2", CDBField.FieldTypes.cftDate, vExistingValidTo, vWhereOperator)
          ElseIf vExistingSource.Length > 0 AndAlso vExistingValidFrom.Length > 0 Then
            .Add(mvClassFields.Item(CategoryFields.Source).Name, CDBField.FieldTypes.cftCharacter, vExistingSource, CDBField.FieldWhereOperators.fwoCloseBracket)
          End If
        End If
      End With

      If pParams.HasValue("ExamUnitLinkId") Then vFields.Add("exam_unit_link_id", pParams("ExamUnitLinkId").Value)
      If pParams.HasValue("ExamCentreId") Then vFields.Add("exam_centre_id", pParams("ExamCentreId").Value)
      If pParams.HasValue("ExamCentreUnitId") Then vFields.Add("exam_centre_unit_id", pParams("ExamCentreUnitId").Value)
      If pParams.HasValue("WorkstreamId") Then vFields.Add("workstream_id", pParams("WorkstreamId").Value)

      Return GetExistingCategories(pParams, vFields, Nothing).Rows.Count > 0
    End Function


    ''' <summary>
    ''' Check if the Categogy is existing then Update it else 
    ''' </summary>
    ''' <param name="pStyle"></param>
    ''' <param name="pParams"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Exists(ByVal pStyle As ActivityEntryStyles, ByVal pParams As CDBParameters, ByRef pCategoryId As Integer) As CategoryUpdated
      Dim vContactCategoryUpdated As CategoryUpdated
      Dim vWhereFields As New CDBFields

      If pParams.HasValue("ExamUnitLinkId") Then vWhereFields.Add("exam_unit_link_id", pParams("ExamUnitLinkId").IntegerValue)
      If pParams.HasValue("ExamCentreId") Then vWhereFields.Add("exam_centre_id", pParams("ExamCentreId").IntegerValue)
      If pParams.HasValue("ExamCentreUnitId") Then vWhereFields.Add("exam_centre_unit_id", pParams("ExamCentreUnitId").IntegerValue)
      If pParams.HasValue("WorkstreamId") Then vWhereFields.Add("workstream_id", pParams("WorkstreamId").IntegerValue)
      If pParams.HasValue("Activity") Then vWhereFields.Add("activity", pParams("Activity").Value)
      If pParams.HasValue("ActivityValue") Then vWhereFields.Add("activity_value", pParams("ActivityValue").Value)


      Dim vValidFrom As String
      If pParams.HasValue("ValidFrom") AndAlso pParams("ValidFrom").Value.Length > 0 Then
        vValidFrom = pParams("ValidFrom").Value
        If (pStyle And ActivityEntryStyles.aesCheckDateRange) > 0 Then
          vWhereFields.Add("valid_to", CDBField.FieldTypes.cftDate, vValidFrom, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
        Else
          vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, vValidFrom)
        End If
      Else
        vValidFrom = TodaysDate()
      End If

      Dim vValidTo As String
      If pParams.HasValue("ValidTo") AndAlso pParams("ValidTo").Value.Length > 0 Then
        vValidTo = pParams("ValidTo").Value
        If (pStyle And ActivityEntryStyles.aesCheckDateRange) > 0 Then
          If (pStyle And ActivityEntryStyles.aesSmartClient) > 0 Then vWhereFields.Add("valid_from", CDBField.FieldTypes.cftDate, vValidTo, CDBField.FieldWhereOperators.fwoLessThanEqual)
        End If
      Else
        vValidTo = TodaysDate()
      End If

      'If we allow multiple activity entries for the same activities
      'but different sources then search only for the specified source
      If (pStyle And ActivityEntryStyles.aesAllowMultipleSource) > 0 AndAlso Not mvEnv.GetConfigOption("activity_exclude_source_check", False) Then vWhereFields.Add("source", pParams("Source").Value)

      Dim vDoInsert As Boolean
      If (pStyle And ActivityEntryStyles.aesIgnoreExisting) = 0 Then
        Dim vDT As DataTable = GetExistingCategories(pParams, vWhereFields, pStyle)
        If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
          pCategoryId = CInt(vDT.Rows(0).Item("category_id").ToString)
        End If
        If pParams.HasValue("ExamUnitLinkId") Then vWhereFields.Remove("exam_unit_link_id")
        If pParams.HasValue("ExamCentreId") Then vWhereFields.Remove("exam_centre_id")
        If pParams.HasValue("ExamCentreUnitId") Then vWhereFields.Remove("exam_centre_unit_id")
        If pParams.HasValue("WorkstreamId") Then vWhereFields.Remove("workstream_id")

        vWhereFields.Add("category_id", pCategoryId)

        InitWithPrimaryKey(vWhereFields)

        If Existing Then
          Dim vUpdateValidFrom As Boolean = False
          Dim vUpdateValidTo As Boolean = True
          If (pStyle And ActivityEntryStyles.aesCheckDateRange) > 0 Then
            If CDate(vValidFrom) <= CDate(ValidFrom) Then vUpdateValidFrom = True
            'if new Valid To <= original Valid To there's no need to update the record
            If CDate(vValidTo) <= CDate(ValidTo) Then vUpdateValidTo = False
          End If
          If vUpdateValidFrom = True Or vUpdateValidTo = True Then
            vContactCategoryUpdated = CategoryUpdated.Update
          Else
            vContactCategoryUpdated = CategoryUpdated.Duplicate
          End If
        Else
          vDoInsert = True
        End If
      Else
        With vWhereFields
          If .Exists("valid_from") Then .Remove("valid_from")
          If .Exists("valid_to") Then .Remove("valid_to")
          .Add("valid_from", CDBField.FieldTypes.cftDate, vValidFrom)
          .Add("valid_to", CDBField.FieldTypes.cftDate, vValidTo)
        End With
        vDoInsert = mvEnv.Connection.GetCount(mvClassFields.DatabaseTableName, vWhereFields) = 0
      End If
      If vDoInsert Then
        vContactCategoryUpdated = CategoryUpdated.Insert
      Else
        If vContactCategoryUpdated <> CategoryUpdated.Update Then
          vContactCategoryUpdated = CategoryUpdated.Duplicate
        End If
      End If
      Return vContactCategoryUpdated
    End Function

    Public Function IsValidForUpdate(ByVal pParams As CDBParameters) As Boolean
      Dim vValid As Boolean = True
      With mvClassFields
        If .Item(CategoryFields.Activity).ValueChanged OrElse .Item(CategoryFields.ActivityValue).ValueChanged OrElse _
           .Item(CategoryFields.Source).ValueChanged OrElse .Item(CategoryFields.ValidFrom).ValueChanged OrElse _
           .Item(CategoryFields.ValidTo).ValueChanged Then
          vValid = (IsValid(pParams, Activity, ActivityValue, Source, ValidFrom, ValidTo) = False)
        End If
      End With
      Return vValid
    End Function

    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      ValidateHistoricActivity()
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
    End Sub

    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer, pForceAmendmentHistory As Boolean)
      ValidateHistoricActivity()
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
    End Sub

    Private Sub ValidateHistoricActivity()
      If mvClassFields("activity").ValueChanged Then
        Dim vValidator As New AllowHistoricActivityValidator(mvEnv, Me.Activity)
        If Not vValidator.Validate() Then
          RaiseError(DataAccessErrors.daeRecordIsHistoric, "Activity")
        End If
      End If
      If mvClassFields("activity_value").ValueChanged Then
        Dim vValidator As New AllowHistoricActivityValueValidator(mvEnv, Me.Activity, Me.ActivityValue)
        If Not vValidator.Validate() Then
          RaiseError(DataAccessErrors.daeRecordIsHistoric, "Activity Value")
        End If
      End If
    End Sub

#End Region

  End Class
End Namespace