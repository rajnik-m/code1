Namespace Access

  Public Class ExamCentreUnit
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ExamCentreUnitFields
      AllFields = 0
      ExamCentreUnitId
      ExamCentreId
      ExamUnitId
      ExamUnitLinkId
      LocalName
      AccreditationStatus
      AccreditationValidFrom
      AccreditationValidTo
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("exam_centre_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_centre_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_id", CDBField.FieldTypes.cftInteger)
        .Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
        .Add("local_name")
        .Add("accreditation_status")
        .Add("accreditation_valid_from", CDBField.FieldTypes.cftDate)
        .Add("accreditation_valid_to", CDBField.FieldTypes.cftDate)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(ExamCentreUnitFields.ExamCentreUnitId).PrimaryKey = True
        .Item(ExamCentreUnitFields.ExamCentreUnitId).PrefixRequired = True
        .SetControlNumberField(ExamCentreUnitFields.ExamCentreUnitId, "XCU")

        .Item(ExamCentreUnitFields.CreatedBy).PrefixRequired = True
        .Item(ExamCentreUnitFields.CreatedOn).PrefixRequired = True

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
          .Item(ExamCentreUnitFields.LocalName).InDatabase = True
          .Item(ExamCentreUnitFields.ExamUnitLinkId).InDatabase = True
          .Item(ExamCentreUnitFields.AccreditationStatus).InDatabase = True
          .Item(ExamCentreUnitFields.AccreditationValidFrom).InDatabase = True
          .Item(ExamCentreUnitFields.AccreditationValidTo).InDatabase = True
        End If

      End With
    End Sub

    Public Overrides Sub AddDeleteCheckItems()
      AddDeleteCheckItem("document_log_links", "exam_centre_unit_id", "a document link")
    End Sub
    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ecu"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_centre_units"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'IRecordCreate
    '--------------------------------------------------
    Public Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New ExamCentreUnit(mvEnv)
    End Function
    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property ExamCentreUnitId() As Integer
      Get
        Return mvClassFields(ExamCentreUnitFields.ExamCentreUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamCentreId() As Integer
      Get
        Return mvClassFields(ExamCentreUnitFields.ExamCentreId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitId() As Integer
      Get
        Return mvClassFields(ExamCentreUnitFields.ExamUnitId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(ExamCentreUnitFields.ExamUnitLinkId).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property LocalName() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.LocalName).Value
      End Get
    End Property
    Public ReadOnly Property AccreditationStatus() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.AccreditationStatus).Value
      End Get
    End Property
    Public ReadOnly Property AccreditationValidFrom() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.AccreditationValidFrom).Value
      End Get
    End Property
    Public ReadOnly Property AccreditationValidTo() As String
      Get
        Return mvClassFields(ExamCentreUnitFields.AccreditationValidTo).Value
      End Get
    End Property

#End Region
#Region "Non AutoGenerated region"
    Private mvCurrentExamUnitsIndex As Integer
    Private mvExamCentreUnits As New List(Of ExamCentreUnit) 'this holds all ExamCentreUnits derived from the values held in mvClassFields  
    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      SaveAccreditationHistory()
      If mvExamCentreUnits.Count <= 1 Then
        'Inserting single record - use existing method 
        MyBase.Save(pAmendedBy, pAudit, pJournalNumber, True)
      Else
        'Inserting multiple records
        Dim vMappedDataTable As DataTable = CARERecord.GetBulkCopyDataTable(mvExamCentreUnits)
        '-- First bulk insert AmendmentHistory
        Dim vClassFieldsList As New List(Of ClassFields)
        For Each vExamCentreUnit In mvExamCentreUnits
          vClassFieldsList.Add(vExamCentreUnit.mvClassFields)
        Next
        mvEnv.BulkInsertAmendmentHistory(CDBEnvironment.AuditTypes.audInsert, DatabaseTableName, ExamCentreUnitFields.ExamCentreUnitId, 0, pAmendedBy, vClassFieldsList, 0)
        '-- now bulk insert ExamCentreUnits
        mvEnv.Connection.BulkCopyData(mvEnv.Connection, DatabaseTableName, vMappedDataTable)
      End If
    End Sub
    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)


      Dim vExamUnitIdValues() As String = pParameterList("ExamUnitId").Value.Split(","c)
      Dim vExamUnitLinkIdValues() As String = pParameterList("ExamUnitLinkId").Value.Split(","c)

      If UBound(vExamUnitIdValues) <> UBound(vExamUnitLinkIdValues) Then
        RaiseError(DataAccessErrors.daeInvalidParameter)
      End If
    End Sub
    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)

      Dim vExamUnitIdValues() As String = pParameterList("ExamUnitId").Value.Split(","c)
      Dim vExamUnitLinkIdValues() As String = pParameterList("ExamUnitLinkId").Value.Split(","c)

      Dim vRecordIndex As Integer = 0

      For vRecordIndex = 0 To UBound(vExamUnitIdValues)
        Dim vExamCentreUnit As New ExamCentreUnit(mvEnv)
        vExamCentreUnit.CopyValues(Me)
        vExamCentreUnit.mvClassFields("exam_unit_id").Value = vExamUnitIdValues(vRecordIndex)
        vExamCentreUnit.mvClassFields("exam_unit_link_id").Value = vExamUnitLinkIdValues(vRecordIndex)
        mvExamCentreUnits.Add(vExamCentreUnit)
      Next vRecordIndex

      mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnExamCentreUnit, mvExamCentreUnits.Count)

      'Default mvClassFields to hold first ExamCentreUnit values
      SetCurrentRecord(0)

    End Sub
    Public Sub SetCurrentRecord(ByVal pRecordIndex As Integer)
      'There may be many ExamCentreUnits passed in the parameters for the create method. This sets the ClassFields with the values from the specified ExamCentreUnit.
      mvCurrentExamUnitsIndex = 0
      mvClassFields = mvExamCentreUnits(mvCurrentExamUnitsIndex).mvClassFields
    End Sub
#Region "Accreditation"
    Private Sub SaveAccreditationHistory()
      If (mvClassFields(ExamCentreUnitFields.AccreditationStatus).ValueChanged And Me.Existing) Or _
      (mvClassFields(ExamCentreUnitFields.AccreditationValidFrom).ValueChanged And Me.Existing) Or _
      (mvClassFields(ExamCentreUnitFields.AccreditationValidTo).ValueChanged And Me.Existing) Then
        Dim vAccreditationHistoryRecord As New ExamAccreditationHistory(mvEnv)
        Dim vParams As New CDBParameters()
        If mvClassFields(ExamCentreUnitFields.AccreditationStatus).SetValue.Length > 0 Then
          vParams.Add("AccreditationStatus", mvClassFields(ExamCentreUnitFields.AccreditationStatus).SetValue)
          vParams.Add("AccreditationValidFrom", mvClassFields(ExamCentreUnitFields.AccreditationValidFrom).SetValue)
          vParams.Add("AccreditationValidTo", mvClassFields(ExamCentreUnitFields.AccreditationValidTo).SetValue)
          vAccreditationHistoryRecord.Create(vParams)
          vAccreditationHistoryRecord.Save()
          If vAccreditationHistoryRecord.AccreditationId > 0 Then
            Dim vAccreditationHistoryLink As New ExamAccreditationHistLink(mvEnv)
            Dim vLinkParams As New CDBParameters()
            vLinkParams.Add("ExamCentreUnitId", mvClassFields(ExamCentreUnitFields.ExamCentreUnitId).Value.ToString)
            vLinkParams.Add("AccreditationId", vAccreditationHistoryRecord.AccreditationId)
            vAccreditationHistoryLink.Create(vLinkParams)
            vAccreditationHistoryLink.Save()
          End If
        End If
      End If
    End Sub
#Region "Accreditation Status"


    ''' <summary>
    ''' Check If the Centers are accredited then only add them when the lookup is called from Trader or Results entry
    ''' </summary>
    ''' <param name="pEnv">Envronment class</param>
    ''' <param name="pDT">DatTable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsCentreUnitAccredited(ByVal pEnv As CDBEnvironment, ByVal pDT As CDBDataTable, ByVal pTrader As Boolean, ByVal pBooking As Boolean) As CDBDataTable

      If pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation).Length > 0 AndAlso _
        pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation) = "Y" Then
        If pDT IsNot Nothing AndAlso pDT.Columns.ContainsKey("AccreditationStatus") Then

          For vRowNumber As Integer = pDT.Rows.Count - 1 To 0 Step -1
            If Not CheckCentreAccreditationStatus(pEnv, pDT.Rows(vRowNumber).IntegerItem("ExamCentreUnitId"), pTrader) Then
              pDT.Rows.RemoveAt(vRowNumber)
            End If
          Next
        End If
      Else
        Return pDT
      End If
      Return pDT
    End Function


    ''' <summary>
    ''' Check If the Centers are accredited then only add them when the lookup is called from Trader or Results entry
    ''' </summary>
    ''' <param name="pEnv">Envronment class</param>
    ''' <param name="pDT">DatTable</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsCentreUnitAccredited(ByVal pEnv As CDBEnvironment, ByVal pDT As CDBDataTable, ByVal pTrader As Boolean) As CDBDataTable

      If pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation).Length > 0 AndAlso _
        pEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlExamCentreUnitAccreditation) = "Y" Then
        If pDT IsNot Nothing AndAlso pDT.Columns.ContainsKey("accreditation_status") Then

          For vRowNumber As Integer = pDT.Rows.Count - 1 To 0 Step -1
            If Not CheckCentreAccreditationStatus(pEnv, pDT.Rows(vRowNumber).IntegerItem("exam_centre_unit_id"), pTrader) Then
              pDT.Rows.RemoveAt(vRowNumber)
            End If
          Next
        End If
      Else
        Return pDT
      End If
      Return pDT
    End Function

    ''' <summary>
    ''' Check If the Centers are accredited then only add them when the lookup is called from Trader or Results entry
    ''' </summary>
    ''' <param name="pEnv">Environment class</param>
    ''' <param name="pCentreUnitId">Centre Unit Id </param>
    ''' <param name="pTrader">Trader flag</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function CheckCentreAccreditationStatus(ByVal pEnv As CDBEnvironment, ByVal pCentreUnitId As Integer, ByVal pTrader As Boolean) As Boolean
      Dim vAnsiJoin As New AnsiJoins
      Dim vFields As String = "ecu.accreditation_status,allow_registration,ignore_accreditation_validity,allow_result_entry,ecu.accreditation_valid_from,ecu.accreditation_valid_to"
      Dim vWhereClause As New CDBFields
      Dim vResult As Boolean = False

      vWhereClause.Add("ecu.exam_centre_unit_id", pCentreUnitId)
      If pTrader Then
        vWhereClause.Add("acs.allow_registration", "Y")
      Else
        vWhereClause.Add("acs.allow_result_entry", "Y")
      End If
      vAnsiJoin.Add("exam_accreditation_statuses acs", "ecu.accreditation_status", "acs.accreditation_status")

      Dim vSql As New SQLStatement(pEnv.Connection, vFields, "exam_centre_units ecu", vWhereClause, "", vAnsiJoin)
      Dim vDataTable As New CDBDataTable
      vDataTable.FillFromSQL(pEnv, vSql)

      If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then

        Dim vIgnoreBooking As Boolean = If(vDataTable.Rows(0).Item("ignore_accreditation_validity").Length > 0 AndAlso vDataTable.Rows(0).Item("ignore_accreditation_validity") = "Y", True, False)

        Dim vValidFrom As String = vDataTable.Rows(0).Item("accreditation_valid_from")
        Dim vValidTo As String = vDataTable.Rows(0).Item("accreditation_valid_to")

        'CheckCentreAccreditationStatus If booking is allowed for centers, this should only be checked for trade application
        If pTrader Then
          If IsAccreditationValid(vValidFrom, vValidTo, vIgnoreBooking) Then vResult = True
        Else
          Return True
        End If

      End If
      Return vResult
    End Function
    ''' <summary>
    ''' Validate the Date range specified for Accreditation
    ''' </summary>
    ''' <param name="pValidFrom">Accreditation Valid from</param>
    ''' <param name="pValidTo">Accreditation valid To</param>
    ''' <param name="pIgnoreValidity"> Ignore Validity</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Shared Function IsAccreditationValid(ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pIgnoreValidity As Boolean) As Boolean
      Dim vResult As Boolean = False

      'Check if the dates are valid
      If Not pIgnoreValidity Then
        If pValidFrom.Length = 0 AndAlso pValidTo.Length = 0 Then
          vResult = False
        ElseIf pValidFrom.Length > 0 AndAlso CDate(pValidFrom) > Date.Today Then
          vResult = False 'future
        ElseIf pValidTo.Length > 0 AndAlso CDate(pValidTo) < Date.Today Then
          vResult = False 'past 
        Else
          vResult = True
        End If
      Else
        vResult = True
      End If
      Return vResult
    End Function

#End Region
#End Region

#End Region
  End Class

End Namespace