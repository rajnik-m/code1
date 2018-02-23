Namespace Access

  Public Class ContactCpdPoint
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ContactCpdPointFields
      AllFields = 0
      ContactCpdPointNumber
      ContactCpdPeriodNumber
      CpdCategoryType
      CpdCategory
      PointsDate
      CpdPoints
      EvidenceSeen
      Notes
      Activity
      ActivityValue
      CpdPoints2
      WebPublish
      CpdItemType
      CpdOutcome
      ContactNumber
      EventSessionCpdNumber
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("contact_cpd_point_number", CDBField.FieldTypes.cftLong)
        .Add("contact_cpd_period_number", CDBField.FieldTypes.cftLong)
        .Add("cpd_category_type")
        .Add("cpd_category")
        .Add("points_date", CDBField.FieldTypes.cftDate)
        .Add("cpd_points", CDBField.FieldTypes.cftLong)
        .Add("evidence_seen")
        .Add("notes")
        .Add("activity")
        .Add("activity_value")
        .Add("cpd_points_2", CDBField.FieldTypes.cftLong)
        .Add("web_publish")
        .Add("cpd_item_type")
        .Add("cpd_outcome")
        .Add("contact_number", CDBField.FieldTypes.cftLong)
        .Add("event_session_cpd_number", CDBField.FieldTypes.cftLong)

        .Item(ContactCpdPointFields.ContactCpdPointNumber).PrimaryKey = True
        .SetControlNumberField(ContactCpdPointFields.ContactCpdPointNumber, "YP")
        .Item(ContactCpdPointFields.CpdPoints2).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPoints2)
        .Item(ContactCpdPointFields.WebPublish).InDatabase = .Item(ContactCpdPointFields.CpdPoints2).InDatabase
        .Item(ContactCpdPointFields.CpdItemType).InDatabase = .Item(ContactCpdPointFields.CpdPoints2).InDatabase
        .Item(ContactCpdPointFields.CpdOutcome).InDatabase = .Item(ContactCpdPointFields.CpdPoints2).InDatabase
        .Item(ContactCpdPointFields.ContactNumber).InDatabase = mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCPDPointsContactNumber)
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ccp"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_cpd_points"
      End Get
    End Property

'--------------------------------------------------
'Default constructor
'--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

'--------------------------------------------------
'Public property procedures
'--------------------------------------------------
    Public ReadOnly Property ContactCpdPointNumber() As Integer
      Get
        Return mvClassFields(ContactCpdPointFields.ContactCpdPointNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactCpdPeriodNumber() As Integer
      Get
        Return mvClassFields(ContactCpdPointFields.ContactCpdPeriodNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property CpdCategoryType() As String
      Get
        Return mvClassFields(ContactCpdPointFields.CpdCategoryType).Value
      End Get
    End Property
    Public ReadOnly Property CpdCategory() As String
      Get
        Return mvClassFields(ContactCpdPointFields.CpdCategory).Value
      End Get
    End Property
    Public ReadOnly Property PointsDate() As String
      Get
        Return mvClassFields(ContactCpdPointFields.PointsDate).Value
      End Get
    End Property
    Public ReadOnly Property CpdPoints() As Double
      Get
        Return mvClassFields(ContactCpdPointFields.CpdPoints).DoubleValue
      End Get
    End Property
    Public ReadOnly Property EvidenceSeen() As String
      Get
        Return mvClassFields(ContactCpdPointFields.EvidenceSeen).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(ContactCpdPointFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property Activity() As String
      Get
        Return mvClassFields(ContactCpdPointFields.Activity).Value
      End Get
    End Property
    Public ReadOnly Property ActivityValue() As String
      Get
        Return mvClassFields(ContactCpdPointFields.ActivityValue).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactCpdPointFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactCpdPointFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property CpdPoints2() As Double
      Get
        Return mvClassFields(ContactCpdPointFields.CpdPoints2).DoubleValue
      End Get
    End Property
    Public ReadOnly Property WebPublish() As String
      Get
        Return mvClassFields(ContactCpdPointFields.WebPublish).Value
      End Get
    End Property
    Public ReadOnly Property CpdItemType() As String
      Get
        Return mvClassFields(ContactCpdPointFields.CpdItemType).Value
      End Get
    End Property
    Public ReadOnly Property CpdOutcome() As String
      Get
        Return mvClassFields(ContactCpdPointFields.CpdOutcome).Value
      End Get
    End Property
    Public ReadOnly Property ContactNumber() As Integer
      Get
        Return mvClassFields(ContactCpdPointFields.ContactNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property EventSessionCPDNumber() As Nullable(Of Integer)
      Get
        Dim vSessionCPDNumber As Nullable(Of Integer)
        If mvClassFields(ContactCpdPointFields.EventSessionCpdNumber).Value.Length > 0 Then vSessionCPDNumber = mvClassFields(ContactCpdPointFields.EventSessionCpdNumber).IntegerValue
        Return vSessionCPDNumber
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      If mvClassFields(ContactCpdPointFields.CpdPoints2).Value.Length = 0 Then mvClassFields(ContactCpdPointFields.CpdPoints2).Value = "0"
      If mvClassFields(ContactCpdPointFields.WebPublish).Value.Length = 0 Then mvClassFields(ContactCpdPointFields.WebPublish).Value = "Y"
      If mvClassFields(ContactCpdPointFields.ContactCpdPeriodNumber).Value.Length = 0 Then mvClassFields(ContactCpdPointFields.ContactCpdPeriodNumber).Value = "0"
      If mvEnv.GetConfigOption("cpd_points_allow_numeric") = False Then
        mvClassFields(ContactCpdPointFields.CpdPoints).IntegerValue = mvClassFields(ContactCpdPointFields.CpdPoints).IntegerValue
        mvClassFields(ContactCpdPointFields.CpdPoints2).IntegerValue = mvClassFields(ContactCpdPointFields.CpdPoints2).IntegerValue
      End If
      If CpdPoints + CpdPoints2 < 0.01 Then RaiseError(DataAccessErrors.daeCPDPointsTotalCannotBeZero)
    End Sub

    Public Overloads Sub Delete(ByVal pAmendedBy As String, ByVal pContactNumber As Integer)
      mvClassFields(ContactCpdPointFields.ContactNumber).IntegerValue = pContactNumber
      Delete(pAmendedBy)
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      If ContactNumber = 0 Then RaiseError(DataAccessErrors.daeParameterNotFound, "ContactNumber")
      Dim vTransaction As Boolean = mvEnv.Connection.StartTransaction()

      'Delete any links to documents
      mvEnv.Connection.DeleteRecords("document_log_links", New CDBFields(New CDBField("contact_cpd_point_number", ContactCpdPointNumber)), False)
      'Delete the CPD Point
      Dim vJournalNumber As Integer = mvEnv.AddJournalRecord(JournalTypes.jnlCPDPoints, JournalOperations.jnlDelete, ContactNumber, 0, ContactCpdPointNumber)
      MyBase.Delete(pAmendedBy, True, vJournalNumber)

      If vTransaction Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Overloads Sub Save(ByVal pAmendedBy As String, ByVal pContactNumber As Integer)
      Save(pAmendedBy)
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Debug.Assert(ContactNumber > 0) 'mvContactNumber should not be zero
      Dim vWhereFields As New CDBFields
      Dim vContactCPDCycleNumber As Integer
      Dim vOriginalPoint As Double
      Dim vOriginalPoint2 As Double
      Dim vCheckPoints2 As Boolean
      Dim vAttrs As String

      If mvEnv.GetConfigOption("cpd_unique_categories", True) Then
        mvClassFields.SetUniqueField(ContactCpdPointFields.ContactCpdPeriodNumber)
        If ContactCpdPeriodNumber = 0 Then
          'Points without a CPD Cycle
          mvClassFields.SetUniqueField(ContactCpdPointFields.ContactNumber)
          mvClassFields.SetUniqueField(ContactCpdPointFields.PointsDate)
        End If
        mvClassFields.SetUniqueField(ContactCpdPointFields.CpdCategoryType)
        mvClassFields.SetUniqueField(ContactCpdPointFields.CpdCategory)
      End If

      If ContactCpdPeriodNumber > 0 Then
        vWhereFields.Add("contact_cpd_period_number", ContactCpdPeriodNumber)
        Dim vSQlStatement As New SQLStatement(mvEnv.Connection, "contact_cpd_cycle_number", "contact_cpd_periods", vWhereFields)
        vContactCPDCycleNumber = IntegerValue(vSQlStatement.GetValue())

        vWhereFields.Clear()
        vWhereFields.Add("contact_cpd_point_number", ContactCpdPointNumber)
        vCheckPoints2 = mvClassFields(ContactCpdPointFields.CpdPoints2).InDatabase
        vAttrs = "cpd_points"
        If vCheckPoints2 Then vAttrs &= ",cpd_points_2"
        Dim vRS As CDBRecordSet = New SQLStatement(mvEnv.Connection, vAttrs, "contact_cpd_points", vWhereFields).GetRecordSet
        If vRS.Fetch = True Then
          vOriginalPoint = vRS.Fields(1).DoubleValue
          If vCheckPoints2 Then vOriginalPoint2 = vRS.Fields(2).DoubleValue
        End If
        vRS.CloseRecordSet()
      End If

      SetValid()

      Dim vJournalNumber As Integer
      Dim vTransaction As Boolean = mvEnv.Connection.StartTransaction()
      If Existing Then
        vJournalNumber = mvEnv.AddJournalRecord(JournalTypes.jnlCPDPoints, JournalOperations.jnlUpdate, ContactNumber, 0, ContactCpdPointNumber)
      Else
        vJournalNumber = mvEnv.AddJournalRecord(JournalTypes.jnlCPDPoints, JournalOperations.jnlInsert, ContactNumber, 0, ContactCpdPointNumber)
      End If
      MyBase.Save(pAmendedBy, True, vJournalNumber)
      If vTransaction Then mvEnv.Connection.CommitTransaction()

      If ContactCpdPeriodNumber > 0 Then
        If vOriginalPoint <> CpdPoints OrElse (vCheckPoints2 AndAlso vOriginalPoint2 <> CpdPoints2) Then
          Dim vAnsiJoins As New AnsiJoins()
          vWhereFields.Clear()
          vWhereFields.Add("ccc.contact_cpd_cycle_number", vContactCPDCycleNumber)
          vAnsiJoins.Add("contact_cpd_periods ccp", "ccc.contact_cpd_cycle_number", "ccp.contact_cpd_cycle_number", AnsiJoin.AnsiJoinTypes.InnerJoin)
          vAnsiJoins.Add("contact_cpd_points ccpo", "ccp.contact_cpd_period_number", "ccpo.contact_cpd_period_number", AnsiJoin.AnsiJoinTypes.InnerJoin)
          vAnsiJoins.Add("cpd_category_types cct", "ccpo.cpd_category_type", "cct.cpd_category_type", AnsiJoin.AnsiJoinTypes.InnerJoin)
          vAnsiJoins.Add("cpd_categories cc", "ccpo.cpd_category_type", "cc.cpd_category_type", "ccpo.cpd_category", "cc.cpd_category")
          vAttrs = "SUM(ccpo.cpd_points),ccc.cpd_cycle_type"
          If vCheckPoints2 Then vAttrs = "SUM(ccpo.cpd_points) + SUM (ccpo.cpd_points_2),ccc.cpd_cycle_type"
          Dim vSQlStatement As New SQLStatement(mvEnv.Connection, vAttrs, "contact_cpd_cycles ccc", vWhereFields, "", vAnsiJoins)
          vSQlStatement.GroupBy = "ccc.cpd_cycle_type"
          Dim vRecordSet As CDBRecordSet
          vRecordSet = vSQlStatement.GetRecordSet()
          While vRecordSet.Fetch() = True
            Dim vSumOfPoints As Integer = vRecordSet.Fields(1).IntegerValue
            vWhereFields.Clear()
            vWhereFields.Add("points_from", vSumOfPoints, CDBField.FieldWhereOperators.fwoLessThanEqual)
            vWhereFields.Add("points_to", vSumOfPoints, CDBField.FieldWhereOperators.fwoGreaterThanEqual)
            vWhereFields.Add("cpd_cycle_type", vRecordSet.Fields("cpd_cycle_type").Value)
            vSQlStatement = New SQLStatement(mvEnv.Connection, "cpd_cycle_status", "cpd_cycle_statuses", vWhereFields)
            'Updating CPD status of the contact 
            Dim vContactCPDCycle As New ContactCpdCycle(mvEnv)
            Dim vParams As New CDBParameters
            Dim vStatus As String = vSQlStatement.GetValue()
            If vStatus.Length > 0 Then
              vParams.Add("CpdCycleStatus", vSQlStatement.GetValue())
              vContactCPDCycle.Init(vContactCPDCycleNumber)
              vContactCPDCycle.Update(vParams)
              vContactCPDCycle.CheckValidity()
              vContactCPDCycle.Save(mvEnv.User.UserID, True)
            End If
          End While
          vRecordSet.CloseRecordSet()
        End If
      End If
    End Sub

    Public Sub InitFromCategory(ByVal pCPDPeriodNumber As Integer, ByVal pCategoryType As String, ByVal pCategory As String)
      Dim vWhereFields As New CDBFields

      If mvEnv.GetConfigOption("cpd_unique_categories", True) Then
        vWhereFields.Add("contact_cpd_period_number", pCPDPeriodNumber)
        vWhereFields.Add("cpd_category_type", pCategoryType)
        vWhereFields.Add("cpd_category", pCategory)
        InitWithPrimaryKey(vWhereFields)
        If Not Existing Then
          Init()
          mvClassFields(ContactCpdPointFields.ContactCpdPeriodNumber).IntegerValue = pCPDPeriodNumber
          mvClassFields(ContactCpdPointFields.CpdCategoryType).Value = pCategoryType
          mvClassFields(ContactCpdPointFields.CpdCategory).Value = pCategory
        End If
      Else
        Init()
      End If
    End Sub

    ''' <summary>Create a new <see cref="ContactCpdPoint">ContactCpdPoint</see> record from an existing <see cref="EventSessionCpd">EventSessionCpd</see>.</summary>
    ''' <param name="pEventSessionCPD">The <see cref="EventSessionCpd">EventSessionCpd</see> to create Points from.</param>
    ''' <param name="pSessionStartDate">Session start date. This will become the <see cref="PointsDate">points date</see>.</param>
    ''' <param name="pContactNumber"><see cref="ContactNumber">Contact number</see> the Points apply to.</param>
    ''' <param name="pCPDPeriodNumber">When the Points are to be linked to a <see cref="ContactCpdPeriod">ContactCpdPeriod</see> the <see cref="ContactCpdPeriodNumber">period number</see> to link to.</param>
    Public Sub CreateFromEventSessionCPD(ByVal pEventSessionCPD As EventSessionCpd, ByVal pSessionStartDate As Date, ByVal pContactNumber As Integer, ByVal pCPDPeriodNumber As Integer)
      Init()
      With mvClassFields
        .Item(ContactCpdPointFields.CpdCategory).Value = pEventSessionCPD.CpdCategory
        .Item(ContactCpdPointFields.CpdCategoryType).Value = pEventSessionCPD.CpdCategoryType
        .Item(ContactCpdPointFields.PointsDate).Value = pSessionStartDate.ToString(CAREDateFormat)
        .Item(ContactCpdPointFields.CpdPoints).DoubleValue = pEventSessionCPD.CpdPoints
        .Item(ContactCpdPointFields.CpdPoints2).DoubleValue = pEventSessionCPD.CpdPoints2
        .Item(ContactCpdPointFields.CpdItemType).Value = pEventSessionCPD.CpdItemType
        .Item(ContactCpdPointFields.CpdOutcome).Value = pEventSessionCPD.CpdOutcome
        .Item(ContactCpdPointFields.WebPublish).Value = pEventSessionCPD.WebPublish
        .Item(ContactCpdPointFields.Notes).Value = pEventSessionCPD.CpdNotes
        .Item(ContactCpdPointFields.ContactNumber).IntegerValue = pContactNumber
        .Item(ContactCpdPointFields.ContactCpdPeriodNumber).IntegerValue = pCPDPeriodNumber
        .Item(ContactCpdPointFields.EventSessionCpdNumber).IntegerValue = pEventSessionCPD.EventSessionCpdNumber
        .Item(ContactCpdPointFields.EvidenceSeen).Value = "N"
      End With
    End Sub
#End Region

  End Class
End Namespace