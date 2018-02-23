Namespace Access

  Public Class ContactCpdPeriod
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum ContactCpdPeriodFields
      AllFields = 0
      ContactCpdPeriodNumber
      ContactCpdCycleNumber
      StartDate
      EndDate
      ContactCpdPeriodNumberDesc
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("contact_cpd_period_number", CDBField.FieldTypes.cftLong)
        .Add("contact_cpd_cycle_number", CDBField.FieldTypes.cftLong)
        .Add("start_date", CDBField.FieldTypes.cftDate)
        .Add("end_date", CDBField.FieldTypes.cftDate)
        .Add("contact_cpd_period_number_desc")

        .Item(ContactCpdPeriodFields.ContactCpdPeriodNumber).PrimaryKey = True
        .SetControlNumberField(ContactCpdPeriodFields.ContactCpdPeriodNumber, "YI")
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
        Return "contact_cpd_periods"
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
    Public ReadOnly Property ContactCpdPeriodNumber() As Integer
      Get
        Return mvClassFields(ContactCpdPeriodFields.ContactCpdPeriodNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property ContactCpdCycleNumber() As Integer
      Get
        Return mvClassFields(ContactCpdPeriodFields.ContactCpdCycleNumber).IntegerValue
      End Get
    End Property
    Public ReadOnly Property StartDate() As String
      Get
        Return mvClassFields(ContactCpdPeriodFields.StartDate).Value
      End Get
    End Property
    Public ReadOnly Property EndDate() As String
      Get
        Return mvClassFields(ContactCpdPeriodFields.EndDate).Value
      End Get
    End Property
    Public ReadOnly Property ContactCpdPeriodNumberDesc() As String
      Get
        Return mvClassFields(ContactCpdPeriodFields.ContactCpdPeriodNumberDesc).Value
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactCpdPeriodFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(ContactCpdPeriodFields.AmendedOn).Value
      End Get
    End Property
#End Region

#Region "Non-AutoGenerated Code"

    Private mvContactCPDPoints As List(Of ContactCpdPoint)
    Private mvContactCPDObjectives As List(Of ContactCpdObjective)

    Protected Overrides Sub ClearFields()
      MyBase.ClearFields()
      mvContactCPDPoints = Nothing
      mvContactCPDObjectives = Nothing
    End Sub

    Protected Overrides Sub SetValid()
      MyBase.SetValid()
      If StartDate.Length > 0 AndAlso EndDate.Length > 0 Then
        ' Set the description as the start month year - end month year
        mvClassFields(ContactCpdPeriodFields.ContactCpdPeriodNumberDesc).Value = MonthName(CDate(StartDate).Month, True) & " " & CDate(StartDate).Year.ToString & " - " & MonthName(CDate(EndDate).Month, True) & " " & CDate(EndDate).Year.ToString
      End If
    End Sub

    Public Sub AddPoint(ByVal pParams As CDBParameters)
      Dim vCPDPoint As New ContactCpdPoint(mvEnv)
      Dim vCreate As Boolean

      pParams.Add("ContactCpdPeriodNumber", ContactCpdPeriodNumber)
      If mvEnv.GetConfigOption("cpd_unique_categories", True) Then
        vCPDPoint.InitFromCategory(ContactCpdPeriodNumber, pParams("CpdCategoryType").Value, pParams("CpdCategory").Value)
        If Not vCPDPoint.Existing Then vCreate = True
      Else
        vCreate = True
      End If
      If vCreate Then
        vCPDPoint.Create(pParams)
      Else
        vCPDPoint.Update(pParams)
      End If
      If mvContactCPDPoints Is Nothing Then mvContactCPDPoints = New List(Of ContactCpdPoint)
      mvContactCPDPoints.Add(vCPDPoint)
    End Sub

    Public Sub AddObjective(ByVal pParams As CDBParameters)
      Dim vCPDObjective As New ContactCpdObjective(mvEnv)

      pParams.Add("ContactCpdPeriodNumber", ContactCpdPeriodNumber)
      vCPDObjective.Create(pParams)
      If mvContactCPDObjectives Is Nothing Then mvContactCPDObjectives = New List(Of ContactCpdObjective)
      mvContactCPDObjectives.Add(vCPDObjective)
    End Sub

    Public Sub SavePoints()
      If Not mvContactCPDPoints Is Nothing Then
        For Each vCPDPoint As ContactCpdPoint In mvContactCPDPoints
          vCPDPoint.Save()
        Next
      End If
    End Sub

    Public Sub SaveObjectives()
      If Not mvContactCPDObjectives Is Nothing Then
        For Each vCPDObjective As ContactCpdObjective In mvContactCPDObjectives
          vCPDObjective.Save(String.Empty, False)
        Next
      End If
    End Sub

#End Region
  End Class
End Namespace