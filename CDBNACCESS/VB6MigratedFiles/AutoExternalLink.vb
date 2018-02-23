

Imports System.Linq

Namespace Access
  ''' <summary>
  ''' Decorator for a Contact External Link that automatically sets its reference based on a pattern
  ''' </summary>
  Public Class AutoExternalLink
    Private mvActualExternalLink As ContactExternalLink
    Private ReadOnly mvEnvironment As CDBEnvironment
    Private ReadOnly mvDataSource As TableMaintenanceData
    Private ReadOnly mvControlNumberTypeCode As String = "AER"

    Private Property ExternalRefFormat As String
    Private Const AUTONUMBER_PLACEHOLDER As String = "[AUTONUMBER]"
    Private mvContact As CARERecord

    Private Property ActualLink As ContactExternalLink
      Get
        If mvActualExternalLink Is Nothing Then
          Dim vDataSource As String = If(Me.DataSource Is Nothing, "", Me.DataSource.FieldValueString("data_source"))
          Me.ActualLink = InitLink(vDataSource)
        End If
        Return mvActualExternalLink
      End Get
      Set(value As ContactExternalLink)
        mvActualExternalLink = value
      End Set
    End Property
    Public Property Parent As CARERecord
      Get
        Return mvContact
      End Get
      Set(value As CARERecord)
        mvContact = value
        If mvActualExternalLink IsNot Nothing AndAlso
          (value Is Nothing OrElse Me.ActualLink.ContactNumber <> value.ClassFields.GetUniquePrimaryKey.IntegerValue) Then
          Me.ActualLink = Nothing 'reset the property.  It will reinitialise itself with the latest values.
        End If
      End Set
    End Property

    Public Property Value As String
      Get
        Return Me.ActualLink.ExternalReference
      End Get
      Protected Friend Set(value As String)
        If Me.ActualLink IsNot Nothing AndAlso Me.DataSource IsNot Nothing Then
          Me.ActualLink.Create(Me.Parent.ClassFields.GetUniquePrimaryKey.IntegerValue, Me.DataSource.FieldValueString("data_source"), value)
        End If
      End Set
    End Property

    Public ReadOnly Property Environment As CDBEnvironment
      Get
        Return mvEnvironment
      End Get
    End Property

    Private ReadOnly Property DataSource As TableMaintenanceData
      Get
        Return mvDataSource
      End Get
    End Property

    ''' <summary>
    ''' Returns True if the DataSource, ContactNumber and ExternalReference properties contain a value
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Ideally it should also validate that the values are valid, i.e. they refer to real objects.  This is currently not required,
    ''' so it not implemented.  However any the code calling this property should be aware that it may be done in the future.</remarks>
    Public ReadOnly Property IsValid As Boolean
      Get
        Return String.IsNullOrWhiteSpace(Me.ActualLink.DataSource) = False AndAlso
          Me.ActualLink.ContactNumber > 0 AndAlso
          String.IsNullOrWhiteSpace(Me.ActualLink.ExternalReference) = False
      End Get
    End Property

    Public ReadOnly Property ControlNumberTypeCode As String
      Get
        Return mvControlNumberTypeCode
      End Get
    End Property

    Public Sub New(pEnv As CDBEnvironment, pDataSource As String, pParent As CARERecord, vFormat As String)
      Me.Parent = pParent
      mvEnvironment = pEnv
      Me.ExternalRefFormat = vFormat

      If If(pDataSource, String.Empty).Length > 0 Then
        Dim vDataSourceInstantiator As Func(Of TableMaintenanceData) = Function() (New TableMaintenanceData(Me.Environment, "data_sources"))
        Dim vDummy As New ContactExternalLink()
        vDummy.Init(Me.Environment)
        vDummy.Create(Me.Parent.UniquePrimaryKey.IntegerValue, pDataSource, Nothing)
        Dim vWhere As CDBFields = vDummy.CreateWhere({ContactExternalLink.ContactExternalLinkFields.celfDataSource})
        mvDataSource = CARERecordFactory.SelectInstance(Of TableMaintenanceData)(Me.Environment, vWhere, vDataSourceInstantiator)
      End If

    End Sub

    Private Function InitLink(pDataSource As String) As ContactExternalLink
      Dim vResult As ContactExternalLink = New ContactExternalLink()
      vResult.Init(Me.Environment)
      vResult.Create(Me.Parent.UniquePrimaryKey.IntegerValue, pDataSource, String.Empty)
      If Not String.IsNullOrWhiteSpace(pDataSource) Then
        Dim vWhere As CDBFields = vResult.CreateWhere(
          {ContactExternalLink.ContactExternalLinkFields.celfContactNumber,
          ContactExternalLink.ContactExternalLinkFields.celfDataSource})
        Dim vSearch As ContactExternalLink = CARERecordFactory.SelectInstance(Of ContactExternalLink)(Me.Environment, vWhere)
        If vSearch IsNot Nothing Then
          vResult = vSearch
        End If
      End If
      Return vResult
    End Function

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      Try
        CheckExternalReference()
        CheckContact()
        Me.ActualLink.Save(pAmendedBy, pAudit)
      Catch ex As CareException
        If ex.ErrorNumber = DataAccessErrors.daeDuplicateRecord Then
          RaiseError(DataAccessErrors.daeExternalReferenceNotUnique, Me.DataSource.FieldValueString("data_source_desc"), Me.Value)
        Else
          Throw
        End If
      End Try
    End Sub

    Private Sub CheckExternalReference()
      If Me.DataSource IsNot Nothing AndAlso
        Me.ActualLink.Existing = False AndAlso
        String.IsNullOrWhiteSpace(Me.ActualLink.ExternalReference) Then
        Me.Value = GenerateExternalReference()
      End If
    End Sub

    Private Function GenerateExternalReference() As String
      Dim vReturn As String = String.Empty
      Dim vFormat As String = Me.ExternalRefFormat
      'The format must always be alpha-numeric, otherwise the Contact TextLookupBox won't be able to differentiate between a contact number and an AutoExternalRef
      vFormat = If(String.IsNullOrWhiteSpace(vFormat), Me.DataSource.FieldValueString("data_source"), vFormat)
      vFormat = If(vFormat.Contains(AUTONUMBER_PLACEHOLDER), vFormat, vFormat + AUTONUMBER_PLACEHOLDER)
      Dim vNewNumber As Integer = Me.Environment.GetControlNumber(Me.ControlNumberTypeCode)
      vReturn = vFormat.Replace(AUTONUMBER_PLACEHOLDER, vNewNumber.ToString())
      Return vReturn
    End Function

    Public Sub SaveIfValid(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      CheckExternalReference()

      If Me.IsValid Then
        Me.Save(pAmendedBy, pAudit)
      End If
    End Sub

    Private Sub CheckContact()
      If Me.Parent.UniquePrimaryKey.IntegerValue > 0 AndAlso Me.ActualLink IsNot Nothing AndAlso Me.ActualLink.ContactNumber <> Me.Parent.UniquePrimaryKey.IntegerValue Then
        Me.ActualLink.Create(Me.Parent.UniquePrimaryKey.IntegerValue, Me.DataSource.FieldValueString("data_source"), Me.Value)
      End If
    End Sub
  End Class
End Namespace
