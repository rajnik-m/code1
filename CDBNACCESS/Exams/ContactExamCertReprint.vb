Public Class ContactExamCertReprint
  Inherits CARERecord
  Implements IEquatable(Of ContactExamCertReprint)

  Public Enum ContactExamCertReprintColumns
    AllFields = 0
    ContactExamCertReprintId
    ContactExamCertId
    ExamCertReprintType
    AmendedBy
    AmendedOn
  End Enum

  Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                        pCertificate As ContactExamCert,
                                        pReprintType As ExamCertReprintType) As ContactExamCertReprint
    Dim vNewInstance As New ContactExamCertReprint(pEnv)
    If pEnv Is Nothing Then
      Throw New ArgumentNullException("pEnv")
    ElseIf ContactExamCertReprint.GetInstance(pEnv, pCertificate) IsNot Nothing Then
      Throw New InvalidOperationException("A Reprint for that Certificate already exists.")
    Else
      vNewInstance.Certificate = pCertificate
      vNewInstance.ReprintType = pReprintType
    End If
    Return vNewInstance
  End Function

  Public Shared Function GetInstance(pEnv As CDBEnvironment, pId As Integer) As ContactExamCertReprint
    Dim vNewInstance As New ContactExamCertReprint(pEnv)
    vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ContactExamCertReprintColumns.ContactExamCertReprintId).Name,
                                                               pId)}))
    Return If(vNewInstance.Existing, vNewInstance, Nothing)
  End Function

  Public Shared Function GetInstance(pEnv As CDBEnvironment, pCertificate As ContactExamCert) As ContactExamCertReprint
    Dim vNewInstance As New ContactExamCertReprint(pEnv)
    vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ContactExamCertReprintColumns.ContactExamCertId).Name,
                                                               pCertificate.Id)}))
    Return If(vNewInstance.Existing, vNewInstance, Nothing)
  End Function


  Public Shared Function GetAll(pEnv As CDBEnvironment) As IEnumerable(Of ContactExamCertReprint)
    Dim vResult As New List(Of ContactExamCertReprint)
    For Each vRow As DataRow In New SQLStatement(pEnv.Connection,
                                                 "*",
                                                 "contact_exam_cert_reprints").GetDataTable.Rows
      vResult.Add(ContactExamCertReprint.GetInstance(pEnv, CInt(vRow("contact_exam_cert_reprint_id"))))
    Next vRow
    Return vResult.AsReadOnly
  End Function

  Private Sub New(ByVal pEnv As CDBEnvironment)
    MyBase.New(pEnv)
    Me.Init()
  End Sub

  Protected Overrides Sub AddFields()
    mvClassFields.Add("contact_exam_cert_reprint_id", CDBField.FieldTypes.cftInteger)
    mvClassFields.Add("contact_exam_cert_id", CDBField.FieldTypes.cftInteger)
    mvClassFields.Add("exam_cert_reprint_type")

    mvClassFields.Item(ContactExamCertReprintColumns.ContactExamCertReprintId).PrimaryKey = True
    mvClassFields.Item(ContactExamCertReprintColumns.ContactExamCertReprintId).PrefixRequired = True
    mvClassFields.Item(ContactExamCertReprintColumns.ContactExamCertId).PrefixRequired = True
    mvClassFields.Item(ContactExamCertReprintColumns.ExamCertReprintType).PrefixRequired = True
    mvClassFields.SetControlNumberField(ContactExamCertReprintColumns.ContactExamCertReprintId, "XDX")
  End Sub

  Protected Overrides ReadOnly Property TableAlias() As String
    Get
      Return "xccr"
    End Get
  End Property

  Protected Overrides ReadOnly Property DatabaseTableName() As String
    Get
      Return "contact_exam_cert_reprints"
    End Get
  End Property

  Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
    Get
      Return True
    End Get
  End Property

  Public ReadOnly Property Id() As Integer
    Get
      Return mvClassFields(ContactExamCertReprintColumns.ContactExamCertReprintId).IntegerValue
    End Get
  End Property

  Private mvContactExamCert As ContactExamCert = Nothing
  Public Property Certificate() As ContactExamCert
    Get
      If mvContactExamCert Is Nothing Then
        mvContactExamCert = ContactExamCert.GetInstance(mvEnv, mvClassFields(ContactExamCertReprintColumns.ContactExamCertId).IntegerValue())
      End If
      Return mvContactExamCert
    End Get
    Private Set(value As ContactExamCert)
      If value Is Nothing Then
        Throw New ArgumentNullException("value")
      ElseIf Not value.Existing Then
        Throw New ArgumentException("Entity must be persisted before it can be used here.", "value")
      Else
        mvContactExamCert = value
        mvClassFields(ContactExamCertReprintColumns.ContactExamCertId).IntegerValue = value.Id
      End If
    End Set
  End Property

  Private mvExamCertReprintType As ExamCertReprintType = Nothing
  Public Property ReprintType() As ExamCertReprintType
    Get
      If mvExamCertReprintType Is Nothing Then
        mvExamCertReprintType = ExamCertReprintType.GetInstance(mvEnv, mvClassFields(ContactExamCertReprintColumns.ExamCertReprintType).Value())
      End If
      Return mvExamCertReprintType
    End Get
    Private Set(value As ExamCertReprintType)
      If value Is Nothing Then
        Throw New ArgumentNullException("value")
      ElseIf Not value.Existing Then
        Throw New ArgumentException("Entity must be persisted before it can be used here.", "value")
      Else
        mvExamCertReprintType = value
        mvClassFields(ContactExamCertReprintColumns.ExamCertReprintType).Value = value.Code
      End If
    End Set
  End Property

  Public ReadOnly Property AmendedBy() As String
    Get
      Return mvClassFields(ContactExamCertReprintColumns.AmendedBy).Value
    End Get
  End Property

  Public ReadOnly Property AmendedOn() As Date
    Get
      Return Date.Parse(mvClassFields(ContactExamCertReprintColumns.AmendedOn).Value)
    End Get
  End Property

  Public Overrides Sub Update(pParameterList As CDBParameters)
    Dim vUnoprocesedParameters As New CDBParameters
    For Each vParameter As CDBParameter In pParameterList
      Select Case vParameter.Name
        Case "ContactExamCertReprintId", "ContactExamCertId", "ExamCertReprintType"
          Throw New ArgumentException("Attempt to upate an immutable property", vParameter.Name)
        Case Else
          vUnoprocesedParameters.Add(vParameter)
      End Select
    Next vParameter
    MyBase.Update(vUnoprocesedParameters)
  End Sub

  Public Overloads Function Equals(pOther As ContactExamCertReprint) As Boolean Implements IEquatable(Of ContactExamCertReprint).Equals
    Return If(pOther Is Nothing, False, pOther.Id = Me.Id)
  End Function

  Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
    Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ContactExamCertReprint) AndAlso Me.Equals(DirectCast(obj, ContactExamCertReprint))
  End Function

  Public Shared Operator =(ByVal pObj1 As ContactExamCertReprint, ByVal pObj2 As ContactExamCertReprint) As Boolean
    Return Object.Equals(pObj1, pObj2)
  End Operator

  Public Shared Operator <>(ByVal pObj1 As ContactExamCertReprint, ByVal pObj2 As ContactExamCertReprint) As Boolean
    Return Not (pObj1 = pObj2)
  End Operator

  Public Overrides Function GetHashCode() As Integer
    Return Me.Id.GetHashCode()
  End Function

End Class
