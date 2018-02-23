Imports System.Linq
Imports Advanced.Data.Merge
Imports CARE.Access.CDBEnvironment.cdbControlConstants

Namespace Access

  Public Class ContactExamCert
    Inherits CARERecord
    Implements IEquatable(Of ContactExamCert)

    Public Enum ContactExamCertColumns
      AllFields = 0
      ContactExamCertId
      ContactNumber
      ExamCertNumberPrefix
      ExamCertNumber
      ExamCertNumberSuffix
      ExamStudentUnitHeaderId
      ExamCertRunId
      IsCertificateRecalled
      AmendedBy
      AmendedOn
    End Enum

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pContactNumber As Integer,
                                          pStudentUnitHeaderId As Integer,
                                          pCertRun As ExamCertRun,
                                          pAttributes As IEnumerable(Of KeyValuePair(Of String, String))) As ContactExamCert
      Dim vTransactionStarted As Boolean = pEnv.Connection.StartTransaction
      Try
        Dim vNewInstance As ContactExamCert = CreateInstance(pEnv,
                                                               pContactNumber,
                                                               pStudentUnitHeaderId,
                                                               pCertRun,
                                                               pAttributes,
                                                               pEnv.GetControlValue(cdbControlExamCertNumberPrefix),
                                                               NextCertificateNumber(pEnv),
                                                               pEnv.GetControlValue(cdbControlExamCertNumberSuffix))
        If vTransactionStarted Then
          pEnv.Connection.CommitTransaction()
        End If
        Return vNewInstance
      Catch vEx As Exception
        If vTransactionStarted Then
          pEnv.Connection.RollbackTransaction()
        End If
        Throw
      End Try
    End Function

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pContactNumber As Integer,
                                          pStudentUnitHeaderId As Integer,
                                          pCertRun As ExamCertRun,
                                          pAttributes As IEnumerable(Of KeyValuePair(Of String, String)),
                                          pCertNumberPrefix As String,
                                          pCertNumber As Integer,
                                          pCertNumberSuffix As String) As ContactExamCert
      Dim vNewInstance As New ContactExamCert(pEnv)
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      ElseIf ContactExamCert.GetInstance(pEnv, pContactNumber, pStudentUnitHeaderId, pCertRun) IsNot Nothing Then
        Throw New InvalidOperationException("A Certificate for that booking unit and certificate run already exists for that contact.")
      Else
        vNewInstance.ContactNumber = pContactNumber
        vNewInstance.StudentUnitHeaderId = pStudentUnitHeaderId
        vNewInstance.CertificateRun = pCertRun
        vNewInstance.CertificateNumberPrefix = pCertNumberPrefix
        vNewInstance.CertificateNumber = pCertNumber
        vNewInstance.CertificateNumberSuffix = pCertNumberSuffix
        vNewInstance.IsRecalled = False
        vNewInstance.mvAttributes = New List(Of ContactExamCertItem)
        For Each vAttribute As KeyValuePair(Of String, String) In pAttributes
          vNewInstance.mvAttributes.Add(ContactExamCertItem.CreateInstance(vNewInstance, vAttribute.Key, vAttribute.Value))
        Next vAttribute
      End If
      Return vNewInstance
    End Function

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pContactNumber As Integer,
                                          pStudentUnitHeaderId As Integer,
                                          pCertRun As ExamCertRun,
                                          pAttributes As IEnumerable(Of KeyValuePair(Of String, String)),
                                          pCertNumberPrefix As String,
                                          pCertNumberSuffix As String) As ContactExamCert
      Dim vTransactionStarted As Boolean = pEnv.Connection.StartTransaction
      Try
        Dim vNewInstance As ContactExamCert = CreateInstance(pEnv,
                                                               pContactNumber,
                                                               pStudentUnitHeaderId,
                                                               pCertRun,
                                                               pAttributes,
                                                               pCertNumberPrefix,
                                                               NextCertificateNumber(pEnv),
                                                               pCertNumberSuffix)
        If vTransactionStarted Then
          pEnv.Connection.CommitTransaction()
        End If
        Return vNewInstance
      Catch vEx As Exception
        If vTransactionStarted Then
          pEnv.Connection.RollbackTransaction()
        End If
        Throw
      End Try
    End Function

    Private Shared ReadOnly Property NextCertificateNumber(pEnv As CDBEnvironment) As Integer
      Get
        Dim vTransactionStarted As Boolean = pEnv.Connection.StartTransaction
        Dim vResult As Integer = IntegerValue(pEnv.GetControlValue(cdbControlExamCertNumber))
        Try
          pEnv.Connection.ExecuteSQL("UPDATE exam_controls SET exam_cert_number = " & vResult & " + 1")
          If vTransactionStarted Then
            pEnv.Connection.CommitTransaction()
          End If
        Catch vEx As Exception
          If vTransactionStarted Then
            pEnv.Connection.RollbackTransaction()
            Throw
          End If
        End Try
        Return vResult
      End Get
    End Property

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pId As Integer) As ContactExamCert
      Dim vNewInstance As New ContactExamCert(pEnv)
      vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ContactExamCertColumns.ContactExamCertId).Name,
                                                                 pId)}))
      Return If(vNewInstance.Existing, vNewInstance, Nothing)
    End Function

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pContactNumber As Integer, pExamStudentUnitHeaderId As Integer, pRun As ExamCertRun) As ContactExamCert
      Dim vNewInstance As New ContactExamCert(pEnv)
      vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ContactExamCertColumns.ContactNumber).Name,
                                                                 pContactNumber),
                                                     New CDBField(vNewInstance.mvClassFields(ContactExamCertColumns.ExamStudentUnitHeaderId).Name,
                                                                 pExamStudentUnitHeaderId),
                                                     New CDBField(vNewInstance.mvClassFields(ContactExamCertColumns.ExamCertRunId).Name,
                                                                  pRun.Id)}))
      Return If(vNewInstance.Existing, vNewInstance, Nothing)
    End Function

    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
      Me.Init()
    End Sub

    Protected Overrides Sub AddFields()
      mvClassFields.Add("contact_exam_cert_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("contact_number", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_cert_number_prefix")
      mvClassFields.Add("exam_cert_number", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_cert_number_suffix")
      mvClassFields.Add("exam_student_unit_header_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_cert_run_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("is_certificate_recalled")

      mvClassFields.Item(ContactExamCertColumns.ContactExamCertId).PrimaryKey = True
      mvClassFields.Item(ContactExamCertColumns.ContactExamCertId).PrefixRequired = True
      mvClassFields.Item(ContactExamCertColumns.ContactNumber).PrefixRequired = True
      mvClassFields.Item(ContactExamCertColumns.ExamCertNumberPrefix).PrefixRequired = True
      mvClassFields.Item(ContactExamCertColumns.ExamCertNumber).PrefixRequired = True
      mvClassFields.Item(ContactExamCertColumns.ExamCertNumberSuffix).PrefixRequired = True
      mvClassFields.Item(ContactExamCertColumns.ExamStudentUnitHeaderId).PrefixRequired = True
      mvClassFields.Item(ContactExamCertColumns.ExamCertRunId).PrefixRequired = True
      mvClassFields.SetControlNumberField(ContactExamCertColumns.ContactExamCertId, "XDC")
    End Sub

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "xcc"
      End Get
    End Property

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "contact_exam_certs"
      End Get
    End Property

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property

    Public ReadOnly Property Id() As Integer
      Get
        Return mvClassFields(ContactExamCertColumns.ContactExamCertId).IntegerValue
      End Get
    End Property

    Public Property ContactNumber As Integer
      Get
        Return Me.mvClassFields(ContactExamCertColumns.ContactNumber).IntegerValue
      End Get
      Private Set(value As Integer)
        If New SQLStatement(mvEnv.Connection,
                    "contact_number",
                    "contacts",
                    New CDBFields({New CDBField("contact_number",
                                                value)})).GetDataTable.Rows.Count < 1 Then
          Throw New ArgumentException("No Contact with the specified number exists.")
        Else
          Me.mvClassFields(ContactExamCertColumns.ContactNumber).IntegerValue = value
          If value = 0 OrElse (mvContact IsNot Nothing AndAlso Me.Contact.ContactNumber <> value) Then
            Me.Contact = Nothing
          End If
        End If
      End Set
    End Property

    Public Property CertificateNumberPrefix As String
      Get
        Return Me.mvClassFields(ContactExamCertColumns.ExamCertNumberPrefix).Value
      End Get
      Private Set(value As String)
        Me.mvClassFields(ContactExamCertColumns.ExamCertNumberPrefix).Value = value
      End Set
    End Property

    Public Property CertificateNumber As Integer
      Get
        Return Me.mvClassFields(ContactExamCertColumns.ExamCertNumber).IntegerValue
      End Get
      Private Set(value As Integer)
        Me.mvClassFields(ContactExamCertColumns.ExamCertNumber).IntegerValue = value
      End Set
    End Property

    Public Property CertificateNumberSuffix As String
      Get
        Return Me.mvClassFields(ContactExamCertColumns.ExamCertNumberSuffix).Value
      End Get
      Private Set(value As String)
        Me.mvClassFields(ContactExamCertColumns.ExamCertNumberSuffix).Value = value
      End Set
    End Property

    Public Property StudentUnitHeaderId As Integer
      Get
        Return Me.mvClassFields(ContactExamCertColumns.ExamStudentUnitHeaderId).IntegerValue
      End Get
      Private Set(value As Integer)
        Me.mvClassFields(ContactExamCertColumns.ExamStudentUnitHeaderId).IntegerValue = value
        If value = 0 OrElse (mvExamStudentUnitHeader IsNot Nothing AndAlso Me.ExamStudentUnitHeader.ExamStudentHeaderId <> value) Then
          Me.ExamStudentUnitHeader = Nothing
        End If
      End Set
    End Property

    Public Property IsRecalled As Boolean
      Get
        Return Me.mvClassFields(ContactExamCertColumns.IsCertificateRecalled).Bool
      End Get
      Set(value As Boolean)
        Me.mvClassFields(ContactExamCertColumns.IsCertificateRecalled).Bool = value
      End Set
    End Property

    Private mvExamCertRun As ExamCertRun = Nothing
    Public Property CertificateRun() As ExamCertRun
      Get
        If mvExamCertRun Is Nothing Then
          mvExamCertRun = ExamCertRun.GetInstance(mvEnv, mvClassFields(ContactExamCertColumns.ExamCertRunId).IntegerValue())
        End If
        Return mvExamCertRun
      End Get
      Private Set(value As ExamCertRun)
        If value Is Nothing Then
          Throw New ArgumentNullException("value")
        ElseIf Not value.Existing Then
          Throw New ArgumentException("Entity must be persisted before it can be used here.", "value")
        Else
          mvExamCertRun = value
          mvClassFields(ContactExamCertColumns.ExamCertRunId).IntegerValue = value.Id
        End If
      End Set
    End Property

    Private mvAttributes As List(Of ContactExamCertItem) = Nothing
    Private mvExamStudentUnitHeader As ExamStudentUnitHeader
    Private mvContact As Contact

    Public ReadOnly Property Attributes As ReadOnlyDictionary(Of String, String)
      Get
        If mvAttributes Is Nothing Then
          mvAttributes = ContactExamCertItem.GetInstances(Me)
        End If
        Return New ReadOnlyDictionary(Of String, String)(From vAttribute As ContactExamCertItem In mvAttributes
                                                         Select New KeyValuePair(Of String, String)(vAttribute.Attribute, vAttribute.Value))
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ContactExamCertColumns.AmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ContactExamCertColumns.AmendedOn).Value)
      End Get
    End Property

    <MergeParent()>
    Public Property ExamStudentUnitHeader As ExamStudentUnitHeader
      Get
        If mvExamStudentUnitHeader Is Nothing AndAlso Me.StudentUnitHeaderId > 0 Then
          Me.ExamStudentUnitHeader = Me.GetRelatedInstance(Of ExamStudentUnitHeader)({ContactExamCertColumns.ExamStudentUnitHeaderId})
        End If
        Return mvExamStudentUnitHeader
      End Get
      Set(value As ExamStudentUnitHeader)
        mvExamStudentUnitHeader = value
        If value IsNot Nothing AndAlso value.ExamStudentHeaderId <> Me.StudentUnitHeaderId Then
          Me.StudentUnitHeaderId = value.ExamStudentHeaderId
        End If
      End Set
    End Property

    <MergeParent()>
    Public Property Contact As Contact
      Get
        If mvContact Is Nothing AndAlso Me.ContactNumber > 0 Then
          Me.Contact = Me.GetRelatedInstance(Of Contact)({ContactExamCertColumns.ContactNumber})
        End If
        Return mvContact
      End Get
      Set(value As Contact)
        mvContact = value
        If value IsNot Nothing AndAlso value.ContactNumber <> Me.ContactNumber Then
          Me.ContactNumber = value.ContactNumber
        End If
      End Set
    End Property

    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      If mvAttributes IsNot Nothing Then
        For Each vAttribute As ContactExamCertItem In mvAttributes
          vAttribute.Save(pAmendedBy, pAudit, pJournalNumber)
        Next vAttribute
      End If
    End Sub

    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer, pForceAmendmentHistory As Boolean)
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
      If mvAttributes IsNot Nothing Then
        For Each vAttribute As ContactExamCertItem In mvAttributes
          vAttribute.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
        Next vAttribute
      End If
    End Sub

    Public Overrides Sub Update(pParameterList As CDBParameters)
      Dim vUnoprocesedParameters As New CDBParameters
      For Each vParameter As CDBParameter In pParameterList
        Select Case vParameter.Name
          Case "ContactExamCertId", "ContactNumber", "ExamCertNumberPrefix", "ExamCertNumber", "ExamCertNumberSuffix", "ExamBookingUnitId", "ExamCertRunId"
            Throw New ArgumentException("Attempt to update an immutable property", vParameter.Name)
          Case Else
            vUnoprocesedParameters.Add(vParameter)
        End Select
      Next vParameter
      MyBase.Update(vUnoprocesedParameters)
    End Sub

    Public Overrides Sub Delete(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      Throw New NotSupportedException("A certificate cannot be deleted")
    End Sub

    Public Overloads Function Equals(pOther As ContactExamCert) As Boolean Implements IEquatable(Of ContactExamCert).Equals
      Return If(pOther Is Nothing, False, pOther.Id = Me.Id)
    End Function

    Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
      Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ContactExamCert) AndAlso Me.Equals(DirectCast(obj, ContactExamCert))
    End Function

    Public Shared Operator =(ByVal pObj1 As ContactExamCert, ByVal pObj2 As ContactExamCert) As Boolean
      Return Object.Equals(pObj1, pObj2)
    End Operator

    Public Shared Operator <>(ByVal pObj1 As ContactExamCert, ByVal pObj2 As ContactExamCert) As Boolean
      Return Not (pObj1 = pObj2)
    End Operator

    Public Overrides Function GetHashCode() As Integer
      Return Me.Id.GetHashCode()
    End Function

    Private Class ContactExamCertItem
      Inherits CARERecord
      Implements IEquatable(Of ContactExamCertItem)

      Public Enum ContactExamCertItemColumns
        AllFields = 0
        ContactExamCertItemId
        ContactExamCertId
        ExamCertAttribute
        ExamCertAttributeValue
        AmendedBy
        AmendedOn
      End Enum

      Friend Shared Function CreateInstance(pCertificate As ContactExamCert,
                                            pAttribute As String,
                                            pValue As String) As ContactExamCertItem
        If pCertificate Is Nothing Then
          Throw New ArgumentNullException("pCertificate")
        ElseIf Not String.IsNullOrWhiteSpace(pAttribute) AndAlso ContactExamCertItem.GetInstance(pCertificate, pAttribute) IsNot Nothing Then
          Throw New InvalidOperationException("That attribute is already listed for that Certificate")
        Else
          Dim vNewInstance As New ContactExamCertItem(pCertificate.Environment)
          vNewInstance.Certificate = pCertificate
          vNewInstance.Attribute = pAttribute
          vNewInstance.Value = pValue
          Return vNewInstance
        End If
      End Function

      Friend Shared Function GetInstance(pCertificate As ContactExamCert, pAttribute As String) As ContactExamCertItem
        If pCertificate Is Nothing Then
          Throw New ArgumentNullException("pReprintType")
        Else
          Dim vNewInstance As New ContactExamCertItem(pCertificate.Environment)
          vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ContactExamCertItemColumns.ContactExamCertId).Name,
                                                                      pCertificate.Id),
                                                         New CDBField(vNewInstance.mvClassFields(ContactExamCertItemColumns.ExamCertAttribute).Name,
                                                                      pAttribute)}))
          Return If(vNewInstance.Existing, vNewInstance, Nothing)
        End If
      End Function

      Friend Shared Function GetInstances(pCertificate As ContactExamCert) As List(Of ContactExamCertItem)
        If pCertificate Is Nothing Then
          Throw New ArgumentNullException("pCertificate")
        Else
          Dim vAttributeList As New List(Of ContactExamCertItem)
          For Each vRow As DataRow In New SQLStatement(pCertificate.Environment.Connection,
                                                       "contact_exam_cert_item_id,contact_exam_cert_id,exam_cert_attribute,exam_cert_attribute_value,amended_by,amended_on",
                                                       "contact_exam_cert_items",
                                                       New CDBField("contact_exam_cert_id",
                                                                    pCertificate.Id)).GetDataTable.AsEnumerable
            Dim vNewInstance As New ContactExamCertItem(pCertificate.Environment)
            vNewInstance.InitFromDataRow(vRow, False)
            vAttributeList.Add(vNewInstance)
          Next vRow
          Return vAttributeList
        End If
      End Function

      Private Sub New(ByVal pEnv As CDBEnvironment)
        MyBase.New(pEnv)
        Me.Init()
      End Sub

      Protected Overrides Sub AddFields()
        mvClassFields.Add("contact_exam_cert_item_id", CDBField.FieldTypes.cftInteger)
        mvClassFields.Add("contact_exam_cert_id", CDBField.FieldTypes.cftInteger)
        mvClassFields.Add("exam_cert_attribute")
        mvClassFields.Add("exam_cert_attribute_value")

        mvClassFields.Item(ContactExamCertItemColumns.ContactExamCertItemId).PrimaryKey = True
        mvClassFields.Item(ContactExamCertItemColumns.ContactExamCertItemId).PrefixRequired = True
        mvClassFields.Item(ContactExamCertItemColumns.ContactExamCertId).PrefixRequired = True
        mvClassFields.Item(ContactExamCertItemColumns.ExamCertAttribute).PrefixRequired = True
        mvClassFields.Item(ContactExamCertItemColumns.ExamCertAttributeValue).PrefixRequired = True
        mvClassFields.SetControlNumberField(ContactExamCertColumns.ContactExamCertId, "XDI")
      End Sub

      Protected Overrides ReadOnly Property DatabaseTableName() As String
        Get
          Return "contact_exam_cert_items"
        End Get
      End Property

      Protected Overrides ReadOnly Property TableAlias() As String
        Get
          Return "xcci"
        End Get
      End Property

      Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
        Get
          Return True
        End Get
      End Property

      Friend ReadOnly Property Id() As Integer
        Get
          If mvCertificate Is Nothing Then
            mvCertificate = ContactExamCert.GetInstance(mvEnv, mvClassFields(ContactExamCertItemColumns.ContactExamCertId).IntegerValue)
          End If
          Return mvClassFields(ContactExamCertItemColumns.ContactExamCertId).IntegerValue
        End Get
      End Property

      Private mvCertificate As ContactExamCert = Nothing
      Friend Property Certificate() As ContactExamCert
        Get
          If mvCertificate Is Nothing Then
            mvCertificate = ContactExamCert.GetInstance(mvEnv, mvClassFields(ContactExamCertItemColumns.ContactExamCertId).IntegerValue)
          End If
          Return mvCertificate
        End Get
        Private Set(value As ContactExamCert)
          If value Is Nothing Then
            Throw New ArgumentNullException("value")
          Else
            mvCertificate = value
            mvClassFields(ContactExamCertItemColumns.ContactExamCertId).IntegerValue = value.Id
          End If
        End Set
      End Property

      Friend Property Attribute() As String
        Get
          Return mvClassFields(ContactExamCertItemColumns.ExamCertAttribute).Value
        End Get
        Private Set(value As String)
          If String.IsNullOrWhiteSpace(value) Then
            Throw New ArgumentNullException("value")
          Else
            mvClassFields(ContactExamCertItemColumns.ExamCertAttribute).Value = value
          End If
        End Set
      End Property

      Friend Property Value() As String
        Get
          Return mvClassFields(ContactExamCertItemColumns.ExamCertAttributeValue).Value
        End Get
        Private Set(value As String)
          mvClassFields(ContactExamCertItemColumns.ExamCertAttributeValue).Value = value
        End Set
      End Property

      Friend ReadOnly Property AmendedBy() As String
        Get
          Return mvClassFields(ContactExamCertColumns.AmendedBy).Value
        End Get
      End Property

      Friend ReadOnly Property AmendedOn() As Date
        Get
          Return Date.Parse(mvClassFields(ContactExamCertColumns.AmendedOn).Value)
        End Get
      End Property

      Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
        mvClassFields(ContactExamCertItemColumns.ContactExamCertId).IntegerValue = Me.Certificate.Id
        MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      End Sub

      Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer, pForceAmendmentHistory As Boolean)
        mvClassFields(ContactExamCertItemColumns.ContactExamCertId).IntegerValue = Me.Certificate.Id
        MyBase.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
      End Sub
      Friend Overloads Function Equals(pOther As ContactExamCertItem) As Boolean Implements IEquatable(Of ContactExamCertItem).Equals
        Return If(pOther Is Nothing, False, pOther.Certificate = Me.Certificate And pOther.Attribute = Me.Attribute)
      End Function

      Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
        Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ContactExamCert) AndAlso Me.Equals(DirectCast(obj, ContactExamCert))
      End Function

      Public Shared Operator =(ByVal pObj1 As ContactExamCertItem, ByVal pObj2 As ContactExamCertItem) As Boolean
        Return Object.Equals(pObj1, pObj2)
      End Operator

      Public Shared Operator <>(ByVal pObj1 As ContactExamCertItem, ByVal pObj2 As ContactExamCertItem) As Boolean
        Return Not (pObj1 = pObj2)
      End Operator

      Public Overrides Function GetHashCode() As Integer
        Return Me.Certificate.GetHashCode() Xor Me.Attribute.GetHashCode
      End Function
    End Class
  End Class

End Namespace