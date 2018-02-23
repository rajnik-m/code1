Namespace Access

  Public Class ExamCertRun
    Inherits CARERecord
    Implements IEquatable(Of ExamCertRun)

    Public Enum ExamCertRunColumns
      AllFields = 0
      ExamCertRunId
      ExamUnitCertRunTypeId
      ExamCertRunTimestamp
      ExamCertReprintType
      AmendedBy
      AmendedOn
    End Enum

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pExamUnitCertRunType As ExamUnitCertRunType) As ExamCertRun
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      Else
        Dim vNewInstance As New ExamCertRun(pEnv)
        vNewInstance.RunType = pExamUnitCertRunType
        vNewInstance.Timestamp = Date.Now
        Return vNewInstance
      End If
    End Function

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pExamUnitCertRunType As ExamUnitCertRunType,
                                          pReprintType As ExamCertReprintType) As ExamCertRun
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      Else
        Dim vNewInstance As New ExamCertRun(pEnv)
        vNewInstance.RunType = pExamUnitCertRunType
        vNewInstance.Timestamp = Date.Now
        vNewInstance.ReprintType = pReprintType
        Return vNewInstance
      End If
    End Function

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pRunId As Integer) As ExamCertRun
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      Else
        Dim vNewInstance As New ExamCertRun(pEnv)
        vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ExamCertRunColumns.ExamCertRunId).Name,
                                                                    pRunId)}))
        Return If(vNewInstance.Existing, vNewInstance, Nothing)
      End If
    End Function

    Private Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
      Me.Init()
    End Sub

    Protected Overrides Sub AddFields()
      mvClassFields.Add("exam_cert_run_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_unit_cert_run_type_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_cert_run_timestamp", CDBField.FieldTypes.cftTime)
      mvClassFields.Add("exam_cert_reprint_type")

      mvClassFields.Item(ExamCertRunColumns.ExamCertRunId).PrimaryKey = True
      mvClassFields.Item(ExamCertRunColumns.ExamUnitCertRunTypeId).PrefixRequired = True
      mvClassFields.Item(ExamCertRunColumns.ExamCertRunTimestamp).PrefixRequired = True
      mvClassFields.Item(ExamCertRunColumns.ExamCertReprintType).PrefixRequired = True
      mvClassFields.SetControlNumberField(ExamCertRunColumns.ExamCertRunId, "XDR")
    End Sub

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "xcr"
      End Get
    End Property

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_cert_runs"
      End Get
    End Property

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property

    Public ReadOnly Property Id() As Integer
      Get
        Return mvClassFields(ExamCertRunColumns.ExamCertRunId).IntegerValue
      End Get
    End Property

    Private mvExamUnitCertRunType As ExamUnitCertRunType = Nothing
    Public Property RunType() As ExamUnitCertRunType
      Get
        If mvExamUnitCertRunType Is Nothing Then
          mvExamUnitCertRunType = ExamUnitCertRunType.GetInstance(mvEnv, mvClassFields(ExamCertRunColumns.ExamUnitCertRunTypeId).IntegerValue())
        End If
        Return mvExamUnitCertRunType
      End Get
      Private Set(value As ExamUnitCertRunType)
        If value Is Nothing Then
          Throw New ArgumentNullException("value")
        Else
          mvExamUnitCertRunType = value
          mvClassFields(ExamCertRunColumns.ExamUnitCertRunTypeId).IntegerValue = value.Id
        End If
      End Set
    End Property

    Private mvExamCertReprintType As ExamCertReprintType = Nothing
    Public Property ReprintType() As ExamCertReprintType
      Get
        If mvExamCertReprintType Is Nothing Then
          ExamCertReprintType.GetInstance(mvEnv, mvClassFields(ExamCertRunColumns.ExamCertReprintType).Value)
        End If
        Return mvExamCertReprintType
      End Get
      Private Set(value As ExamCertReprintType)
        If mvEnv Is Nothing Then
          Throw New ArgumentNullException("pEnv")
        Else
          mvExamCertReprintType = value
          mvClassFields(ExamCertRunColumns.ExamCertReprintType).Value = value.Code
        End If
      End Set
    End Property

    Public Property Timestamp() As Date
      Get
        Return CDate(mvClassFields(ExamCertRunColumns.ExamCertRunTimestamp).Value)
      End Get
      Private Set(value As Date)
        mvClassFields(ExamCertRunColumns.ExamCertRunTimestamp).Value = value.ToString(CAREDateTimeFormat)
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamCertRunColumns.AmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ExamCertRunColumns.AmendedOn).Value)
      End Get
    End Property

    Public Overrides Sub Update(pParameterList As CDBParameters)
      Dim vUnoprocesedParameters As New CDBParameters
      For Each vParameter As CDBParameter In pParameterList
        Select Case vParameter.Name
          Case "ExamCertRunId", "ExamUnitCertRunTypeId", "ExamCertRunTimestamp", "ExamCertReprintType"
            Throw New ArgumentException("Attempt to upate an immutable property", vParameter.Name)
          Case Else
            vUnoprocesedParameters.Add(vParameter)
        End Select
      Next vParameter
      MyBase.Update(vUnoprocesedParameters)
    End Sub

    Public Overrides Sub Delete(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      Throw New NotSupportedException("A certificate run cannot be deleted")
    End Sub

    Public Overloads Function Equals(pOther As ExamCertRun) As Boolean Implements IEquatable(Of ExamCertRun).Equals
      Return If(pOther Is Nothing, False, pOther.Id = Me.Id)
    End Function

    Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
      Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ExamCertRun) AndAlso Me.Equals(DirectCast(obj, ExamCertRun))
    End Function

    Public Shared Operator =(ByVal pObj1 As ExamCertRun, ByVal pObj2 As ExamCertRun) As Boolean
      Return Object.Equals(pObj1, pObj2)
    End Operator

    Public Shared Operator <>(ByVal pObj1 As ExamCertRun, ByVal pObj2 As ExamCertRun) As Boolean
      Return Not (pObj1 = pObj2)
    End Operator

    Public Overrides Function GetHashCode() As Integer
      Return Me.Id.GetHashCode()
    End Function
  End Class

End Namespace