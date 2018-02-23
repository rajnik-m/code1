Namespace Access

  Public Class ExamCertRunType
    Inherits CARERecord
    Implements IEquatable(Of ExamCertRunType)

    Public Enum ExamCertRunTypeColumns
      AllFields = 0
      ExamCertRunType
      ExamCertRunTypeDesc
      AmendedBy
      AmendedOn
    End Enum

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pCode As String,
                                          pDescription As String) As ExamCertRunType
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      ElseIf ExamCertRunType.GetInstance(pEnv, pCode) IsNot Nothing Then
        Throw New InvalidOperationException("A Certificate Run Type with that code already exists")
      Else
        Dim vNewInstance As New ExamCertRunType(pEnv)
        vNewInstance.Code = pCode
        vNewInstance.Description = pDescription
        Return vNewInstance
      End If
    End Function

    Public Shared Function GetInstance(pEnv As CDBEnvironment,
                                       pCode As String) As ExamCertRunType
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      Else
        Dim vNewInstance As New ExamCertRunType(pEnv)
        vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ExamCertRunTypeColumns.ExamCertRunType).Name,
                                                                    pCode)}))
        Return If(vNewInstance.Existing, vNewInstance, Nothing)
      End If
    End Function

    Private Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
      Me.Init()
    End Sub

    Protected Overrides Sub AddFields()
      mvClassFields.Add("exam_cert_run_type")
      mvClassFields.Add("exam_cert_run_type_desc")

      mvClassFields.Item(ExamCertRunTypeColumns.ExamCertRunType).PrimaryKey = True
      mvClassFields.Item(ExamCertRunTypeColumns.ExamCertRunType).PrefixRequired = True
      mvClassFields.Item(ExamCertRunTypeColumns.ExamCertRunTypeDesc).PrefixRequired = True
    End Sub

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_cert_run_types"
      End Get
    End Property

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "xcrt"
      End Get
    End Property

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property

    Public Property Code() As String
      Get
        Return mvClassFields(ExamCertRunTypeColumns.ExamCertRunType).Value
      End Get
      Private Set(value As String)
        If String.IsNullOrWhiteSpace(value) Then
          Throw New ArgumentNullException("value")
        Else
          mvClassFields(ExamCertRunTypeColumns.ExamCertRunType).Value = value
        End If
      End Set
    End Property

    Public Property Description() As String
      Get
        Return mvClassFields(ExamCertRunTypeColumns.ExamCertRunTypeDesc).Value
      End Get
      Set(value As String)
        If String.IsNullOrWhiteSpace(value) Then
          Throw New ArgumentNullException("value")
        Else
          mvClassFields(ExamCertRunTypeColumns.ExamCertRunTypeDesc).Value = value
        End If
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamCertRunTypeColumns.AmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ExamCertRunTypeColumns.AmendedOn).Value)
      End Get
    End Property

    Public Overrides Sub Update(pParameterList As CDBParameters)
      Dim vUnoprocesedParameters As New CDBParameters
      For Each vParameter As CDBParameter In pParameterList
        Select vParameter.Name
          Case "ExamCertRunType"
            Throw New ArgumentException("Attempt to upate an immutable property", vParameter.Name)
          Case "ExamCertRunTypeDesc"
            Me.Description = vParameter.Value
          Case Else
            vUnoprocesedParameters.Add(vParameter)
        End Select
      Next vParameter
      MyBase.Update(vUnoprocesedParameters)
    End Sub

    Public Overrides Sub Delete(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      If New SQLStatement(mvEnv.Connection,
                          "exam_unit_cert_run_type_id",
                          "exam_unit_cert_run_types",
                          New CDBFields({New CDBField("exam_cert_run_type",
                                                     Me.Code)})).GetDataTable.Rows.Count > 0 Then
        Throw New InvalidOperationException("Attempt to delete an exam certificate run type that is used by a unit certificate run type")
      Else
        MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
      End If
    End Sub

    Public Overloads Function Equals(pOther As ExamCertRunType) As Boolean Implements IEquatable(Of ExamCertRunType).Equals
      Return If(pOther Is Nothing, False, pOther.Code = Me.Code)
    End Function

    Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
      Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ExamCertRunType) AndAlso Me.Equals(DirectCast(obj, ExamCertRunType))
    End Function

    Public Shared Operator =(ByVal pObj1 As ExamCertRunType,
                             ByVal pObj2 As ExamCertRunType) As Boolean
      Return Object.Equals(pObj1, pObj2)
    End Operator

    Public Shared Operator <>(ByVal pObj1 As ExamCertRunType,
                              ByVal pObj2 As ExamCertRunType) As Boolean
      Return Not (pObj1 = pObj2)
    End Operator

    Public Overrides Function GetHashCode() As Integer
      Return Me.Code.GetHashCode()
    End Function
  End Class

End Namespace