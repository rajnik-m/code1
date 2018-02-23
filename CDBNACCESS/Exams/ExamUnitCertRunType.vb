Namespace Access

  Public Class ExamUnitCertRunType
    Inherits CARERecord
    Implements IEquatable(Of ExamUnitCertRunType)

    Public Enum ExamUnitCertRunTypeColumns
      AllFields = 0
      ExamUnitCertRunTypeId
      ExamUnitLinkId
      ExamCertRunType
      IncludeView
      ExcludeView
      StandardDocument
      AmendedBy
      AmendedOn
    End Enum

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pExamUnitLinkId As Integer,
                                          pRunType As ExamCertRunType,
                                          pDocument As String) As ExamUnitCertRunType
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      ElseIf ExamUnitCertRunType.GetInstance(pEnv, pExamUnitLinkId, pRunType) IsNot Nothing Then
        Throw New InvalidOperationException("That certificate run type is already defined for that exam unit link.")
      Else
        Dim vNewInstance As New ExamUnitCertRunType(pEnv)
        vNewInstance.ExamUnitLinkId = pExamUnitLinkId
        vNewInstance.RunType = pRunType
        vNewInstance.Document = pDocument
        Return vNewInstance
      End If
    End Function

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pRunTypeId As Integer) As ExamUnitCertRunType
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      Else
        Dim vNewInstance As New ExamUnitCertRunType(pEnv)
        vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ExamUnitCertRunTypeColumns.ExamUnitCertRunTypeId).Name,
                                                                      pRunTypeId)}))
        Return If(vNewInstance.Existing, vNewInstance, Nothing)
      End If
    End Function

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pExamUnitLinkId As Integer, pRunType As ExamCertRunType) As ExamUnitCertRunType
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      Else
        Dim vNewInstance As New ExamUnitCertRunType(pEnv)
        vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ExamUnitCertRunTypeColumns.ExamUnitLinkId).Name,
                                                                      pExamUnitLinkId),
                                                       New CDBField(vNewInstance.mvClassFields(ExamUnitCertRunTypeColumns.ExamCertRunType).Name,
                                                                      pRunType.Code)}))
        Return If(vNewInstance.Existing, vNewInstance, Nothing)
      End If
    End Function

    Private Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
      Me.Init()
    End Sub

    Protected Overrides Sub AddFields()
      mvClassFields.Add("exam_unit_cert_run_type_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
      mvClassFields.Add("exam_cert_run_type")
      mvClassFields.Add("include_view")
      mvClassFields.Add("exclude_view")
      mvClassFields.Add("standard_document")

      mvClassFields.Item(ExamUnitCertRunTypeColumns.ExamUnitCertRunTypeId).PrimaryKey = True
      mvClassFields.Item(ExamUnitCertRunTypeColumns.ExamUnitLinkId).PrefixRequired = True
      mvClassFields.Item(ExamUnitCertRunTypeColumns.ExamCertRunType).PrefixRequired = True
      mvClassFields.Item(ExamUnitCertRunTypeColumns.IncludeView).PrefixRequired = True
      mvClassFields.Item(ExamUnitCertRunTypeColumns.ExcludeView).PrefixRequired = True
      mvClassFields.Item(ExamUnitCertRunTypeColumns.StandardDocument).PrefixRequired = True
      mvClassFields.SetControlNumberField(ExamUnitCertRunTypeColumns.ExamUnitCertRunTypeId, "XCR")
    End Sub

    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "xucr"
      End Get
    End Property

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_unit_cert_run_types"
      End Get
    End Property

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property

    Public ReadOnly Property Id() As Integer
      Get
        Return mvClassFields(ExamUnitCertRunTypeColumns.ExamUnitCertRunTypeId).IntegerValue
      End Get
    End Property

    Public Property ExamUnitLinkId() As Integer
      Get
        Return mvClassFields(ExamUnitCertRunTypeColumns.ExamUnitLinkId).IntegerValue
      End Get
      Private Set(value As Integer)
        If New SQLStatement(mvEnv.Connection,
                           "Count(*)",
                           "exam_unit_links",
                           New CDBFields({New CDBField("exam_unit_link_id",
                                                       value)})).GetDataTable.Rows.Count < 1 Then
          Throw New ArgumentException("No Exam Unit Link with the specified ID exists.")
        Else
          mvClassFields(ExamUnitCertRunTypeColumns.ExamUnitLinkId).IntegerValue = value
        End If
      End Set
    End Property

    Private mvRunType As ExamCertRunType = Nothing
    Public Property RunType() As ExamCertRunType
      Get
        If mvRunType Is Nothing Then
          mvRunType = ExamCertRunType.GetInstance(mvEnv, mvClassFields(ExamUnitCertRunTypeColumns.ExamCertRunType).Value)
        End If
        Return mvRunType
      End Get
      Private Set(value As ExamCertRunType)
        If value Is Nothing Then
          Throw New ArgumentNullException("value")
        ElseIf Not value.Existing Then
          Throw New ArgumentException("Entity must be persisted before it can be used here.", "pExamCertReprintType")
        Else
          mvRunType = value
          mvClassFields(ExamUnitCertRunTypeColumns.ExamCertRunType).Value = value.Code
        End If
      End Set
    End Property

    Public Property IncludeView() As String
      Get
        Return mvClassFields(ExamUnitCertRunTypeColumns.IncludeView).Value
      End Get
      Set(value As String)
        If value.Length > 0 AndAlso
           New SQLStatement(mvEnv.Connection,
                            "Count(*)",
                            "view_names",
                            New CDBFields({New CDBField("view_name",
                                                        value),
                                           New CDBField("view_type",
                                                        "M")})).GetDataTable.Rows.Count < 1 Then
          Throw New ArgumentException("No view with the specified name and a type of ""M"" exists.")
        Else
          If Not String.IsNullOrWhiteSpace(value) AndAlso Not mvEnv.Connection.AttributeExists(value, "exam_student_unit_header_id") Then 'BR20762 do not execute Query code here, just check that the field exists.
            Throw New ArgumentException("The specified include view does not contain the column ""exam_student_unit_header_id.")
          End If
          mvClassFields(ExamUnitCertRunTypeColumns.IncludeView).Value = value
        End If
      End Set
    End Property

    Public Property ExcludeView() As String
      Get
        Return mvClassFields(ExamUnitCertRunTypeColumns.ExcludeView).Value
      End Get
      Set(value As String)
        If value.Length > 0 AndAlso
           New SQLStatement(mvEnv.Connection,
                            "Count(*)",
                            "view_names",
                            New CDBFields({New CDBField("view_name",
                                                        value),
                                           New CDBField("view_type",
                                                        "M")})).GetDataTable.Rows.Count < 1 Then
          Throw New ArgumentException("No view with the specified name and a type of ""M"" exists.")
        Else
          If Not String.IsNullOrWhiteSpace(value) AndAlso
             Not New SQLStatement(mvEnv.Connection,
                                  "*",
                                  value).GetDataTable.Columns.Contains("exam_student_unit_header_id") Then
            Throw New ArgumentException("The specified exclude view does not contain the column ""exam_student_unit_header_id"".")
          End If
          mvClassFields(ExamUnitCertRunTypeColumns.ExcludeView).Value = value
        End If
      End Set
    End Property

    Public Property Document() As String
      Get
        Return mvClassFields(ExamUnitCertRunTypeColumns.StandardDocument).Value
      End Get
      Set(value As String)
        If New SQLStatement(mvEnv.Connection,
                            "Count(*)",
                            "standard_documents",
                            New CDBFields({New CDBField("standard_document",
                                                        value)})).GetDataTable.Rows.Count < 1 Then
          Throw New ArgumentException("No Standard Document with the specified code exists.")
        Else
          mvClassFields(ExamUnitCertRunTypeColumns.StandardDocument).Value = value
        End If
      End Set
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamUnitCertRunTypeColumns.AmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ExamUnitCertRunTypeColumns.AmendedOn).Value)
      End Get
    End Property

    Public Overrides Sub Update(pParameterList As CDBParameters)
      Dim vUnoprocesedParameters As New CDBParameters
      For Each vParameter As CDBParameter In pParameterList
        Select Case vParameter.Name
          Case "ExamUnitCertRunTypeId"
            If vParameter.IntegerValue <> Me.Id Then
              Throw New ArgumentException("Attempt to update an immutable property", vParameter.Name)
            End If
          Case "ExamUnitLinkId"
            If vParameter.IntegerValue <> Me.ExamUnitLinkId Then
              Throw New ArgumentException("Attempt to upate an immutable property", vParameter.Name)
            End If
          Case "ExamCertRunType"
            If vParameter.Value <> Me.RunType.Code Then
              Throw New ArgumentException("Attempt to upate an immutable property", vParameter.Name)
            End If
          Case "IncludeView"
            Me.IncludeView = vParameter.Value
          Case "ExcludeView"
            Me.ExcludeView = vParameter.Value
          Case "StandardDocument"
            Me.Document = vParameter.Value
          Case Else
            vUnoprocesedParameters.Add(vParameter)
        End Select
      Next vParameter
      MyBase.Update(vUnoprocesedParameters)
    End Sub

    Public Overrides Sub Delete(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      If New SQLStatement(mvEnv.Connection,
                          "exam_cert_run_id",
                          "exam_cert_runs",
                          New CDBFields({New CDBField("exam_unit_cert_run_type_id",
                                                     Me.Id)})).GetDataTable.Rows.Count > 0 Then
        Throw New InvalidOperationException("Attempt to delete an exam unit certificate run type that is used by a certificate run")
      Else
        MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
      End If
    End Sub

    Public Overloads Function Equals(pOther As ExamUnitCertRunType) As Boolean Implements IEquatable(Of ExamUnitCertRunType).Equals
      Return If(pOther Is Nothing, False, pOther.Id = Me.Id)
    End Function

    Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
      Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ExamUnitCertRunType) AndAlso Me.Equals(DirectCast(obj, ExamUnitCertRunType))
    End Function

    Public Shared Operator =(ByVal pObj1 As ExamUnitCertRunType, ByVal pObj2 As ExamUnitCertRunType) As Boolean
      Return Object.Equals(pObj1, pObj2)
    End Operator

    Public Shared Operator <>(ByVal pObj1 As ExamUnitCertRunType, ByVal pObj2 As ExamUnitCertRunType) As Boolean
      Return Not (pObj1 = pObj2)
    End Operator

    Public Overrides Function GetHashCode() As Integer
      Return Me.Id.GetHashCode()
    End Function
  End Class

End Namespace