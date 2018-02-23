Imports System.Linq

Namespace Access

  Public Class ExamCertReprintType
    Inherits CARERecord
    Implements IEquatable(Of ExamCertReprintType)

    Public Enum ExamCertReprintTypeColumns
      AllFields = 0
      ExamCertReprintType
      ExamCertReprintTypeDesc
      UseOriginalCertNumber
      ExamCertNumberPrefix
      ExamCertNumberSuffix
      AllowReprint
      AmendedBy
      AmendedOn
    End Enum

    Public Shared Function CreateInstance(pEnv As CDBEnvironment,
                                          pCode As String,
                                          pDescription As String) As ExamCertReprintType
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      ElseIf Not String.IsNullOrWhiteSpace(pCode) AndAlso ExamCertReprintType.GetInstance(pEnv, pCode) IsNot Nothing Then
        Throw New InvalidOperationException("A Certificate Reprint Type with that code already exists")
      Else
        Dim vNewInstance As New ExamCertReprintType(pEnv)
        vNewInstance.Code = pCode
        vNewInstance.Description = pDescription
        Return vNewInstance
      End If
    End Function

    Public Shared Function GetInstance(pEnv As CDBEnvironment, pCode As String) As ExamCertReprintType
      If pEnv Is Nothing Then
        Throw New ArgumentNullException("pEnv")
      ElseIf String.IsNullOrWhiteSpace(pCode) Then
        Throw New ArgumentNullException("pCode")
      Else
        Dim vNewInstance As New ExamCertReprintType(pEnv)
        vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ExamCertReprintTypeColumns.ExamCertReprintType).Name,
                                                                    pCode)}))
        Return If(vNewInstance.Existing, vNewInstance, Nothing)
      End If
    End Function

    Private Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
      Me.Init()
    End Sub

    Protected Overrides Sub AddFields()
      mvClassFields.Add("exam_cert_reprint_type")
      mvClassFields.Add("exam_cert_reprint_type_desc")
      mvClassFields.Add("use_original_cert_number")
      mvClassFields.Add("exam_cert_number_prefix")
      mvClassFields.Add("exam_cert_number_suffix")
      mvClassFields.Add("allow_reprint")

      mvClassFields.Item(ExamCertReprintTypeColumns.ExamCertReprintType).PrimaryKey = True
      mvClassFields.Item(ExamCertReprintTypeColumns.ExamCertReprintType).PrefixRequired = True
      mvClassFields.Item(ExamCertReprintTypeColumns.ExamCertReprintTypeDesc).PrefixRequired = True
      mvClassFields.Item(ExamCertReprintTypeColumns.UseOriginalCertNumber).PrefixRequired = True
      mvClassFields.Item(ExamCertReprintTypeColumns.ExamCertNumberPrefix).PrefixRequired = True
      mvClassFields.Item(ExamCertReprintTypeColumns.ExamCertNumberSuffix).PrefixRequired = True
      mvClassFields.Item(ExamCertReprintTypeColumns.AllowReprint).PrefixRequired = True
    End Sub

    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "exam_cert_reprint_types"
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
        Return mvClassFields(ExamCertReprintTypeColumns.ExamCertReprintType).Value
      End Get
      Private Set(value As String)
        If String.IsNullOrWhiteSpace(value) Then
          Throw New ArgumentNullException("value")
        Else
          mvClassFields(ExamCertReprintTypeColumns.ExamCertReprintType).Value = value
        End If
      End Set
    End Property

    Public Property Description() As String
      Get
        Return mvClassFields(ExamCertReprintTypeColumns.ExamCertReprintTypeDesc).Value
      End Get
      Set(value As String)
        If String.IsNullOrWhiteSpace(value) Then
          Throw New ArgumentNullException("value")
        Else
          mvClassFields(ExamCertReprintTypeColumns.ExamCertReprintTypeDesc).Value = value
        End If
      End Set
    End Property

    Public Property UseOriginalNumber() As Boolean
      Get
        Return mvClassFields(ExamCertReprintTypeColumns.UseOriginalCertNumber).Value = "Y"
      End Get
      Set(value As Boolean)
        mvClassFields(ExamCertReprintTypeColumns.ExamCertReprintTypeDesc).Value = If(value, "Y", "N")
      End Set
    End Property

    Public Property CertificateNumberPrefix() As String
      Get
        Return mvClassFields(ExamCertReprintTypeColumns.ExamCertNumberPrefix).Value
      End Get
      Set(value As String)
        If String.IsNullOrWhiteSpace(value) Then
          Throw New ArgumentNullException("value")
        Else
          mvClassFields(ExamCertReprintTypeColumns.ExamCertNumberPrefix).Value = value
        End If
      End Set
    End Property

    Public Property CertificateNumberSuffix() As String
      Get
        Return mvClassFields(ExamCertReprintTypeColumns.ExamCertNumberSuffix).Value
      End Get
      Set(value As String)
        If String.IsNullOrWhiteSpace(value) Then
          Throw New ArgumentNullException("value")
        Else
          mvClassFields(ExamCertReprintTypeColumns.ExamCertNumberSuffix).Value = value
        End If
      End Set
    End Property

    Private mvAttributes As List(Of ExamCertReprintTypeItem) = Nothing
    Public ReadOnly Property Attributes() As IList(Of String)
      Get
        If mvAttributes Is Nothing Then
          mvAttributes = ExamCertReprintTypeItem.GetInstances(Me)
        End If
        Return New List(Of String)((From vAttribute As ExamCertReprintTypeItem In ExamCertReprintTypeItem.GetInstances(Me)
                                    Select vAttribute.Attribute).AsEnumerable).AsReadOnly
      End Get
    End Property

    Public Sub AddAttribute(pAttribute As String)
      Me.mvAttributes.Add(ExamCertReprintTypeItem.CreateInstance(Me, pAttribute))
    End Sub

    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(ExamCertReprintTypeColumns.AmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As Date
      Get
        Return Date.Parse(mvClassFields(ExamCertReprintTypeColumns.AmendedOn).Value)
      End Get
    End Property

    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer)
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      If mvAttributes IsNot Nothing Then
        For Each vAttribute As ExamCertReprintTypeItem In mvAttributes
          vAttribute.Save(pAmendedBy, pAudit, pJournalNumber)
        Next vAttribute
      End If
    End Sub

    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer, pForceAmendmentHistory As Boolean)
      MyBase.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
      If mvAttributes IsNot Nothing Then
        For Each vAttribute As ExamCertReprintTypeItem In mvAttributes
          vAttribute.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory)
        Next vAttribute
      End If
    End Sub

    Public Overrides Sub Update(pParameterList As CDBParameters)
      Dim vUnoprocesedParameters As New CDBParameters
      For Each vParameter As CDBParameter In pParameterList
        Select Case vParameter.Name
          Case "ExamCertReprintType"
            Throw New ArgumentException("Attempt to upate an immutable property", vParameter.Name)
          Case "ExamCertReprintTypeDesc"
            Me.Description = vParameter.Value
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
                          New CDBFields({New CDBField("exam_cert_reprint_type",
                                                     Me.Code)})).GetDataTable.Rows.Count > 0 Then
        Throw New InvalidOperationException("Attempt to delete an exam cwertificate reprint type that is used by a certificate run")
      Else
        MyBase.Delete(pAmendedBy, pAudit, pJournalNumber)
      End If
    End Sub

    Public Overloads Function Equals(pOther As ExamCertReprintType) As Boolean Implements IEquatable(Of ExamCertReprintType).Equals
      Return If(pOther Is Nothing, False, pOther.Code = Me.Code)
    End Function

    Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
      Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ExamCertReprintType) AndAlso Me.Equals(DirectCast(obj, ExamCertReprintType))
    End Function

    Public Shared Operator =(ByVal pObj1 As ExamCertReprintType, ByVal pObj2 As ExamCertReprintType) As Boolean
      Return Object.Equals(pObj1, pObj2)
    End Operator

    Public Shared Operator <>(ByVal pObj1 As ExamCertReprintType, ByVal pObj2 As ExamCertReprintType) As Boolean
      Return Not (pObj1 = pObj2)
    End Operator

    Public Overrides Function GetHashCode() As Integer
      Return Me.Code.GetHashCode()
    End Function

    Private Class ExamCertReprintTypeItem
      Inherits CARERecord
      Implements IEquatable(Of ExamCertReprintTypeItem)

      Public Enum ExamCertReprintTypeItemColumns
        AllFields = 0
        ExamCertReprintType
        ExamCertAttribute
        AmendedBy
        AmendedOn
      End Enum

      Friend Shared Function CreateInstance(pReprintType As ExamCertReprintType,
                                            pAttribute As String) As ExamCertReprintTypeItem

        If pReprintType Is Nothing Then
          Throw New ArgumentNullException("pReprintType")
        ElseIf Not String.IsNullOrWhiteSpace(pAttribute) AndAlso ExamCertReprintTypeItem.GetInstance(pReprintType, pAttribute) IsNot Nothing Then
          Throw New InvalidOperationException("That attribute is already listed for that Certificate Reprint Type")
7:      Else
          Dim vNewInstance As New ExamCertReprintTypeItem(pReprintType.Environment)
          vNewInstance.ReprintType = pReprintType
          vNewInstance.Attribute = pAttribute
          Return vNewInstance
        End If
      End Function

      Friend Shared Function GetInstance(pReprintType As ExamCertReprintType, pAttribute As String) As ExamCertReprintTypeItem
        If pReprintType Is Nothing Then
          Throw New ArgumentNullException("pReprintType")
        Else
          Dim vNewInstance As New ExamCertReprintTypeItem(pReprintType.Environment)
          vNewInstance.InitWithPrimaryKey(New CDBFields({New CDBField(vNewInstance.mvClassFields(ExamCertReprintTypeItemColumns.ExamCertReprintType).Name,
                                                                      pReprintType.Code),
                                                         New CDBField(vNewInstance.mvClassFields(ExamCertReprintTypeItemColumns.ExamCertAttribute).Name,
                                                                      pAttribute)}))
          Return If(vNewInstance.Existing, vNewInstance, Nothing)
        End If
      End Function

      Friend Shared Function GetInstances(pReprintType As ExamCertReprintType) As List(Of ExamCertReprintTypeItem)
        If pReprintType Is Nothing Then
          Throw New ArgumentNullException("pReprintType")
        Else
          Dim vAttributeList As New List(Of ExamCertReprintTypeItem)
          For Each vRow As DataRow In New SQLStatement(pReprintType.Environment.Connection,
                                                       "exam_cert_reprint_type,exam_cert_attribute,amended_by,amended_on",
                                                       "exam_cert_reprint_type_items",
                                                       New CDBField("exam_cert_reprint_type",
                                                                    pReprintType.Code)).GetDataTable.AsEnumerable
            Dim vNewInstance As New ExamCertReprintTypeItem(pReprintType.Environment)
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
        mvClassFields.Add("exam_cert_reprint_type")
        mvClassFields.Add("exam_cert_attribute")

        mvClassFields.Item(ExamCertReprintTypeItemColumns.ExamCertReprintType).PrimaryKey = True
        mvClassFields.Item(ExamCertReprintTypeItemColumns.ExamCertReprintType).PrefixRequired = True
        mvClassFields.Item(ExamCertReprintTypeItemColumns.ExamCertAttribute).PrimaryKey = True
        mvClassFields.Item(ExamCertReprintTypeItemColumns.ExamCertAttribute).PrefixRequired = True
      End Sub

      Protected Overrides ReadOnly Property DatabaseTableName() As String
        Get
          Return "exam_cert_reprint_type_items"
        End Get
      End Property

      Protected Overrides ReadOnly Property TableAlias() As String
        Get
          Return "xcrti"
        End Get
      End Property

      Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
        Get
          Return True
        End Get
      End Property

      Private mvReprintType As ExamCertReprintType = Nothing
      Friend Property ReprintType() As ExamCertReprintType
        Get
          If mvReprintType Is Nothing Then
            mvReprintType = ExamCertReprintType.GetInstance(mvEnv, mvClassFields(ExamCertReprintTypeItemColumns.ExamCertReprintType).Value)
          End If
          Return mvReprintType
        End Get
        Private Set(value As ExamCertReprintType)
          If value Is Nothing Then
            Throw New ArgumentNullException("value")
          Else
            mvReprintType = value
            mvClassFields(ExamCertReprintTypeItemColumns.ExamCertReprintType).Value = value.Code
          End If
        End Set
      End Property

      Friend Property Attribute() As String
        Get
          Return mvClassFields(ExamCertReprintTypeItemColumns.ExamCertAttribute).Value
        End Get
        Set(value As String)
          If String.IsNullOrWhiteSpace(value) Then
            Throw New ArgumentNullException("value")
          Else
            mvClassFields(ExamCertReprintTypeItemColumns.ExamCertAttribute).Value = value
          End If
        End Set
      End Property

      Friend ReadOnly Property AmendedBy() As String
        Get
          Return mvClassFields(ExamCertReprintTypeColumns.AmendedBy).Value
        End Get
      End Property

      Friend ReadOnly Property AmendedOn() As Date
        Get
          Return Date.Parse(mvClassFields(ExamCertReprintTypeColumns.AmendedOn).Value)
        End Get
      End Property

      Friend Overloads Function Equals(pOther As ExamCertReprintTypeItem) As Boolean Implements IEquatable(Of ExamCertReprintTypeItem).Equals
        Return If(pOther Is Nothing, False, pOther.ReprintType = Me.ReprintType And pOther.Attribute = Me.Attribute)
      End Function

      Public NotOverridable Overrides Function Equals(obj As Object) As Boolean
        Return obj IsNot Nothing AndAlso obj.GetType Is GetType(ExamCertReprintType) AndAlso Me.Equals(DirectCast(obj, ExamCertReprintType))
      End Function

      Public Shared Operator =(ByVal pObj1 As ExamCertReprintTypeItem, ByVal pObj2 As ExamCertReprintTypeItem) As Boolean
        Return Object.Equals(pObj1, pObj2)
      End Operator

      Public Shared Operator <>(ByVal pObj1 As ExamCertReprintTypeItem, ByVal pObj2 As ExamCertReprintTypeItem) As Boolean
        Return Not (pObj1 = pObj2)
      End Operator

      Public Overrides Function GetHashCode() As Integer
        Return Me.ReprintType.GetHashCode() Xor Me.Attribute.GetHashCode
      End Function
    End Class
  End Class

End Namespace