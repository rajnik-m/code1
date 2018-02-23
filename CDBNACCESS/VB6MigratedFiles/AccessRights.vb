Namespace Access
  Public Class AccessRights

    Public Enum DocumentAccessRights
      darHeader = 1
      darView = 2
      darPrint = 4
      darEdit = 8
      darDelete = 16
      darEditHeader = 32
      darEditClass = 64
      darAllRights = 127
    End Enum

    Private mvEnv As CDBEnvironment

    Public Sub Init(ByRef pEnv As CDBEnvironment)
      mvEnv = pEnv
    End Sub

    Public Function GetDocumentRights(ByVal pDocumentNumber As Integer) As DocumentAccessRights
      Dim vRecordSet As CDBRecordSet
      Dim vRights As DocumentAccessRights

      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT department, created_by, cl.document_class, " & AccessAttributes() & " FROM communications_log cl, document_classes dc WHERE communications_log_number = " & pDocumentNumber & " AND cl.document_class = dc.document_class")
      If vRecordSet.Fetch() = True Then
        vRights = GetAccessRights(vRecordSet, vRecordSet.Fields("department").Value, vRecordSet.Fields("created_by").Value)
        vRights = GetAdditionalDocumentRights(vRecordSet, vRights, pDocumentNumber)
      End If
      vRecordSet.CloseRecordSet()
      GetDocumentRights = vRights
    End Function

    Function GetAdditionalDocumentRights(ByRef pRecordSet As CDBRecordSet, ByVal pRights As DocumentAccessRights, ByVal pDocumentNumber As Integer) As DocumentAccessRights
      Dim vRights As DocumentAccessRights
      Dim vType As String
      Dim vUserNumber As Integer
      Dim vRecordSet As CDBRecordSet

      vUserNumber = mvEnv.User.ContactNumber
      vRights = pRights
      If vRights <> DocumentAccessRights.darAllRights And vUserNumber > 0 Then
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT link_type FROM communications_log_links WHERE communications_log_number = " & pDocumentNumber & " AND contact_number = " & vUserNumber & " AND link_type IN('A','C','D')")
        While vRecordSet.Fetch() = True
          vType = vRecordSet.Fields(1).Value
          With pRecordSet
            Select Case vType
              Case "A"
                vRights = vRights Or (DocumentAccessRights.darView Or DocumentAccessRights.darHeader)
                If .Fields("addressee_edit_header").Bool Then vRights = vRights Or DocumentAccessRights.darEditHeader
                If .Fields("addressee_edit_class").Bool Then vRights = vRights Or DocumentAccessRights.darEditClass
                If .Fields("addressee_print").Bool Then vRights = vRights Or DocumentAccessRights.darPrint
                If .Fields("addressee_edit").Bool Then vRights = vRights Or DocumentAccessRights.darEdit
                If .Fields("addressee_delete").Bool Then vRights = vRights Or DocumentAccessRights.darDelete
              Case "C"
                vRights = vRights Or (DocumentAccessRights.darView Or DocumentAccessRights.darHeader)
                If .Fields("copied_edit_header").Bool Then vRights = vRights Or DocumentAccessRights.darEditHeader
                If .Fields("copied_edit_class").Bool Then vRights = vRights Or DocumentAccessRights.darEditClass
                If .Fields("copied_print").Bool Then vRights = vRights Or DocumentAccessRights.darPrint
                If .Fields("copied_edit").Bool Then vRights = vRights Or DocumentAccessRights.darEdit
                If .Fields("copied_delete").Bool Then vRights = vRights Or DocumentAccessRights.darDelete
              Case "D"
                vRights = vRights Or (DocumentAccessRights.darView Or DocumentAccessRights.darHeader)
                If .Fields("distributed_edit_header").Bool Then vRights = vRights Or DocumentAccessRights.darEditHeader
                If .Fields("distributed_edit_class").Bool Then vRights = vRights Or DocumentAccessRights.darEditClass
                If .Fields("distributed_print").Bool Then vRights = vRights Or DocumentAccessRights.darPrint
                If .Fields("distributed_edit").Bool Then vRights = vRights Or DocumentAccessRights.darEdit
                If .Fields("distributed_delete").Bool Then vRights = vRights Or DocumentAccessRights.darDelete
            End Select
          End With
        End While
        vRecordSet.CloseRecordSet()
      End If
      GetAdditionalDocumentRights = vRights
    End Function

    Public Function AccessAttributes() As String
      Dim vAttrs As String

      vAttrs = "creator_header,creator_view,creator_print,creator_edit,creator_delete, department_header,department_view,department_print,department_edit,department_delete, public_header,public_view,public_print,public_edit,public_delete"
      vAttrs = vAttrs & ",creator_edit_header,creator_edit_class,department_edit_header,department_edit_class,public_edit_header,public_edit_class"
      vAttrs = vAttrs & ",addressee_edit_header,addressee_edit_class,addressee_print,addressee_edit,addressee_delete"
      vAttrs = vAttrs & ",copied_edit_header,copied_edit_class,copied_print,copied_edit,copied_delete"
      vAttrs = vAttrs & ",distributed_edit_header,distributed_edit_class,distributed_print,distributed_edit,distributed_delete"
      AccessAttributes = vAttrs
    End Function

    Public Function GetAccessRights(ByRef pRecordSet As CDBRecordSet, ByVal pDepartment As String, ByVal pCreator As String) As DocumentAccessRights
      Dim vAccessRights As DocumentAccessRights

      With pRecordSet
        If pCreator = mvEnv.User.Logname Then
          'read in the creators access rights for this class
          If .Fields("creator_header").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darHeader
          If .Fields("creator_view").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darView
          If .Fields("creator_print").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darPrint
          If .Fields("creator_edit").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEdit
          If .Fields("creator_delete").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darDelete
          If .Fields("creator_edit_header").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEditHeader
          If .Fields("creator_edit_class").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEditClass
        ElseIf pDepartment = mvEnv.User.Department Then
          'read in departments access rights for this class
          If .Fields("department_header").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darHeader
          If .Fields("department_view").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darView
          If .Fields("department_print").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darPrint
          If .Fields("department_edit").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEdit
          If .Fields("department_delete").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darDelete
          If .Fields("department_edit_header").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEditHeader
          If .Fields("department_edit_class").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEditClass
        Else
          'read in public access rights for this class
          If .Fields("public_header").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darHeader
          If .Fields("public_view").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darView
          If .Fields("public_print").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darPrint
          If .Fields("public_edit").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEdit
          If .Fields("public_delete").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darDelete
          If .Fields("public_edit_header").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEditHeader
          If .Fields("public_edit_class").Bool Then vAccessRights = vAccessRights Or DocumentAccessRights.darEditClass
        End If
        If mvEnv.User.AccessLevel = CDBUser.UserAccessLevel.ualReadOnly Then vAccessRights = vAccessRights And (DocumentAccessRights.darHeader Or DocumentAccessRights.darView Or DocumentAccessRights.darPrint)
      End With
      GetAccessRights = vAccessRights
    End Function

    Public Function GetClassRights(ByVal pDocumentClass As String, ByVal pDepartment As String, ByVal pCreatedBy As String) As DocumentAccessRights
      Dim vRecordSet As CDBRecordSet

      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & AccessAttributes() & " FROM document_classes WHERE document_class = '" & pDocumentClass & "'")
      If vRecordSet.Fetch() = True Then GetClassRights = GetAccessRights(vRecordSet, pDepartment, pCreatedBy)
      vRecordSet.CloseRecordSet()
    End Function
  End Class
End Namespace
