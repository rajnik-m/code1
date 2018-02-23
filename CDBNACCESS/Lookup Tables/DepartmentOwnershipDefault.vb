Namespace Access

  Public Class DepartmentOwnershipDefault
    Inherits CARERecord

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum DepartmentOwnershipDefaultFields
      AllFields = 0
      Department
      OwnershipGroup
      OwnershipAccessLevel
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("department")
        .Add("ownership_group")
        .Add("ownership_access_level")

        .Item(DepartmentOwnershipDefaultFields.Department).PrimaryKey = True

        .Item(DepartmentOwnershipDefaultFields.OwnershipGroup).PrimaryKey = True
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "dod"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "department_ownership_defaults"
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
    Public ReadOnly Property Department() As String
      Get
        Return mvClassFields(DepartmentOwnershipDefaultFields.Department).Value
      End Get
    End Property
    Public ReadOnly Property OwnershipGroup() As String
      Get
        Return mvClassFields(DepartmentOwnershipDefaultFields.OwnershipGroup).Value
      End Get
    End Property
    Public ReadOnly Property OwnershipAccessLevel() As CDBEnvironment.OwnershipAccessLevelTypes
      Get
        Return CDBEnvironment.GetOwnershipAccessLevel(mvClassFields(DepartmentOwnershipDefaultFields.OwnershipAccessLevel).Value)
      End Get
    End Property
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(DepartmentOwnershipDefaultFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(DepartmentOwnershipDefaultFields.AmendedOn).Value
      End Get
    End Property
#End Region

    Public Sub AddForUser(ByVal pDept As String, ByVal pLogname As String)
      Dim vRecordSet As CDBRecordSet
      Dim vDOD As New DepartmentOwnershipDefault(mvEnv)
      Dim vOGU As New OwnershipGroupUser(mvEnv)

      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields() & " FROM " & mvClassFields.TableNameAndAlias & " WHERE department = '" & pDept & "'")
      While vRecordSet.Fetch
        vDOD.InitFromRecordSet(vRecordSet)
        vOGU.Init()
        vOGU.InitFromDepartment(vDOD.OwnershipGroup, pLogname, vDOD.OwnershipAccessLevel)
        vOGU.Save()
      End While
      vRecordSet.CloseRecordSet()

      'Now find all ownership groups with no departmental default and add browse access
      'This should not be done if we want to support NO ownership of some contacts
      vRecordSet = mvEnv.Connection.GetRecordSet("SELECT ownership_group FROM ownership_groups WHERE ownership_group NOT IN( SELECT dod.ownership_group FROM " & mvClassFields.TableNameAndAlias & " WHERE department = '" & pDept & "')")
      While vRecordSet.Fetch
        vOGU.Init()
        vOGU.InitFromDepartment(vRecordSet.Fields("ownership_group").Value, pLogname, CDBEnvironment.OwnershipAccessLevelTypes.oaltBrowse)
        vOGU.Save()
      End While
      vRecordSet.CloseRecordSet()
    End Sub

  End Class
End Namespace
