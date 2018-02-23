Namespace Access

  Public Class EntityGroups
    Inherits CollectionList(Of EntityGroup)

    Private mvInvalidConGroup As Boolean
    Private mvInvalidOrgGroup As Boolean
    Private mvEventGroupCount As Integer

    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(1)
      Dim vWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet

      Dim vConn As CDBConnection = pEnv.Connection

      'Read all contact groups
      Dim vLastConGroup As EntityGroup = Nothing
      Dim vContactGroup As ContactGroup = New ContactGroup(pEnv)
      vRecordSet = vContactGroup.SelectGroupsSQL.GetRecordSet
      While vRecordSet.Fetch()
        vContactGroup.InitFromRecordSet(vRecordSet)
        If vContactGroup.EntityGroupDesc = "Events" Then
          mvInvalidConGroup = True
        ElseIf vContactGroup.EntityGroupCode = OrganisationGroup.DefaultGroupCode Then
          mvInvalidConGroup = True
        Else
          If ContainsKey(vContactGroup.EntityGroupCode) Then
            mvInvalidConGroup = True
          Else
            Add(vContactGroup.EntityGroupCode, vContactGroup)
            vLastConGroup = vContactGroup
            vContactGroup = New ContactGroup(pEnv)
          End If
        End If
      End While
      vRecordSet.CloseRecordSet()
      If Not ContainsKey(ContactGroup.DefaultGroupCode) Then
        vContactGroup = New ContactGroup(pEnv)  'Setup a default group
        vContactGroup.Init()
        Insert(1, vContactGroup.EntityGroupCode, vContactGroup)  'Now add it before the others
        vLastConGroup = vContactGroup
      End If

      'Read all organisation groups
      Dim vLastOrgGroup As EntityGroup = Nothing
      Dim vOrganisationGroup As OrganisationGroup = New OrganisationGroup(pEnv)
      vRecordSet = vOrganisationGroup.SelectGroupsSQL.GetRecordSet
      While vRecordSet.Fetch()
        vOrganisationGroup.InitFromRecordSet(vRecordSet)
        If vOrganisationGroup.EntityGroupDesc = "Events" Then
          mvInvalidOrgGroup = True
        ElseIf vOrganisationGroup.EntityGroupCode = ContactGroup.DefaultGroupCode Then
          mvInvalidOrgGroup = True
        Else
          If ContainsKey(vOrganisationGroup.EntityGroupCode) Then
            mvInvalidOrgGroup = True
          Else
            Add(vOrganisationGroup.EntityGroupCode, vOrganisationGroup)
            vLastOrgGroup = vOrganisationGroup
            vOrganisationGroup = New OrganisationGroup(pEnv)
          End If
        End If
      End While
      vRecordSet.CloseRecordSet()

      If Not ContainsKey(OrganisationGroup.DefaultGroupCode) Then
        vOrganisationGroup = New OrganisationGroup(pEnv)  'Setup a default group
        vOrganisationGroup.Init()
        Insert(IndexOf(vLastConGroup) + 1, vOrganisationGroup.EntityGroupCode, vOrganisationGroup)    'Now add it before the others
        vLastOrgGroup = vOrganisationGroup
      End If

      If pEnv.GetConfigOption("option_events", False) Then
        'Add an events entity group
        Dim vEventGroup As New EventGroup(pEnv)
        vRecordSet = vEventGroup.SelectGroupsSQL.GetRecordSet
        While vRecordSet.Fetch()
          vEventGroup.InitFromRecordSet(vRecordSet)
          If vEventGroup.EntityGroupCode = OrganisationGroup.DefaultGroupCode Then
            mvInvalidOrgGroup = True
          ElseIf vEventGroup.EntityGroupCode = ContactGroup.DefaultGroupCode Then
            mvInvalidOrgGroup = True
          Else
            If ContainsKey(vEventGroup.EntityGroupCode) Then
              mvInvalidOrgGroup = True
            Else
              Add(vEventGroup.EntityGroupCode, vEventGroup)
              mvEventGroupCount += 1
              vEventGroup = New EventGroup(pEnv)
            End If
          End If
        End While
        vRecordSet.CloseRecordSet()
        If Not ContainsKey(EventGroup.DefaultGroupCode) Then
          vEventGroup = New EventGroup(pEnv)    'Setup a default group
          vEventGroup.Init()
          Insert(IndexOf(vLastOrgGroup) + 1, vEventGroup.EntityGroupCode, vEventGroup)  'Now add it before the others
          mvEventGroupCount += 1
        End If
      End If
      For Each vEntityGroup As EntityGroup In Me
        vEntityGroup.GetColorPreference()
      Next
    End Sub

    Public ReadOnly Property EventGroupCount() As Integer
      Get
        Return mvEventGroupCount
      End Get
    End Property

    Public ReadOnly Property GroupFromCode(ByVal pType As EntityGroup.EntityGroupTypes, ByVal pCode As String) As EntityGroup
      Get
        If MyBase.ContainsKey(pCode) Then
          Return MyBase.Item(pCode)
        Else
          Return DefaultGroup(pType)
        End If
      End Get
    End Property

    Public ReadOnly Property DefaultGroup(ByVal pType As EntityGroup.EntityGroupTypes) As EntityGroup
      Get
        For Each vEntityGroup As EntityGroup In Me
          If vEntityGroup.EntityGroupType = pType And vEntityGroup.DefaultGroup Then
            Return vEntityGroup
          End If
        Next vEntityGroup
        Return Nothing
      End Get
    End Property
  End Class
End Namespace