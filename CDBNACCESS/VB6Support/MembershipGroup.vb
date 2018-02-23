Namespace Access

  Partial Public Class MembershipGroup

    Public Enum MembershipGroupRecordSetTypes 'These are bit values
      mgrtAll = &HFFS
      'ADD additional recordset types here
      mgrtBranch = &H100S
    End Enum

    Private mvBranchCode As String
    Private mvBranchSet As Boolean

    Protected Overrides Sub ClearFields()
      mvBranchCode = ""
      mvBranchSet = False
    End Sub

    Public Overloads Function GetRecordSetFields(ByVal pRSType As MembershipGroupRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If (pRSType And MembershipGroupRecordSetTypes.mgrtAll) = MembershipGroupRecordSetTypes.mgrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "mg")
      End If
      If (pRSType And MembershipGroupRecordSetTypes.mgrtBranch) = MembershipGroupRecordSetTypes.mgrtBranch Then vFields = vFields & ", b.branch"
      Return vFields
    End Function

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As MembershipGroupRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(MembershipGroupFields.MembershipGroupNumber, vFields)
        '.SetItem mgfOrganisationNumber, vFields
        'Modify below to handle each recordset type as required
        If (pRSType And MembershipGroupRecordSetTypes.mgrtAll) = MembershipGroupRecordSetTypes.mgrtAll Then
          .SetItem(MembershipGroupFields.MembershipNumber, vFields)
          .SetItem(MembershipGroupFields.OrganisationNumber, vFields)
          .SetItem(MembershipGroupFields.DefaultGroup, vFields)
          .SetItem(MembershipGroupFields.ValidFrom, vFields)
          .SetItem(MembershipGroupFields.ValidTo, vFields)
          .SetItem(MembershipGroupFields.IsCurrent, vFields)
          .SetItem(MembershipGroupFields.AmendedBy, vFields)
          .SetItem(MembershipGroupFields.AmendedOn, vFields)
        End If
        If (pRSType And MembershipGroupRecordSetTypes.mgrtBranch) = MembershipGroupRecordSetTypes.mgrtBranch Then
          mvBranchCode = vFields("branch").Value
          mvBranchSet = True 'Branch code could be null so use this flag to show that it was selected
        End If
      End With
    End Sub


    Friend Sub SetHistoric(ByVal pNewOrganisationNumber As Integer)
      'Only used when this is the DefaultGroup
      'Set this record as no longer current and add MembershipGroupHistory
      mvClassFields.Item(MembershipGroupFields.ValidTo).Value = TodaysDate()
      mvClassFields.Item(MembershipGroupFields.DefaultGroup).Bool = False
      mvClassFields.Item(MembershipGroupFields.IsCurrent).Bool = False

      Dim vHistory As New MembershipGroupHistory(mvEnv)
      vHistory.Init()
      vHistory.Create(MembershipNumber, OrganisationNumber, pNewOrganisationNumber)
      Dim vTrans As Boolean
      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If
      Save()
      vHistory.Save()
      If vTrans Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public Sub InitFromMemberAndBranch(ByVal pEnv As CDBEnvironment, ByVal pMembershipNumber As Integer, ByVal pBranchCode As String, Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "")
      InitFromBranchOrOrganisation(pEnv, pMembershipNumber, pBranchCode, 0, pValidFrom, pValidTo, 0)
    End Sub

    Public Sub InitFromMemberAndOrganisation(ByVal pEnv As CDBEnvironment, ByVal pMembershipNumber As Integer, ByVal pOrganisationNumber As Integer, Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "", Optional ByVal pExistingMembershipGroupNumber As Integer = 0)
      InitFromBranchOrOrganisation(pEnv, pMembershipNumber, "", pOrganisationNumber, pValidFrom, pValidTo, pExistingMembershipGroupNumber)
    End Sub

    Private Sub InitFromBranchOrOrganisation(ByVal pEnv As CDBEnvironment, ByVal pMembershipNumber As Integer, ByVal pBranchCode As String, ByVal pOrganisationNumber As Integer, ByVal pValidFrom As String, ByVal pValidTo As String, ByVal pExistingMembershipGroupNumber As Integer)
      'pValidFrom & pValidTo are used to limit the selection to records within that date range
      'pExistingMembershipGroupNumber is used to specifically exclude that Group
      Dim vRecordSet As CDBRecordSet
      Dim vSQL As String

      mvEnv = pEnv
      If pMembershipNumber > 0 Then
        vSQL = "SELECT " & GetRecordSetFields() & " FROM membership_groups mg, organisations org"
        If pBranchCode.Length > 0 Then vSQL = vSQL & ", branches b"
        vSQL = vSQL & " WHERE mg.membership_number = " & pMembershipNumber
        If IsDate(pValidFrom) = True And pExistingMembershipGroupNumber > 0 Then vSQL = vSQL & " AND mg.membership_group_number <> " & pExistingMembershipGroupNumber 'Attempting to update this record and we need to know whether the update would cause an over-lap
        If pOrganisationNumber > 0 Then vSQL = vSQL & " AND mg.organisation_number = " & pOrganisationNumber
        If IsDate(pValidFrom) Then
          'Find MembershipGroup that was valid on this date
          vSQL = vSQL & " AND ((valid_from" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pValidFrom) & " AND (valid_to IS NULL OR (valid_to IS NOT NULL AND valid_to" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, pValidFrom) & ")))"
          vSQL = vSQL & " OR (valid_from" & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, pValidFrom)
          If IsDate(pValidTo) Then vSQL = vSQL & " AND valid_from" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, pValidTo) '& ")"   '(valid_to IS NULL OR (valid_to IS NOT NULL AND valid_to" & mvEnv.Connection.SQLLiteral("<=", cftDate, pValidTo) & ")))"
          vSQL = vSQL & "))"
        ElseIf pBranchCode.Length > 0 Then
          'When looking for a particular Branch, only check for current or default MembershipGroups
          vSQL = vSQL & " AND (mg.is_current = 'Y' OR mg.default_group = 'Y')"
        End If
        vSQL = vSQL & " AND mg.organisation_number = org.organisation_number AND org.organisation_group = '" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMemberOrganisationGroup) & "'"
        If pBranchCode.Length > 0 Then vSQL = vSQL & " AND org.organisation_number = b.organisation_number AND b.branch = '" & pBranchCode & "'"

        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
        If vRecordSet.Fetch() Then
          InitFromRecordSet(vRecordSet)
        Else
          InitClassFields()
          SetDefaults()
        End If
        vRecordSet.CloseRecordSet()
      Else
        InitClassFields()
        SetDefaults()
      End If
    End Sub

    Public Sub DedupNewGroup()
      'Dedup the addition of a new MembershipGroup to ensure that ValidFrom & ValidTo dates do not overlap
      'This class should be initialised for the new MembershipGroup but not saved
      Dim vRS As CDBRecordSet
      Dim vMemberGroups As New CollectionList(Of MembershipGroup)
      Dim vMG As New MembershipGroup(mvEnv)
      Dim vNewMG As MembershipGroup
      Dim vUpdateParams As New CDBParameters
      Dim vSave As Boolean
      Dim vSQL As String
      Dim vTrans As Boolean
      Dim vUpdate As Boolean

      vSQL = "SELECT " & GetRecordSetFields() & " FROM membership_groups mg WHERE membership_number = " & MembershipNumber
      vSQL = vSQL & " AND organisation_number = " & OrganisationNumber
      vSQL = vSQL & " AND ((valid_from" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, ValidFrom) & " AND (valid_to IS NULL OR (valid_to IS NOT NULL AND valid_to" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, ValidFrom) & ")))"
      vSQL = vSQL & " OR (valid_from" & mvEnv.Connection.SQLLiteral(">", CDBField.FieldTypes.cftDate, ValidFrom)
      If IsDate(ValidTo) Then vSQL = vSQL & " AND valid_from" & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, ValidTo)
      vSQL = vSQL & ")) ORDER BY valid_from, valid_to" & IIf(mvEnv.Connection.NullsSortAtEnd = True, "", " DESC").ToString 'Sort by valid_to null at the end
      vRS = mvEnv.Connection.GetRecordSet(vSQL)
      While vRS.Fetch
        vMG = New MembershipGroup(mvEnv)
        vMG.InitFromRecordSet(vRS)
        vMemberGroups.Add(vMG.MembershipGroupNumber.ToString, vMG)
      End While
      vRS.CloseRecordSet()
      vSave = True
      vUpdate = False
      If mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTrans = True
      End If
      If vMemberGroups.Count > 0 Then
        'First deal with this MembershipGroup being the DefaultGroup
        If DefaultGroup Then
          For Each vMG In vMemberGroups
            vUpdate = False
            vUpdateParams.Clear()
            If (CDate(vMG.ValidFrom) < CDate(ValidFrom)) Then
              vUpdateParams.Add("ValidTo", CDBField.FieldTypes.cftDate, CDate(ValidFrom).AddDays(-1).ToString(CAREDateFormat))
              vUpdate = True
            ElseIf (CDate(vMG.ValidFrom) >= CDate(ValidFrom)) Then
              If IsDate(vMG.ValidTo) Then
                If IsDate(ValidTo) Then
                  If (CDate(vMG.ValidTo) = CDate(ValidTo)) Then
                    vMG.Delete()
                  ElseIf (CDate(vMG.ValidTo) > CDate(ValidTo)) Then
                    vUpdateParams.Add("ValidFrom", CDBField.FieldTypes.cftDate, CDate(ValidTo).AddDays(1).ToString(CAREDateFormat))
                    vUpdate = True
                  Else
                    vMG.Delete()
                  End If
                Else
                  vMG.Delete()
                End If
              Else
                If IsDate(ValidTo) Then
                  vUpdateParams.Add("ValidFrom", CDBField.FieldTypes.cftDate, CDate(ValidTo).AddDays(1).ToString(CAREDateFormat))
                Else
                  vMG.Delete()
                End If
              End If
            End If
            If vUpdate Then
              vMG.Update(vUpdateParams)
              vMG.Save()
            End If
          Next vMG
        Else
          vUpdateParams.Clear()
          For Each vMG In vMemberGroups
            If vMG.DefaultGroup Then
              If (CDate(vMG.ValidFrom) <= CDate(ValidFrom)) Then
                If IsDate(vMG.ValidTo) Then
                  mvClassFields.Item(MembershipGroupFields.ValidFrom).Value = CDate(vMG.ValidTo).AddDays(1).ToString(CAREDateFormat)
                Else
                  vSave = False
                End If
              ElseIf IsDate(vMG.ValidTo) Then
                If IsDate(ValidTo) Then
                  If CDate(vMG.ValidTo) >= CDate(ValidTo) Then
                    mvClassFields.Item(MembershipGroupFields.ValidTo).Value = CDate(vMG.ValidFrom).AddDays(-1).ToString(CAREDateFormat)
                  Else
                    vNewMG = New MembershipGroup(mvEnv)
                    With vUpdateParams
                      .Add("ValidFrom", CDBField.FieldTypes.cftDate, CDate(vMG.ValidTo).AddDays(1).ToString(CAREDateFormat))
                      .Add("ValidTo", CDBField.FieldTypes.cftDate, ValidTo)
                      .Add("MembershipNumber", MembershipNumber)
                      .Add("OrganisationNumber", OrganisationNumber)
                    End With
                    vNewMG.Init()
                    vNewMG.Create(vUpdateParams)
                    vNewMG.Save()
                    mvClassFields.Item(MembershipGroupFields.ValidTo).Value = CDate(vMG.ValidFrom).AddDays(-1).ToString(CAREDateFormat)
                  End If
                Else
                  mvClassFields.Item(MembershipGroupFields.ValidFrom).Value = CDate(vMG.ValidTo).AddDays(1).ToString(CAREDateFormat)
                End If
              Else
                vSave = False
              End If
            Else
              If (CDate(vMG.ValidFrom) <= CDate(ValidFrom)) Then
                mvClassFields.Item(MembershipGroupFields.ValidFrom).Value = vMG.ValidFrom
                If IsDate(vMG.ValidTo) Then
                  If IsDate(ValidTo) Then
                    If (CDate(vMG.ValidTo) >= CDate(ValidTo)) Then
                      mvClassFields.Item(MembershipGroupFields.ValidTo).Value = vMG.ValidTo
                    End If
                  End If
                Else
                  mvClassFields.Item(MembershipGroupFields.ValidTo).Value = ""
                End If
                vMG.Delete()
              Else
                If IsDate(vMG.ValidTo) Then
                  If IsDate(ValidTo) Then
                    If (CDate(vMG.ValidTo) > CDate(ValidTo)) Then
                      mvClassFields.Item(MembershipGroupFields.ValidTo).Value = vMG.ValidTo
                    End If
                  End If
                Else
                  mvClassFields.Item(MembershipGroupFields.ValidTo).Value = ""
                End If
                vMG.Delete()
              End If
            End If
          Next vMG
        End If
      End If
      If vSave Then Save()
      If vTrans Then mvEnv.Connection.CommitTransaction()
    End Sub

    Friend Sub CloneForCMT(ByVal pNewMembershipNumber As Integer, ByVal pOldMembersBranchCode As String, ByVal pNewMembersBranchCode As String)
      'For CMT - (i)  If original DefaultGroup was for Members Branch, then set new DefaultGroup to be Members new Branch
      '          (ii) For all other records, copy across to the new Membership
      Dim vBranch As Branch = Nothing
      Dim vResetOrganisation As Boolean = False
      If DefaultGroup = True And mvBranchSet = True Then
        If mvBranchCode.length > 0 And (mvBranchCode = pOldMembersBranchCode) Then
          'New record needs to be created with the new OrganisationNumber
          vBranch = New Branch
          vBranch.Init(mvEnv, pNewMembersBranchCode)
          vResetOrganisation = True
        End If
      End If
      With mvClassFields
        .ClearSetValues()
        .Item(MembershipGroupFields.MembershipGroupNumber).Value = CStr(mvEnv.GetControlNumber("MG"))
        .Item(MembershipGroupFields.MembershipNumber).Value = CStr(pNewMembershipNumber)
        If vResetOrganisation Then .Item(MembershipGroupFields.OrganisationNumber).Value = CStr(vBranch.OrganisationNumber)
        .Item(MembershipGroupFields.ValidFrom).Value = TodaysDate()
        'ValidTo & DefaultGroup will remain unchanged
      End With
      SetCurrent()
      mvExisting = False
    End Sub

    Private Sub SetCurrent()
      Dim vIsCurrent As Boolean = True
      If CDate(ValidFrom) > Today Then
        vIsCurrent = False
      Else
        If IsDate(ValidTo) Then
          If CDate(ValidTo) < CDate(Today.ToString(CAREDateFormat)) Then vIsCurrent = False
        End If
      End If
      mvClassFields.Item(MembershipGroupFields.IsCurrent).Bool = vIsCurrent
    End Sub
  End Class

End Namespace