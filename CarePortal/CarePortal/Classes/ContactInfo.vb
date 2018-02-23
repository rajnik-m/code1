<CLSCompliant(True)> _
Public Class ContactInfo

  Public Enum ContactTypes
    ctOrganisation = 1
    ctContact
    ctJoint
  End Enum

  Public Enum AddressTypes
    ataContact = 1
    ataOrganisation = 2
  End Enum

  Public Enum OwnershipAccessLevels
    oalNone
    oalBrowse
    oalRead
    oalWrite
  End Enum

  Private mvInitialised As Boolean
  Private mvContactNumber As Integer
  Private mvAddressNumber As Integer
  Private mvOrganisationNumber As Integer
  Private mvContactType As ContactTypes
  Private mvContactGroup As String
  Private mvAccessLevel As OwnershipAccessLevels
  Private mvContactName As String
  Private mvEMailAddress As String
  Private mvEMailAddresses As String
  Private mvEMailAddressesValid As Boolean
  Private mvPhoneNumber As String
  Private mvSelectedDocumentNumber As Integer
  Private mvSelectedActionNumber As Integer
  Private mvStatus As String
  Private mvStatusDesc As String
  Private mvOwnershipGroup As String
  Private mvAccessLevelDesc As String
  Private mvPrincipalDepartmentDesc As String
  Private mvPrincipalUserName As String
  Private mvVATCategory As String
  Private mvDefaultContactNumber As Integer
  Private mvBranchCode As String
  Private mvDateOfBirth As String
  Private mvDOBEstimated As Boolean
  Private mvJointContactNumber1 As Integer
  Private mvJointContactNumber2 As Integer
  Private mvAddressValidTo As String
  'BR11482/PP1165
  Private mvAddressType As AddressTypes
  Private mvCPDCycleType As String
  Private mvContactCPDCycleNumber As Integer
  Private mvLegacyNumberValid As Boolean
  Private mvLegacyNumber As Integer
  Private mvLegacyResidueValid As Boolean
  Private mvLegacyResidueAmount As Double
  Private mvBypassContactPrompt As Boolean
  Private mvSurname As String
  Private mvInitials As String
  Private mvPreferred As String
  Private mvSelectedContactNumbers As String      'a comma separated list of all selected contact numbers in a display grid

  Public SelectedAddressNumber As Integer
  Public SelectedContactNumber2 As Integer
  Public SelectedContactPositionNumber As Integer
  Public SelectedContactPositionValidFrom As String
  Public SelectedContactPositionValidTo As String
  Public SelectedCommunicationNumber As Integer
  Public CreateAtOrganisationNumber As Integer
  Public CreateAtAddressNumber As Integer
  Public ContactCreated As Boolean
  Public SelectedMembershipNumber As Integer
  Public SelectedBequestNumber As Integer
  Public RelatedContact As ContactInfo

  Private Sub Initialise()
  End Sub

  Public Shared Function JointSalutation(ByVal pTitle As String, ByVal pForenames As String, ByVal pSurname As String) As String
    Dim vTitleCode As String
    Dim vTitleCode1 As String = ""
    Dim vTitleCode2 As String = ""
    Dim vSurname As String
    Dim vSurname1 As String = ""
    Dim vSurname2 As String = ""
    Dim vForename As String
    Dim vForename1 As String = ""
    Dim vForename2 As String = ""
    Dim vSalutationBuilder As New StringBuilder
    Dim vSalutation As String = ""
    Dim vRow As DataRow

    vTitleCode = JointItem(pTitle, pTitle, vTitleCode1, vTitleCode2, False)
    If vTitleCode.Length > 0 Then
      vSurname = JointItem(pTitle, pSurname, vSurname1, vSurname2, False)
      vForename = JointItem(pTitle, pForenames, vForename1, vForename2, False)
      vRow = JointTitleRow(pTitle)
      If vRow IsNot Nothing AndAlso vRow("JointTitle").ToString = "Y" Then vSalutation = vRow("Salutation").ToString
      If vSalutation.Length > 0 Then
        vSalutation = vSalutation.Replace("title1", vTitleCode1)
        vSalutation = vSalutation.Replace("title2", vTitleCode2)
        vSalutation = vSalutation.Replace("surname1", vSurname1)
        If vSurname2.Length = 0 Then vSurname2 = vSurname1
        vSalutation = vSalutation.Replace("surname2", vSurname2)
        vSalutation = vSalutation.Replace("forename1", vForename1)
        vSalutation = vSalutation.Replace("forename2", vForename2)
        vSalutation = vSalutation.Replace("  ", " ")
      Else
        Dim vList As New ArrayListEx
        vList.Add("Dear")
        vList.Add(vTitleCode1)
        If vSurname2.Length > 0 Then vList.Add(vSurname1)
        If vTitleCode1.Length + vSurname1.Length > 0 Then vList.Add("&")
        vList.Add(vTitleCode2)
        If vSurname2.Length > 0 Then
          vList.Add(vSurname2)
        Else
          vList.Add(vSurname1)
        End If
        vSalutation = vList.SSNonBlankList
      End If
    End If
    Return vSalutation.Trim
  End Function


  Public Shared Function JointTitleRow(ByVal pTitle As String) As DataRow
    Dim vTable As DataTable = TitlesTable()
    If pTitle.Trim.Length > 0 Then
      Return vTable.Rows.Find(pTitle)
    Else
      Return Nothing
    End If
  End Function

  Private Shared Function TitlesTable() As DataTable
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtTitles)
    Dim vPrimaryColumns(0) As DataColumn
    vPrimaryColumns(0) = vTable.Columns("Title")
    vTable.PrimaryKey = vPrimaryColumns
    Return vTable
  End Function

  Public Shared Function JointItem(ByVal pTitle As String, ByVal pString As String, ByRef pFirst As String, ByRef pSecond As String, ByVal pValidateTitles As Boolean) As String
    Dim vPos As Integer
    Dim vSep As String = "&"
    Dim vCheckJoint As Boolean
    Dim vResultString As String = ""
    Dim vRow As DataRow = Nothing

    Dim vTable As DataTable = TitlesTable()
    If pTitle.Length > 0 Then
      vRow = vTable.Rows.Find(pTitle)
    End If
    If vRow IsNot Nothing Then
      If vRow("JointTitle").ToString = "Y" Then
        If pTitle = pString Then
          'Check for und et e & +
          If pString.IndexOf(" and ") > 0 Then
            vSep = " and "
          ElseIf pString.IndexOf(" und ") > 0 Then
            vSep = " und "
          ElseIf pString.IndexOf(" et ") > 0 Then
            vSep = " et "
          ElseIf pString.IndexOf(" e ") > 0 Then
            vSep = " e "
          ElseIf pString.IndexOf(" + ") > 0 Then
            vSep = "+"
          Else
            vSep = "&"
          End If
        Else
          If pString.IndexOf(" and ") > 0 Then
            vSep = " and "
          ElseIf pString.IndexOf(" + ") > 0 Then
            vSep = "+"
          ElseIf pString.IndexOf(" und ") > 0 Then
            vSep = " und "
          ElseIf pString.IndexOf(" et ") > 0 Then
            vSep = " et "
          ElseIf pString.IndexOf(" e ") > 0 Then
            vSep = " e "
          Else
            vSep = "&"
          End If
        End If
        vCheckJoint = True
      End If
    Else
      'vCheckJoint = AppValues.ConfigurationOption(AppValues.ConfigurationOptions.cd_joint_contact_support, True)
      vCheckJoint = DataHelper.ConfigurationOption(DataHelper.ConfigurationOptions.cd_joint_contact_support)
    End If
    vPos = pString.IndexOf(vSep)
    If vPos > 0 And vCheckJoint Then
      pFirst = pString.Substring(0, vPos).Trim
      pSecond = pString.Substring(vPos + vSep.Length).Trim
      Dim vRowFirst As DataRow = Nothing
      Dim vRowSecond As DataRow = Nothing
      If pValidateTitles Then
        vRowFirst = vTable.Rows.Find(pFirst)
        If vRowFirst IsNot Nothing Then pFirst = vRowFirst("Title").ToString
        vRowSecond = vTable.Rows.Find(pSecond)
        If vRowSecond IsNot Nothing Then pSecond = vRowSecond("Title").ToString

        Dim vValidTitle As Boolean = True
        Dim vRowJoint As DataRow = Nothing
        vRowJoint = vTable.Rows.Find(pFirst & " & " & pSecond)
        If vRowJoint IsNot Nothing OrElse pFirst.Length = 0 OrElse pSecond.Length = 0 Then
          vValidTitle = False
        End If

        If vValidTitle = True Then
          If vRowFirst IsNot Nothing AndAlso vRowSecond IsNot Nothing Then
            If vRow Is Nothing Then       'If we haven't already found the joint in the table
              Dim vNewRow As DataRow = vTable.NewRow
              vNewRow("Title") = String.Format("{0} {1} {2}", pFirst, vSep.Trim, pSecond)
              vNewRow("JointTitle") = "Y"
              vNewRow("Sex") = "U"
              vNewRow("Salutation") = ""
              vTable.Rows.Add(vNewRow)
            End If
          End If
        End If
      End If
      If Not pValidateTitles OrElse (vRowFirst IsNot Nothing AndAlso vRowSecond IsNot Nothing) Then
        vResultString = String.Format("{0} {1} {2}", pFirst, vSep.Trim, pSecond)
      End If
    Else
      pFirst = pString.Trim
      pSecond = ""
      If Not pValidateTitles OrElse (pString.Length > 0 AndAlso vTable.Rows.Find(pString) IsNot Nothing) Then vResultString = pString
    End If
    Return vResultString
  End Function


End Class

