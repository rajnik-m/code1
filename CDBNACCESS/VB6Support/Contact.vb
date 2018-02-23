Imports System.Linq
Imports Advanced.Data.Merge

Namespace Access

  Partial Public Class Contact

    Public Enum ContactRecordSetTypes 'These are bit values
      crtNumber = 1
      crtName = 2
      crtAddress = 4
      crtAddressCountry = 8
      crtVAT = 16
      crtGroup = 32
      crtPhone = 64
      crtDetail = 128 + 64 + 32 'Include phone and group
      crtDefaultAddressNumber = 256
    End Enum

    Public Enum ContactSex
      cscFemale
      cscMale
      cscUnknown
    End Enum

    Public Enum ContactRelationshipLinkTypes
      crltAll
      crltJoint
      crltReal
      crltFrom
      crltTo
    End Enum

    Public Enum ContactGiftAidMergeDates
      cgamdEqual
      cgamdLessThan
      cgamdGreaterThan
    End Enum

    Private Class ContactMergeInfo
      Public TableName As String = ""
      Public ContactAttr As String = ""
      Public AddressAttr As String = ""
      Public SetAmend As Boolean
      Public UniqueContact As Boolean
      Public UniqueAttrs As String = ""
    End Class

    Private mvContactLinks As CollectionList(Of ContactLink)
    Private mvGiftAidDeclarations As CollectionList(Of GiftAidDeclaration)
    Private mvMergeInfo As List(Of ContactMergeInfo)
    Private mvJointContact1 As Contact
    Private mvJointContact2 As Contact

    Public Overloads Function GetRecordSetFields(ByVal pRSType As ContactRecordSetTypes) As String
      Dim vFields As String = ""
      Dim vAddressType As Address.AddressRecordSetTypes

      CheckClassFields()
      vFields = "c.contact_number,"
      If (pRSType And ContactRecordSetTypes.crtName) > 0 Then
        vFields = vFields & "title,forenames,initials,surname,honorifics,salutation,label_name,preferred_forename,contact_type,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataIrishGiftAid) Then vFields = vFields & "ni_number,"
        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDutchSupport) Then vFields = vFields & "prefix_honorifics,surname_prefix,informal_salutation,"
      End If
      If (pRSType And ContactRecordSetTypes.crtVAT) > 0 Then vFields = vFields & "contact_vat_category,"
      If (pRSType And ContactRecordSetTypes.crtGroup) > 0 Then vFields = vFields & "c.contact_group,"
      If (pRSType And ContactRecordSetTypes.crtPhone) > 0 Then vFields = vFields & "c.dialling_code,c.std_code,c.telephone,c.ex_directory,"
      If (pRSType And ContactRecordSetTypes.crtDetail) = ContactRecordSetTypes.crtDetail Then
        vFields = vFields & "sex,c.source,c.source_date,name_gathering_source,date_of_birth,c.status,c.status_date,c.status_reason,c.department,c.notes,dob_estimated,c.amended_on,c.amended_by,"
        If mvClassFields(ContactFields.OwnershipGroup).InDatabase Then vFields = vFields & "c.ownership_group,"
        If mvClassFields(ContactFields.ResponseChannel).InDatabase Then vFields = vFields & "c.response_channel,"
        If mvClassFields(ContactFields.ContactReference).InDatabase Then vFields &= "c.contact_reference,"
      End If
      If (pRSType And ContactRecordSetTypes.crtAddress) > 0 Then
        vAddressType = Address.AddressRecordSetTypes.artNumber Or Address.AddressRecordSetTypes.artDetails
        If (pRSType And ContactRecordSetTypes.crtAddressCountry) > 0 Then vAddressType = vAddressType Or Address.AddressRecordSetTypes.artCountrySortCode
        If mvCurrentAddress Is Nothing Then mvCurrentAddress = New Address(mvEnv)
        vFields = vFields & mvCurrentAddress.GetRecordSetFields(vAddressType)
      End If
      If (pRSType And ContactRecordSetTypes.crtDefaultAddressNumber) > 0 Then
        If Right(RTrim(vFields), 1) <> "," Then vFields = vFields & ","
        vFields = vFields & "c.address_number AS default_address_number"
      End If
      If Right(vFields, 1) = "," Then vFields = Left(vFields, Len(vFields) - 1)
      Return vFields
    End Function

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As ContactRecordSetTypes)
      Dim vAddressType As Address.AddressRecordSetTypes
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always grab the unique key, 'cos you need it for saving
        .SetItem(ContactFields.ContactNumber, vFields)
        If (pRSType And ContactRecordSetTypes.crtName) > 0 Then
          .SetItem(ContactFields.Title, vFields)
          .SetItem(ContactFields.Forenames, vFields)
          .SetItem(ContactFields.Initials, vFields)
          .SetItem(ContactFields.Surname, vFields)
          .SetItem(ContactFields.Honorifics, vFields)
          .SetItem(ContactFields.Salutation, vFields)
          .SetItem(ContactFields.LabelName, vFields)
          .SetItem(ContactFields.PreferredForename, vFields)
          .SetItem(ContactFields.ContactType, vFields)
          If vFields.Exists("ni_number") Then .SetItem(ContactFields.NiNumber, vFields)
          If vFields.Exists("prefix_honorifics") Then .SetOptionalItem(ContactFields.PrefixHonorifics, vFields)
          If vFields.Exists("surname_prefix") Then .SetOptionalItem(ContactFields.SurnamePrefix, vFields)
          If vFields.Exists("informal_salutation") Then .SetOptionalItem(ContactFields.InformalSalutation, vFields)
          If vFields.Exists("label_name_format_code") Then .SetOptionalItem(ContactFields.LabelNameFormatCode, vFields)
          mvLabelNameValid = True
          mvSalutationValid = True
          mvInformalSalutationValid = True
        End If
        If (pRSType And ContactRecordSetTypes.crtVAT) > 0 Then
          .SetItem(ContactFields.ContactVatCategory, vFields)
        End If
        If (pRSType And ContactRecordSetTypes.crtAddress) > 0 Then
          .SetItem(ContactFields.AddressNumber, vFields)
          vAddressType = Access.Address.AddressRecordSetTypes.artNumber Or Access.Address.AddressRecordSetTypes.artDetails
          If (pRSType And ContactRecordSetTypes.crtAddressCountry) > 0 Then vAddressType = vAddressType Or Access.Address.AddressRecordSetTypes.artCountrySortCode
          mvCurrentAddress.InitFromRecordSet(mvEnv, pRecordSet, vAddressType)
        End If
        If (pRSType And ContactRecordSetTypes.crtGroup) > 0 Then
          .SetItem(ContactFields.ContactGroup, vFields)
        End If
        If (pRSType And ContactRecordSetTypes.crtPhone) > 0 Then
          .SetItem(ContactFields.DiallingCode, vFields)
          .SetItem(ContactFields.StdCode, vFields)
          .SetItem(ContactFields.Telephone, vFields)
          .SetItem(ContactFields.ExDirectory, vFields)
        End If
        If (pRSType And ContactRecordSetTypes.crtDetail) = ContactRecordSetTypes.crtDetail Then
          .SetItem(ContactFields.Sex, vFields)
          .SetItem(ContactFields.Source, vFields)
          .SetItem(ContactFields.SourceDate, vFields)
          .SetItem(ContactFields.NameGatheringSource, vFields)
          .SetItem(ContactFields.DateOfBirth, vFields)
          .SetItem(ContactFields.Status, vFields)
          .SetItem(ContactFields.StatusDate, vFields)
          .SetItem(ContactFields.StatusReason, vFields)
          .SetItem(ContactFields.Department, vFields)
          .SetOptionalItem(ContactFields.OwnershipGroup, vFields)
          .SetOptionalItem(ContactFields.ResponseChannel, vFields)
          .SetOptionalItem(ContactFields.ContactReference, vFields)
          .SetItem(ContactFields.Notes, vFields)
          .SetItem(ContactFields.DobEstimated, vFields)
          .SetItem(ContactFields.AmendedOn, vFields)
          .SetItem(ContactFields.AmendedBy, vFields)
          mvSourceDateValid = True
          mvStatusDateValid = True
        End If
      End With
      If (pRSType And ContactRecordSetTypes.crtDefaultAddressNumber) > 0 Then
        mvClassFields(ContactFields.AddressNumber).SetValue = pRecordSet.Fields("default_address_number").Value
      End If
      mvOwnershipValid = False
    End Sub

    Public Sub InitRecordSetType(ByVal pEnv As CDBEnvironment, ByVal pRSType As ContactRecordSetTypes, ByVal pContactNumber As Integer, Optional ByVal pAddressNumber As Integer = 0)
      Dim vSQL As String
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      InitClassFields()
      SetDefaults()
      If pAddressNumber > 0 Then
        vSQL = "SELECT " & GetRecordSetFields(pRSType) & " FROM contacts c, contact_addresses ca, addresses a"
      Else
        vSQL = "SELECT " & GetRecordSetFields(pRSType) & " FROM contacts c, addresses a"
      End If
      If CBool(pRSType & ContactRecordSetTypes.crtAddressCountry) Then vSQL = vSQL & ", countries co"
      If pAddressNumber > 0 Then
        vSQL = vSQL & " WHERE c.contact_number = " & pContactNumber & " AND ca.contact_number = c.contact_number AND ca.address_number = " & pAddressNumber & " AND a.address_number = ca.address_number"
      Else
        vSQL = vSQL & " WHERE c.contact_number = " & pContactNumber & " AND a.address_number = c.address_number"
      End If
      If CBool(pRSType & ContactRecordSetTypes.crtAddressCountry) Then vSQL = vSQL & " AND a.country = co.country"
      vRecordSet = pEnv.Connection.GetRecordSet(vSQL)
      If vRecordSet.Fetch() Then
        InitFromRecordSet(pEnv, vRecordSet, pRSType)
      End If
      vRecordSet.CloseRecordSet()
    End Sub

    Public ReadOnly Property Owners() As CDBParameters
      Get
        If mvOwners Is Nothing Then InitOwners()
        Owners = mvOwners
      End Get
    End Property

    Private Sub InitOwners()
      mvOwners = New CDBParameters
      Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT cu.department, department_desc FROM contact_users cu, departments d WHERE contact_number = " & ContactNumber & " AND cu.department = d.department")
      With vRS
        While .Fetch()
          mvOwners.Add((.Fields(1).Value), .Fields(2).Value)
        End While
        .CloseRecordSet()
      End With
    End Sub

    Public Function ProcessJointContact(ByVal pContact2 As Contact, ByVal pSource As String) As Integer
      Dim vRealToReal As String
      Dim vRealToJoint As String
      Dim vJoint As Contact
      Dim vCA As New ContactAddress(mvEnv)

      vJoint = GetJointContact(pContact2)
      If vJoint.ContactNumber = 0 Then
        'Didn't find an existing joint contact
        With pContact2
          If Surname = .Surname Then 'Same surname
            vJoint.Surname = Surname
            If Len(TitleName) > 0 And Len(.TitleName) > 0 Then 'Salutation is Dear xx & xx surname
              vJoint.Salutation = "Dear " & TitleName & " & " & .TitleName & " " & Surname
            Else
              vJoint.Salutation = "Dear Members"
            End If
            'Label name is Title Initials & Title Initials Surname
            vJoint.LabelName = LTrim(TitleName & " ") & LTrim(Initials & " ") & "& " & LTrim(.TitleName & " ") & LTrim(.Initials & " ") & Surname
          Else 'Different surname
            vJoint.Surname = Surname & " & " & .Surname
            If Len(TitleName) > 0 And Len(.TitleName) > 0 Then 'Salutation is Dear xx surname & xx joint surname
              vJoint.Salutation = "Dear " & TitleName & " " & Surname & " & " & .TitleName & " " & .Surname
            Else
              vJoint.Salutation = "Dear Members"
            End If
            'Label name is Title Initials Surname & Title Initials Surname
            vJoint.LabelName = LTrim(TitleName & " ") & LTrim(Initials & " ") & Surname & " & " & LTrim(.TitleName & " ") & LTrim(.Initials & " ") & .Surname
          End If
        End With
        With vJoint
          .ContactType = ContactTypes.ctcJoint
          .Sex = ContactSex.cscUnknown
          .Source = pSource
          .SourceDate = TodaysDate()
          .Department = mvEnv.User.Department
          .AddressNumber = AddressNumber
          .ContactGroupCode = ContactGroupCode
          .VATCategory = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefConVatCat)
          If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups And mvClassFields(ContactFields.OwnershipGroup).InDatabase And .OwnershipGroup.Length = 0 Then
            .OwnershipGroup = OwnershipGroup
          End If
          .Save()
          .AddUser(.Department, True) 'Add contact_user
          vCA.Create(.ContactNumber, .AddressNumber, "N", TodaysDate, "")
          vCA.Save(mvEnv.User.Logname, True)
        End With
        'Add Membership Relationship Links
        vRealToReal = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToRealLink)
        vRealToJoint = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink)
        vJoint.AddLinks(vRealToJoint, Me)
        vJoint.AddLinks(vRealToJoint, pContact2)
        AddLinks(vRealToReal, pContact2)

        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesNormal, ContactTypes.ctcContact, ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlsDerivedSuppression))
        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesNormal, ContactTypes.ctcContact, pContact2.ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlsDerivedSuppression))
        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesNormal, ContactTypes.ctcContact, vJoint.ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlJointSuppression))
      End If
      ProcessJointContact = vJoint.ContactNumber
    End Function

    Private Sub GetPositionInfo()
      Dim vRecordSet As CDBRecordSet
      Dim vRecordSet2 As CDBRecordSet
      Dim vAttrs As String

      If Not mvPositionValid Then
        mvPositionCurrent = False
        If ContactType = ContactTypes.ctcOrganisation Then
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT o.organisation_number,name,o.contact_number,o.status,o.status_date,o.status_reason FROM organisations o WHERE o.organisation_number = " & ContactNumber)
          If vRecordSet.Fetch() Then
            mvOrganisationNumber = vRecordSet.Fields(1).LongValue
            mvOrganisationName = vRecordSet.Fields(2).Value
            mvOrgContactNumber = vRecordSet.Fields(3).LongValue
            mvOrgStatus = vRecordSet.Fields(4).Value
            mvOrgStatusDate = vRecordSet.Fields(5).Value
            mvOrgStatusReason = vRecordSet.Fields(6).MultiLine
          End If
          vRecordSet.CloseRecordSet()
        Else
          If mvCurrentAddress.AddressType = Address.AddressTypes.ataOrganisation Then
            'Assume there is a current position record first
            vAttrs = "position,position_location,o.organisation_number,name,o.contact_number,contact_position_number"
            vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vAttrs & " FROM contact_positions cp, organisation_addresses oa, organisations o WHERE cp.contact_number = " & ContactNumber & " AND cp.address_number = " & mvCurrentAddress.AddressNumber & " AND " & mvEnv.Connection.DBSpecialCol("cp", "current") & " = 'Y' AND cp.address_number = oa.address_number AND oa.organisation_number = o.organisation_number")
            If vRecordSet.Fetch() Then
              mvPosition = vRecordSet.Fields(1).Value
              mvPositionLocation = vRecordSet.Fields(2).Value
              mvOrganisationNumber = vRecordSet.Fields(3).LongValue
              mvOrganisationName = vRecordSet.Fields(4).Value
              mvOrgContactNumber = vRecordSet.Fields(5).LongValue
              mvContactPositionNumber = vRecordSet.Fields(6).LongValue
              mvPositionCurrent = True
            Else
              vRecordSet2 = mvEnv.Connection.GetRecordSet("SELECT o.organisation_number,name,o.contact_number FROM organisation_addresses oa, organisations o WHERE oa.address_number = " & mvCurrentAddress.AddressNumber & " AND oa.organisation_number = o.organisation_number")
              If vRecordSet2.Fetch() Then
                mvOrganisationNumber = vRecordSet2.Fields(1).LongValue
                mvOrganisationName = vRecordSet2.Fields(2).Value
                mvOrgContactNumber = vRecordSet2.Fields(3).LongValue
              End If
              vRecordSet2.CloseRecordSet()
            End If
            vRecordSet.CloseRecordSet()
          End If
        End If
        mvPositionValid = True
      End If
    End Sub

    Public Sub AddLinks(ByVal pRelationship As String, ByVal pContact2 As Contact)
      Dim vLink As New ContactLink(mvEnv)
      vLink.Create(ContactNumber, pContact2.ContactNumber, pRelationship, "")

      'We moved validation from the XML layer to the business object layer, which now prevents overlaps by default.  Previously in this code, overlaps were allowed so we're continuing for now.  
      'This will need to be reviewed at some point, as overlaps should really not be allowed.  However the only alternative at the moment is calling MergeLinks, and that deletes existing data!  Not a viable option.
      vLink.AllowsOverlaps = True

      vLink.Save()
    End Sub

    Public Sub AddUser(ByVal pDepartment As String, Optional ByVal pNew As Boolean = False)
      Dim vFields As New CDBFields
      vFields.Add("contact_number", ContactNumber)
      vFields.Add("department", pDepartment)
      Dim vAdd As Boolean
      If pNew Then
        vAdd = True
      Else
        vAdd = mvEnv.Connection.GetCount("contact_users", vFields) = 0
      End If
      If vAdd Then
        vFields.AddAmendedOnBy(mvEnv.User.UserID)
        mvEnv.Connection.InsertRecord("contact_users", vFields)
      End If
    End Sub

    Public Function GetJointContact(ByRef pContact2 As Contact) As Contact
      Dim vRealToReal As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToRealLink)
      Dim vRealToJoint As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink)

      Dim vSQL As String = "SELECT c.contact_number FROM contact_links cl1, contact_links cl2, contacts c WHERE"
      If pContact2 Is Nothing Then
        vSQL = vSQL & " cl1.contact_number_1 = " & ContactNumber
      Else
        vSQL = vSQL & " cl1.contact_number_1 IN (" & ContactNumber & "," & pContact2.ContactNumber & ")"
        vSQL = vSQL & " AND cl1.contact_number_2 IN (" & ContactNumber & "," & pContact2.ContactNumber & ")"
      End If
      vSQL = vSQL & " AND cl1.contact_number_1 <> cl1.contact_number_2"
      vSQL = vSQL & " AND cl1.relationship = '%1'"
      vSQL = vSQL & " AND cl2.contact_number_2 = cl1.contact_number_1"
      vSQL = vSQL & " AND cl2.relationship = '%2'"
      vSQL = vSQL & " AND c.contact_number = cl2.contact_number_1"
      vSQL = vSQL & " AND c.contact_type = 'J'"

      Dim vSQL2 As String = Replace(vSQL, "%1", vRealToReal)
      vSQL2 = Replace(vSQL2, "%2", vRealToJoint)

      Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet(vSQL2)
      Dim vJoint As Contact = Nothing
      Dim vFound As Boolean
      If vRecordSet.Fetch() Then
        vJoint = New Contact(mvEnv)
        vJoint.Init((vRecordSet.Fields(1).LongValue))
        vFound = True
      End If
      vRecordSet.CloseRecordSet()
      If Not vFound Then
        'Didn't find a joint using the membership relationship codes so try the marketing relationship codes
        vSQL2 = vSQL
        vSQL2 = Replace(vSQL2, "%1", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToDerivedLink))
        vSQL2 = Replace(vSQL2, "%2", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink))
        vRecordSet = mvEnv.Connection.GetRecordSet(vSQL2)
        If vRecordSet.Fetch() Then
          vJoint = New Contact(mvEnv)
          vJoint.Init(vRecordSet.Fields(1).LongValue)
          vFound = True
        End If
        vRecordSet.CloseRecordSet()
        If Not vFound Then
          vJoint = New Contact(mvEnv)
          vJoint.Init()
        End If
      End If
      Return vJoint
    End Function

    Private Sub ReadGAConfig()
      If mvGAConfig.Length = 0 Then
        mvGAConfig = mvEnv.GetConfig("cd_gone_away_marker").ToUpper
        mvGAStatus = InStr(mvGAConfig, "STATUS") > 0
        mvGASuppression = InStr(mvGAConfig, "SUPPRESSION") > 0
        If Not mvGAStatus And Not mvGASuppression Then mvGAStatus = True
        If mvGASuppression And mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMailingSupp).Length = 0 Then
          RaiseError(DataAccessErrors.daeGASuppressionNotSet)
        End If
        If mvGAStatus And mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAStatus).Length = 0 Then
          RaiseError(DataAccessErrors.daeGAStatusNotSet)
        End If
      End If
    End Sub

    Private Sub SetLabelName(ByVal pJunior As Boolean, Optional ByVal pLabelNameFormatCode As String = "")
      Dim vFormat As String = ""
      Dim vLabelName As String = ""
      Dim vItem As String = ""
      Dim vJoint As Boolean
      Dim vTitle1 As String = ""
      Dim vTitle2 As String = ""
      Dim vInitials1 As String = ""
      Dim vInitials2 As String = ""
      Dim vSurname1 As String = ""
      Dim vSurname2 As String = ""
      Dim vContainsInitials As Boolean
      Dim vContainsFornames As Boolean
      Dim vHonorifics1 As String = ""
      Dim vHonorifics2 As String = ""
      Dim vForenames1 As String = ""
      Dim vForenames2 As String = ""
      Dim vWords() As String
      Dim vIndex As Integer
      Dim vJointTitle As String = ""
      Dim vLabelNameFormatCode As Boolean

      Dim vTitle As String = mvClassFields.Item(ContactFields.Title).Value
      Dim vInitials As String = mvClassFields.Item(ContactFields.Initials).Value
      Dim vSurname As String = mvClassFields.Item(ContactFields.Surname).Value
      Dim vHonorifics As String = mvClassFields.Item(ContactFields.Honorifics).Value
      Dim vForenames As String = mvClassFields.Item(ContactFields.Forenames).Value
      Dim vPreferred As String = mvClassFields.Item(ContactFields.PreferredForename).Value
      Dim vPrefixHonorifics As String = mvClassFields.Item(ContactFields.PrefixHonorifics).Value
      Dim vSurnamePrefix As String = mvClassFields.Item(ContactFields.SurnamePrefix).Value

      Dim vListDict As New Dictionary(Of String, String) From {{"title", vTitle}, {"initials", vInitials}, {"forenames", vForenames},
                                                                   {"preferred_forename", vPreferred}, {"surname", vSurname}, {"prefix_honorifics", vPrefixHonorifics},
                                                                   {"honorifics", vHonorifics}}

      If pJunior Then
        If Len(mvJnrLabelNameFormat) = 0 Then
          mvJnrLabelNameFormat = mvEnv.GetConfig("jnr_label_name_format")
          If Len(mvJnrLabelNameFormat) = 0 Then mvJnrLabelNameFormat = "forenames surname honorifics" 'NoTranslate
        End If
        vFormat = mvJnrLabelNameFormat
      Else
        If mvEnv.GetConfigOption("cd_joint_contact_support", True) Then
          If ContactType = ContactTypes.ctcJoint Then
            vJoint = True
          ElseIf InStr(vTitle, " & ") > 0 Then
            vJoint = True
          ElseIf InStr(vInitials, " & ") > 0 Then
            vJoint = True
          ElseIf InStr(vSurname, " & ") > 0 Then
            vJoint = True
          End If
        End If
      End If
      Dim vSelectedTitle As New Title
      vSelectedTitle.Init(mvEnv, vTitle)
      If vJoint AndAlso (ContactType <> ContactTypes.ctcJoint OrElse vSelectedTitle.LabelName.Length = 0) Then
        vItem = JointItem(vTitle, vTitle1, vTitle2)
        vItem = JointItem(vInitials, vInitials1, vInitials2)
        vItem = JointItem(vSurname, vSurname1, vSurname2)
        If vTitle1.Length > 0 Then vLabelName = vTitle1 & " "
        If vInitials1.Length > 0 Then vLabelName = vLabelName & vInitials1 & " "
        If vSurname2.Length > 0 Then vLabelName = vLabelName & vSurname1 & " "
        If vLabelName.Length > 0 Then vLabelName = RTrim(vLabelName) & " & "
        If vTitle2.Length > 0 Then vLabelName = vLabelName & vTitle2 & " "
        If vInitials2.Length > 0 Then vLabelName = vLabelName & vInitials2 & " "
        If vSurname2.Length > 0 Then
          vLabelName = vLabelName & vSurname2
        Else
          vLabelName = vLabelName & vSurname1
        End If
      Else
        If Not pJunior Then
          If Not pLabelNameFormatCode.Length > 0 Then
            If mvLabelNameFormat.Length = 0 Then
              mvLabelNameFormat = vSelectedTitle.LabelName
              If Len(mvLabelNameFormat) = 0 Then mvLabelNameFormat = mvEnv.GetConfig("label_name_format")
              If Len(mvLabelNameFormat) = 0 Then mvLabelNameFormat = "title initials surname honorifics" 'NoTranslate
            End If
          Else
            'get label name format from the label name format code
            Dim vRS As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT label_name_format, joint_title FROM label_name_format_codes WHERE label_name_format_code = '" & pLabelNameFormatCode & "'")
            While vRS.Fetch() = True
              mvLabelNameFormat = vRS.Fields("label_name_format").Value
              vJointTitle = vRS.Fields("joint_title").Value
              vLabelNameFormatCode = True
            End While
            vRS.CloseRecordSet()
          End If
          vFormat = mvLabelNameFormat
        End If
        vContainsInitials = InStr(vFormat, " initials ") > 0
        vContainsFornames = InStr(vFormat, " fornames ") > 0
        vWords = Split(vFormat, " ")
        For vIndex = 0 To UBound(vWords)
          Dim vContinue As Boolean = True
          Dim vStart As Integer = 1
          'BR17812 check for IsNull
          If InStr(vStart, UCase(vWords(vIndex).ToString()), "ISNULL(") > 0 Then
            vWords(vIndex) = ProcessIsNullLabelFormat(vWords(vIndex), vListDict)
          End If

          Dim vThisWord As String = vWords(vIndex)
          Dim vThisWordSuffix As String = String.Empty
          If vThisWord.Length > 0 AndAlso
             vThisWord.Substring(vThisWord.Length - 1).Equals(",") Then
            vThisWordSuffix = vThisWord.Substring(vThisWord.Length - 1)
            vThisWord = vThisWord.Substring(0, vThisWord.Length - 1)
          End If

          Select Case vThisWord
            Case "title1", "title2", "initials1", "initials2", "forenames1", "forenames2", "surname1", "surname2", "honorifics1", "honorifics2"
              If pLabelNameFormatCode.Length = 0 Then
                If vSelectedTitle.JointTitle = False Then
                  vLabelName = vLabelName & vThisWord & vThisWordSuffix & " "
                  vContinue = False
                End If
              ElseIf vLabelNameFormatCode And vJointTitle <> "Y" Then
                vLabelName = vLabelName & vThisWord & vThisWordSuffix & " "
                vContinue = False
              End If
            Case "title", "forenames", "initials", "surname", "honorifics", "preferred_forename", "prefix_honorifics"
              If Len(pLabelNameFormatCode) = 0 Then
                If vSelectedTitle.JointTitle = True Then
                  vLabelName = vLabelName & vThisWord & vThisWordSuffix & " "
                  vContinue = False
                End If
              ElseIf vLabelNameFormatCode And vJointTitle = "Y" Then
                vLabelName = vLabelName & vThisWord & vThisWordSuffix & " "
                vContinue = False
              End If
          End Select
          If vContinue Then
            Select Case vThisWord
              Case "title1"
                vItem = JointItem(vTitle, vTitle1, vTitle2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vTitle1.Length > 0 Then vLabelName = vLabelName & vTitle1 & " "
              Case "title2"
                vItem = JointItem(vTitle, vTitle1, vTitle2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vTitle2.Length > 0 Then vLabelName = vLabelName & vTitle2 & " "
              Case "title"
                If vTitle.Length > 0 Then vLabelName = vLabelName & vTitle & " "
              Case "initials1"
                vItem = JointItem(vInitials, vInitials1, vInitials2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vInitials1.Length > 0 Then vLabelName = vLabelName & vInitials1 & " "
              Case "initials2"
                vItem = JointItem(vInitials, vInitials1, vInitials2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vInitials2.Length > 0 Then vLabelName = vLabelName & vInitials2 & " "
              Case "initials"
                If vInitials.Length > 0 Then vLabelName = vLabelName & vInitials & " "
              Case "surname1"
                vItem = JointItem(vSurname, vSurname1, vSurname2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vSurname1.Length > 0 Then vLabelName = vLabelName & vSurname1 & " "
              Case "surname2"
                vItem = JointItem(vSurname, vSurname1, vSurname2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vSurname2.Length > 0 Then vLabelName = vLabelName & vSurname2 & " "
              Case "surname"
                If vSurnamePrefix.Length > 0 Then
                  If (vInitials.Length > 0 And vContainsInitials) Or (vForenames.Length > 0 And vContainsFornames) Then
                    vLabelName = vLabelName & vSurnamePrefix & " "
                  Else
                    vLabelName = vLabelName & SurnamePrefixCapitalised & " "
                  End If
                End If
                If vSurname.Length > 0 Then vLabelName = vLabelName & vSurname & " "
              Case "honorifics1"
                vItem = JointItem(vHonorifics, vHonorifics1, vHonorifics2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vHonorifics1.Length > 0 Then vLabelName = vLabelName & vHonorifics1 & " "
              Case "honorifics2"
                vItem = JointItem(vHonorifics, vHonorifics1, vHonorifics2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vHonorifics2.Length > 0 Then vLabelName = vLabelName & vHonorifics2 & " "
              Case "honorifics"
                If vHonorifics.Length > 0 Then vLabelName = vLabelName & vHonorifics & " "
              Case "forenames1"
                vItem = JointItem(vForenames, vForenames1, vForenames2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vForenames1.Length > 0 Then vLabelName = vLabelName & vForenames1 & " "
              Case "forenames2"
                vItem = JointItem(vForenames, vForenames1, vForenames2, vSelectedTitle.TitleName, , vSelectedTitle)
                If vForenames2.Length > 0 Then vLabelName = vLabelName & vForenames2 & " "
              Case "forenames"
                If vForenames.Length > 0 Then vLabelName = vLabelName & vForenames & " "
              Case "preferred_forename"
                If vPreferred.Length > 0 Then vLabelName = vLabelName & vPreferred & " "
              Case "prefix_honorifics"
                If vPrefixHonorifics.Length > 0 Then vLabelName = vLabelName & vPrefixHonorifics & " "
              Case Else
                'unknown string so add it?
                vItem = FirstWord(vWords(vIndex))
                vLabelName = vLabelName & vItem & vThisWordSuffix & " "
            End Select
          End If
        Next
      End If
      mvClassFields.Item(ContactFields.LabelName).Value = vLabelName.TrimEnd
      mvLabelNameValid = True
    End Sub

    Private Function ProcessIsNullLabelFormat(ByVal vItem As String, ByVal vListDict As Dictionary(Of String, String)) As String
      'looptest
      Dim start As Integer
      Dim at As Integer
      Dim count As Integer
      Dim [end] As Integer
      Dim iLastBracket As Integer
      Dim vComma As Integer
      Dim sFirstWord As String
      Dim sSecondWord As String
      Dim vTheString As String
      Dim vFound As Boolean = False
      Dim vStart As Integer = 1
      start = vItem.ToString().Length - 1
      [end] = 0

      count = 0
      at = 0
      Try
        While start > -1 And at > -1
          count = start - [end] 'Count must be within the substring.
          at = UCase(vItem.ToString()).LastIndexOf("ISNULL(", start, count)
          If at = -1 Then 'one last check
            at = InStr(vStart, UCase(vItem.ToString()), "ISNULL(")

          End If
          If at > -1 And at <> 0 Then
            'occurance of last isnull(
            iLastBracket = InStr(at, UCase(vItem.ToString()), ")")
            vComma = InStr(at + 1, vItem.ToString(), ",")
            If at = 1 Then at = 0
            vTheString = vItem.ToString().Substring(at, iLastBracket - at)
            sFirstWord = vItem.ToString().Substring(at + 7, vComma - at - 8)
            sSecondWord = vItem.ToString().Substring(vComma, iLastBracket - vComma - 1)

            For Each d As KeyValuePair(Of String, String) In vListDict
              If d.Key = sFirstWord Then
                If (String.IsNullOrEmpty(d.Value.ToString())) Then
                  vItem = vItem.Replace(vTheString, sSecondWord)
                  vFound = True
                Else
                  vItem = vItem.Replace(vTheString, sFirstWord)
                  vFound = True
                End If

                Exit For
              End If
            Next
            If vFound Then
              start = vItem.ToString().Length
            Else
              Exit While
            End If
          Else
            Exit While
          End If
        End While
        If vFound = False Then
          vItem = "Error"
        End If

        Return vItem

      Catch vException As Exception
        If vFound = False Then
          vItem = "Error"
        End If

        Return vItem
      End Try
    End Function

    Public Function JointItem(ByVal pString As String, ByRef pFirst As String, ByRef pSecond As String, Optional ByRef pTitleName As String = "", Optional ByVal pTitles As SortedList(Of String, Title) = Nothing, Optional ByVal pTitle As Title = Nothing) As String
      Dim vPos As Integer
      Dim vSep As String = "&"
      Dim vCheckJoint As Boolean
      Dim vJointTitle As Boolean
      If pTitleName.Length > 0 Then
        If pTitle Is Nothing Then
          If pTitles.ContainsKey(pTitleName) Then
            vJointTitle = pTitles(pTitleName).JointTitle
          End If
        Else
          vJointTitle = pTitle.JointTitle
        End If
        If vJointTitle Then
          If pTitleName = pString Then
            'Check for und et e & +
            If InStr(pString, " and ") > 0 Then
              vSep = " and "
            ElseIf InStr(pString, " und ") > 0 Then
              vSep = " und "
            ElseIf InStr(pString, " et ") > 0 Then
              vSep = " et "
            ElseIf InStr(pString, " e ") > 0 Then
              vSep = " e "
            ElseIf InStr(pString, " + ") > 0 Then
              vSep = "+"
            Else
              vSep = "&"
            End If
          Else
            If InStr(pString, " and ") > 0 Then
              vSep = " and "
            ElseIf InStr(pString, " und ") > 0 Then
              vSep = " und "
            ElseIf InStr(pString, " et ") > 0 Then
              vSep = " et "
            ElseIf InStr(pString, " e ") > 0 Then
              vSep = " e "
            ElseIf InStr(pString, " + ") > 0 Then
              vSep = "+"
            Else
              vSep = "&"
            End If
          End If
          vCheckJoint = True
        Else
          vCheckJoint = mvEnv.GetConfigOption("cd_joint_contact_support", True)
        End If
      Else
        vCheckJoint = mvEnv.GetConfigOption("cd_joint_contact_support", True)
      End If
      vPos = InStr(pString, vSep)
      If vPos > 0 And vCheckJoint Then
        pFirst = Trim(Left(pString, vPos - 1))
        pSecond = Trim(Mid(pString, vPos + Len(vSep)))
        Return pFirst & " " & Trim(vSep) & " " & pSecond
      Else
        pFirst = Trim(pString)
        Return pString
      End If
    End Function

    Public Sub ClearDefaultPhoneNumber()
      mvClassFields.Item(ContactFields.DiallingCode).Value = ""
      mvClassFields.Item(ContactFields.StdCode).Value = ""
      mvClassFields.Item(ContactFields.Telephone).Value = ""
      mvClassFields.Item(ContactFields.ExDirectory).Bool = False
    End Sub

    Public Sub DeleteAddress(ByRef pAddressNumber As Integer, ByRef pFromPosition As Boolean)
      Dim vWhereFields As New CDBFields
      Dim vContactsAtAddress As Integer

      'This routine assumes the contact has been initialised with the address that is to be deleted
      If ContactType <> Contact.ContactTypes.ctcOrganisation Then
        'Check if this is the default address for the contact
        If pAddressNumber = AddressNumber Then
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, ProjectText.String20704) 'This Address is the default for the Contact\r\n\r\nAddress cannot be deleted
        End If
        DeleteAddressCheck(pFromPosition)
        Dim vContactAddress As New ContactAddress(mvEnv)
        vContactAddress.Init(ContactNumber, pAddressNumber)
        'Check if an organisation type address
        If Address.AddressType = Address.AddressTypes.ataOrganisation Then
          'Check if the default contact for an organisation - If coming from position delete we will already have checked this
          If (ContactNumber = OrganisationContactNumber) And (pFromPosition = False) Then
            RaiseError(DataAccessErrors.daeCannotDeleteAddress, ProjectText.String20705) 'This Contact is the default contact for the Organisation\r\n\r\nLink cannot be deleted
          Else
            If mvEnv.AuditStyle = CDBEnvironment.AuditStyleTypes.ausExtended Then
              Dim vPositions As List(Of ContactPosition) = GetAddressPositions(pAddressNumber)
              Dim vRoles As List(Of ContactRole) = GetOrganisationRoles()
              If Not pFromPosition Then
                mvEnv.Connection.StartTransaction()
                For Each vPosition As ContactPosition In vPositions
                  vPosition.DeleteOnly(mvEnv.User.UserID, True, 0)
                Next
              End If
              If ContactNumber <> OrganisationContactNumber Then
                mvEnv.Connection.StartTransaction()
                For Each vRole As ContactRole In vRoles
                  vRole.Delete(mvEnv.User.UserID, True)
                Next
              End If
            Else
              mvEnv.Connection.StartTransaction()
              vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
              If Not pFromPosition Then
                vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
                mvEnv.Connection.DeleteRecords("contact_positions", vWhereFields, False)
                vWhereFields.Remove(2)
              End If
              'If this contact is the organisation default and we get here then we must be deleting a position
              'and the contact must have another position at this organisation in which case we should leave the roles
              If ContactNumber <> OrganisationContactNumber Then
                vWhereFields.Add("organisation_number", CDBField.FieldTypes.cftLong, OrganisationNumber)
                mvEnv.Connection.DeleteRecords("contact_roles", vWhereFields, False)
              End If
            End If
          End If
        Else
          vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
          'Count the number of contacts records linked to this address
          vContactsAtAddress = mvEnv.Connection.GetCount("contact_addresses", vWhereFields)
          'If only one contact using this address then delete it
          If vContactsAtAddress = 1 Then
            Dim vAddress As New Address(mvEnv)
            vAddress.Init(pAddressNumber)
            mvEnv.Connection.StartTransaction()
            vAddress.Delete(mvEnv.User.UserID, True)
          End If
        End If
        mvEnv.Connection.StartTransaction()
        vWhereFields.Clear()
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
        mvEnv.Connection.DeleteRecords("contact_address_usages", vWhereFields, False)
        vContactAddress.Delete(mvEnv.User.UserID, True)
        SwitchCurrentAddressToDefault()
        UpdateGoneAway(True)
      Else
        'Check if this is the default address for the organisation
        If pAddressNumber = AddressNumber Then
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, ProjectText.String20706) 'This Address is the default for the Organisation\r\n\r\nAddress cannot be deleted
        Else
          DeleteAddressCheck(False)
          vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
          'Count the number of contacts records linked to this address
          vContactsAtAddress = mvEnv.Connection.GetCount("contact_addresses", vWhereFields)
          If vContactsAtAddress > 0 Then
            RaiseError(DataAccessErrors.daeCannotDeleteAddress, ProjectText.String20707) 'Contacts reference this Address\r\n\r\nAddress cannot be deleted
          Else
            Dim vOrgAddress As New OrganisationAddress(mvEnv)
            vOrgAddress.Init(ContactNumber, pAddressNumber)
            Dim vAddress As New Address(mvEnv)
            vAddress.Init(pAddressNumber)
            mvEnv.Connection.StartTransaction()
            vAddress.Delete(mvEnv.User.UserID, True)
            vWhereFields.Clear()
            vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddressNumber)
            vWhereFields.Add("organisation_number", CDBField.FieldTypes.cftLong, ContactNumber)
            mvEnv.Connection.DeleteRecords("organisation_address_usages", vWhereFields, False)
            vOrgAddress.Delete(mvEnv.User.UserID, True)
            SwitchCurrentAddressToDefault()
          End If
        End If
      End If
      If Not pFromPosition Then mvEnv.Connection.CommitTransaction()
    End Sub

    Private Function GetAddressPositions(ByVal pAddressNumber As Integer) As List(Of ContactPosition)
      Dim vPositions As New List(Of ContactPosition)
      Dim vPosition As New ContactPosition(mvEnv)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", ContactNumber)
      vWhereFields.Add("address_number", pAddressNumber)
      Return vPosition.GetList(Of ContactPosition)(vPosition, vWhereFields)
    End Function

    Private Function GetOrganisationRoles() As List(Of ContactRole)
      Dim vRoles As New List(Of ContactRole)
      Dim vRole As New ContactRole(mvEnv)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", ContactNumber)
      vWhereFields.Add("organisation_number", OrganisationNumber)
      Return vRole.GetList(Of ContactRole)(vRole, vWhereFields)
    End Function

    Friend Sub DeleteAddressCheck(ByVal pPosition As Boolean)
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
      vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, Address.AddressNumber)
      If mvEnv.Connection.GetCount("contact_mailings", vWhereFields) > 0 Then 'Check if there any contact mailing records for this address
        If pPosition Then
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16520)) 'Mailing History records relate to this Position\r\n\r\nThe Position cannot be deleted
        Else
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16521)) 'Mailing History records relate to this Address\r\n\r\nThe Address cannot be deleted
        End If
      ElseIf mvEnv.Connection.GetCount("communications", vWhereFields) > 0 Then  'Check if there any communications records for this address
        If pPosition Then
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16522)) 'Communications records relate to this Position\r\n\r\nThe Position cannot be deleted
        Else
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16523)) 'Communications records relate to this Address\r\n\r\nThe Address cannot be deleted
        End If
      ElseIf mvEnv.Connection.GetCount("communications_log", vWhereFields) > 0 Then  'Check if there any communications_log records for this address
        If pPosition Then
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16524)) 'Communications Log records relate to this Position\r\n\r\nThe Position cannot be deleted
        Else
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16525)) 'Communications Log records relate to this Address\r\n\r\nThe Address cannot be deleted
        End If
      ElseIf mvEnv.Connection.GetCount("financial_history", vWhereFields) > 0 Then  'Check if there any financial_history for this address
        If pPosition Then
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16526)) 'Financial History records relate to this Position\r\n\r\nThe Position cannot be deleted
        Else
          RaiseError(DataAccessErrors.daeCannotDeleteAddress, (ProjectText.String16527)) 'Financial History records relate to this Address\r\n\r\nThe Address cannot be deleted
        End If
      End If
    End Sub

    Public Sub SwitchCurrentAddress(ByVal pOldAddress As Address, ByVal pNewAddress As Address, ByVal pUpdateMembers As Boolean)
      'Update subsidiary tables and set their address_number attributes to the specified address
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vChangeBranch As Boolean
      Dim vStartTrans As Boolean
      Dim vTableList As String
      Dim vTables() As String
      Dim vIndex As Integer
      Dim vSQL As String
      Dim vEventSQL As String
      Dim vRecordSet As CDBRecordSet
      Dim vAddressChangeWithBranch As String

      'Optimise this if setting the default address as historic as the update will not have any effect
      If pOldAddress.AddressNumber <> pNewAddress.AddressNumber Then
        vAddressChangeWithBranch = mvEnv.GetConfig("cd_address_change_with_branch")
        If Len(vAddressChangeWithBranch) = 0 Then vAddressChangeWithBranch = "N"
        If (vAddressChangeWithBranch <> "N" And pUpdateMembers) Then vChangeBranch = True
        If mvEnv.GetConfig("me_synchronise_branch") = "Y" _
          And (mvCurrentAddress.AddressNumber = pOldAddress.AddressNumber) Then
          'i.e. synchronising branch on default address is required and the default address is being replaced.
          vChangeBranch = True
        End If
        If Not mvEnv.Connection.InTransaction Then
          mvEnv.Connection.StartTransaction()
          vStartTrans = True
        End If
        vTableList = "back_order_details,bankers_orders,direct_debits,covenants,subscriptions,"
        vTableList = vTableList & "gift_aid_donations,selected_contacts,new_orders,selected_contacts_temp,"
        vTableList = vTableList & "credit_card_authorities,credit_customers,selection_set_data,invoices"
        If Not (vChangeBranch And vAddressChangeWithBranch = "P") Then vTableList = vTableList & ",orders"
        vTables = Split(vTableList, ",")
        vUpdateFields.Add("address_number", CDBField.FieldTypes.cftLong, pNewAddress.AddressNumber)
        For vIndex = 0 To UBound(vTables)
          vWhereFields.Clear()
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
          vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pOldAddress.AddressNumber)
          Select Case vTables(vIndex)
            Case "selected_contacts", "selected_contacts_temp", "credit_customers", "selection_set_data"
              'dont use cancellation reason
            Case "gift_aid_donations"
              vWhereFields.Add("claim_number", CDBField.FieldTypes.cftLong) 'Claim number is null
            Case "new_orders"
              vWhereFields.Add("date_fulfilled", CDBField.FieldTypes.cftDate) 'Date Fulfilled is null
            Case "back_order_details"
              vWhereFields.Add("ordered", CDBField.FieldTypes.cftLong, "issued", CDBField.FieldWhereOperators.fwoNotEqual) 'ordered <> issued
            Case "invoices"
              vWhereFields.Add("reprint_count", -1)
            Case Else
              vWhereFields.Add("cancellation_reason") 'Cancellation Reason is null
          End Select
          mvEnv.Connection.UpdateRecords(vTables(vIndex), vUpdateFields, vWhereFields, False)
        Next
        'Special handling for members
        vWhereFields.Clear()
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pOldAddress.AddressNumber)
        vWhereFields.Add("cancellation_reason") 'Cancellation Reason is null

        If vChangeBranch = True And pOldAddress.Branch <> "" And pNewAddress.Branch <> "" Then
          Member.UpdateMembersBranch(mvEnv, pOldAddress.AddressNumber, pNewAddress.AddressNumber, pOldAddress.Branch, pNewAddress.Branch)
        Else
          'Just update the address
          mvEnv.Connection.UpdateRecords("members", vUpdateFields, vWhereFields, False)
        End If
        'Special handling for orders
        If vChangeBranch And vAddressChangeWithBranch = "P" Then mvEnv.Connection.UpdateRecords("orders", vUpdateFields, vWhereFields, False)

        'Special case for order_details - only update if the paymentplan is not cancelled
        vSQL = "UPDATE order_details SET address_number = " & pNewAddress.AddressNumber
        vSQL = vSQL & " WHERE contact_number = " & ContactNumber & " AND address_number = " & pOldAddress.AddressNumber
        vSQL = vSQL & " AND order_number IN (SELECT o.order_number FROM order_details od, orders o WHERE od.contact_number = " & ContactNumber
        vSQL = vSQL & " AND od.address_number = " & pOldAddress.AddressNumber & " AND o.order_number = od.order_number"
        vSQL = vSQL & " AND o.cancellation_reason IS NULL)"
        mvEnv.Connection.ExecuteSQL(vSQL, CDBConnection.cdbExecuteConstants.sqlIgnoreError)

        'Special cases for Events related tables
        'Only updated if the Event is not in the past or cancelled, old address = default address
        'SDT 6/12/2004 BR8822 Removed test for old address made historic to keep consistent with other commitments
        vEventSQL = "SELECT DISTINCT ev.event_number FROM events ev, sessions s WHERE ev.cancellation_reason IS NULL"
        vEventSQL = vEventSQL & " AND s.event_number = ev.event_number AND s.end_date" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (TodaysDate()))

        'First update room_booking_links table (this SQL is much more complicated!!)
        vSQL = "UPDATE room_booking_links SET address_number = " & pNewAddress.AddressNumber
        vSQL = vSQL & " WHERE contact_number = " & ContactNumber & " AND address_number = " & pOldAddress.AddressNumber
        vSQL = vSQL & " AND room_booking_number IN (SELECT DISTINCT crb.room_booking_number FROM contact_room_bookings crb, events ev, sessions s"
        vSQL = vSQL & " WHERE crb.event_number = ev.event_number AND crb.contact_number = " & ContactNumber & " AND crb.cancellation_reason IS NULL AND ev.cancellation_reason IS NULL"
        vSQL = vSQL & " AND s.event_number = ev.event_number AND s.end_date" & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, (TodaysDate())) & ")"
        mvEnv.Connection.ExecuteSQL(vSQL, CDBConnection.cdbExecuteConstants.sqlIgnoreError)

        'SPECIAL HANDLING FOR BACK_ORDERS as it is very difficult to device a generic SQL statement
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT distinct batch_number, transaction_number FROM back_order_details WHERE contact_number = " & ContactNumber & " AND ordered <> issued")
        vUpdateFields.Clear()
        vUpdateFields.Add("address_number", CDBField.FieldTypes.cftLong, pNewAddress.AddressNumber)
        vWhereFields.Clear()
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pOldAddress.AddressNumber)
        vWhereFields.Add("batch_number", CDBField.FieldTypes.cftLong)
        vWhereFields.Add("transaction_number", CDBField.FieldTypes.cftLong)
        With vRecordSet
          While .Fetch() = True
            vWhereFields("batch_number").Value = .Fields("batch_number").Value
            vWhereFields("transaction_number").Value = .Fields("transaction_number").Value
            mvEnv.Connection.UpdateRecords("back_orders", vUpdateFields, vWhereFields, False)
          End While
          .CloseRecordSet()
        End With
        'Now the rest of the tables
        vTableList = "contact_room_bookings,delegates,event_bookings,event_submissions"
        vTables = Split(vTableList, ",")
        For vIndex = 0 To UBound(vTables)
          vSQL = "UPDATE " & vTables(vIndex) & " SET address_number = " & pNewAddress.AddressNumber
          vSQL = vSQL & " WHERE contact_number = " & ContactNumber & " AND address_number = " & pOldAddress.AddressNumber
          vSQL = vSQL & " AND event_number IN (" & vEventSQL & ")"
          Select Case vTables(vIndex)
            Case "contact_room_bookings", "event_bookings"
              vSQL = vSQL & " AND cancellation_reason IS NULL"
          End Select
          mvEnv.Connection.ExecuteSQL(vSQL, CDBConnection.cdbExecuteConstants.sqlIgnoreError)
        Next
        'BR21065 - Update for Exam Booking Mailings
        Dim vExamBooking As New ExamBooking(mvEnv)
        vWhereFields.Clear()
        vWhereFields.Add("eb.contact_number", ContactNumber)
        vWhereFields.Add("eb.address_number", pOldAddress.AddressNumber)
        vWhereFields.Add("eb.cancellation_reason", "")
        Dim vExamJoin As New AnsiJoins
        vExamJoin.AddLeftOuterJoin("exam_sessions es", "es.exam_session_id", "eb.exam_session_id")
        vWhereFields.Add("es.valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrLessThanEqual)
        vWhereFields.Add("es.valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoNullOrGreaterThanEqual)
        Dim vExamSQL As New SQLStatement(mvEnv.Connection, vExamBooking.GetRecordSetFields.Replace("exam_session_id", "eb.exam_session_id"), "exam_bookings eb", vWhereFields, "", vExamJoin)
        Dim vRS As CDBRecordSet = vExamSQL.GetRecordSet
        While vRS.Fetch
          vExamBooking.InitFromRecordSet(vRS)
          Dim vParams As New CDBParameters
          vParams.Add("AddressNumber", pNewAddress.AddressNumber)
          vExamBooking.Update(vParams)
          vExamBooking.Save()
        End While
        vRS.CloseRecordSet()

        If vStartTrans Then mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Private Sub SwitchCurrentAddressToDefault()
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields
      Dim vTables() As String
      Dim vIndex As Integer
      Dim vTableNames As String

      'Update subsidiary tables and set their address_number attribute to the new address
      'This routine may be called due to an address having been deleted or
      'A contact position record may have been deleted
      '--------------------------------------------------------------------------------------------------
      If AddressNumber <> Address.AddressNumber Then
        'Removed communications, communications_log, financial_history, contact_mailings
        'If the address is being deleted there cannot be any records in these tables
        vTableNames = "bankers_orders,caf_voucher_transactions,contact_address_usages,covenants,contact_header,direct_debits,members," & "orders,order_details,selected_contacts,subscriptions,credit_customers,proforma_invoices,proforma_invoice_details,"
        'The following tables are only updated in the CHUI if the related batch has not yet been processed
        vTableNames = vTableNames & "back_orders,back_order_details,batch_transaction_analysis,batch_transactions,credit_sales," & "despatch_notes,event_bookings,gift_aid_donations,invoices,receipts,thank_you_letters,contact_journals,"
        'The following tables were not updated at all in the CHUI
        vTableNames = vTableNames & "selected_subscriptions,selected_members,selected_orders,contact_exports,new_orders,organisers," & "venues,delegates,event_personnel,loan_items,event_submissions,credit_card_authorities"
        vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, Address.AddressNumber)
        vUpdateFields.Add("address_number", CDBField.FieldTypes.cftLong, AddressNumber)
        vTables = Split(vTableNames, ",")
        For vIndex = 0 To UBound(vTables)
          mvEnv.Connection.UpdateRecords(vTables(vIndex), vUpdateFields, vWhereFields, False)
        Next
      End If
    End Sub

    Public Sub UpdateGoneAway(ByRef pSaveContact As Boolean)
      Dim vWhereFields As New CDBFields

      If ContactType <> Contact.ContactTypes.ctcOrganisation Then
        vWhereFields.Add("contact_number", ContactNumber)
        vWhereFields.Add("historical", "N")
        If mvEnv.Connection.GetCount("contact_addresses", vWhereFields) > 0 Then
          RemoveGoneAway(pSaveContact)
        Else
          SetGoneAway(pSaveContact)
        End If
      End If
    End Sub

    Public Sub RemoveGoneAway(ByRef pSaveContact As Boolean)
      RemoveGoneAway(pSaveContact, False, Nothing)
    End Sub

    Public Sub RemoveGoneAway(ByRef pSaveContact As Boolean, ByVal pUpdateAddresses As Boolean, ByVal pAddressValidToDate As Date)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields

      If ContactType <> Contact.ContactTypes.ctcOrganisation Then
        ReadGAConfig()
        If mvGAStatus Then
          If Status.Length > 0 Then
            If Status = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAStatus) Then 'Remove GA Status
              Status = ""
              StatusDate = ""
              StatusReason = ""
              If pSaveContact Then Save()
            End If
          End If
        End If
        If mvGASuppression Then 'Set the GA Suppression as no longer valid
          With vUpdateFields
            .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate)
            .AddAmendedOnBy(mvEnv.User.UserID)
          End With
          With vWhereFields
            .Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
            .Add("mailing_suppression", CDBField.FieldTypes.cftCharacter, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMailingSupp))
          End With
          mvEnv.Connection.UpdateRecords("contact_suppressions", vUpdateFields, vWhereFields, False)
        End If
        If pUpdateAddresses Then
          For Each vAddress As ContactAddress In ContactAddresses()
            If vAddress.Historical AndAlso (ContactAddresses.Count = 1 OrElse Date.Compare(DateValue(vAddress.ValidTo), pAddressValidToDate) = 0) Then
              'If Address was historic and only 1 address Or valid to date matches valid to end date check then un-set historical flag and clear valid to date
              vAddress.ValidTo = String.Empty
              vAddress.Historical = False
              vAddress.Save()
            End If
          Next
        End If
      End If
    End Sub

    Public Function GetJointLinks(Optional ByVal pIncludeHistoricalLinks As Boolean = False) As CollectionList(Of ContactLink)
      Dim vContactLink As New ContactLink(mvEnv)
      Dim vRS As CDBRecordSet
      Dim vSQL As String

      If mvContactLinks Is Nothing Then
        mvContactLinks = New CollectionList(Of ContactLink)
        vContactLink.Init()
        'Build basic SELECT statement
        vSQL = "SELECT " & vContactLink.GetRecordSetFields() & " FROM contact_links cl WHERE "
        'Add the bits that are type specific
        If ContactType = ContactTypes.ctcJoint Then
          'If the current contact is a joint contact then retrieve those individuals in the joint relationship
          vSQL = vSQL & "contact_number_1 = " & ContactNumber
        Else
          'If the current contact is one of the individuals in a joint relationship then retrieve the joint contact
          vSQL = vSQL & "contact_number_2 = " & ContactNumber
        End If
        vSQL = vSQL & " AND relationship IN ('" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlRealToJointLink) & "','" & mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink) & "')"
        'Add the bits to retrieve only current links
        If Not pIncludeHistoricalLinks Then
          vSQL = vSQL & " AND ((valid_from IS NULL OR valid_from " & mvEnv.Connection.SQLLiteral("<=", CDBField.FieldTypes.cftDate, TodaysDate()) & ")"
          vSQL = vSQL & " AND (valid_to IS NULL OR valid_to " & mvEnv.Connection.SQLLiteral(">=", CDBField.FieldTypes.cftDate, TodaysDate()) & ")"
          vSQL = vSQL & " AND (historical IS NULL OR historical = 'N'))"
        End If
        'Add any ordering or grouping
        vSQL = vSQL & " ORDER BY contact_number_1, contact_number_2"
        'Get the links and build collection
        vRS = mvEnv.Connection.GetRecordSet(vSQL)
        With vRS
          While .Fetch()
            vContactLink = New ContactLink(mvEnv)
            vContactLink.InitFromRecordSet(vRS)
            mvContactLinks.Add((vContactLink.ContactNumber1 & vContactLink.ContactNumber2 & vContactLink.RelationshipCode), vContactLink)
          End While
          .CloseRecordSet()
        End With
      End If
      Return mvContactLinks
    End Function

    Public ReadOnly Property ContactAddresses() As Collection
      Get
        Dim vContactAddress As ContactAddress
        If mvAddresses Is Nothing Then
          mvAddresses = New Collection
          Dim vAddress As Address
          vAddress = New Address(mvEnv)
          vContactAddress = New ContactAddress(mvEnv)
          vAddress.Init()
          vContactAddress.Init()
          Dim vRecordSet As CDBRecordSet
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vContactAddress.GetRecordSetFields() & ", " & vAddress.GetRecordSetFieldsCountry.Replace("a.address_number,", "") & " FROM contact_addresses ca, addresses a, countries co WHERE contact_number = " & ContactNumber & " AND ca.address_number = a.address_number AND a.country = co.country ORDER BY historical, town")
          With vRecordSet
            While .Fetch()
              vAddress = New Address(mvEnv)
              vContactAddress = New ContactAddress(mvEnv)
              vAddress.InitFromRecordSetCountry(vRecordSet)
              vContactAddress.InitFromRecordSet(vRecordSet)
              vContactAddress.Address = vAddress
              mvAddresses.Add(vContactAddress, vContactAddress.AddressNumber.ToString)
            End While
            .CloseRecordSet()
          End With
        End If
        Return mvAddresses
      End Get
    End Property

    Public Sub MarkAsGoneAway(Optional ByVal pBatchNumber As Integer = 0, Optional ByVal pTransactionNumber As Integer = 0, Optional ByVal pSaveContact As Boolean = True, Optional ByVal pSetStatusReason As Boolean = True)
      Dim vContactAddress As ContactAddress
      If ContactType <> ContactTypes.ctcOrganisation Then
        SetGoneAway(pSaveContact, pSetStatusReason)
        For Each vContactAddress In ContactAddresses
          With vContactAddress
            If IsDate(.ValidFrom) Then
              If CDate(.ValidFrom) > Today Then .ValidFrom = TodaysDate()
            End If
            .Historical = True
            .ValidTo = TodaysDate()
            .Save()
          End With
        Next vContactAddress
        mvEnv.AddJournalRecord(JournalTypes.jnlGoneAway, JournalOperations.jnlUpdate, ContactNumber, AddressNumber, 0, 0, 0, pBatchNumber, pTransactionNumber)
      End If
    End Sub

    Public Sub SetGoneAway(ByRef pSaveContact As Boolean)
      SetGoneAway(pSaveContact, True)
    End Sub

    Public Sub SetGoneAway(ByRef pSaveContact As Boolean, ByVal pSetStatusReason As Boolean)
      If ContactType <> ContactTypes.ctcOrganisation _
      AndAlso Not Status.Equals(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDeceasedStatus)) Then
        ReadGAConfig()
        If mvGAStatus Then
          Status = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAStatus)
          If pSetStatusReason Then
            StatusReason = "automatic"
          End If
          If pSaveContact Then Save()
        End If
        If mvGASuppression Then
          Dim vSuppression As New ContactSuppression(mvEnv)
          vSuppression.SaveSuppression(ContactSuppression.SuppressionEntryStyles.sesNormal, ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMailingSupp), "", "", "")
        End If
      End If
    End Sub

    Public Sub SaveChanges(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      MyBase.Save(pAmendedBy, pAudit)
    End Sub

    Public ReadOnly Property Position() As String
      Get
        GetPositionInfo()
        Return mvPosition
      End Get
    End Property

    Public ReadOnly Property PositionLocation() As String
      Get
        GetPositionInfo()
        PositionLocation = mvPositionLocation
      End Get
    End Property

    Public ReadOnly Property OrganisationName() As String
      Get
        GetPositionInfo()
        Return mvOrganisationName
      End Get
    End Property

    Public ReadOnly Property OrganisationNumber() As Integer
      Get
        GetPositionInfo()
        Return mvOrganisationNumber
      End Get
    End Property

    Public ReadOnly Property OrganisationContactNumber() As Integer
      Get
        GetPositionInfo()
        Return mvOrgContactNumber
      End Get
    End Property

    Public ReadOnly Property ContactMemberStatus() As ContactMemberStatuses
      Get
        ContactMemberStatus = ContactMemberStatuses.cmsNone
        Dim vRecordSet As CDBRecordSet = mvEnv.Connection.GetRecordSet("SELECT single_membership FROM members m, membership_types mt WHERE contact_number = " & ContactNumber & " AND cancellation_reason IS NULL AND m.membership_type = mt.membership_type")
        While vRecordSet.Fetch()
          If vRecordSet.Fields(1).Bool Then
            Return ContactMemberStatuses.cmsSingle
          Else
            If ContactMemberStatus = ContactMemberStatuses.cmsNone Then Return ContactMemberStatuses.cmsMember
          End If
        End While
        vRecordSet.CloseRecordSet()
      End Get
    End Property

    Public ReadOnly Property GiftAidDeclarations() As CollectionList(Of GiftAidDeclaration)
      Get
        Dim vGAD As GiftAidDeclaration
        Dim vRecordSet As CDBRecordSet
        Dim vContactLink As ContactLink
        Dim vContactNos As String = ""

        If mvGiftAidDeclarations Is Nothing Then
          If ContactType = ContactTypes.ctcJoint Then
            'Need to retrieve any Declarations for the individuals
            For Each vContactLink In GetJointLinks(True)
              vContactNos = vContactNos & "," & vContactLink.ContactNumber2
            Next vContactLink
          End If
          If vContactNos.Length > 0 Then
            vContactNos = " IN (" & ContactNumber & vContactNos & ")"
          Else
            vContactNos = " = " & ContactNumber
          End If
          mvGiftAidDeclarations = New CollectionList(Of GiftAidDeclaration)
          vGAD = New GiftAidDeclaration
          vGAD.Init(mvEnv, pRaiseNoGAControlError:=False)
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vGAD.GetRecordSetFields(GiftAidDeclaration.GiftAidDeclarationRecordSetTypes.gadrtAll) & " FROM gift_aid_declarations WHERE contact_number " & vContactNos & " ORDER BY start_date")
          While vRecordSet.Fetch()
            vGAD = New GiftAidDeclaration
            vGAD.InitFromRecordSet(mvEnv, vRecordSet, GiftAidDeclaration.GiftAidDeclarationRecordSetTypes.gadrtAll)
            mvGiftAidDeclarations.Add(vGAD.DeclarationNumber.ToString, vGAD)
          End While
          vRecordSet.CloseRecordSet()
        End If
        Return mvGiftAidDeclarations
      End Get
    End Property

    Public Sub DoContactMerge(ByRef pJob As JobSchedule, ByVal pConn As CDBConnection, ByRef pDContact As Contact, ByVal pOAddress As Integer, ByVal pDAddress As Integer, ByVal pDBANotes As String)
      Dim vWhereFields As New CDBFields
      Dim vInsertFields As New CDBFields
      Dim vGADec As New GiftAidDeclaration

      If GiftAidDeclarations.Count() > 0 Then
        If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason).Length = 0 Then
          RaiseError(DataAccessErrors.daeGAMergeCancellationNotSet)
        End If
      End If
      'BC 4851:Remove any Principal users for Duplicate, we don't want to merge
      'this onto the original
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        vWhereFields.Add("contact_number", pDContact.ContactNumber, CDBField.FieldWhereOperators.fwoEqual)
        mvEnv.Connection.DeleteRecords("principal_users", vWhereFields, False)
        vWhereFields.Clear()
      End If

      'BR 11575: Do not perform GA Merge if either Contact is Joint
      If ContactType <> ContactTypes.ctcJoint And pDContact.ContactType <> ContactTypes.ctcJoint Then
        'If duplicate Contact has Declarations then need to update unclaimed lines at end
        If pDContact.GiftAidDeclarations.Count() > 0 Then
          mvUpdateGAData = True
          pJob.InfoMessage = "Check for overlapping Gift Aid Declarations"
          DoGAMerge(pDContact)
        End If

        If Me.GiftAidDeclarations.Count() > 0 Then
          'Flag to process transactions of duplicate contact and its contact_links (e.g. joint contact) against the gift aid declarations  
          mvUpdateGAData = True
        End If
      End If

      'Sort out any Positions before merging the data
      pJob.InfoMessage = ProjectText.String16507 'Checking Contact Positions
      DoContactPositionMerge(pConn, pDContact)

      'Check for overlapping Supressions and merge them
      DoSuppressionsMerge(pDContact)

      pJob.InfoMessage = ProjectText.String20708 'Check for overlapping Categories
      DoCategoriesMerge(pDContact)

      DoContactCommunicationsDeDup(pConn, ContactNumber, pDContact.ContactNumber) 'BR17095 Clear duplicate device default.

      GetContactMergeInfo(pConn, True)
      'Change the contact number - and address number if appropriate
      For Each vMergeInfo As ContactMergeInfo In mvMergeInfo
        ChangeContact(pJob, pConn, vMergeInfo.TableName, ContactNumber, pDContact.ContactNumber, pOAddress, pDAddress, vMergeInfo.SetAmend, vMergeInfo.ContactAttr, vMergeInfo.AddressAttr, vMergeInfo.UniqueAttrs)
      Next

      'now do the default address and the contact and the dba note
      pJob.InfoMessage = ProjectText.String31257 'Deleting Duplicate Contact
      pConn.StartTransaction()

      Me.DoExamsMerge(pDContact)

      'Find if there are any other links to the duplicate contacts default address if not then delete it
      'NB The confusing thing here is that if duplicate contact address passed to this routine is 0
      'there will always be at least one record found and the address will not be deleted!
      If pDContact.AddressNumber > 0 Then
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pDContact.AddressNumber)
        If pConn.GetCount("contact_addresses", vWhereFields) = 0 Then
          Dim vDupAddress As Address = pDContact.Address
          If Not (vDupAddress IsNot Nothing AndAlso vDupAddress.AddressNumber.Equals(pDContact.AddressNumber)) Then
            'Shouldn't actually happen, but just in case ....
            vDupAddress = New Address(Me.Environment)
            vDupAddress.Init(pDContact.AddressNumber)
          End If
          If vDupAddress IsNot Nothing AndAlso vDupAddress.AddressNumber.Equals(pDContact.AddressNumber) Then
            vDupAddress.Delete(Me.Environment.User.UserID, True)
          End If
        End If
      End If

      'Merge the contact details if necessary
      If TitleName = "" And pDContact.TitleName <> "" Then
        TitleName = pDContact.TitleName
      End If
      If Forenames = "" And pDContact.Forenames <> "" Then
        Forenames = pDContact.Forenames
      End If
      If Initials = "" And pDContact.Initials <> "" Then
        Initials = pDContact.Initials
      End If
      If Honorifics = "" And pDContact.Honorifics <> "" Then
        Honorifics = pDContact.Honorifics
        LabelName = LabelName & " " & Honorifics
      End If
      If DateOfBirth = "" And pDContact.DateOfBirth <> "" Then
        DateOfBirth = pDContact.DateOfBirth
        DobEstimated = pDContact.DobEstimated
      End If
      If Telephone = "" And pDContact.Telephone <> "" Then
        DiallingCode = pDContact.DiallingCode
        StdCode = pDContact.StdCode
        Telephone = pDContact.Telephone
        ExDirectory = pDContact.ExDirectory
      End If
      'BR19025 if control set then use source_date and source code from the oldest contact on the master contact
      If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMergeUseOldestSource) = "Y" Then
        'since both contact records are already set we only need to update to use duplicate contact source date and code if duplicate date is later than master
        If Date.Compare(DateValue(SourceDate), DateValue(pDContact.SourceDate)) > 0 Then
          SourceDate = pDContact.SourceDate
          Source = pDContact.Source
        End If
      End If
      SaveChanges()

      'Delete existing duplicate contact records
      vWhereFields.Clear()
      vWhereFields.Add("contact_number_1", CDBField.FieldTypes.cftLong, ContactNumber)
      vWhereFields.Add("contact_number_2", CDBField.FieldTypes.cftLong, pDContact.ContactNumber)
      pConn.DeleteRecords("duplicate_contacts", vWhereFields, False)
      vWhereFields(1).Value = CStr(pDContact.ContactNumber)
      vWhereFields(2).Value = CStr(ContactNumber)
      pConn.DeleteRecords("duplicate_contacts", vWhereFields, False)

      'Insert dba notes
      vInsertFields.Add("master", CDBField.FieldTypes.cftLong, ContactNumber)
      vInsertFields.Add("duplicate", CDBField.FieldTypes.cftLong, pDContact.ContactNumber)
      vInsertFields.Add("notes", CDBField.FieldTypes.cftMemo, pDBANotes)
      vInsertFields.Add("merged_on", CDBField.FieldTypes.cftDate, TodaysDate())
      pConn.InsertRecord("dba_notes", vInsertFields)

      vWhereFields.Clear()
      If pDContact IsNot Nothing AndAlso pDContact.ContactNumber > 0 Then
        pDContact.DeleteMergedContact(Me.Environment.User.UserID, True, 0)
      End If

      pConn.CommitTransaction()

      'Now, if we have merged Financial History data or duplicate Contact had a Declaration
      'we need to update the Gift Aid unclaimed lines
      If mvUpdateGAData Then
        pJob.InfoMessage = "Updating Gift Aid Declarations unclaimed lines"
        mvGiftAidDeclarations = Nothing 'Force reselection of GA Declarations following merge
        For Each vGADec In GiftAidDeclarations
          vGADec.PopulateUnclaimedLines(True)
        Next vGADec
      End If
    End Sub

    Public Sub DoExamsMerge(pDuplicateContact As Contact)

      Dim vExamMerge As IMergeOperation = New ContactExamMergeData(Me.Environment, Me, pDuplicateContact)

      vExamMerge.ExecuteOperation()

    End Sub


    Public Sub DoOrgMerge(ByRef pJob As JobSchedule, ByVal pConn As CDBConnection, ByVal pDOrg As Contact, ByVal pOAddress As Integer, ByVal pDAddress As Integer, ByVal pDBANotes As String, ByVal pDelete As Boolean)
      Dim vWhereFields As New CDBFields
      Dim vInsertFields As New CDBFields

      Dim vDupOrganisation As New Organisation(Me.Environment)
      vDupOrganisation.InitWithAddress(Me.Environment, pDOrg.ContactNumber, pDAddress)

      pJob.InfoMessage = ProjectText.String31254 'Processing Dummy Contact

      pConn.StartTransaction()
      'Remove links to the dummy contact which are not required
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pDOrg.ContactNumber)
      pConn.DeleteRecords("contact_roles", vWhereFields, False)
      pConn.DeleteRecords("contact_positions", vWhereFields, False)
      pConn.DeleteRecords("contact_addresses", vWhereFields, False)
      'BC 4851:Remove any Principal users for Duplicate, we don't want to merge
      'this onto the original
      If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipGroups Then
        mvEnv.Connection.DeleteRecords("principal_users", vWhereFields, False)
      End If
      pConn.CommitTransaction()

      'Sort out any Positions before merging the data
      pJob.InfoMessage = ProjectText.String16507 'Checking Contact Positions
      DoContactPositionMerge(pConn, pDOrg, False, Not (pDelete))

      DoSuppressionsMerge(pDOrg)

      pJob.InfoMessage = ProjectText.String20708 'Check for overlapping Categories
      DoCategoriesMerge(pDOrg)

      'Change the contact number - and address number if appropriate on tables that might link to the dummy contact
      GetContactMergeInfo(pConn, True)
      For Each vMergeInfo As ContactMergeInfo In mvMergeInfo
        If vMergeInfo.TableName <> "organisations" Then
          ChangeContact(pJob, pConn, vMergeInfo.TableName, ContactNumber, (pDOrg.ContactNumber), pOAddress, pDAddress, vMergeInfo.SetAmend, vMergeInfo.ContactAttr, vMergeInfo.AddressAttr, vMergeInfo.UniqueAttrs)
        End If
      Next

      If pDelete Then 'If deleting the duplicate organisations address
        'Process change address number
        For Each vMergeInfo As ContactMergeInfo In mvMergeInfo
          If vMergeInfo.TableName <> "organisations" Then
            If vMergeInfo.TableName = "contact_addresses" Then
              ' Contact_addresses need to be deduplicated BR17546
              DeleteDuplicateContactAddress(pJob, pConn, pOAddress, pDAddress)
            End If
            If vMergeInfo.AddressAttr.Length > 0 Then
              ChangeOrgAddress(pJob, pConn, vMergeInfo.TableName, pOAddress, pDAddress)
            End If
          End If
        Next
        'Since the contacts table will not be done above do it here
        ChangeOrgAddress(pJob, pConn, "contacts", pOAddress, pDAddress)

        pJob.InfoMessage = ProjectText.String31255 'Deleting Duplicate Organisation Address
        pConn.StartTransaction()
        vWhereFields.Clear()
        vWhereFields.Add("organisation_number", CDBField.FieldTypes.cftLong, pDOrg.ContactNumber)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pDAddress)
        pConn.DeleteRecords("organisation_addresses", vWhereFields)

        Dim vDupAddress As Address = vDupOrganisation.Address
        If Not (vDupAddress IsNot Nothing AndAlso vDupAddress.AddressNumber.Equals(pDAddress)) Then
          'Shouldn't actually happen, but just in case ....
          vDupAddress = New Address(Me.Environment)
          vDupAddress.Init(pDAddress)
        End If
        If vDupAddress IsNot Nothing AndAlso vDupAddress.AddressNumber.Equals(pDAddress) Then
          vDupAddress.Delete(Me.Environment.User.UserID, True)
        End If

        pConn.CommitTransaction()
      End If

      'Change the organisation number - and address number if appropriate on other tables that might link to the duplicate organisation
      mvMergeInfoValid = False
      GetContactMergeInfo(pConn, False)
      mvMergeInfoValid = False
      For Each vMergeInfo As ContactMergeInfo In mvMergeInfo
        ChangeOrg(pJob, pConn, vMergeInfo.TableName, pDOrg.ContactNumber, vMergeInfo.ContactAttr, vMergeInfo.UniqueAttrs)
      Next
      pJob.InfoMessage = ProjectText.String31256 'Deleting Organisation Record
      pConn.StartTransaction()

      'BR19025 if control set then use source_date and source code from the oldest contact on the master contact
      If mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlMergeUseOldestSource) = "Y" Then
        'since both contact records are already set we only need to update to use duplicate contact source date and code if duplicate date is later than master
        If Date.Compare(DateValue(SourceDate), DateValue(pDOrg.SourceDate)) > 0 Then

          'update both tables with same data update fields
          Dim vUpdateFields As New CDBFields
          vUpdateFields.Add("source_date", CDBField.FieldTypes.cftDate, pDOrg.SourceDate)
          vUpdateFields.Add("source", pDOrg.Source)

          'update organisations table
          vWhereFields.Clear()
          vWhereFields.Add("organisation_number", CDBField.FieldTypes.cftLong, OrganisationNumber)
          pConn.UpdateRecords("organisations", vUpdateFields, vWhereFields)

          'update contacts table
          vWhereFields.Clear()
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, OrganisationNumber)
          pConn.UpdateRecords("contacts", vUpdateFields, vWhereFields)
        End If
      End If

      'Deleting from duplicate_contacts
      vWhereFields.Clear()
      vWhereFields.Add("contact_number_1", CDBField.FieldTypes.cftLong, ContactNumber)
      vWhereFields.Add("contact_number_2", CDBField.FieldTypes.cftLong, pDOrg.ContactNumber)
      pConn.DeleteRecords("duplicate_contacts", vWhereFields, False)
      vWhereFields(1).Value = CStr(pDOrg.ContactNumber)
      vWhereFields(2).Value = CStr(ContactNumber)
      pConn.DeleteRecords("duplicate_contacts", vWhereFields, False)

      'Deleting from organisations & dummy contact
      Dim vDupOrgNumber As Integer = vDupOrganisation.OrganisationNumber
      pDOrg.DeleteMergedContact(Me.Environment.User.UserID, True, 0)
      vDupOrganisation.DeleteMergedOrganisation(Me.Environment.User.UserID, True, 0)

      'Insert dba notes
      vInsertFields.Add("master", CDBField.FieldTypes.cftLong, ContactNumber)
      vInsertFields.Add("duplicate", CDBField.FieldTypes.cftLong, vDupOrgNumber)
      vInsertFields.Add("notes", CDBField.FieldTypes.cftMemo, pDBANotes)
      vInsertFields.Add("merged_on", CDBField.FieldTypes.cftDate, TodaysDate())
      pConn.InsertRecord("dba_notes", vInsertFields)
      pConn.CommitTransaction()
    End Sub

    Private Sub DoGAMerge(ByRef pDContact As Contact)
      'This will run through all the Gift Aid Declarations linked to both the primary Contact
      'and duplicate Contact check whether any of these Declarations will overlap once they are merged.
      'Any potential overlaps will be handled by a combination of changing the start and/or end dates,
      'creating new Declarations and cancelling Declarations
      Dim vGADec As GiftAidDeclaration 'Other GAD object used during processing
      Dim vOrigGADec As GiftAidDeclaration 'GAD for primary Contact
      Dim vDupGADec As GiftAidDeclaration 'GAD for duplicate Contact
      Dim vNewGADec As GiftAidDeclaration 'New GAD created due to an overlap
      Dim vNewGADs As Collection 'Collection of vNewGADec objects
      Dim vOrigDecType As GiftAidDeclaration.GiftAidDeclarationTypes
      Dim vCancelDGAD As Boolean 'Cancel the duplicate GAD
      Dim vCancelOGAD As Boolean 'Cancel the original GAD
      Dim vContinue As Boolean
      Dim vDateChange As Boolean 'GAD dates have been changed
      Dim vGADChanged As Boolean 'Used to see if vNewGAD dates have been amended before the Save
      Dim vMergeCancelRsn As String
      Dim vNewGADNos As String 'GAD numbers of new GAD's created
      Dim vNewEndDate As String = "" 'Amended GAD end date
      Dim vNewStartDate As String = "" 'Amended GAD start date
      Dim vNotes As String = ""
      Dim vOverlap As Boolean 'The GAD's had an overlap which may require the type updating
      Dim vTrans As Boolean 'A database Transaction has been started

      vMergeCancelRsn = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlGAMergeCancellationReason)

      For Each vDupGADec In pDContact.GiftAidDeclarations
        vCancelDGAD = False
        If vDupGADec.CancellationReason <> vMergeCancelRsn Then 'If the GAD was cancelled because it was previously merged from another contact, then don't process it
          For Each vOrigGADec In GiftAidDeclarations
            vOverlap = False
            vDateChange = False
            vCancelOGAD = False
            vContinue = True
            vNewGADs = Nothing
            vNewGADec = Nothing
            If ((vOrigGADec.BatchNumber > 0 Or vOrigGADec.PaymentPlanNumber > 0) Or (vDupGADec.BatchNumber > 0 Or vDupGADec.PaymentPlanNumber > 0)) Then
              'These GAD's are specifically linked and overlaps are allowed
              vContinue = False
            ElseIf vOrigGADec.CancellationReason = vMergeCancelRsn Then
              'The GAD was cancelled because it was previously merged from another contact, so don't process it
              vContinue = False
            End If
            If vContinue Then

              '(1) Duplicate GAD starts before original GAD and overlaps it
              If GADateCompare(vDupGADec.StartDate, vOrigGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan And GADateCompare(vDupGADec.EndDate, vOrigGADec.StartDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                '(a) Both GAD's are live
                If (vOrigGADec.CancellationReason.Length = 0 And vDupGADec.CancellationReason.Length = 0) Then
                  'Extend Original GAD to cover start of Duplicate GAD
                  vNewStartDate = vDupGADec.StartDate
                  vNewEndDate = vOrigGADec.EndDate
                  vDateChange = True
                  vOverlap = True
                  vCancelDGAD = True 'Cancel the duplicate GAD

                  If GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                    'Duplicate GAD ends after original GAD so extend the end date
                    vNewEndDate = vDupGADec.EndDate
                  End If

                  '(b) Duplicate GAD has been cancelled
                ElseIf (vOrigGADec.CancellationReason.Length = 0 And vDupGADec.CancellationReason.Length > 0) Then
                  If vDupGADec.CancellationReason = vMergeCancelRsn Then
                    'Cancellation reason is merge cancellation reason so overlap is OK
                  Else
                    vNewStartDate = vOrigGADec.StartDate
                    vNewEndDate = vOrigGADec.EndDate
                    If GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                      'Cancelled duplicate GAD ends before live original GAD
                      If vOrigGADec.HasTaxClaims Then
                        'Live original GAD needs to be cancelled and new GAD created for non-overlapping period
                        If GADateCompare(vOrigGADec.EarliestClaimedPaymentDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                          'Earliest claimed payment date after end date of cancelled GAD
                          vNewStartDate = CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                          vDateChange = True
                          vOverlap = True
                        Else
                          'Need to create a new GAD to cover the period from the end date of the duplicate to the end date of the original
                          vNewGADs = New Collection
                          vNewGADec = New GiftAidDeclaration
                          vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat), vOrigGADec.EndDate)
                          vNewGADs.Add(vNewGADec)
                          If GADateCompare(vOrigGADec.LatestClaimedPaymentDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                            'Latest claimed payment dated before end date of cancelled GAD
                            'So change end date of original GAD to be the end date of the cancelled duplicate GAD
                            vNewEndDate = vDupGADec.EndDate
                            vDateChange = True
                          Else 'If GADateCompare(vOrigGADec.LatestClaimedPaymentDate, vOrigGADec.EndDate) = cgamdLessThan Then
                            'Latest claimed payment dated on/before end date of original GAD
                            'so change end date of original GAD to be the date of the last claimed payment
                            vNewEndDate = vOrigGADec.LatestClaimedPaymentDate
                            vDateChange = True
                          End If
                          vOverlap = True
                          vCancelOGAD = True 'Cancel the original GAD
                        End If
                      Else
                        'No tax claims so amend the start date
                        vNewStartDate = CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                        vDateChange = True
                        vOverlap = True
                      End If
                    Else
                      'Cancelled duplicate GAD covers wider date range than live GAD
                      'So cancel live original GAD as it is not required
                      vCancelOGAD = True
                      vOverlap = True
                    End If
                  End If

                  '(c) Original GAD has been cancelled
                ElseIf vOrigGADec.CancellationReason.Length > 0 And vDupGADec.CancellationReason.Length = 0 Then
                  If vOrigGADec.CancellationReason = vMergeCancelRsn Then
                    'Cancellation reason is merge cancellation reason so overlap is OK
                  Else
                    vNewStartDate = vDupGADec.StartDate
                    vNewEndDate = vDupGADec.EndDate
                    If vDupGADec.HasTaxClaims Then
                      If GADateCompare(vDupGADec.EarliestClaimedPaymentDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                        'First tax claim dated after the end date of the original GAD
                        'Set start date on duplicte GAD to be after end of original GAD and create new GAD for the start
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, vDupGADec.StartDate, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vOrigGADec.StartDate))))
                        vNewGADs.Add(vNewGADec)
                        vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOrigGADec.EndDate)))
                        vDateChange = True
                        vOverlap = True
                      ElseIf GADateCompare(vDupGADec.EarliestClaimedPaymentDate, vOrigGADec.StartDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                        'First tax claim dated during the overlap with the original GAD
                        'Change dates of duplicate GAD and create new GAD for the start overlap
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, vDupGADec.StartDate, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vOrigGADec.StartDate))))
                        vNewGADs.Add(vNewGADec)
                        vNewStartDate = vDupGADec.EarliestClaimedPaymentDate
                        vDateChange = True
                        vOverlap = True
                        vCancelDGAD = True 'Cancel the duplicate GAD
                        'Create a new GAD for the end overlap
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOrigGADec.EndDate))), vDupGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                      Else
                        'First tax claim dated before start date of original GAD
                        vNewGADs = New Collection
                        If GADateCompare(vDupGADec.LatestClaimedPaymentDate, vOrigGADec.StartDate) <> ContactGiftAidMergeDates.cgamdLessThan Then
                          'Last tax claim dated on/after start date of original GAD
                          vNewGADec = New GiftAidDeclaration
                          vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, vDupGADec.StartDate, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vOrigGADec.StartDate))))
                          vNewGADs.Add(vNewGADec)
                          vCancelDGAD = True
                        End If
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOrigGADec.EndDate))), vDupGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                        vNewEndDate = If(GADateCompare(vDupGADec.LatestClaimedPaymentDate, vOrigGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan, CDate(vOrigGADec.StartDate).AddDays(-1).ToString(CAREDateFormat), vDupGADec.LatestClaimedPaymentDate)
                        vDateChange = True
                        vOverlap = True
                      End If
                    Else
                      'No tax claims
                      If GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                        'Duplicate GAD covers longer period than original GAD
                        'So create new GAD
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOrigGADec.EndDate))), vDupGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                      End If
                      'Change end date of duplicate GAD
                      vNewEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vOrigGADec.StartDate)))
                      vDateChange = True
                      vOverlap = True
                    End If
                  End If
                End If
              End If

              '(2) Duplicate GAD starts after original GAD but before original GAD ends
              If (GADateCompare(vDupGADec.StartDate, vOrigGADec.StartDate) = ContactGiftAidMergeDates.cgamdGreaterThan And GADateCompare(vDupGADec.StartDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan) Then
                '(a) Both GAD's are live
                If (vOrigGADec.CancellationReason.Length = 0 And vDupGADec.CancellationReason.Length = 0) Then
                  vNewStartDate = vOrigGADec.StartDate
                  vNewEndDate = vOrigGADec.EndDate
                  If GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                    'Duplicate GAD ends after original GAD
                    vNewEndDate = vDupGADec.EndDate
                    vDateChange = True
                  End If
                  vOverlap = True
                  vCancelDGAD = True 'Cancel the duplicate GAD

                  '(b) Duplicate GAD Is cancelled
                ElseIf (vDupGADec.CancellationReason.Length > 0 And vOrigGADec.CancellationReason.Length = 0) Then
                  vNewStartDate = vOrigGADec.StartDate
                  vNewEndDate = vOrigGADec.EndDate
                  If GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                    'Duplicate GAD ends after original GAD
                    If vOrigGADec.HasTaxClaims Then
                      If GADateCompare(vOrigGADec.LatestClaimedPaymentDate, vDupGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                        'Last claim before duplicate GAD start
                        vNewEndDate = CDate(vDupGADec.StartDate).AddDays(-1).ToString(CAREDateFormat)
                        vDateChange = True
                      Else
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, vOrigGADec.StartDate, CDate(vDupGADec.StartDate).AddDays(-1).ToString(CAREDateFormat))
                        vNewGADs.Add(vNewGADec)
                        If GADateCompare(vOrigGADec.LatestClaimedPaymentDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                          'Last original GAD tax claim before end of original GAD
                          'So set original GAD end date to be last claim date
                          vNewEndDate = vOrigGADec.LatestClaimedPaymentDate
                          vDateChange = True
                        End If
                        vCancelOGAD = True 'Cancel the original GAD
                      End If
                    Else
                      'No tax claims
                      vNewEndDate = CDate(vDupGADec.StartDate).AddDays(-1).ToString(CAREDateFormat)
                      vDateChange = True
                    End If
                    vOverlap = True
                  ElseIf GADateCompare(vOrigGADec.EndDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                    'Original GAD ends after duplicate GAD
                    If vOrigGADec.HasTaxClaims Then
                      If GADateCompare(vOrigGADec.LatestClaimedPaymentDate, vDupGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                        'Last original GAD tax claim before start date of duplicate GAD
                        vNewEndDate = CDate(vDupGADec.StartDate).AddDays(-1).ToString(CAREDateFormat)
                        vDateChange = True
                        vOverlap = True
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat), vOrigGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                      ElseIf GADateCompare(vOrigGADec.EarliestClaimedPaymentDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                        'First original GAD tax claim after end date of duplicate GAD
                        vNewStartDate = CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                        vDateChange = True
                        'Add new GAD to cover the period before the start date of the duplicate GAD
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, vOrigGADec.StartDate, CDate(vDupGADec.StartDate).AddDays(-1).ToString(CAREDateFormat))
                        vNewGADs.Add(vNewGADec)
                      Else
                        'Last original GAD tax claim after start date of duplicate GAD
                        'Create a new GAD to cover the period before the cancelled GAD
                        vNewEndDate = vOrigGADec.LatestClaimedPaymentDate
                        vDateChange = True
                        vCancelOGAD = True
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, vOrigGADec.StartDate, CDate(vDupGADec.StartDate).AddDays(-1).ToString(CAREDateFormat))
                        vNewGADs.Add(vNewGADec)
                        'And, create a new GAD to cover the period after the cancelled GAD
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat), vOrigGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                      End If
                    Else
                      'No tax claims so update end date and create new GAD
                      vNewEndDate = CDate(vDupGADec.StartDate).AddDays(-1).ToString(CAREDateFormat)
                      vDateChange = True
                      vOverlap = True
                      vNewGADs = New Collection
                      vNewGADec = New GiftAidDeclaration
                      vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat), vOrigGADec.EndDate)
                      vNewGADs.Add(vNewGADec)
                    End If
                  End If

                  '(c) Original GAD cancelled
                ElseIf (vOrigGADec.CancellationReason.Length > 0 And vDupGADec.CancellationReason.Length = 0) Then
                  vNewStartDate = vDupGADec.StartDate
                  vNewEndDate = vDupGADec.EndDate
                  If GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                    'Duplicate GAD ends after original GAD
                    If vDupGADec.HasTaxClaims Then
                      If GADateCompare(vDupGADec.EarliestClaimedPaymentDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                        'First claim after end date of original GAD
                        vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOrigGADec.EndDate)))
                        vDateChange = True
                        vOverlap = True
                      Else
                        'Some claims before original GAD ends
                        'If GADateCompare(vDupGADec.LatestClaimedPaymentDate, vDupGADec.EndDate) = cgamdLessThan Then
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOrigGADec.EndDate))), vDupGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                        'End If
                        vNewEndDate = vDupGADec.LatestClaimedPaymentDate
                        vDateChange = True
                        vOverlap = True
                        vCancelDGAD = True
                      End If
                    Else
                      'No tax claims
                      vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOrigGADec.EndDate)))
                      vDateChange = True
                    End If
                  ElseIf GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                    'Duplicate GAD ends before original GAD
                    'The entire period of the live GAD is covered by the cancelled GAD
                    vCancelDGAD = True 'Cancel the duplicate GAD
                    vOverlap = True
                  End If
                End If
              End If

              '(3) Both the original and duplicate GAD's start on the same date
              If GADateCompare(vOrigGADec.StartDate, vDupGADec.StartDate) = ContactGiftAidMergeDates.cgamdEqual Then
                '(a) Both GAD's are live
                If (vOrigGADec.CancellationReason.Length = 0 And vDupGADec.CancellationReason.Length = 0) Then
                  vNewStartDate = vOrigGADec.StartDate
                  vNewEndDate = vOrigGADec.EndDate
                  If GADateCompare(vDupGADec.EndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                    vNewEndDate = vDupGADec.EndDate
                    vDateChange = True
                  End If
                  vCancelDGAD = True
                  vOverlap = True

                  '(b) Duplicate GAD is cancelled
                ElseIf (vDupGADec.CancellationReason.Length > 0 And vOrigGADec.CancellationReason.Length = 0) Then
                  vNewStartDate = vOrigGADec.StartDate
                  vNewEndDate = vOrigGADec.EndDate
                  If vOrigGADec.HasTaxClaims Then
                    If GADateCompare(vOrigGADec.EarliestClaimedPaymentDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                      'First claim after end of duplicate GAD
                      vNewStartDate = CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                      vDateChange = True
                    Else
                      'Last claim on/before end date of original GAD
                      If GADateCompare(vOrigGADec.EndDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                        'If original GAD ends after duplicate GAD then add new GAD
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vOrigGADec, CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat), vOrigGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                        vNewEndDate = If(GADateCompare(vOrigGADec.LatestClaimedPaymentDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan, vOrigGADec.LatestClaimedPaymentDate, vDupGADec.EndDate)
                        vDateChange = True
                      End If
                      vCancelOGAD = True
                      vOverlap = True
                    End If
                  Else
                    'No tax claims
                    If GADateCompare(vOrigGADec.EndDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                      'Original GAD ends after duplicate GAD
                      'So change start date of
                      vNewStartDate = CDate(vDupGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                      vDateChange = True
                    Else
                      vCancelOGAD = True
                    End If
                    vOverlap = True
                  End If

                  '(c) Original GAD is cancelled
                ElseIf (vOrigGADec.CancellationReason.Length > 0 And vDupGADec.CancellationReason.Length = 0) Then
                  vNewStartDate = vDupGADec.StartDate
                  vNewEndDate = vDupGADec.EndDate
                  If vDupGADec.HasTaxClaims Then
                    If GADateCompare(vDupGADec.EarliestClaimedPaymentDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                      'First claim after end of original GAD
                      vNewStartDate = CDate(vOrigGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                      vDateChange = True
                      vOverlap = True
                    Else
                      If GADateCompare(vDupGADec.LatestClaimedPaymentDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                        'Last claim before end date of duplicate GAD
                        vNewGADs = New Collection
                        vNewGADec = New GiftAidDeclaration
                        vNewGADec.InitNewFromMergedDeclaration(mvEnv, vDupGADec, CDate(vDupGADec.LatestClaimedPaymentDate).AddDays(1).ToString(CAREDateFormat), vDupGADec.EndDate)
                        vNewGADs.Add(vNewGADec)
                      End If
                      vNewEndDate = vDupGADec.LatestClaimedPaymentDate
                      vDateChange = True
                      vCancelDGAD = True
                    End If
                  Else
                    'No tax claims
                    vNewStartDate = CDate(vOrigGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                    vDateChange = True
                    vOverlap = True
                  End If
                End If
              End If

              '(4) Update the Declarations
              If vOverlap = True Or vCancelDGAD = True Or vCancelOGAD = True Then
                mvUpdateGAData = True 'After the Contact merge has finished, update unclaimed lines for all GAD's
              End If

              If vDateChange Or vCancelDGAD Or vCancelOGAD Then
                If mvEnv.Connection.InTransaction = False Then
                  mvEnv.Connection.StartTransaction()
                  vTrans = True
                End If

                '(a) Neither GAD is cancelled OR only duplicate GAD is cancelled
                If (vOrigGADec.CancellationReason.Length = 0 And vDupGADec.CancellationReason.Length = 0) Or (vDupGADec.CancellationReason.Length > 0 And vOrigGADec.CancellationReason.Length = 0) Then
                  If vOverlap Or vDateChange Then
                    If vDateChange Then
                      'Before we change the dates of vOrigGADec check they will not overlap another GAD
                      For Each vGADec In GiftAidDeclarations
                        If (vGADec.DeclarationNumber <> vOrigGADec.DeclarationNumber) And (vGADec.BatchNumber = 0 And vGADec.PaymentPlanNumber = 0) Then
                          If GADateCompare(vNewStartDate, vOrigGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan And GADateCompare(vOrigGADec.StartDate, vGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan And GADateCompare(vNewStartDate, vGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                            'The StartDate has been moved backwards so that it now overlaps the EndDate of another GAD
                            vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                          End If
                          If GADateCompare(vNewEndDate, vOrigGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan And GADateCompare(vOrigGADec.EndDate, vGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan And GADateCompare(vNewEndDate, vGADec.StartDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                            'The new EndDate has been moved fowards so that it now overlaps the StartDate of another GAD
                            vNewEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vGADec.StartDate)))
                          End If
                        End If
                      Next vGADec
                      vOrigDecType = vOrigGADec.DeclarationType
                      If vOverlap = True And vCancelOGAD = False Then
                        'Upgrade DeclarationType if required
                        GACheckTypeUpgrade(vOrigDecType, vDupGADec)
                        vOverlap = False
                      End If
                      vOrigGADec.Update(vNewStartDate, vNewEndDate, vOrigDecType, vOrigGADec.Notes)
                    End If

                    If Not (vNewGADs Is Nothing) Then
                      'If we had to create a new GAD then check these dates are OK
                      vNewGADNos = ""
                      For Each vNewGADec In vNewGADs
                        vNewStartDate = vNewGADec.StartDate
                        vNewEndDate = vNewGADec.EndDate
                        vGADChanged = False
                        For Each vGADec In GiftAidDeclarations
                          If (vGADec.BatchNumber = 0 And vGADec.PaymentPlanNumber = 0) And (vGADec.DeclarationNumber <> vOrigGADec.DeclarationNumber) Then
                            If (GADateCompare(vNewGADec.StartDate, vGADec.EndDate) <> ContactGiftAidMergeDates.cgamdGreaterThan And GADateCompare(vNewGADec.EndDate, vGADec.EndDate) <> ContactGiftAidMergeDates.cgamdLessThan) Then
                              'The StartDate overlaps the EndDate of another GAD
                              'Note: EndDate could be null
                              If vGADec.EndDate.Length > 0 Then
                                vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                              ElseIf vNewGADec.EndDate.Length > 0 Then
                                vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vNewGADec.EndDate)))
                              Else
                                vNewStartDate = CStr(CDate("01/01/9999")) 'Just set to something that is after a null GAD end date
                              End If
                              vGADChanged = True
                            End If
                            If GADateCompare(vNewGADec.StartDate, vGADec.StartDate) = ContactGiftAidMergeDates.cgamdEqual Then
                              'New GAD starts on same day as an existing GAD
                              If GADateCompare(vNewGADec.EndDate, vGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                                'Start new GAD after existing GAD
                                vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                              Else
                                'Total overlap so new GAD not required
                                'Note: EndDate could be null
                                If vGADec.EndDate.Length > 0 Then
                                  vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                                ElseIf vNewGADec.EndDate.Length > 0 Then
                                  vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vNewGADec.EndDate)))
                                Else
                                  vNewStartDate = CStr(CDate("01/01/9999")) 'Just set to something that is after a null GAD end date
                                End If
                              End If
                              vGADChanged = True
                            End If
                            If (GADateCompare(vNewGADec.StartDate, vGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan And GADateCompare(vNewGADec.EndDate, vGADec.StartDate) <> ContactGiftAidMergeDates.cgamdLessThan) Then
                              'The EndDate overlaps the StartDate of another GAD
                              vNewEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vGADec.StartDate)))
                              vGADChanged = True
                            End If
                          End If
                        Next vGADec

                        'As Start/End dates may have changed, now do an Update
                        If vGADChanged Then
                          If GADateCompare(vNewStartDate, vNewEndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                            'These dates are now invalid and so no new GAD is required
                          Else
                            vNewGADec.Update(vNewStartDate, vNewEndDate, vNewGADec.DeclarationType, vNewGADec.Notes)
                            vNewGADec.SaveChanges()
                          End If
                        Else
                          vNewGADec.SaveChanges()
                        End If
                        If vNewGADec.Existing Then
                          If vNewGADNos.Length > 0 Then vNewGADNos = vNewGADNos & ", "
                          vNewGADNos = vNewGADNos & CStr(vNewGADec.DeclarationNumber)
                        End If
                      Next vNewGADec

                      vNotes = vOrigGADec.Notes
                      If vNotes.Length > 0 Then vNotes = vNotes & vbCrLf & vbCrLf
                      vNotes = vNotes & String.Format(ProjectText.String16506, vNewGADNos) 'Superceded by Gift Aid Declaration(s) %s following Contact merge.
                      With vOrigGADec
                        .Update(.StartDate, .EndDate, .DeclarationType, vNotes)
                      End With
                    End If

                    If vCancelOGAD Then
                      vOrigGADec.Cancel(vMergeCancelRsn, "", "", False, 0, 0, vOrigGADec.EndDate, True)
                      vOrigGADec.RemoveUnclaimedLines()
                      If vOverlap Then
                        'May need to upgrade the DeclarationType
                        vOrigDecType = vDupGADec.DeclarationType
                        If GACheckTypeUpgrade(vOrigDecType, vOrigGADec) Then
                          With vDupGADec
                            .Update(.StartDate, .EndDate, vOrigDecType, .Notes)
                          End With
                        End If
                      End If
                      vOverlap = False
                    End If

                    If vOverlap Then
                      'May need to upgrade the DeclarationType
                      vOrigDecType = vOrigGADec.DeclarationType
                      If GACheckTypeUpgrade(vOrigDecType, vDupGADec) Then
                        With vOrigGADec
                          .Update(.StartDate, .EndDate, vOrigDecType, .Notes)
                        End With
                      End If
                    End If
                    vOrigGADec.SaveChanges()
                  End If

                  '(b) Only original GAD is cancelled
                ElseIf (vOrigGADec.CancellationReason.Length > 0 And vDupGADec.CancellationReason.Length = 0) Then
                  'Before we change the dates of vDupGADec check they will not overlap another GAD
                  If vOverlap Or vDateChange Then

                    If vDateChange Then
                      For Each vGADec In pDContact.GiftAidDeclarations
                        vGADChanged = False
                        vOrigDecType = vGADec.DeclarationType
                        If (vDupGADec.DeclarationNumber <> vGADec.DeclarationNumber) And (vGADec.BatchNumber = 0 And vGADec.PaymentPlanNumber = 0) Then
                          If GADateCompare(vNewStartDate, vDupGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan And GADateCompare(vNewStartDate, vGADec.EndDate) = ContactGiftAidMergeDates.cgamdLessThan Then
                            'Moving start date backwards to before end date of another GAD
                            vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                            If GACheckTypeUpgrade(vOrigDecType, vOrigGADec) Then vGADChanged = True
                          End If
                          If GADateCompare(vNewEndDate, vDupGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan And GADateCompare(vNewEndDate, vGADec.StartDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                            'Moving end date forwards to after start date of another GAD
                            vNewEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vGADec.StartDate)))
                            If GACheckTypeUpgrade(vOrigDecType, vOrigGADec) Then vGADChanged = True
                          End If
                        End If
                        If vGADChanged Then
                          With vGADec
                            .Update(.StartDate, .EndDate, vOrigDecType, .Notes)
                            .SaveChanges()
                          End With
                        End If
                      Next vGADec

                      vDupGADec.Update(vNewStartDate, vNewEndDate, vDupGADec.DeclarationType, vDupGADec.Notes)

                      If Not (vNewGADs Is Nothing) Then
                        vNewGADNos = ""
                        For Each vNewGADec In vNewGADs
                          vNewStartDate = vNewGADec.StartDate
                          vNewEndDate = vNewGADec.EndDate
                          vGADChanged = False
                          For Each vGADec In GiftAidDeclarations
                            If (vGADec.BatchNumber = 0 And vGADec.PaymentPlanNumber = 0) And (vGADec.DeclarationNumber <> vDupGADec.DeclarationNumber) Then
                              If (GADateCompare(vNewGADec.StartDate, vGADec.EndDate) <> ContactGiftAidMergeDates.cgamdGreaterThan And GADateCompare(vNewGADec.EndDate, vGADec.EndDate) <> ContactGiftAidMergeDates.cgamdLessThan) Then
                                'The StartDate overlaps the EndDate of another GAD
                                'Note: EndDate could be null
                                If vGADec.EndDate.Length > 0 Then
                                  vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                                ElseIf vNewGADec.EndDate.Length > 0 Then
                                  vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vNewGADec.EndDate)))
                                Else
                                  vNewStartDate = DateSerial(9998, 1, 1).ToString(CAREDateFormat) 'Just set to something that is after a null GAD end date
                                End If
                                vGADChanged = True
                              End If
                              If GADateCompare(vNewGADec.StartDate, vGADec.StartDate) = ContactGiftAidMergeDates.cgamdEqual Then
                                'New GAD starts on same day as an existing GAD
                                If GADateCompare(vNewGADec.EndDate, vGADec.EndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                                  'Start new GAD after existing GAD
                                  vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                                Else
                                  'Total overlap so new GAD not required
                                  'Note: EndDate could be null
                                  If vGADec.EndDate.Length > 0 Then
                                    vNewStartDate = CDate(vGADec.EndDate).AddDays(1).ToString(CAREDateFormat)
                                  ElseIf vNewGADec.EndDate.Length > 0 Then
                                    vNewStartDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vNewGADec.EndDate)))
                                  Else
                                    vNewStartDate = DateSerial(9998, 1, 1).ToString(CAREDateFormat) 'Just set to something that is after a null GAD end date
                                  End If
                                End If
                                vGADChanged = True
                              End If
                              If (GADateCompare(vNewGADec.StartDate, vGADec.StartDate) = ContactGiftAidMergeDates.cgamdLessThan And GADateCompare(vNewGADec.EndDate, vGADec.StartDate) <> ContactGiftAidMergeDates.cgamdLessThan) Then
                                'The EndDate overlaps the StartDate of another GAD
                                vNewEndDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vGADec.StartDate)))
                                vGADChanged = True
                              End If
                            End If
                          Next vGADec

                          If vGADChanged Then
                            If GADateCompare(vNewStartDate, vNewEndDate) = ContactGiftAidMergeDates.cgamdGreaterThan Then
                              'These dates are now invalid and so no new GAD is required
                            Else
                              vNewGADec.Update(vNewStartDate, vNewEndDate, vNewGADec.DeclarationType, vNewGADec.Notes)
                              vNewGADec.SaveChanges()
                            End If
                          Else
                            vNewGADec.SaveChanges()
                          End If
                          If vNewGADec.Existing Then
                            If vNewGADNos.Length > 0 Then vNewGADNos = vNewGADNos & ", "
                            vNewGADNos = vNewGADNos & CStr(vNewGADec.DeclarationNumber)
                          End If
                        Next vNewGADec

                        vNotes = vDupGADec.Notes
                        If vNotes.Length > 0 Then vNotes = vNotes & vbCrLf & vbCrLf
                        vNotes = vNotes & String.Format(ProjectText.String16506, vNewGADNos) 'Superceded by Gift Aid Declaration(s) %s following Contact merge.
                        vDupGADec.Update(vDupGADec.StartDate, vDupGADec.EndDate, vDupGADec.DeclarationType, vNotes)
                      End If
                    End If

                    If vOverlap Then
                      vOrigDecType = vOrigGADec.DeclarationType
                      If GACheckTypeUpgrade(vOrigDecType, vDupGADec) Then
                        With vOrigGADec
                          .Update(.StartDate, .EndDate, vOrigDecType, .Notes)
                          .SaveChanges()
                        End With
                      End If
                    End If
                  End If
                End If

                If vTrans Then
                  mvEnv.Connection.CommitTransaction()
                  vTrans = False
                End If
              End If
            End If
          Next vOrigGADec
        End If
        If vCancelDGAD Then
          'Cancel the duplicate GAD here after we have run through all the original GAD's
          'as the date range of the duplicate could cover more than one original
          vDupGADec.Cancel(vMergeCancelRsn, "", "", False, 0, 0, vDupGADec.EndDate, True)
          vDupGADec.RemoveUnclaimedLines()
        Else
          'Just save in case anything has changed
          vDupGADec.SaveChanges()
        End If
      Next vDupGADec
    End Sub

    Public Function GADateCompare(ByVal pDate1 As String, ByVal pDate2 As String) As ContactGiftAidMergeDates
      'Compare First Date with Second, return string to indicate whether Less, Equal or Greater Date
      'Because Gift Aid end dates may be null (open-ended) but CVDate needs a value.
      'e.g. if pDate1 < pDate2, return cgamdLessThan ("<")
      'Note: This is used by Contact Merge and Data Updates
      If pDate1.Length = 0 Then
        If pDate2.Length = 0 Then
          Return ContactGiftAidMergeDates.cgamdEqual
        Else
          Return ContactGiftAidMergeDates.cgamdGreaterThan
        End If
      Else
        If pDate2.Length = 0 Then
          Return ContactGiftAidMergeDates.cgamdLessThan
        Else
          If CDate(pDate1) > CDate(pDate2) Then
            Return ContactGiftAidMergeDates.cgamdGreaterThan
          ElseIf CDate(pDate1) < CDate(pDate2) Then
            Return ContactGiftAidMergeDates.cgamdLessThan
          Else
            Return ContactGiftAidMergeDates.cgamdEqual
          End If
        End If
      End If
    End Function

    Public Function GACheckTypeUpgrade(ByRef pOriginalDeclarationType As GiftAidDeclaration.GiftAidDeclarationTypes, ByVal pDuplicateDeclaration As GiftAidDeclaration) As Boolean
      'If Duplicate Gift Aid Declaration covers more than the original,
      'upgrade Type of Original
      Select Case pOriginalDeclarationType
        Case GiftAidDeclaration.GiftAidDeclarationTypes.gadtMember
          If pDuplicateDeclaration.DeclarationType = GiftAidDeclaration.GiftAidDeclarationTypes.gadtDonation Or pDuplicateDeclaration.DeclarationType = GiftAidDeclaration.GiftAidDeclarationTypes.gadtAll Then
            pOriginalDeclarationType = GiftAidDeclaration.GiftAidDeclarationTypes.gadtAll
            Return True
          End If
        Case GiftAidDeclaration.GiftAidDeclarationTypes.gadtDonation
          If pDuplicateDeclaration.DeclarationType = GiftAidDeclaration.GiftAidDeclarationTypes.gadtMember Or pDuplicateDeclaration.DeclarationType = GiftAidDeclaration.GiftAidDeclarationTypes.gadtAll Then
            pOriginalDeclarationType = GiftAidDeclaration.GiftAidDeclarationTypes.gadtAll
            Return True
          End If
      End Select
    End Function

    Private Sub ChangeContact(ByVal pJob As JobSchedule, ByVal pConn As CDBConnection, ByVal pTable As String, ByVal pConTo As Integer, ByVal pConFrom As Integer, ByVal pAddTo As Integer, ByVal pAddFrom As Integer, ByVal pSetAmend As Boolean, ByVal pConAttr As String, ByVal pAddAttr As String, ByVal pUnique As String)
      'Change_contact takes a table name and two contact numbers. all records
      'linked to the second contact number are moved to the first contact number
      Dim vArray() As String
      Dim vValue As String
      Dim vDone As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vRecordSet2 As CDBRecordSet
      Dim vUpdateFields As New CDBFields
      Dim vUpdate3Fields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vWhere2Fields As New CDBFields
      Dim vWhere3Fields As New CDBFields

      pJob.InfoMessage = String.Format(ProjectText.String31251, StrConv(Replace(pTable, "_", " "), VbStrConv.ProperCase)) 'Transferring: %s
      pConn.StartTransaction()
      If pTable = "invoices" Then
        vRecordSet = pConn.GetRecordSet("SELECT DISTINCT company FROM invoices WHERE contact_number = " & pConFrom)
        While vRecordSet.Fetch() And vDone = False
          vRecordSet2 = pConn.GetRecordSet("SELECT sales_ledger_account FROM credit_customers WHERE contact_number = " & pConTo & " AND company " & pConn.SQLLiteral("=", CDBField.FieldTypes.cftCharacter, (vRecordSet.Fields(1).Value)))
          If vRecordSet2.Fetch() Then
            vUpdateFields.Clear()
            vUpdateFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConTo)
            'vUpdateFields.Add "sales_ledger_account", cftCharacter, vRecordSet2.Fields(1).Value
            If Len(pAddAttr) > 0 And pAddTo > 0 Then vUpdateFields.Add(pAddAttr, CDBField.FieldTypes.cftLong, pAddTo)
            vWhereFields.Clear()
            vWhereFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConFrom)
            vWhereFields.Add("company", CDBField.FieldTypes.cftCharacter, vRecordSet.Fields(1).Value)
            pConn.UpdateRecords(pTable, vUpdateFields, vWhereFields)
          End If
          vRecordSet2.CloseRecordSet()
        End While
        vRecordSet.CloseRecordSet()
      Else
        If pTable = "financial_history" Then
          'If there are Financial History records we will need to update the unclaimed lines at the end
          vWhereFields.Clear()
          vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, pConFrom)
          If mvEnv.Connection.GetCount("financial_history", vWhereFields) > 0 Then mvUpdateGAData = True
          vWhereFields.Clear()
        End If
        If pUnique.Length > 0 Then
          If Len(pAddAttr) > 0 And pTable <> "contact_positions" Then
            If InStr(pUnique, pAddAttr) = 0 Then pUnique = pUnique & "," & pAddAttr
          End If
          If pTable = "contact_search_names" Then
            mvEnv.Connection.ExecuteSQL("DELETE FROM contact_search_names WHERE contact_number = " & pConFrom & " AND search_name IN (SELECT search_name FROM contact_search_names WHERE contact_number = " & pConTo & ")")
          End If
          vRecordSet = pConn.GetRecordSet("SELECT " & pUnique & " FROM " & pTable & " WHERE " & pConAttr & " = " & pConFrom)
          While vRecordSet.Fetch()
            vUpdateFields.Clear()
            vUpdateFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConTo)
            If Len(pAddAttr) > 0 And pTable <> "contact_positions" And pAddTo > 0 Then
              If vRecordSet.Fields(pAddAttr).LongValue = pAddFrom Then vUpdateFields.Add(pAddAttr, CDBField.FieldTypes.cftLong, pAddTo)
            End If
            If pTable = "contact_search_names" Then vUpdateFields.Add("is_active", CDBField.FieldTypes.cftCharacter, "N")
            vWhereFields.Clear()
            vArray = Split(pUnique, ",")
            For Each vAttr As String In vArray
              vValue = vRecordSet.Fields(vAttr).Value
              Select Case vAttr
                Case "contact_number", "contact_number_1", "contact_number_2", "address_number", "organisation_number", "selection_set", "revision", "booking_number", "contact_position_number"
                  vWhereFields.Add(vAttr, CDBField.FieldTypes.cftLong, vValue)
                Case "event_number"
                  vWhereFields.Add(vAttr, CDBField.FieldTypes.cftInteger, vValue)
                Case "created_on"
                  vWhereFields.Add(vAttr, CDBField.FieldTypes.cftTime, vValue)
                Case "contact_position_number"
                  'Do nothing
                Case Else
                  vWhereFields.Add(vAttr, CDBField.FieldTypes.cftCharacter, vValue)
              End Select
            Next
            If pConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False) = 0 Then
              If pConn.IsLastErrorDuplicate() Then
                'The update would have created a duplicate and there is a unique index so delete it
                pConn.DeleteRecords(pTable, vWhereFields, False)
              End If
            End If
          End While
          vRecordSet.CloseRecordSet()
        Else
          If Len(pAddAttr) > 0 And pAddTo > 0 Then
            vUpdateFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConTo)
            vUpdateFields.Add(pAddAttr, CDBField.FieldTypes.cftLong, pAddTo)
            vWhereFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConFrom)
            vWhereFields.Add(pAddAttr, CDBField.FieldTypes.cftLong, pAddFrom)
            pConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
          End If
          vUpdateFields.Clear()
          vUpdateFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConTo)
          vWhereFields.Clear()
          vWhereFields.Add(pConAttr, CDBField.FieldTypes.cftLong, pConFrom)
          If pTable = "sticky_notes" Then
            vWhereFields.Add("record_type", If(ContactType = ContactTypes.ctcOrganisation, "O", "C"))
          End If
          If pTable <> "communications" And (Len(pAddAttr) > 0 And pAddTo > 0) Then vWhereFields.Add(pAddAttr, pAddFrom, CDBField.FieldWhereOperators.fwoNotEqual)
          If pTable = "contact_accounts" Then CheckContactAccountsMerge(pConn, pConTo, vUpdateFields)
          pConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
        End If
      End If
      pConn.CommitTransaction()
    End Sub

    Private Sub ChangeOrg(ByRef pJob As JobSchedule, ByVal pConn As CDBConnection, ByRef pTable As String, ByRef pOrgFrom As Integer, ByRef pOrgAttr As String, ByVal pUnique As String)
      Dim vArray() As String
      Dim vValue As String
      Dim vRecordSet As CDBRecordSet
      Dim vUpdateFields As New CDBFields
      Dim vUpdate3Fields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vWhere2Fields As New CDBFields
      Dim vWhere3Fields As New CDBFields

      pJob.InfoMessage = String.Format(ProjectText.String31251, ProperName(pTable)) 'Transferring: %s
      pConn.StartTransaction()
      If pUnique.Length > 0 Then
        vRecordSet = pConn.GetRecordSet("SELECT " & pUnique & " FROM " & pTable & " WHERE " & pOrgAttr & " = " & pOrgFrom)
        While vRecordSet.Fetch()
          vUpdateFields.Clear()
          vUpdateFields.Add(pOrgAttr, CDBField.FieldTypes.cftLong, ContactNumber)
          vArray = Split(pUnique, ",")
          vWhereFields.Clear()
          For Each vAttr As String In vArray
            vValue = vRecordSet.Fields(vAttr).Value
            Select Case vAttr
              Case "contact_number", "address_number", "organisation_number", "contact_position_number"
                vWhereFields.Add(vAttr, CDBField.FieldTypes.cftLong, vValue)
              Case Else
                vWhereFields.Add(vAttr, CDBField.FieldTypes.cftCharacter, vValue)
            End Select
          Next
          If pConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False) = 0 Then
            If pConn.IsLastErrorDuplicate() Then
              'The update would have created a duplicate and there is a unique index so delete it
              pConn.DeleteRecords(pTable, vWhereFields, False)
            End If
          End If
        End While
        vRecordSet.CloseRecordSet()
      Else
        vUpdateFields.Add(pOrgAttr, CDBField.FieldTypes.cftLong, ContactNumber)
        vWhereFields.Add(pOrgAttr, CDBField.FieldTypes.cftLong, pOrgFrom)
        pConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
      End If
      pConn.CommitTransaction()
    End Sub

    Private Sub ChangeOrgAddress(ByRef pJob As JobSchedule, ByVal pConn As CDBConnection, ByVal pTable As String, ByVal pAddTo As Integer, ByVal pAddFrom As Integer)
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields

      pJob.InfoMessage = String.Format(ProjectText.String31251, ProperName(pTable)) 'Transferring: %s
      ' pConn.StartTransaction()
      If pTable = "purchase_order_payments" Then
        vUpdateFields.Add("payee_address_number", CDBField.FieldTypes.cftLong, pAddTo)
        vWhereFields.Add("payee_address_number", CDBField.FieldTypes.cftLong, pAddFrom)
      ElseIf Not pTable = "service_bookings" Then
        vUpdateFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddTo)
        vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, pAddFrom)
      Else
        vUpdateFields.Add("booking_address_number", CDBField.FieldTypes.cftLong, pAddTo)
        vWhereFields.Add("booking_address_number", CDBField.FieldTypes.cftLong, pAddFrom)
      End If
      pConn.UpdateRecords(pTable, vUpdateFields, vWhereFields, False)
      'Delete any that are left behind (code removed)
      ' pConn.CommitTransaction()
    End Sub

    Private Sub CheckContactAccountsMerge(ByVal pConn As CDBConnection, ByVal pOrigContactNumber As Integer, ByRef pUpdateFields As CDBFields)
      'Check to see if original Contact has Contact Accounts with the DefaultAccount flag set
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDefaultBankAccount) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add("contact_number", pOrigContactNumber)
        vWhereFields.Add("default_account", "Y")
        'If original Contact has ContactAccount records with the DefaultAccount flag set, then update any records for Duplicate Contact to not have the flag set
        If pConn.GetCount("contact_accounts", vWhereFields) > 0 Then pUpdateFields.Add("default_account", "N")
      End If
    End Sub

    Private Sub DoSuppressionsMerge(ByVal pDContact As Contact)
      Dim vDSuppression As ContactSuppression
      If pDContact.ContactType = ContactTypes.ctcOrganisation Then
        vDSuppression = New OrganisationSuppression(mvEnv)
      Else
        vDSuppression = New ContactSuppression(mvEnv)
      End If
      Dim vWhereFields As New CDBFields
      Dim vTableName As String
      If ContactType = Contact.ContactTypes.ctcOrganisation Then
        vWhereFields.Add("organisation_number", pDContact.OrganisationNumber)
        vTableName = "organisation_suppressions"
      Else
        vWhereFields.Add("contact_number", pDContact.ContactNumber)
        vTableName = "contact_suppressions"
      End If
      Dim vSQL As New SQLStatement(mvEnv.Connection, vDSuppression.GetRecordSetFields, vTableName & " cs", vWhereFields)
      Dim vRS As CDBRecordSet = vSQL.GetRecordSet
      While vRS.Fetch
        vDSuppression.InitFromRecordSet(vRS)
        vDSuppression.SaveSuppression(ContactSuppression.SuppressionEntryStyles.sesNormal, ContactNumber,
                      vDSuppression.MailingSuppression, vDSuppression.ValidFrom, vDSuppression.ValidTo,
                      vDSuppression.AmendedOn, vDSuppression.AmendedBy, False,
                      vDSuppression.Source, vDSuppression.Notes, vDSuppression.ResponseChannel)
      End While
      mvEnv.Connection.DeleteRecords(vTableName, vWhereFields, False)
    End Sub

    Private Sub DoCategoriesMerge(ByVal pDContact As Contact, ByVal pOPositionNumber As Integer, ByVal pDPositionNumber As Integer, ByVal pCheckDates As Boolean)
      Dim vRS As CDBRecordSet
      Dim vKey As String = ""
      Dim vParams As CDBParameters = Nothing
      Dim vValidFrom As String = ""
      Dim vValidTo As String = ""
      Dim vCol As New Collection
      Dim vWhereFields As New CDBFields
      Dim vNotes As String = ""
      Dim vTable As String = ""
      Dim vAttr As String = ""
      Dim vAttrValue As Integer
      Dim vType As ContactCategory.ContactCategoryTypes
      Dim vBothAttrs As String = ""

      If pOPositionNumber > 0 Then
        'Expect both pOPositionNumber & pDPositionNumber to be set
        vType = ContactCategory.ContactCategoryTypes.cctPosition
        vTable = "contact_position_activities"
        vAttr = "contact_position_number"
      ElseIf pDContact.ContactType = ContactTypes.ctcOrganisation Then
        vType = ContactCategory.ContactCategoryTypes.cctOrganisation
        vTable = "organisation_categories"
        vAttr = "organisation_number"
      Else
        vType = ContactCategory.ContactCategoryTypes.cctContact
        vTable = "contact_categories"
        vAttr = "contact_number"
      End If

      If vType = ContactCategory.ContactCategoryTypes.cctPosition Then
        vBothAttrs = pOPositionNumber & "," & pDPositionNumber
        vAttrValue = pOPositionNumber
      Else
        vBothAttrs = ContactNumber & "," & pDContact.ContactNumber
        vAttrValue = ContactNumber
      End If

      Dim vCategoryAttrs As String = vAttr & ",activity,activity_value,quantity,source,valid_from,valid_to,amended_by,amended_on,notes,activity_date,response_channel"
      If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbResponseChannel) = False Then vCategoryAttrs = vCategoryAttrs.Replace("response_channel", "'' AS response_channel")
      vWhereFields.Add(vAttr, CDBField.FieldTypes.cftLong, If(vType = ContactCategory.ContactCategoryTypes.cctPosition, pDPositionNumber, pDContact.ContactNumber))
      If mvEnv.Connection.GetCount(vTable, vWhereFields) > 0 Then 'BR13384: Contact Categories only updated if duplicate contact has categories otherwise primary contact categories remain
        With vWhereFields
          .Clear()
          If pCheckDates Then
            'This is OrganisationCategories only
            .Add(vAttr, vAttrValue, CType(CDBField.FieldWhereOperators.fwoOpenBracketTwice + CDBField.FieldWhereOperators.fwoCloseBracket, CDBField.FieldWhereOperators))
            .Add(vAttr & "#2", pDContact.ContactNumber, CType(CDBField.FieldWhereOperators.fwoOR + CDBField.FieldWhereOperators.fwoOpenBracket, CDBField.FieldWhereOperators))
            .Add("valid_from", CDBField.FieldTypes.cftDate, TodaysDate, CType(CDBField.FieldWhereOperators.fwoGreaterThanEqual + CDBField.FieldWhereOperators.fwoCloseBracketTwice, CDBField.FieldWhereOperators))
          Else
            .Add(vAttr, vBothAttrs, CDBField.FieldWhereOperators.fwoIn)
          End If
        End With
        Dim vSQLStatement As New SQLStatement(mvEnv.Connection, vCategoryAttrs, vTable, vWhereFields, "activity, activity_value,source, valid_from,valid_to")
        vRS = vSQLStatement.GetRecordSet()
        With vRS
          While vRS.Fetch
            vValidFrom = .Fields("valid_from").Value
            vValidTo = .Fields("valid_to").Value
            If vKey = "" Or vKey <> .Fields("activity").Value & "|" & .Fields("activity_value").Value & "|" & .Fields("source").Value Then
              'new activity: setup the Contact category to the first of these
              If vKey <> "" Then
                If vNotes.Length > 0 Then vParams("Notes").Value = vParams("Notes").Value & vbCrLf & vbCrLf & "Merge Data:" & vbCrLf & vNotes
                vCol.Add(vParams)
              End If
              vParams = New CDBParameters
              ActivityParamsFromRS(vParams, vRS, vAttr)
              If Not vParams.Exists("ContactNumber") Then vParams.Add("ContactNumber", ContactNumber)
              vKey = .Fields("activity").Value & "|" & .Fields("activity_value").Value & "|" & .Fields("source").Value
              vNotes = ""
            Else
              'same activity: check the dates to see which one we will be using
              If CDate(vValidFrom) >= CDate(vParams("ValidFrom").Value) Then
                If CDate(vValidFrom) <= CDate(vParams("ValidTo").Value) Then
                  'the second valid from is between the first valid from and to hence we will keep the first valid from
                  If CDate(vValidTo) > CDate(vParams("ValidTo").Value) Then
                    vParams("ValidTo").Value = vValidTo
                  End If
                  If .Fields(vAttr).LongValue = vAttrValue Then
                    'this is a primary contact activity. Update the params to this one and store the duplicate's details onto the notes
                    If vType = ContactCategory.ContactCategoryTypes.cctPosition Then
                      vNotes = vNotes & "Position Number: " & vParams("ContactPositionNumber").Value & vbCrLf
                    Else
                      vNotes = vNotes & "Contact Number: " & vParams(ProperName(vAttr)).Value & vbCrLf
                    End If
                    vNotes = vNotes & "Quantity: " & vParams("Quantity").Value & vbCrLf
                    vNotes = vNotes & "Activity Date: " & vParams("ActivityDate").Value & vbCrLf
                    vNotes = vNotes & "Notes : " & vParams("Notes").Value & vbCrLf
                    vParams("Quantity").Value = .Fields("quantity").Value
                    vParams("ActivityDate").Value = .Fields("activity_date").Value
                    vParams("Notes").Value = .Fields("notes").Value
                  Else
                    'duplicate contact, store the values into the notes field
                    If vType = ContactCategory.ContactCategoryTypes.cctPosition Then
                      vNotes = vNotes & "Position Number: " & .Fields(vAttr).Value & vbCrLf
                    Else
                      vNotes = vNotes & "Contact Number: " & .Fields(vAttr).Value & vbCrLf
                    End If
                    vNotes = vNotes & "Quantity: " & .Fields("quantity").Value & vbCrLf
                    vNotes = vNotes & "Activity Date: " & .Fields("activity_date").Value & vbCrLf
                    vNotes = vNotes & "Notes : " & .Fields("notes").Value & vbCrLf
                  End If
                Else
                  'the start date of the 2nd record is more than the end date of the first one, hence no overlap so we need to create a new activity
                  If Len(vNotes) > 0 Then vParams("Notes").Value = vParams("Notes").Value & vbCrLf & vbCrLf & "Merge Data:" & vbCrLf & vNotes
                  vCol.Add(vParams)
                  vParams = New CDBParameters
                  ActivityParamsFromRS(vParams, vRS, vAttr)
                  vNotes = ""
                End If
              End If
            End If
          End While
          If Not vParams Is Nothing Then
            If vNotes.Length > 0 Then vParams("Notes").Value = vParams("Notes").Value & vbCrLf & vbCrLf & "Merge Data:" & vbCrLf & vNotes
            vCol.Add(vParams)
          End If

          If vCol.Count() > 0 Then
            'we have a few ativities to insert, delete all the existing ones and then insert these
            Dim vKeepCol As New Collection
            If pCheckDates Then
              'First select all other Categories for pDContact with ValidTo < Today and add them to a collection
              With vWhereFields
                .Clear()
                .Add(vAttr, pDContact.ContactNumber)
                .Add("valid_to", CDBField.FieldTypes.cftDate, TodaysDate, CDBField.FieldWhereOperators.fwoLessThan)
              End With
              vSQLStatement = New SQLStatement(mvEnv.Connection, vCategoryAttrs, vTable, vWhereFields)
              Dim vCatRS As CDBRecordSet = vSQLStatement.GetRecordSet()
              While vCatRS.Fetch
                vParams = New CDBParameters()
                ActivityParamsFromRS(vParams, vCatRS, vAttr)
                vParams("AmendedBy").Value = vCatRS.Fields("amended_by").Value
                vParams("AmendedOn").Value = vCatRS.Fields("amended_on").Value
                vKeepCol.Add(vParams)
              End While
              vCatRS.CloseRecordSet()
            End If
            vWhereFields = New CDBFields
            vWhereFields.Add(vAttr, CDBField.FieldTypes.cftLong, vBothAttrs, CDBField.FieldWhereOperators.fwoIn)
            mvEnv.Connection.StartTransaction()
            mvEnv.Connection.DeleteRecords(vTable, vWhereFields, False)
            For Each vParams In vCol
              vParams(ProperName(vAttr)).Value = vAttrValue.ToString
              vParams.Add("MergeActivity", CDBField.FieldTypes.cftCharacter, "Y") 'BR12443: Pass this parameter to stop validation
              If vType = ContactCategory.ContactCategoryTypes.cctOrganisation Then
                Dim vOC As New OrganisationCategory(mvEnv)
                vOC.Init()
                vOC.Create(vParams)
                vOC.Save(vParams("AmendedBy").Value)
              Else
                Dim vCC As ContactCategory
                If pOPositionNumber > 0 Then
                  vCC = New PositionCategory(mvEnv) 'BR17378, If a Postion Category merge use Position Category class
                Else
                  vCC = New ContactCategory(mvEnv)
                End If
                vCC.Init()
                vCC.Create(vParams)
                vCC.Save(vParams("AmendedBy").Value)
              End If
            Next vParams
            For Each vParams In vKeepCol
              vParams.Add("MergeActivity", "Y")   'Pass this parameter to stop validation
              If vType = ContactCategory.ContactCategoryTypes.cctOrganisation Then
                Dim vOC As New OrganisationCategory(mvEnv)
                vOC.Init()
                vOC.Create(vParams)
                vOC.Save(vParams("AmendedBy").Value)
              Else
                Dim vCC As ContactCategory
                If pOPositionNumber > 0 Then
                  vCC = New PositionCategory(mvEnv) 'BR17378, If a Position Category merge use Position Category clas
                Else
                  vCC = New ContactCategory(mvEnv)
                End If
                vCC.Init()
                vCC.Create(vParams)
                vCC.Save(vParams("AmendedBy").Value)
              End If
            Next
            mvEnv.Connection.CommitTransaction()
          End If
          .CloseRecordSet()
        End With
      End If
    End Sub

    Private Sub ActivityParamsFromRS(ByRef pParams As CDBParameters, ByRef pRS As CDBRecordSet, ByRef pAttr As String)
      With pParams
        .Add(ProperName(pAttr), CDBField.FieldTypes.cftLong, pRS.Fields(pAttr).Value)
        .Add("Activity", CDBField.FieldTypes.cftCharacter, pRS.Fields("activity").Value)
        .Add("ActivityValue", CDBField.FieldTypes.cftCharacter, pRS.Fields("activity_value").Value)
        .Add("Quantity", CDBField.FieldTypes.cftLong, pRS.Fields("quantity").Value)
        .Add("Source", CDBField.FieldTypes.cftCharacter, pRS.Fields("source").Value)
        .Add("ValidFrom", CDBField.FieldTypes.cftCharacter, pRS.Fields("valid_from").Value)
        .Add("ValidTo", CDBField.FieldTypes.cftCharacter, pRS.Fields("valid_to").Value)
        .Add("AmendedBy", CDBField.FieldTypes.cftCharacter, "ContMerge")
        .Add("AmendedOn", CDBField.FieldTypes.cftDate, TodaysDate)
        .Add("Notes", CDBField.FieldTypes.cftCharacter, pRS.Fields("notes").Value)
        .Add("ActivityDate", CDBField.FieldTypes.cftDate, pRS.Fields("activity_date").Value)
        .Add("ResponseChannel", CDBField.FieldTypes.cftCharacter, pRS.Fields("response_channel").Value)
      End With
    End Sub

    Private Sub DoContactPositionMerge(ByVal pConn As CDBConnection, ByVal pDContact As Contact, Optional ByVal pContactMerge As Boolean = True, Optional ByVal pCheckPositionAddress As Boolean = True)
      Dim vWhereFields As New CDBFields
      Dim vPosition As ContactPosition
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vOverlap As Boolean 'Positions overlap
      Dim vUpdateDupPosition As Boolean 'Update duplicate Position

      Dim vAllOrigPositions As New List(Of ContactPosition)    'All Positions against the original
      Dim vAllDupPositions As New List(Of ContactPosition)     'All Positions against the duplicate
      'A Contact can only have 1 current Position at an Organisation at any time (irrespective of site)
      'From Organisation Merge
      ' - If the Contacts are the same then check for overlapping Position dates
      'From Contact Merge
      ' - If the Organisations are the same then check for overlapping Position dates
      'Remember, there could be different sites

      'Select ContactPositions for original
      Dim vOrigPosition As New ContactPosition(mvEnv) 'Original Position being processed
      vOrigPosition.Init()
      vSQL = "SELECT " & vOrigPosition.GetRecordSetFields() & " FROM contact_positions cp WHERE "
      If pContactMerge Then
        vSQL = vSQL & "contact_number = %1 ORDER BY organisation_number,"
      Else
        vSQL = vSQL & "organisation_number = %1 AND contact_number <> %1 ORDER BY contact_number,"
      End If
      'Order with 'started' nulls first and 'finished' nulls last
      vSQL = vSQL & " started" & pConn.DBSortByNullsFirst & ", finished"
      If pConn.NullsSortAtEnd = False Then vSQL = vSQL & " DESC"
      vRS = pConn.GetRecordSet(Replace(vSQL, "%1", ContactNumber.ToString))
      While vRS.Fetch
        vOrigPosition = New ContactPosition(mvEnv)
        vOrigPosition.InitFromRecordSet(vRS)
        vAllOrigPositions.Add(vOrigPosition)
      End While
      vRS.CloseRecordSet()

      'Select ContactPositions for duplicate
      Dim vDupPosition As ContactPosition 'Duplicate Position being processed
      If vAllOrigPositions.Count > 0 Then
        vRS = pConn.GetRecordSet(Replace(vSQL, "%1", pDContact.ContactNumber.ToString))
        While vRS.Fetch
          vDupPosition = New ContactPosition(mvEnv)
          vDupPosition.InitFromRecordSet(vRS)
          vAllDupPositions.Add(vDupPosition)
        End While
        vRS.CloseRecordSet()
      End If

      Dim vDPStarted As String 'Duplicate Position started
      Dim vDPFinished As String 'Duplicate Position finished
      Dim vOPStarted As String 'Original Position started
      Dim vOPFinished As String 'Original Position finished
      Dim vNewStarted As String 'New original started
      Dim vNewFinished As String 'New original finished
      Dim vNewDupStarted As String 'New duplicate started
      Dim vNewDupFinished As String 'New duplicate finished
      Dim vDeletePositions As New List(Of ContactPosition)  'Positions to be deleted
      Dim vSavePositions As New List(Of ContactPosition)    'For each Position to be deleted, this collection will hold the original ContactPositionNumber

      If vAllOrigPositions.Count > 0 And vAllDupPositions.Count > 0 Then
        For Each vDupPosition In vAllDupPositions
          For Each vOrigPosition In vAllOrigPositions
            If (pContactMerge = True And (vOrigPosition.OrganisationNumber = vDupPosition.OrganisationNumber)) Or (pContactMerge = False And (vOrigPosition.ContactNumber = vDupPosition.ContactNumber)) Then
              'Contact Merge - Both Positions are for the same Organisation
              'Organisation Merge - Both Positions are for the same Contact
              'A Contact can only have 1 current Position at an Organisation
              vOverlap = False
              vUpdateDupPosition = False
              vDPStarted = If(vDupPosition.Started.Length > 0, vDupPosition.Started, DateSerial(1900, 1, 1).ToString(CAREDateFormat))
              vDPFinished = If(vDupPosition.Finished.Length > 0, vDupPosition.Finished, DateSerial(9998, 12, 31).ToString(CAREDateFormat))
              vOPStarted = If(vOrigPosition.Started.Length > 0, vOrigPosition.Started, DateSerial(1900, 1, 1).ToString(CAREDateFormat))
              vOPFinished = If(vOrigPosition.Finished.Length > 0, vOrigPosition.Finished, DateSerial(9998, 12, 31).ToString(CAREDateFormat))
              vNewStarted = vOrigPosition.Started
              vNewFinished = vOrigPosition.Finished
              vNewDupStarted = vDupPosition.Started
              vNewDupFinished = vDupPosition.Finished

              '1) If we need to account for different sites then try and update duplicate to not overlap
              If pCheckPositionAddress Then
                If (pContactMerge = False) Or (pContactMerge = True And (vOrigPosition.AddressNumber <> vDupPosition.AddressNumber)) Then
                  If (CDate(vDPStarted) < CDate(vOPStarted)) And (CDate(vDPFinished) > CDate(vOPStarted)) Then
                    'Duplicate starts before original
                    If CDate(vDPFinished) >= DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOPFinished)) Then
                      'Duplicate finishes after original - change the start date
                      vUpdateDupPosition = True
                      vOverlap = True
                      vNewStarted = vDupPosition.Started
                      vNewDupStarted = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOPFinished)))
                    ElseIf CDate(vDPStarted) <= DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vOPStarted)) Then
                      'Change the finished date
                      vUpdateDupPosition = True
                      vNewDupFinished = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vOPStarted)))
                    End If
                  ElseIf (CDate(vDPFinished) > CDate(vOPFinished)) And (CDate(vDPStarted) < CDate(vOPFinished)) Then
                    'Duplicate ends after original
                    If DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOPFinished)) <= CDate(vDPFinished) Then
                      'Change start date
                      vUpdateDupPosition = True
                      vNewDupStarted = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vOPFinished)))
                    End If
                  End If
                  If vUpdateDupPosition Then
                    With vDupPosition
                      .Update(.Position, .Mail, .Current, vNewDupStarted, vNewDupFinished, .PositionLocation, 0, .PositionFunction, .PositionSeniority)
                      .Save()
                    End With
                    vDPStarted = If(vDupPosition.Started.Length > 0, vDupPosition.Started, DateSerial(1900, 1, 1).ToString(CAREDateFormat))
                    vDPFinished = If(vDupPosition.Finished.Length > 0, vDupPosition.Finished, DateSerial(9998, 12, 31).ToString(CAREDateFormat))
                  End If
                End If
              End If

              '2) Original & Duplicate overlap with Original starting before Duplicate
              If CDate(vOPStarted) < CDate(vDPStarted) Then
                'Original started before Duplicate started
                If CDate(vOPFinished) > CDate(vDPStarted) Then
                  'Original finished after Duplicate started
                  vOverlap = True
                  'Duplicate finished after original?
                  If CDate(vDPFinished) > CDate(vOPFinished) Then vNewFinished = vDupPosition.Finished
                End If
              End If

              '3) Original & Duplicate overlap with Duplicate starting before Original
              If CDate(vOPStarted) > CDate(vDPStarted) Then
                'Original started after Duplicate
                If CDate(vDPFinished) > CDate(vOPStarted) Then
                  'Original started before Duplicate finished
                  vOverlap = True
                  'Duplicate finished after original?
                  vNewStarted = vDupPosition.Started
                  If CDate(vDPFinished) > CDate(vOPFinished) Then vNewFinished = vDupPosition.Finished
                End If
              End If

              '4) Original & Duplicate overlap with both starting on same date
              If CDate(vOPStarted) = CDate(vDPStarted) Then
                'Both start on same date
                vOverlap = True
                If CDate(vDPFinished) > CDate(vOPFinished) Then vNewFinished = vDupPosition.Finished
              End If

              If vOverlap Then
                If vUpdateDupPosition = False Then
                  vDeletePositions.Add(vDupPosition) 'Add this to a collection of Positions to be deleted
                  vSavePositions.Add(vOrigPosition)
                End If

                'Check that changing Position dates will not cause an overlap
                For Each vPosition In vAllOrigPositions
                  If vPosition.ContactPositionNumber <> vOrigPosition.ContactPositionNumber Then
                    If (pContactMerge = True And (vOrigPosition.OrganisationNumber = vPosition.OrganisationNumber)) Or (pContactMerge = False And (vOrigPosition.ContactNumber = vPosition.ContactNumber)) Then
                      If IsDate(vPosition.Started) Or IsDate(vPosition.Finished) Then
                        If IsDate(vPosition.Started) And IsDate(vNewStarted) Then
                          If CDate(vNewStarted) < CDate(vPosition.Started) Then
                            If IsDate(vNewFinished) Then
                              If CDate(vNewFinished) > CDate(vPosition.Started) Then
                                vNewFinished = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vPosition.Started)))
                              End If
                            Else
                              vNewFinished = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vPosition.Started)))
                            End If
                          Else
                            If IsDate(vPosition.Finished) Then
                              If CDate(vPosition.Finished) >= CDate(vNewStarted) Then
                                vNewStarted = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vPosition.Finished)))
                              End If
                            End If
                          End If
                        ElseIf IsDate(vPosition.Started) Then
                          If IsDate(vNewFinished) Then
                            If CDate(vNewFinished) >= CDate(vPosition.Started) Then
                              vNewFinished = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vPosition.Started)))
                            End If
                          Else
                            vNewFinished = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, -1, CDate(vPosition.Started)))
                          End If
                        ElseIf IsDate(vNewStarted) Then
                          If IsDate(vPosition.Finished) Then
                            If CDate(vNewStarted) <= CDate(vPosition.Finished) Then
                              vNewStarted = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vPosition.Finished)))
                            End If
                          End If
                        Else
                          vNewStarted = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 1, CDate(vPosition.Finished)))
                        End If
                      Else
                        '
                      End If
                    End If
                  End If
                Next vPosition
                With vOrigPosition
                  If vUpdateDupPosition Then
                    .Update(.Position, .Mail, .Current, vNewStarted, vNewFinished, .PositionLocation, 0, .PositionFunction, .PositionSeniority)
                  Else
                    'Duplicate Position is to be deleted so if Position/Location/Function/Seniority not set then set them from the duplicate
                    .Update(If(.Position.Length > 0, .Position, vDupPosition.Position), .Mail, .Current, vNewStarted, vNewFinished, If(.PositionLocation.Length > 0, .PositionLocation, vDupPosition.PositionLocation), 0, If(.PositionFunction.Length > 0, .PositionFunction, vDupPosition.PositionFunction), If(.PositionSeniority.Length > 0, .PositionSeniority, vDupPosition.PositionSeniority))
                  End If
                End With
              End If
            End If
          Next vOrigPosition
        Next vDupPosition

        'Now update / delete everything as required
        If vDeletePositions.Count > 0 Then
          pConn.StartTransaction()
          vWhereFields.Add("contact_position_number", CDBField.FieldTypes.cftLong)
          For Each vPosition In vDeletePositions
            vWhereFields(1).Value = CStr(vPosition.ContactPositionNumber)
            pConn.DeleteRecords("contact_positions", vWhereFields, False)
          Next vPosition
          For Each vPosition In vAllOrigPositions
            vPosition.Save()
          Next vPosition
          pConn.CommitTransaction()

          'Once the Positions are merged, deal with the ContactPositionActivities
          For vIndex As Integer = 0 To vDeletePositions.Count - 1
            DoCategoriesMerge(pDContact, vSavePositions.Item(vIndex).ContactPositionNumber, vDeletePositions.Item(vIndex).ContactPositionNumber)
          Next
        End If
      End If
    End Sub
    ''' <summary>
    ''' De duplicate communications BR14606
    ''' </summary>
    ''' <param name="pConn">The database connection</param>
    ''' <param name="pPContactNumber">The primary contact</param>
    ''' <param name="pDContactNumber">The duplicate contact</param>
    ''' <remarks>Does not merge. Ensures that there is only one default device per device</remarks>
    Private Sub DoContactCommunicationsDeDup(pConn As CDBConnection, ByVal pPContactNumber As Integer, ByVal pDContactNumber As Integer)

      Dim vContactNumbers As List(Of Integer) = New List(Of Integer) ' List of primary and duplicate contact numbers
      Dim vCommunicationNumbersToDelete As List(Of Integer) = New List(Of Integer) ' communication numbers that need to be deleted, duplcates
      Dim vDuplicateDeviceDefault As List(Of Integer) = New List(Of Integer) ' communication numbers that need device default set to N
      ' Fields that are not compared when determining whether a communication is a duplicate
      Dim vExcludeFields As List(Of String) = New List(Of String)(New String() {"address_number", "contact_number", "amended_by", "amended_on", "communication_number"})

      Dim vSQLStatement As SQLStatement
      Dim vFields As String = "address_number,contact_number,device,ex_directory,dialling_code,std_code," & pConn.DBSpecialCol("", "number") & ",extension,notes,cli_number,amended_by,amended_on,communication_number,valid_from,valid_to,is_active,mail,device_default,preferred_method,archive,communication_usage"
      Dim vTable As String = "communications"
      Dim vWhereFieldsSelect As CDBFields
      Dim vWhereFieldsUpdate As CDBFields
      Dim vWhereFieldsDelete As CDBFields

      Dim vDatatable As DataTable ' Datatable containing communications for both primary and duplicate contacts
      Dim vPCommunications() As DataRow ' Primary Contacts communications
      Dim vDCommunications() As DataRow ' Duplicate contacts communications

      Dim vUpdateFields As CDBFields

      vContactNumbers.Add(pPContactNumber)
      vContactNumbers.Add(pDContactNumber)

      vWhereFieldsSelect = New CDBFields(New CDBField("contact_number", vContactNumbers))
      vSQLStatement = New SQLStatement(pConn, vFields, vTable, vWhereFieldsSelect)
      vDatatable = vSQLStatement.GetDataTable
      vPCommunications = vDatatable.Select("contact_number=" & pPContactNumber.ToString)
      vDCommunications = vDatatable.Select("contact_number=" & pDContactNumber.ToString)

      For Each vdrPCommunications As DataRow In vPCommunications
        For Each vdrDCommunications As DataRow In vDCommunications
          If vdrPCommunications.FilteredEquals(vdrDCommunications, vExcludeFields) Then
            vCommunicationNumbersToDelete.Add(CInt(vdrDCommunications("communication_number")))
          Else
            If DirectCast(vdrPCommunications("device"), String) = DirectCast(vdrDCommunications("device"), String) And DirectCast(vdrPCommunications("device_default"), String) = "Y" And DirectCast(vdrDCommunications("device_default"), String) = "Y" Then
              'Use CInt instead of DirectCast as the db field is a number in Oracle but an integer in MS SQL database and DirectCast will error when
              'assigning a number into an integer. 
              vDuplicateDeviceDefault.Add(CInt(vdrDCommunications("communication_number")))
            End If
          End If
        Next
      Next
      If vCommunicationNumbersToDelete.Count > 0 Then
        vWhereFieldsDelete = New CDBFields(New CDBField("communication_number", vCommunicationNumbersToDelete)) ' Create an IN Clause fromn the list
        pConn.DeleteRecords(vTable, vWhereFieldsDelete)
      End If
      If vDuplicateDeviceDefault.Count > 0 Then
        vUpdateFields = New CDBFields(New CDBField("device_default", "N")) '
        vWhereFieldsUpdate = New CDBFields(New CDBField("communication_number", vDuplicateDeviceDefault)) ' Create an IN Clause fromn the list
        pConn.UpdateRecords(vTable, vUpdateFields, vWhereFieldsUpdate, False)
      End If
    End Sub

    Private Sub GetContactMergeInfo(ByVal pConn As CDBConnection, ByRef pContact As Boolean)
      Dim vTable As String
      Dim vTables As String = ""
      Dim vAttr As String
      Dim vIgnore As Boolean
      Dim vRecordSet As CDBRecordSet
      Dim vMergeInfo As ContactMergeInfo

      If Not mvMergeInfoValid Then
        If pContact Then
          vAttr = "contact_number"
        Else
          vAttr = "organisation_number"
        End If
        'Create an array element for each instance of contact/organisation number
        mvMergeInfo = New List(Of ContactMergeInfo)
        vRecordSet = pConn.GetRecordSet("SELECT table_name,primary_key FROM maintenance_attributes WHERE attribute_name = '" & vAttr & "' AND is_base_table_attribute = 'Y' ORDER BY table_name")
        While vRecordSet.Fetch()
          vTable = vRecordSet.Fields(1).Value
          vIgnore = False
          If pContact Then
            If vTable = "contacts" Then vIgnore = True
          Else
            If vTable = "organisations" Then vIgnore = True
          End If
          If Left(vTable, 5) = "temp_" Then vIgnore = True
          If Left(vTable, 4) = "ext_" Then vIgnore = True
          Select Case vTable
            Case "contact_categories", "organisation_categories"
              vIgnore = True
            Case "contact_suppressions", "organisation_suppressions"
              vIgnore = True
            Case "exam_bookings", "exam_booking_units", "exam_student_header"
              vIgnore = True
          End Select
          If Not vIgnore Then
            vMergeInfo = New ContactMergeInfo
            vMergeInfo.TableName = vTable
            vMergeInfo.ContactAttr = vAttr
            vMergeInfo.UniqueContact = vRecordSet.Fields(2).Bool
            vMergeInfo.SetAmend = False 'We no longer update amended fields - If we needed to we would have to check maintenance_attribute to see if they were present
            If mvMergeInfo.Count > 0 Then
              vTables = vTables & ",'" & vTable & "'"
            Else
              vTables = "'" & vTable & "'"
            End If
            Select Case vTable 'Handle special cases
              Case "contact_addresses"
                vMergeInfo.UniqueAttrs = "contact_number,address_number"
              Case "contact_credit_cards"
                vMergeInfo.UniqueAttrs = "contact_number,credit_card_number"
              Case "contact_positions"
                vMergeInfo.UniqueAttrs = "contact_position_number"
                '        Case "contact_categories"
                '          mvMergeInfo.UniqueAttrs = "contact_number,activity,activity_value,source"
              Case "organisation_categories"
                vMergeInfo.UniqueAttrs = "organisation_number,activity,activity_value,source"
            End Select
            mvMergeInfo.Add(vMergeInfo)
          End If
        End While
        vRecordSet.CloseRecordSet()

        If pContact Then
          'For contacts find the tables that have address number as well
          vRecordSet = pConn.GetRecordSet("SELECT table_name FROM maintenance_attributes WHERE table_name IN (" & vTables & ") AND attribute_name = 'address_number' AND is_base_table_attribute = 'Y' ORDER BY table_name")
          While vRecordSet.Fetch()
            vTable = vRecordSet.Fields(1).Value
            For Each vMergeInfo In mvMergeInfo
              If vMergeInfo.TableName = vTable Then
                'BR14137: Don't set AddressAttr for organisations as this address is not related to the duplicate contact
                If vTable <> "organisations" Then vMergeInfo.AddressAttr = "address_number"
                Exit For
              End If
            Next
          End While
          vRecordSet.CloseRecordSet()
        End If

        'Build a list of all tables where the primary key includes the attribute we are searching for
        vTables = ""
        For Each vMergeInfo In mvMergeInfo
          If vMergeInfo.UniqueContact Then
            If vTables.Length > 0 Then
              vTables = vTables & ",'" & vMergeInfo.TableName & "'"
            Else
              vTables = "'" & vMergeInfo.TableName & "'"
              If vMergeInfo.AddressAttr.Length > 0 Then vTables = vTables & ",'" & vMergeInfo.AddressAttr & "'"
            End If
          End If
        Next

        'Get the primary key attributes for the tables
        vRecordSet = pConn.GetRecordSet("SELECT table_name,attribute_name FROM maintenance_attributes WHERE table_name IN (" & vTables & ") AND is_base_table_attribute = 'Y' AND primary_key = 'Y' ORDER BY table_name, sequence_number")
        While vRecordSet.Fetch()
          vTable = vRecordSet.Fields(1).Value
          For Each vMergeInfo In mvMergeInfo
            If vMergeInfo.TableName = vTable Then
              If vMergeInfo.UniqueAttrs.Length > 0 Then vMergeInfo.UniqueAttrs = vMergeInfo.UniqueAttrs & ","
              vMergeInfo.UniqueAttrs = vMergeInfo.UniqueAttrs & vRecordSet.Fields(2).Value
              Exit For
            End If
          Next
        End While
        vRecordSet.CloseRecordSet()

        'Add additional elements for other tables which include the numbers we are changing
        If pContact Then
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "contact_links"
          vMergeInfo.ContactAttr = "contact_number_2"
          vMergeInfo.UniqueAttrs = "contact_number_1,contact_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "contact_links"
          vMergeInfo.ContactAttr = "contact_number_1"
          vMergeInfo.UniqueAttrs = "contact_number_1,contact_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "organisation_links"
          vMergeInfo.ContactAttr = "organisation_number_2"
          vMergeInfo.UniqueAttrs = "organisation_number_1,organisation_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "organisation_links"
          vMergeInfo.ContactAttr = "organisation_number_1"
          vMergeInfo.UniqueAttrs = "organisation_number_1,organisation_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "proforma_invoices"
          vMergeInfo.ContactAttr = "order_contact_number"
          vMergeInfo.AddressAttr = "order_address_number"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "purchase_invoices"
          vMergeInfo.ContactAttr = "payee_contact_number"
          vMergeInfo.AddressAttr = "payee_address_number"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "purchase_orders"
          vMergeInfo.ContactAttr = "payee_contact_number"
          vMergeInfo.AddressAttr = "payee_address_number"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "sticky_notes"
          vMergeInfo.ContactAttr = "unique_id"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "department_notes"
          vMergeInfo.ContactAttr = "unique_id"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "sundry_costs"
          vMergeInfo.ContactAttr = "unique_id"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "purchase_order_payments"
          vMergeInfo.ContactAttr = "payee_contact_number"
          vMergeInfo.AddressAttr = "payee_address_number"
          mvMergeInfo.Add(vMergeInfo)
        Else
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "organisation_links"
          vMergeInfo.ContactAttr = "organisation_number_2"
          vMergeInfo.UniqueAttrs = "organisation_number_1,organisation_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "organisation_links"
          vMergeInfo.ContactAttr = "organisation_number_1"
          vMergeInfo.UniqueAttrs = "organisation_number_1,organisation_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "contact_links"
          vMergeInfo.ContactAttr = "contact_number_2"
          vMergeInfo.UniqueAttrs = "contact_number_1,contact_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "contact_links"
          vMergeInfo.ContactAttr = "contact_number_1"
          vMergeInfo.UniqueAttrs = "contact_number_1,contact_number_2,relationship"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "sticky_notes"
          vMergeInfo.ContactAttr = "unique_id"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "department_notes"
          vMergeInfo.ContactAttr = "unique_id"
          mvMergeInfo.Add(vMergeInfo)
          vMergeInfo = New ContactMergeInfo
          vMergeInfo.TableName = "sundry_costs"
          vMergeInfo.ContactAttr = "unique_id"
          mvMergeInfo.Add(vMergeInfo)
        End If
        'tables common to contacts and organisations
        vMergeInfo = New ContactMergeInfo
        vMergeInfo.TableName = "batch_transactions"
        vMergeInfo.ContactAttr = "mailing_contact_number"
        vMergeInfo.AddressAttr = "mailing_address_number"
        mvMergeInfo.Add(vMergeInfo)
        vMergeInfo = New ContactMergeInfo
        vMergeInfo.TableName = "caf_voucher_transactions"
        vMergeInfo.ContactAttr = "mailing_contact_number"
        vMergeInfo.AddressAttr = "mailing_address_number"
        mvMergeInfo.Add(vMergeInfo)
        vMergeInfo = New ContactMergeInfo
        vMergeInfo.TableName = "financial_links"
        vMergeInfo.ContactAttr = "donor_contact_number"
        mvMergeInfo.Add(vMergeInfo)
        vMergeInfo = New ContactMergeInfo
        vMergeInfo.TableName = "service_bookings"
        vMergeInfo.ContactAttr = "booking_contact_number"
        vMergeInfo.AddressAttr = "booking_address_number"
        mvMergeInfo.Add(vMergeInfo)
        'Add one more entry for BTA table to update the Deceased contact as 
        'contact merge always merge only updated contact and address number
        vMergeInfo = New ContactMergeInfo
        vMergeInfo.TableName = "batch_transaction_analysis"
        vMergeInfo.ContactAttr = "deceased_contact_number"
        mvMergeInfo.Add(vMergeInfo)
        mvMergeInfoValid = True
      End If
    End Sub

    Public Function SpacePadInitials(ByRef pString As String) As String
      Dim vInitials As String = ""
      Dim vWords() As String
      Dim vItems() As String
      Dim vItems2() As String
      Dim vInitial As String = ""
      Dim vAddStyle As Boolean

      If Not mvInitialsStyleRead Then
        mvInitialsStyle = mvEnv.GetConfig("initials_format")
        If mvInitialsStyle.Length > 2 Then
          mvInitialsStyle = Mid(mvInitialsStyle, 2, Len(mvInitialsStyle) - 2)
        ElseIf Len(mvInitialsStyle) = 2 Then
          mvInitialsStyle = ""
        Else
          mvInitialsStyle = " "
        End If
        mvInitialsStyleRead = True
      End If

      If pString.Length > 0 Then
        vWords = Split(pString, " ")
        vAddStyle = False
        For Each vIndexWord As String In vWords
          vItems = Split(vIndexWord, "-")
          For Each vIndexWord2 As String In vItems
            vItems2 = Split(vIndexWord2, ".")
            For Each vIndexWord3 As String In vItems2
              If vIndexWord3.ToLower.Length > 0 Then
                Select Case vIndexWord3.ToLower
                  Case "+", "et", "und", "and"
                    vInitial = "+"
                    If vAddStyle Then
                      vInitials = vInitials & mvInitialsStyle
                      If InStr(mvInitialsStyle, " ") > 0 Then vInitial = "+ "
                      If mvEnv.GetConfigOption("cd_joint_contact_support", True) Then vAddStyle = False
                    End If
                  Case "&"
                    vInitial = "&"
                    If vAddStyle Then
                      vInitials = vInitials & mvInitialsStyle
                      If InStr(mvInitialsStyle, " ") > 0 Then vInitial = "& "
                      If mvEnv.GetConfigOption("cd_joint_contact_support", True) Then vAddStyle = False
                    End If
                  Case Else
                    vInitial = vIndexWord3.Substring(0, 1).ToUpper
                    If vAddStyle Then
                      vInitials = vInitials & mvInitialsStyle
                    Else
                      vAddStyle = True
                    End If
                End Select
                vInitials = vInitials & vInitial
              End If
            Next
          Next
        Next
      End If
      If vInitials.Length > 0 Then
        If Not vInitials.EndsWith(mvInitialsStyle) Then
          Return (vInitials & mvInitialsStyle).Trim
        Else
          Return vInitials.Trim
        End If
      End If
      Return ""
    End Function

    Public Function CheckDeleteRights(ByVal pCheckContactPersonalDetails As Boolean) As String
      Dim vRecordSet As CDBRecordSet
      Dim vConn As CDBConnection
      Dim vFiveYearsAgo As String = ""
      Dim vSixYearsAgo As String = ""
      Dim vError As Boolean
      Dim vMsg As String = ""
      Dim vSQL As String
      Dim vDBName As String
      Dim vExtRecordSet As CDBRecordSet

      vConn = mvEnv.Connection
      If pCheckContactPersonalDetails Then
        'Check referential integrity on custom forms
        If mvEnv.GetConfigOption("option_custom_data", False) Then
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT referential_sql, db_name FROM custom_forms WHERE client = '" & mvEnv.ClientCode & "' AND custom_form >= " & mvEnv.FirstCustomFormNumber & " AND custom_form <= " & mvEnv.LastCustomFormNumber & " AND referential_sql IS NOT NULL")
          While vRecordSet.Fetch() = True
            vSQL = vRecordSet.Fields(1).Value
            vSQL = Replace(vSQL, "?", CStr(ContactNumber))
            vSQL = Replace(vSQL, "#", CStr(ContactNumber))
            vDBName = vRecordSet.Fields(2).Value
            vExtRecordSet = mvEnv.GetConnection(vDBName).GetRecordSet(vSQL)
            If vExtRecordSet.Fetch() = True Then
              vMsg = ProjectText.String16554 'The Contact exists in an external database and may not be deleted
              vError = True
            End If
            vExtRecordSet.CloseRecordSet()
          End While
          vRecordSet.CloseRecordSet()
        End If
        'Check for rights to delete
        If Not vError Then
          If mvEnv.OwnershipMethod = CDBEnvironment.OwnershipMethods.omOwnershipDepartments And Department <> mvEnv.User.Department Then
            Return ProjectText.String16503 'You must be in the Owner Department to delete this Contact
          ElseIf mvEnv.User.HasItemAccessRights(CDBUser.AccessControlItems.aciContactDelete) = False Then
            Return ProjectText.String31258 'You do not have access to delete this Contact
          Else
            'Check if the default contact for an organisation
            If mvEnv.Connection.GetCount("organisations", Nothing, "contact_number = " & ContactNumber) > 0 Then
              Return ProjectText.String16517 'The Contact is a default Mailing Contact\r\n\r\nContact cannot be deleted
            End If
          End If
        End If
      End If
      If Not vError Then
        vFiveYearsAgo = Today.AddYears(-5).ToString(CAREDateFormat)
        vSixYearsAgo = Today.AddYears(-6).ToString(CAREDateFormat)
        '--------------------------------------------------------------------------------------------------
        'Check for rights to delete from the communications log
        vRecordSet = vConn.GetRecordSet("SELECT public_delete, department_delete, department, creator_delete, created_by, communications_log_number FROM communications_log cl, document_classes dc WHERE contact_number = " & ContactNumber & " AND dc.document_class = cl.document_class")
        With vRecordSet
          While .Fetch() = True AndAlso vError = False
            If ((.Fields(1).Bool) Or (.Fields(2).Bool And .Fields(3).Value = mvEnv.User.Department) Or (.Fields(4).Bool And .Fields(5).Value = mvEnv.User.Logname)) Then
              'OK to delete
            Else
              vMsg = String.Format(ProjectText.String16528, .Fields(6).Value) 'No delete privilege for Communications Log number %s\r\n\r\nContact cannot be deleted
              vError = True
            End If
          End While
          .CloseRecordSet()
        End With
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for current membership etc..
      If Not vError Then
        vRecordSet = vConn.GetRecordSet("SELECT membership_number, m.cancellation_reason, o.contact_number, m.cancelled_on FROM members m, orders o WHERE m.contact_number = " & ContactNumber & " AND o.order_number = m.order_number")
        With vRecordSet
          While .Fetch() = True And (vError = False)
            If Len(.Fields(2).Value) = 0 Then
              vMsg = String.Format(ProjectText.String16530, .Fields(1).Value) 'Membership %s is still current\r\n\r\nContact cannot be deleted
              vError = True
            Else
              If .Fields(3).IntegerValue <> ContactNumber And CDate(.Fields(4).Value) > CDate(vFiveYearsAgo) Then
                vMsg = String.Format(ProjectText.String16531, .Fields(1).Value) 'Membership %s has a separate payer and expired less than 5 years ago\r\n\r\nContact cannot be deleted
                vError = True
              End If
            End If
          End While
          .CloseRecordSet()
        End With
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for payer of memberships
      If Not vError Then
        If vConn.GetCount("orders o, members m", Nothing, "o.contact_number = " & ContactNumber & " AND o.order_type = 'M' AND m.order_number = o.order_number AND (m.cancellation_reason IS NULL OR m.cancelled_on" & vConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vFiveYearsAgo) & ")") > 0 Then
          vMsg = (ProjectText.String16533) 'The Contact is the Payer of Memberships which are either current or expired less than 5 years ago\r\n\r\nContact cannot be deleted
          vError = True
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for payer of covenants
      If Not vError Then
        If vConn.GetCount("orders o, covenants c", Nothing, "o.contact_number = " & ContactNumber & " AND c.order_number = o.order_number AND (c.cancellation_reason IS NULL OR c.cancelled_on" & vConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vSixYearsAgo) & ")") > 0 Then
          vMsg = (ProjectText.String16534) 'The Contact is the Payer of an Order attached to Covenants which are either current or expired less than 6 years ago\r\n\r\nContact cannot be deleted
          vError = True
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for covenants
      If Not vError Then
        If vConn.GetCount("covenants", Nothing, "contact_number = " & ContactNumber & " AND (cancellation_reason IS NULL or cancelled_on" & vConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vSixYearsAgo) & ")") > 0 Then
          vMsg = (ProjectText.String16535) 'The Contact has Covenants which are either current or expired less than 6 years ago\r\n\r\nContact cannot be deleted
          vError = True
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'If gift aid donation has been claimed, should have been ocer 6 years ago
      'otherwise financial history  record should have been posted over six years ago
      If Not vError Then
        vRecordSet = vConn.GetRecordSet("SELECT claim_number, batch_number, transaction_number FROM gift_aid_donations WHERE contact_number = " & ContactNumber)
        While vRecordSet.Fetch() = True And (vError = False)
          If vRecordSet.Fields(1).IntegerValue > 0 Then
            If vConn.GetCount("tax_claims", Nothing, "claim_number = " & vRecordSet.Fields(1).IntegerValue & " AND claim_generated" & vConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vSixYearsAgo)) > 0 Then
              vMsg = (ProjectText.String16536) 'The Contact has Gift Aid Donations which were claimed less than 6 years ago\r\n\r\nContact cannot be deleted
              vError = True
            End If
          Else
            If vConn.GetCount("financial_history", Nothing, "batch_number = " & vRecordSet.Fields(2).IntegerValue & " AND transaction_number = " & vRecordSet.Fields(3).IntegerValue & " AND posted" & vConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vSixYearsAgo)) > 0 Then
              vMsg = (ProjectText.String16537) 'The Contact has Gift Aid Donations which were posted less than 6 years ago\r\n\r\nContact cannot be deleted
              vError = True
            End If
          End If
        End While
        vRecordSet.CloseRecordSet()
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for caf vouchers
      If Not vError Then
        If vConn.GetCount("caf_voucher_transactions", Nothing, "contact_number = " & ContactNumber) > 0 Then
          Return (ProjectText.String16539) 'The Contact has CAF Voucher Transactions awaiting posting\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for cash receipts
      If Not vError Then
        If vConn.GetCount("batch_transactions bt, batches b", Nothing, "contact_number = " & ContactNumber & " AND b.batch_number = bt.batch_number AND b.posted_to_nominal = 'N'") > 0 Then
          Return (ProjectText.String16540) 'The Contact has Cash Receipts awaiting posting\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for orders
      If Not vError Then
        If vConn.GetCount("orders", Nothing, "contact_number = " & ContactNumber & " AND cancellation_reason IS NULL AND renewal_date" & vConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vFiveYearsAgo)) > 0 Then
          Return (ProjectText.String16541) 'The Contact has orders which are either not cancelled or have a renewal date of less than 5 years ago\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for financial history
      If Not vError Then
        If vConn.GetCount("financial_history", Nothing, "contact_number = " & ContactNumber & " AND transaction_date" & vConn.SQLLiteral(">", CDBField.FieldTypes.cftDate, vFiveYearsAgo)) > 0 Then
          Return (ProjectText.String16542) 'The Contact has financial history records which are less than 5 years old\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for purchase orders
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", ContactNumber)
      vWhereFields.Add("cancellation_reason")
      If Not vError Then
        If vConn.GetCount("purchase_orders", vWhereFields) > 0 Then
          Return ProjectText.String16556  'The Contact has purchase orders which are not cancelled\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for payee contact on purchase order
      If Not vError Then
        vWhereFields(1).Name = "payee_contact_number"
        If vConn.GetCount("purchase_orders", vWhereFields) > 0 Then
          Return ProjectText.String16557  'The Contact is a payee for a purchase orders which are not cancelled\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for purchase invoices
      If Not vError Then
        vWhereFields(1).Name = "contact_number"
        vWhereFields.Remove(2)
        If vConn.GetCount("purchase_invoices", vWhereFields) > 0 Then
          Return ProjectText.String16558  'The Contact has purchase invoices\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for payee contact on purchase invoice
      If Not vError Then
        vWhereFields(1).Name = "payee_contact_number"
        If vConn.GetCount("purchase_invoices", vWhereFields) > 0 Then
          Return ProjectText.String16559  'The Contact is a payee for a purchase invoice\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for payee contact on purchase order payments
      If Not vError Then
        vWhereFields.Add("cheque_produced_on", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoOpenBracket)
        vWhereFields.Add("authorised_by", CDBField.FieldTypes.cftDate, "", CDBField.FieldWhereOperators.fwoNotEqual Or CDBField.FieldWhereOperators.fwoCloseBracket)
        If vConn.GetCount("purchase_order_payments", vWhereFields) > 0 Then
          Return ProjectText.String16560  'The Contact is a payee for a purchase order payment which has been authorised or paid\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for event booker
      If Not vError Then
        If vConn.GetCount("event_bookings", Nothing, "contact_number = " & ContactNumber) > 0 Then
          Return (ProjectText.String33071) 'The Contact has event bookings\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for event delegate
      If Not vError Then
        If vConn.GetCount("delegates", Nothing, "contact_number = " & ContactNumber) > 0 Then
          Return (ProjectText.String33072) 'The Contact is an event delegate\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for exam bookings
      If Not vError Then
        If vConn.GetCount("exam_bookings", Nothing, "contact_number = " & ContactNumber) > 0 Then
          Return (ProjectText.String33073) 'The Contact has exam bookings\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for batch_transaction_analysis (e.g. used as delivery contact)
      If Not vError Then
        If vConn.GetCount("batch_transaction_analysis", Nothing, "contact_number = " & ContactNumber) > 0 Then
          Return ProjectText.String33074  'The Contact is on transaction details e.g. as delivery contact\r\n\r\nContact cannot be deleted
        End If
      End If
      '--------------------------------------------------------------------------------------------------
      'Check for financial links
      If Not vError Then
        If vConn.GetCount("financial_links", Nothing, "contact_number = " & ContactNumber) > 0 Then
          Return ProjectText.String33076  'The Contact has financial links\r\n\r\nContact cannot be deleted
        End If
      End If

      Return vMsg
    End Function

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Dim vMsg As String = DeleteContact()
      If vMsg.Length > 0 Then Throw New Exception(vMsg)
    End Sub

    Public Function DeleteContact() As String
      Dim vMsg As String
      Dim vInTransaction As Boolean

      vInTransaction = mvEnv.Connection.InTransaction
      If Not vInTransaction Then
        GetContactMergeInfo(mvEnv.Connection, True)
        mvEnv.Connection.StartTransaction()
      End If
      vMsg = DeleteMembership()
      If vMsg.Length = 0 Then vMsg = DeleteCovenants()
      If vMsg.Length = 0 Then vMsg = DeleteFinancial()
      If vMsg.Length = 0 Then vMsg = DeleteContactData()
      If Not vInTransaction Then
        If vMsg.Length = 0 Then
          mvEnv.Connection.CommitTransaction()
        Else
          mvEnv.Connection.RollbackTransaction()
        End If
      End If
      If mvMergeInfoValid Then
        For Each vMergeInfo As ContactMergeInfo In mvMergeInfo
          If vMergeInfo.TableName.Length > 0 Then
            Debug.Print("Table Not Deleted From: " & vMergeInfo.TableName)
          End If
        Next
      End If
      Return vMsg
    End Function

    Private Function DeleteContactData() As String
      Dim vConn As CDBConnection
      Dim vWhereFields As New CDBFields
      Dim vAGRWhereFields As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vMsg As String = ""
      Dim vTables As String

      vConn = mvEnv.Connection
      'For each address linked to this contact
      vRecordSet = vConn.GetRecordSet("SELECT ca.address_number,postcode FROM contact_addresses ca, addresses a WHERE ca.contact_number = " & ContactNumber & " AND a.address_number = ca.address_number")
      With vRecordSet
        While .Fetch() = True
          'Check if there are any other users of this address
          vWhereFields.Clear()
          vWhereFields.Add("address_number", .Fields(1).IntegerValue)
          vWhereFields.Add("contact_number", ContactNumber, CDBField.FieldWhereOperators.fwoNotEqual)
          If vConn.GetCount("contact_addresses", vWhereFields) = 0 Then
            'If there are no other users of this address and it is not O type then delete it
            vWhereFields.Clear()
            vWhereFields.Add("address_number", .Fields(1).IntegerValue)
            vWhereFields.Add("address_type", "O", CDBField.FieldWhereOperators.fwoNotEqual)
            vConn.DeleteRecords("addresses", vWhereFields, False)
            vAGRWhereFields.Clear()
            vAGRWhereFields.Add("postcode", .Fields(2).Value, CDBField.FieldWhereOperators.fwoEqual)
            vAGRWhereFields.Add("address_number", .Fields(1).IntegerValue, CDBField.FieldWhereOperators.fwoNotEqual)
            If vConn.GetCount("addresses", vAGRWhereFields) = 0 Then
              vAGRWhereFields.Remove((2))
              vConn.DeleteRecords("address_geographical_regions", vAGRWhereFields, False)
            End If
          End If
        End While
        .CloseRecordSet()
      End With
      '----------------------------------------------------------------------------------------------
      'Delete from communications log history, communications log subjects and optionally communications_log_links
      '            where the log number matches the communications log
      vWhereFields.Clear()
      vWhereFields.Add("communications_log_number", CDBField.FieldTypes.cftLong, "SELECT cl.communications_log_number FROM communications_log cl WHERE cl.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.DeleteRecordsMultiTable("communications_log_history,communications_log_subjects,communications_log_links", vWhereFields)
      vWhereFields(1).Name = "unique_id"
      vWhereFields.Add("record_type", CDBField.FieldTypes.cftCharacter, "D")
      vConn.DeleteRecords("sticky_notes", vWhereFields, False)
      ClearMergeInfo("communications_log_links")
      '--------------------------------------------------------------------------------------------------
      'Delete from sticky_notes
      vWhereFields.Clear()
      vWhereFields.Add("unique_id", CDBField.FieldTypes.cftLong, ContactNumber)
      vWhereFields.Add("record_type", CDBField.FieldTypes.cftCharacter, "C")
      vConn.DeleteRecords("sticky_notes", vWhereFields, False)
      ClearMergeInfo("sticky_notes")
      '--------------------------------------------------------------------------------------------------
      'Delete from contact_links
      vWhereFields.Clear()
      vWhereFields.Add("contact_number_1", CDBField.FieldTypes.cftLong, ContactNumber)
      vConn.DeleteRecords("contact_links", vWhereFields, False)
      vWhereFields.Clear()
      vWhereFields.Add("contact_number_2", CDBField.FieldTypes.cftLong, ContactNumber)
      vConn.DeleteRecords("contact_links", vWhereFields, False)
      ClearMergeInfo("contact_links")
      'Delete from organisation_links (if contact is linked to organisation)
      vWhereFields.Clear()
      vWhereFields.Add("organisation_number_2", CDBField.FieldTypes.cftLong, ContactNumber)
      vConn.DeleteRecords("organisation_links", vWhereFields, False)
      '----------------------------------------------------------------------------------------------
      'Delete from contact_addresses, contact_address_usages, communications, communications log
      '            contact_categories, contact_positions, contact_roles, contact_mailings
      '            contact_suppressions, contact_users, selected_contacts, contacts, contact_external_links, contact_actions
      '            where the contact_number matches
      vWhereFields.Clear()
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
      vTables = "contact_addresses,contact_address_usages,communications,communications_log,"
      vTables = vTables & "contact_categories,contact_positions,contact_roles,contact_mailings,"
      vTables = vTables & "contact_suppressions,contact_users,selected_contacts,contacts,contact_external_links,"
      vTables = vTables & "contact_appointments,contact_journals,contact_actions,registered_users,"
      vTables = vTables & "contact_header,contact_expenditure,contact_performances,contact_scores"
      vConn.DeleteRecordsMultiTable(vTables, vWhereFields)
      ClearMergeInfo(vTables)
      vConn.DeleteRecords("contact_cpd_items", vWhereFields, False)
      ClearMergeInfo("contact_cpd_items")
      vConn.DeleteRecords("principal_users", vWhereFields, False)
      ClearMergeInfo("principal_users")
      vConn.DeleteRecords("contact_search_names", vWhereFields, False)
      ClearMergeInfo("contact_search_names")
      '--------------------------------------------------------------------------------------------------
      'Delete from department_notes
      vWhereFields.Clear()
      vWhereFields.Add("unique_id", CDBField.FieldTypes.cftLong, ContactNumber)
      vWhereFields.Add("record_type", CDBField.FieldTypes.cftCharacter, "C")
      vConn.DeleteRecords("department_notes", vWhereFields, False)
      Return vMsg
    End Function

    Private Function DeleteCovenants() As String
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vConn As CDBConnection
      Dim vTables As String

      vConn = mvEnv.Connection
      vWhereFields.Add("covenant_number", CDBField.FieldTypes.cftLong, "SELECT covenant_number FROM covenants c WHERE c.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.DeleteRecords("tax_claim_lines", vWhereFields, False)

      vUpdateFields.Clear()
      vUpdateFields.Add("covenant", CDBField.FieldTypes.cftCharacter, "N")
      vWhereFields.Clear()
      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, "SELECT order_number FROM covenants c WHERE c.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.UpdateRecords("orders", vUpdateFields, vWhereFields, False)

      vWhereFields.Clear()
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
      vTables = "covenants,gift_aid_donations"
      vConn.DeleteRecordsMultiTable(vTables, vWhereFields)
      ClearMergeInfo(vTables)
      Return ""
    End Function

    Private Function DeleteFinancial() As String
      Dim vUpdateFields As New CDBFields
      Dim vWhereFields As New CDBFields
      Dim vConn As CDBConnection
      Dim vRecordSet As CDBRecordSet
      Dim vMsg As String = ""
      Dim vTables As String

      vConn = mvEnv.Connection
      'Update financial history records to show no bank details number
      vUpdateFields.Add("bank_details_number", CDBField.FieldTypes.cftLong)
      vWhereFields.Add("bank_details_number", CDBField.FieldTypes.cftLong, "SELECT ca.bank_details_number FROM contact_accounts ca WHERE ca.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.UpdateRecords("financial_history", vUpdateFields, vWhereFields, False)
      '----------------------------------------------------------------------------------------------
      'Update orders to show no bankers order
      vUpdateFields.Clear()
      vUpdateFields.Add("bankers_order", CDBField.FieldTypes.cftCharacter, "N")
      vWhereFields.Clear()
      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, "SELECT bo.order_number FROM bankers_orders bo WHERE bo.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.UpdateRecords("orders", vUpdateFields, vWhereFields, False)
      '----------------------------------------------------------------------------------------------
      vUpdateFields.Clear()
      vUpdateFields.Add("direct_debit", CDBField.FieldTypes.cftCharacter, "N")
      vWhereFields.Clear()
      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, "SELECT dd.order_number FROM direct_debits dd WHERE dd.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.UpdateRecords("orders", vUpdateFields, vWhereFields, False)
      '----------------------------------------------------------------------------------------------
      'Delete financial history details
      vRecordSet = vConn.GetRecordSet("SELECT batch_number, transaction_number FROM financial_history WHERE contact_number = " & ContactNumber)
      With vRecordSet
        While .Fetch() = True
          vConn.DeleteRecords("financial_history_details", .Fields, False)
        End While
        .CloseRecordSet()
      End With
      '----------------------------------------------------------------------------------------------
      'Delete from order details, order payment history, subscriptions  where there is an order
      vWhereFields.Clear()
      vWhereFields.Add("order_number", CDBField.FieldTypes.cftLong, "SELECT o.order_number FROM orders o WHERE o.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.DeleteRecordsMultiTable("order_details,order_payment_history,subscriptions", vWhereFields)

      '----------------------------------------------------------------------------------------------
      'Delete from purchase order details where there is a purchase order (must be cancelled)
      vWhereFields.Clear()
      vWhereFields.Add("purchase_order_number", CDBField.FieldTypes.cftLong, "SELECT po.purchase_order_number FROM purchase_orders po WHERE po.contact_number = " & ContactNumber, CDBField.FieldWhereOperators.fwoIn)
      vConn.DeleteRecordsMultiTable("purchase_order_details,purchase_order_payments", vWhereFields)
      '----------------------------------------------------------------------------------------------
      ' Delete from contact_accounts, contact_credit_cards, receipts, bankers_orders, direct_debits
      '             subscriptions, vat_registration_numbers, order_details, financial_history, orders
      '             purchase_orders
      vWhereFields.Clear()
      vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
      vTables = "contact_accounts,contact_credit_cards,receipts,bankers_orders,direct_debits,subscriptions,vat_registration_numbers,order_details,financial_history,orders,purchase_orders"
      vConn.DeleteRecordsMultiTable(vTables, vWhereFields)
      ClearMergeInfo(vTables)
      Return vMsg
    End Function

    Private Function DeleteMembership() As String
      Dim vConn As CDBConnection
      Dim vWhereFields As New CDBFields
      Dim vWhere2 As New CDBFields
      Dim vRecordSet As CDBRecordSet
      Dim vOrderNumber As Integer
      Dim vMembership As Integer
      Dim vMsg As String = ""

      vConn = mvEnv.Connection
      vWhereFields.Add("contact_number", ContactNumber)
      vRecordSet = vConn.GetRecordSet("SELECT order_number, membership_number FROM members WHERE contact_number = " & ContactNumber)
      With vRecordSet
        While .Fetch() = True
          vOrderNumber = .Fields(1).IntegerValue
          vMembership = .Fields(2).IntegerValue
          'Delete from members
          vWhere2.Clear()
          vWhere2.Add("order_number", CDBField.FieldTypes.cftLong, vOrderNumber)
          vWhere2.Add("membership_number", vMembership, CDBField.FieldWhereOperators.fwoNotEqual)
          vConn.DeleteRecords("members", vWhere2, False)
          'Delete from tax claim lines
          vWhere2.Clear()
          vWhere2.Add("covenant_number", CDBField.FieldTypes.cftLong, "SELECT covenant_number FROM covenants c WHERE c.order_number = " & vOrderNumber, CDBField.FieldWhereOperators.fwoIn)
          vConn.DeleteRecords("tax_claim_lines", vWhere2, False)
          'Delete from orders, order details, subscriptions,order payment history, bankers orders, direct debits
          vWhere2.Clear()
          vWhere2.Add("order_number", CDBField.FieldTypes.cftLong, vOrderNumber)
          vConn.DeleteRecordsMultiTable("orders,order_details,subscriptions,order_payment_history,bankers_orders,direct_debits", vWhere2)
        End While
        vRecordSet.CloseRecordSet()
      End With
      vConn.DeleteRecords("members", vWhereFields, False)
      ClearMergeInfo("members")
      Return vMsg
    End Function

    Private Sub ClearMergeInfo(ByRef pTables As String)
      Dim vTables() As String

      If mvMergeInfoValid Then
        vTables = Split(pTables, ",")
        For Each vTableName As String In vTables
          For Each vMergeInfo As ContactMergeInfo In mvMergeInfo
            If vMergeInfo.TableName = vTableName Then vMergeInfo.TableName = ""
          Next
        Next
      End If
    End Sub

    Public ReadOnly Property PrincipalUser() As PrincipalUser
      Get
        If Not mvPrincipalUserValid Then
          mvPrincipalUser = New PrincipalUser
          If mvExisting Then
            mvPrincipalUser.Init(mvEnv, ContactNumber)
          Else
            mvPrincipalUser.Init(mvEnv)
          End If
          mvPrincipalUserValid = True
        End If
        Return mvPrincipalUser
      End Get
    End Property

    Public Property VATNumber() As String
      Get
        If Not mvVATNumberValid Then
          If mvExisting Then mvVATNumber = mvEnv.Connection.GetValue(mvEnv.Connection.GetSelectSQLCSC & "vat_registration_number FROM vat_registration_numbers WHERE contact_number = " & ContactNumber)
          mvVATNumberValid = True
        End If
        VATNumber = mvVATNumber
      End Get
      Set(ByVal Value As String)
        Dim vWhereFields As New CDBFields
        Dim vFields As New CDBFields
        If Value <> VATNumber Then
          vWhereFields.Add("contact_number", ContactNumber)
          If Value.Length > 0 Then
            vFields.Add("vat_registration_number", Value)
            If mvVATNumber.Length > 0 Then
              mvEnv.Connection.UpdateRecords("vat_registration_numbers", vFields, vWhereFields)
            Else
              vFields.Add("contact_number", ContactNumber)
              mvEnv.Connection.InsertRecord("vat_registration_numbers", vFields)
            End If
          Else
            mvEnv.Connection.DeleteRecords("vat_registration_numbers", vWhereFields)
          End If
        End If
        mvVATNumberValid = False 'Make sure the VAT Number is reloaded in case of rollback
      End Set
    End Property

    Public ReadOnly Property WebAddress() As String
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vSQL As String

        If Not mvWebAddressValid Then
          If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataDevicesWWWAddress) Then
            vSQL = "SELECT " & mvEnv.Connection.DBSpecialCol("", "number") & " FROM communications co, devices d WHERE co.contact_number = " & ContactNumber & " AND co.device = d.device AND www_address = 'Y'"
            If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then
              vSQL = vSQL & " AND is_active = 'Y' ORDER BY device_default DESC, co.device"
            Else
              vSQL = vSQL & " ORDER BY co.device"
            End If
            vRecordSet = mvEnv.Connection.GetRecordSet(vSQL)
            If vRecordSet.Fetch() Then
              mvWebAddress = vRecordSet.Fields(1).Value
            End If
            vRecordSet.CloseRecordSet()
          End If
          mvWebAddressValid = True
        End If
        Return mvWebAddress
      End Get
    End Property

    Private ReadOnly Property ActionCount() As Integer
      Get
        Dim vWhereFields As New CDBFields
        Dim vStatuses As New CDBParameters

        vStatuses.Add(Action.GetActionStatusCode(Action.ActionStatuses.astDefined))
        vStatuses.Add(Action.GetActionStatusCode(Action.ActionStatuses.astScheduled))
        vStatuses.Add(Action.GetActionStatusCode(Action.ActionStatuses.astOverdue))

        If ContactType = ContactTypes.ctcOrganisation Then
          vWhereFields.Add("organisation_number", ContactNumber)
          vWhereFields.Add("oa.action_number", CDBField.FieldTypes.cftLong, "a.action_number")
          vWhereFields.Add("action_status", CDBField.FieldTypes.cftCharacter, vStatuses.InList, CDBField.FieldWhereOperators.fwoIn)
          Return mvEnv.Connection.GetCount("organisation_actions oa, actions a", vWhereFields)
        Else
          vWhereFields.Add("contact_number", ContactNumber)
          vWhereFields.Add("ca.action_number", CDBField.FieldTypes.cftLong, "a.action_number")
          vWhereFields.Add("action_status", CDBField.FieldTypes.cftCharacter, vStatuses.InList, CDBField.FieldWhereOperators.fwoIn)
          Return mvEnv.Connection.GetCount("contact_actions ca, actions a", vWhereFields)
        End If
      End Get
    End Property

    Public ReadOnly Property PreferredCommunication() As String
      Get
        Dim vRecordSet As CDBRecordSet
        Dim vComms As New Communication(mvEnv)
        Dim vWhereFields As New CDBFields
        Dim vPreferredCommunication As String = ""

        If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCommunicationsUsages) Then
          If ContactType = ContactTypes.ctcOrganisation Then
            vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong)
            vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, "( SELECT address_number FROM organisation_addresses WHERE organisation_number = " & ContactNumber & ")", CDBField.FieldWhereOperators.fwoIn)
          Else
            vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
          End If
          vComms.Init()
          vWhereFields.Add("preferred_method", CDBField.FieldTypes.cftCharacter, "Y")
          vRecordSet = mvEnv.Connection.GetRecordSet("SELECT " & vComms.GetRecordSetFields() & " FROM communications com, devices d WHERE " & mvEnv.Connection.WhereClause(vWhereFields) & " AND com.device = d.device")
          If vRecordSet.Fetch() Then
            vComms.InitFromRecordSet(vRecordSet)
            vPreferredCommunication = vComms.PhoneNumber
          Else
            vPreferredCommunication = ProjectText.String16555 'Not Set
          End If
          vRecordSet.CloseRecordSet()
        End If
        Return vPreferredCommunication
      End Get
    End Property

    Public Sub UpdateOwner()
      Dim vOriginalDepartment As String = mvClassFields.Item(ContactFields.Department).SetValue
      'We may have changed the owning department and need to update the contact_users records
      If Department <> vOriginalDepartment Then
        If Owners.Exists(Department) Then
          'We are just swapping ownership and will leave the original owner as a co-owner
        Else
          'We are adding a new department as the owner - remove the original (assuming it existed)
          AddOwner(Department)
          If Owners.Exists(vOriginalDepartment) Then RemoveOwner(Owners(vOriginalDepartment))
        End If
      End If
    End Sub

    Public Sub AddOwner(ByVal pOwner As String)
      Dim vFields As New CDBFields
      vFields.AddAmendedOnBy(mvEnv.User.Logname, TodaysDate)
      vFields.Add("contact_number", ContactNumber)
      vFields.Add("department", pOwner)
      mvEnv.Connection.InsertRecord("contact_users", vFields)
    End Sub

    Public Sub RemoveOwner(ByVal pOwner As CDBParameter)
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("contact_number", ContactNumber)
      vWhereFields.Add("department", pOwner.Name)
      mvEnv.Connection.DeleteRecords("contact_users", vWhereFields, If(ContactType = ContactTypes.ctcOrganisation, False, True)) 'Don't error if department doesn't exist for the dummy contact
    End Sub

    Public ReadOnly Property StatusChangedAction() As Integer
      Get
        StatusChangedAction = mvStatusChangedAction
      End Get
    End Property
    Public Function GetPositions(Optional ByVal pAddressNumber As Integer = 0, Optional ByVal pOrganisationNumber As Integer = 0, Optional ByVal pPosition As String = "", Optional ByVal pStarted As String = "", Optional ByVal pFinished As String = "", Optional ByVal pCurrent As ContactPosition.CurrentSettingTypes = ContactPosition.CurrentSettingTypes.cstNone) As List(Of ContactPosition)
      Dim vRS As CDBRecordSet
      Dim vSQL As String
      Dim vFields As New CDBFields

      If mvContactPositions Is Nothing Then
        Dim vContactPosition As New ContactPosition(mvEnv)
        mvContactPositions = New List(Of ContactPosition)
        vContactPosition.Init()
        'Build basic SELECT statement
        vSQL = "SELECT " & vContactPosition.GetRecordSetFields() & " FROM contact_positions cp WHERE "
        With vFields
          If ContactNumber > 0 Then .Add("cp.contact_number", ContactNumber)
          If pAddressNumber > 0 Then .Add("cp.address_number", pAddressNumber)
          If pOrganisationNumber > 0 Then .Add("cp.organisation_number", pOrganisationNumber)

          If pPosition.Length > 0 Then .Add("position", CDBField.FieldTypes.cftCharacter, pPosition)
          If pStarted.Length > 0 Then .Add("started", CDBField.FieldTypes.cftDate, pStarted)
          If pFinished.Length > 0 Then .Add("finished", CDBField.FieldTypes.cftDate, pFinished)
          If pCurrent <> ContactPosition.CurrentSettingTypes.cstNone Then
            .Add("current", If(pCurrent = ContactPosition.CurrentSettingTypes.cstCurrent, "Y", "N"))
            .Item("current").SpecialColumn = True
            .TableAlias = "cp"
          End If
        End With
        vRS = mvEnv.Connection.GetRecordSet(vSQL & " " & mvEnv.Connection.WhereClause(vFields))
        With vRS
          While .Fetch()
            vContactPosition = New ContactPosition(mvEnv)
            vContactPosition.InitFromRecordSet(vRS)
            mvContactPositions.Add(vContactPosition)
          End While
          .CloseRecordSet()
        End With
      End If
      Return mvContactPositions
    End Function

    Public Sub SetDefaultAddress(ByVal pNewDefaultAddress As Address, ByVal pSetOldHistoric As Boolean, ByVal pSwitchAddresses As Boolean, ByVal pUpdateMembers As Boolean, Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "", Optional ByVal pOldAddressValidTo As String = "")
      'This method is designed to set the default address of a contact to the address
      'passed to this method - An assumption is made that the contact class has been
      'correctly initialised and the current default address has been read
      ' BR12131 - Added optional parameter to allow us to pass in a valid_to date for the old address. Only used when pSetOldHistoric is true

      Dim vContactAddress As New ContactAddress(mvEnv)
      Dim vStartTrans As Boolean
      Dim vWhereFields As New CDBFields
      Dim vUpdateFields As New CDBFields

      Dim vOldAddressNumber As Integer = AddressNumber
      If pNewDefaultAddress.AddressNumber <> vOldAddressNumber Then
        If Not mvEnv.Connection.InTransaction Then
          mvEnv.Connection.StartTransaction()
          vStartTrans = True
        End If
        'First update the contact - Either the organisation dummy contact or the actual contact
        mvClassFields.Item(ContactFields.AddressNumber).IntegerValue = pNewDefaultAddress.AddressNumber
        RemoveGoneAway(False) 'Remove the Gone Away status
        SaveChanges(mvEnv.User.Logname, True)
        With vContactAddress
          If ContactType <> ContactTypes.ctcOrganisation Then
            'Now check the contact address record for the new address
            .InitFromContactAndAddress(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, ContactNumber, (pNewDefaultAddress.AddressNumber))
            If IsDate(pValidFrom) Then
              If CDate(pValidFrom) <= Today Then .ValidFrom = pValidFrom
            End If
            If IsDate(pValidTo) Then
              If CDate(pValidTo) >= Today Then .ValidTo = pValidTo
            End If
            'If one exists then make sure it is no longer historic
            If .Existing Then
              If .Historical = True Or IsDate(.ValidTo) Then
                .Historical = False
                If CDate(.ValidTo) < Today Then .ValidTo = ""
              End If
            End If
            .Save(mvEnv.User.UserID, True) 'Save the new one

            If pSetOldHistoric Then
              .InitFromContactAndAddress(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, ContactNumber, vOldAddressNumber)
              If .Existing Then
                .Historical = True
                If pOldAddressValidTo <> "" And IsDate(pOldAddressValidTo) Then
                  .ValidTo = CDate(pOldAddressValidTo).ToString(CAREDateFormat)
                Else
                  If .ValidTo = "" Then .ValidTo = TodaysDate()
                End If
                .Save(mvEnv.User.UserID, True) 'Save the old one
              End If
            End If
          Else
            .InitFromContactAndAddress(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, ContactNumber, vOldAddressNumber)
            ' BR 11752 - Need to ensure these are populated
            If IsDate(pValidFrom) Then
              .ValidFrom = pValidFrom
            Else
              .ValidFrom = ""
            End If
            If IsDate(pValidTo) Then
              .ValidTo = pValidTo
            Else
              .ValidTo = ""
            End If

            If IsDate(.ValidTo) Then
              .Historical = (CDate(.ValidTo) < Today)
            Else
              .Historical = False
            End If

            If .Existing Then
              .AddressNumber = pNewDefaultAddress.AddressNumber
              .Save(mvEnv.User.UserID)
            End If
            'Now update the dummy contact position
            vWhereFields.Add("contact_number", CDBField.FieldTypes.cftLong, ContactNumber)
            vWhereFields.Add("address_number", CDBField.FieldTypes.cftLong, vOldAddressNumber)
            vUpdateFields.Add("address_number", CDBField.FieldTypes.cftLong, pNewDefaultAddress.AddressNumber)
            mvEnv.Connection.UpdateRecords("contact_positions", vUpdateFields, vWhereFields, False)
          End If
        End With
        If pSwitchAddresses Then SwitchCurrentAddress(Address, pNewDefaultAddress, pUpdateMembers)
        If vStartTrans Then mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Public Function MovePosition(ByRef pCurrentPosition As ContactPosition, ByRef pNewPosition As ContactPosition) As Boolean
      Dim vAdjustRoles As Boolean
      Dim vChangedDefaultAddress As Boolean
      Dim vPosActivityRecordSet As CDBRecordSet = Nothing
      Dim vPosLinkRecordSet As CDBRecordSet = Nothing

      'BR19961 First save the new position so that gone away does not get set for the fleeting moment when you move and change address
      pNewPosition.SaveWithAddressLink(True, mvEnv.User.UserID, True) 'Save here so that contact_addresses not left historic if just changing position
      ' KA (2011/12/21) added true for audit, as no amendment history was being added for the move

      If Not pCurrentPosition Is Nothing Then
        If pNewPosition.OrganisationNumber <> pCurrentPosition.OrganisationNumber Then
          vAdjustRoles = True
        Else
          If DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(pCurrentPosition.Finished), CDate(pNewPosition.Started)) > 1 Then vAdjustRoles = True
        End If
        pCurrentPosition.SaveWithAddressLink(vAdjustRoles, mvEnv.User.UserID, True) ' KA (2011/12/21) added true for audit, as no amendment history was being added for the move
        'Jira 388: If 'cd_position_links_move' config set then Move Current Position Activities/Relationships to New Position
        If mvEnv.GetConfig("cd_position_links_move", "NEVER").ToUpper = "ALWAYS" Then
          pNewPosition.MovePositionData(pCurrentPosition.ContactPositionNumber, vPosActivityRecordSet, vPosLinkRecordSet)
        End If
        'Update PositionActivities/Links attached to the current Position
        pCurrentPosition.UpdatePositionData()
        If AddressNumber = pCurrentPosition.AddressNumber And AddressNumber <> pNewPosition.AddressNumber And pNewPosition.Current Then
          'mvContact.AddressNumber = mvContactPosition.AddressNumber       'Change the default address of the contact
          Dim vAddress As New Address(mvEnv)
          vAddress.Init(pNewPosition.AddressNumber)
          SetDefaultAddress(vAddress, True, True, True)
          vChangedDefaultAddress = True
        End If
      End If

      'If Position Activity records were selected to move in MovePositionData routine then Move to new Position now after Save
      If vPosActivityRecordSet IsNot Nothing Then
        While vPosActivityRecordSet.Fetch
          Dim vActivity As New PositionCategory(mvEnv)
          vActivity.InitFromRecordSet(vPosActivityRecordSet)
          Dim vNewActivity As New PositionCategory(mvEnv)
          mvEnv.Connection.StartTransaction()
          vNewActivity.MovePositionActivities(vActivity, pNewPosition.ContactPositionNumber, pNewPosition.Started, pNewPosition.Finished)
          vNewActivity.Save()
          mvEnv.Connection.CommitTransaction()
        End While
        vPosActivityRecordSet.CloseRecordSet()
      End If
      'If Position Link records were selected to move in MovePositionData routine then Move to new Position now after Save
      If vPosLinkRecordSet IsNot Nothing Then
        While vPosLinkRecordSet.Fetch
          Dim vLink As New PositionLink(mvEnv)
          vLink.InitFromRecordSet(vPosLinkRecordSet)
          Dim vNewLink As New PositionLink(mvEnv)
          mvEnv.Connection.StartTransaction()
          vNewLink.MovePositionLinks(vLink, pNewPosition.ContactPositionNumber, pNewPosition.Started, pNewPosition.Finished)
          vNewLink.Save()
          mvEnv.Connection.CommitTransaction()
        End While
        vPosLinkRecordSet.CloseRecordSet()
      End If

      If pNewPosition.Current Then
        If pNewPosition.Organisation.OrganisationNumber = pNewPosition.Organisation.ContactNumber Then 'Set the organisation default contact if not already set
          pNewPosition.Organisation.ContactNumber = ContactNumber
          pNewPosition.Organisation.Save()
        End If
      End If
      Return vChangedDefaultAddress
    End Function

    Public Sub CreateAtExistingAddress(ByVal pEnv As CDBEnvironment, ByVal pParams As CDBParameters, ByVal pAddressNumber As Integer, ByVal pPosition As String, Optional ByVal pNoCapitalisation As Boolean = False)
      'This function is only used by WEB Services at present
      'Please let me know if you want to change it (SDT)
      Init()
      Address.Init(pAddressNumber)
      mvOrganisationNumber = Address.OrganisationNumber 'Force read outside of transaction
      If Address.Existing Then
        CreateAtAddress(pParams, pPosition, pNoCapitalisation)
      Else
        RaiseError(DataAccessErrors.daeExpectedDataMissing)
      End If
    End Sub

    Public Overloads Sub Create(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters, Optional ByRef pNoCapitalisation As Boolean = False)
      Init()
      Address.Create(pEnv, Me, pParams, pParams("Address").Value = " ") ', pHouseName, pAddress, pTown, pCounty, pPostCode, pCountry, pPaf, pBranch, pBuildingNumber
      CreateAtAddress(pParams, "", pNoCapitalisation)
    End Sub

    Private Sub CreateAtAddress(ByVal pParams As CDBParameters, ByVal pPosition As String, Optional ByVal pNoCapitalisation As Boolean = False)

      If pParams.HasValue("ContactNumber") Then
        'Contact number is being passed.  E.g. via OpenLink
        ValidateContactNumber(MaintenanceTypes.Insert, pParams("ContactNumber").Value)
        SetContactNumber(pParams("ContactNumber").IntegerValue)
      End If

      Forenames = pParams.ParameterExists("Forenames").CapitalisedValue(pNoCapitalisation)
      Dim vInitials As String = pParams.ParameterExists("Initials").Value
      If vInitials.Length = 0 Then
        Initials = SpacePadInitials(Forenames)
      Else
        Initials = vInitials
      End If
      Dim vPreferred As String = pParams.ParameterExists("PreferredForename").Value
      If vPreferred.Length = 0 Then
        SetValid(ContactFields.PreferredForename)
      Else
        PreferredForename = vPreferred
      End If
      If pParams.Exists("SurnamePrefix") Then
        SurnamePrefix = pParams("SurnamePrefix").Value
        If pParams.HasValue("SurnamePrefix") Then pParams("Surname").Value = Trim(Replace(Trim(pParams("Surname").Value), pParams("SurnamePrefix").Value & " ", ""))
      End If
      Surname = pParams("Surname").CapitalisedValue(pNoCapitalisation)
      Honorifics = pParams.ParameterExists("Honorifics").Value

      Notes = pParams.ParameterExists("Notes").Value
      DateOfBirth = pParams.ParameterExists("DateOfBirth").Value
      DobEstimated = pParams.ParameterExists("DOBEstimated").Bool
      JuniorContact = pParams.ParameterExists("JuniorContact").Bool
      Dim vGroup As String = pParams.ParameterExists("ContactGroup").Value
      If vGroup.Length = 0 Then vGroup = mvEnv.EntityGroups.DefaultGroup(EntityGroup.EntityGroupTypes.egtContact).EntityGroupCode
      ContactGroupCode = vGroup
      If pParams.HasValue("Department") Then Department = pParams("Department").Value
      If pParams.HasValue("NiNumber") Then NiNumber = pParams("NiNumber").Value
      VATCategory = pParams.OptionalValue("VatCategory", mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefConVatCat))

      Source = pParams("Source").Value
      SourceDate = pParams.OptionalValue("SourceDate", TodaysDate)

      Status = pParams.ParameterExists("Status").Value
      If pParams.HasValue("StatusDate") Then StatusDate = pParams("StatusDate").Value
      If pParams.Exists("StatusReason") Then StatusReason = pParams("StatusReason").Value

      If pParams.HasValue("OwnershipGroup") Then
        OwnershipGroup = pParams("OwnershipGroup").Value
      Else
        OwnershipGroup = mvEnv.User.OwnershipGroup
      End If

      If pParams.HasValue("PrefixHonorifics") Then PrefixHonorifics = pParams("PrefixHonorifics").Value

      'Set the title
      mvClassFields.Item(ContactFields.Title).Value = CapitaliseWords(pParams.ParameterExists("Title").Value, False, False, True)
      If pParams.HasValue("LabelNameFormatCode") Then LabelNameFormatCode = pParams("LabelNameFormatCode").Value
      Dim vTitle As New Title
      vTitle.Init(mvEnv, TitleName)
      'We need to check for joint contacts here
      CheckForJointContacts(Nothing, vTitle)

      If ContactType = ContactTypes.ctcJoint Then
        Sex = ContactSex.cscUnknown
        If LabelNameFormatCode.Length = 0 Then
          If (Not mvLabelNameValid) Then SetLabelName(False)
        Else
          If (Not mvLabelNameValid) Then SetLabelName(False, LabelNameFormatCode)
        End If
      Else
        TitleName = vTitle.TitleName
        SetTitle(vTitle)
        Dim vSex As String = pParams.ParameterExists("Sex").Value
        If Sex = ContactSex.cscUnknown Then
          If vSex.Length > 0 Then
            If vSex = "M" Then
              Sex = ContactSex.cscMale
            ElseIf vSex = "F" Then
              Sex = ContactSex.cscFemale
            End If
            If Salutation = Surname And Not vTitle.Existing Then 'Title is not known but gender is and the Salutation has been set to the Surname via the call to SetTitle above
              If pParams.Exists("Salutation") Then
                If Not pParams.HasValue("Salutation") Then
                  Salutation = "" 'clear the default Salutation setting
                  pParams.Item("Salutation").Value = Salutation 'let the Salutation property itself determine a value
                End If
              Else
                Salutation = "" 'clear the default Salutation setting
                pParams.Add("Salutation", CDBField.FieldTypes.cftCharacter, Salutation) 'let the Salutation property itself determine a value
              End If
            End If
          End If
        End If
      End If
      If pParams.HasValue("Salutation") Then Salutation = pParams("Salutation").Value
      If pParams.HasValue("LabelName") Then LabelName = pParams("LabelName").Value
      If pParams.HasValue("InformalSalutation") Then InformalSalutation = pParams("InformalSalutation").Value
      If pParams.HasValue("ResponseChannel") Then ResponseChannel = pParams("ResponseChannel").Value
      If pParams.HasValue("ContactReference") Then ContactReference = pParams("ContactReference").Value
      '-----------------------------------------------------------------------------
      'Try and sort out or cache the control numbers prior to starting a transaction
      If Address.AddressNumber = 0 Then Address.SetControlNumber()
      SetControlNumber()
      If Address.AddressType = Address.AddressTypes.ataOrganisation Then mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnPosition, 1)
      mvEnv.CacheControlNumbers(CDBEnvironment.CachedControlNumberTypes.ccnAddressLink, 1)
      '-----------------------------------------------------------------------------

      mvEnv.Connection.StartTransaction()
      If Not Address.Existing Then Address.Save(mvEnv.User.UserID, True)
      AddressNumber = Address.AddressNumber
      Dim vAddressValidFrom As String = TodaysDate()
      If pParams.ContainsKey("ValidFrom") Then vAddressValidFrom = pParams("ValidFrom").Value
      Dim vAddressValidTo As String = ""
      If pParams.ContainsKey("ValidTo") Then vAddressValidTo = pParams("ValidTo").Value
      SaveCheckForJointContact(vAddressValidFrom, vAddressValidTo, mvEnv.User.UserID, True) 'Adds default suppression, user and contact address record
      If Address.AddressType = Address.AddressTypes.ataOrganisation Then
        Dim vCP As New ContactPosition(mvEnv)
        vCP.Create(ContactNumber, AddressNumber, Address.OrganisationNumber, "Y", "Y", pPosition, TodaysDate(), "", pParams.OptionalValue("PositionLocation", ""))
        vCP.Save(mvEnv.User.UserID, True)
        mvOrganisationNumber = Address.OrganisationNumber
        mvOrganisationName = Address.OrganisationName
        mvContactPositionNumber = vCP.ContactPositionNumber
        mvPosition = pPosition
      End If
      If pParams.HasValue("PrincipalUser") Then
        Dim vPU As New PrincipalUser
        vPU.Create(mvEnv, ContactNumber, pParams("PrincipalUser").Value, pParams.ParameterExists("PrincipalUserReason").Value)
      End If
      If pParams.HasValue("VatNumber") Then VATNumber = pParams("VatNumber").Value
      mvPositionValid = True
      mvEnv.Connection.CommitTransaction()
    End Sub

    Private Sub ValidateContactNumber(pType As MaintenanceTypes, vNumber As String)
      Dim vContactNumber As Integer
      If Integer.TryParse(vNumber, vContactNumber) Then
        'First check for uniqueness if inserting
        If pType = MaintenanceTypes.Insert Then
          Dim vContact As New Contact(Me.Environment)
          vContact.Init(vContactNumber)
          If vContact.Existing Then
            RaiseError(DataAccessErrors.daeRecordExists, ProjectText.LangContactNumber)
          End If
        End If
        'Now check for boundaries
        Dim vSQL As New SQLStatement(Me.Environment.Connection, "control_number", "control_numbers", New CDBField("control_number_type", Me.ClassFields.ControlNumberType))
        Dim vMaxNumber As Integer = CType(Me.Environment.Connection.GetValue(vSQL.SQL, True), Integer)
        vMaxNumber -= 1 'the last created contact number should be the control number minus one.  The next Contact number is the Control Number
        If vContactNumber > vMaxNumber Then
          RaiseError(DataAccessErrors.daeRangeExceeded, ProjectText.LangContactNumber, vMaxNumber.ToString())
        End If
      Else
        'The code here will only be reached when a Contact is created with parameters that have bypassed ValidateParameterList, which should be never.
        RaiseError(DataAccessErrors.daeInvalidParameter, ProjectText.LangContactNumber)
      End If
    End Sub

    Public Sub SaveCheckForJointContact(ByVal pAddressValidFrom As String, ByVal pAddressValidTo As String, Optional ByVal pAmendedBy As String = "", Optional ByVal pAudit As Boolean = False)
      Dim vAddress As Address = Nothing
      Dim vDeDup As Boolean
      Dim vContactLink As New ContactLink(mvEnv)

      If ContactType = ContactTypes.ctcJoint Then
        'Setup in case dedup required later
        If AddressNumber <> IntegerValue(mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlNonAddress)) Then 'Dedup if not non-address number
          vAddress = New Address(mvEnv)
          vAddress.Init(AddressNumber)
          vDeDup = True
        End If
        With mvClassFields
          .Item(ContactFields.PreferredForename).Value = ""
          .Item(ContactFields.Honorifics).Value = ""
        End With
        'SetLabelName False
        SaveNewContact(pAmendedBy, pAudit, vAddress, False, pAddressValidFrom, pAddressValidTo)
        mvJointContact1.AddressNumber = AddressNumber
        mvJointContact1.SaveNewContact(pAmendedBy, pAudit, vAddress, vDeDup, pAddressValidFrom, pAddressValidTo)
        mvJointContact2.AddressNumber = AddressNumber
        mvJointContact2.SaveNewContact(pAmendedBy, pAudit, vAddress, vDeDup, pAddressValidFrom, pAddressValidTo)

        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesignoreExisting, ContactTypes.ctcContact, ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlJointSuppression))
        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesignoreExisting, ContactTypes.ctcContact, mvJointContact1.ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlsDerivedSuppression))
        ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesignoreExisting, ContactTypes.ctcContact, mvJointContact2.ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlsDerivedSuppression))

        vContactLink.Init()
        vContactLink.InitNew(Me.Environment, ContactLink.ContactLinkTypes.cltContact, ContactNumber, mvJointContact1.ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink))
        vContactLink.Save()

        vContactLink = New ContactLink(Me.Environment)
        vContactLink.Init()
        vContactLink.InitNew(mvEnv, ContactLink.ContactLinkTypes.cltContact, ContactNumber, mvJointContact2.ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToJointLink))
        vContactLink.Save()

        vContactLink = New ContactLink(Me.Environment)
        vContactLink.Init()
        vContactLink.InitNew(mvEnv, ContactLink.ContactLinkTypes.cltContact, mvJointContact1.ContactNumber, mvJointContact2.ContactNumber, mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDerivedToDerivedLink))
        vContactLink.AllowsOverlaps = True 'validaion of overlaps moved from xml layer to object.  As the code above and below bypasses the xml layer, if allows overlaps, so we have to set this property to continue existing functionality.  Should use merge but at the moment that's a terrible pice of code which deletes data so until we have a good alternative...
        vContactLink.Save()
      Else
        SaveNewContact(pAmendedBy, pAudit, Nothing, False, pAddressValidFrom, pAddressValidTo)
      End If
    End Sub

    Friend Sub SaveNewContact(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pAddress As Address, ByVal pDeDup As Boolean, ByVal pAddressValidFrom As String, ByVal pAddressValidTo As String)

      Dim vSuppression As String = mvEnv.GetControlValue(CDBEnvironment.cdbControlConstants.cdbControlDefMailingSupp)
      Dim vJnrSuppression As String = ""
      Dim vValidFrom As String = ""
      If IsJunior() Then
        vJnrSuppression = mvEnv.GetConfig("cd_junior_suppression")
        If IsDate(DateOfBirth) Then
          vValidFrom = DateOfBirth
        Else
          vValidFrom = TodaysDate()
        End If
      End If
      Dim vInsert As Boolean = True
      If pDeDup Then
        For Each vDupContact As Contact In pAddress.ContactsAtAddress(True)
          If vDupContact.IsPotentialDuplicate(Me, False) Then
            If vDupContact.ContactType = ContactTypes.ctcContact Then
              mvClassFields.Item(ContactFields.ContactNumber).IntegerValue = vDupContact.ContactNumber
              vInsert = False
              Exit For
            End If
          End If
        Next vDupContact
      End If
      If vInsert Then
        Save(pAmendedBy, pAudit, 0) 'Save the contact
        AddUser(Department, True)
        Dim vCA As New ContactAddress(mvEnv)
        vCA.Create(mvEnv, ContactAddress.ContactAddresssLinkTypes.caltContact, ContactNumber, AddressNumber, "N", pAddressValidFrom, pAddressValidTo)
        vCA.Save(pAmendedBy, True)
        If vSuppression.Length > 0 Then ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesignoreExisting, ContactTypes.ctcContact, ContactNumber, vSuppression)
        If vJnrSuppression.Length > 0 Then ContactSuppression.ContactTypeSaveSuppression(mvEnv, ContactSuppression.SuppressionEntryStyles.sesignoreExisting, ContactTypes.ctcContact, ContactNumber, vJnrSuppression, vValidFrom, CDate(vValidFrom).AddYears(mvEnv.JuniorAgeLimit).ToString(CAREDateFormat))
      End If
    End Sub

    Public Function IsPotentialDuplicate(ByVal pContact As Contact, ByVal pUseSoundex As Boolean) As Boolean
      Dim vMainInitials As String
      Dim vSubInitials As String
      Dim vMainILen As Integer
      Dim vSubILen As Integer
      Dim vForename As String
      Dim vPos As Integer
      Dim vFirstMatch As Boolean

      Dim vMainType As ContactTypes = ContactType
      If vMainType = ContactTypes.ctcOrganisation Or vMainType = ContactTypes.ctcJoint Then
        IsPotentialDuplicate = True
      Else
        If Surname.Length > 0 Then
          ' BR 12014
          ' See if we need to try the soundex check
          If pUseSoundex Then
            ' Need to pop in a quick soundex check
            vFirstMatch = ((GetSoundexCode((pContact.SurnameWithoutPrefix)) = GetSoundexCode(SurnameWithoutPrefix)) Or (StrComp(Surname, pContact.Surname, CompareMethod.Text) = 0))
          Else
            vFirstMatch = (StrComp(Surname, pContact.Surname, CompareMethod.Text) = 0)
          End If

          If vFirstMatch Then
            If Len(TitleName) > 0 And StrComp(TitleName, pContact.TitleName, CompareMethod.Text) = 0 Then
              vMainInitials = Initials
              vSubInitials = pContact.Initials
              vMainILen = Len(vMainInitials)
              vSubILen = Len(vSubInitials)
              If (Len(pContact.Forenames) = 0 And vSubILen = 0) Or (Len(Forenames) = 0 And vMainILen = 0) Then
                IsPotentialDuplicate = True
              Else
                If vMainILen > vSubILen Then
                  vMainInitials = Left(vMainInitials, vSubILen)
                ElseIf vSubILen > vMainILen Then
                  vSubInitials = Left(vSubInitials, vMainILen)
                End If
                If StrComp(vMainInitials, vSubInitials, CompareMethod.Text) = 0 Then
                  If Len(pContact.Forenames) = 0 Then
                    IsPotentialDuplicate = True
                  Else
                    vForename = Forenames
                    vPos = InStr(vForename, " ")
                    If vPos > 0 Then vForename = Left(vForename, vPos - 1)
                    If Len(vForename) < 2 Then
                      IsPotentialDuplicate = True
                    Else
                      If InStr(pContact.Forenames, vForename) > 0 Then IsPotentialDuplicate = True
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    End Function

    Public Sub CheckForJointContacts(ByVal pTitles As SortedList(Of String, Title), Optional ByVal pTitle As Title = Nothing)
      Dim vTitle1 As Title = Nothing
      Dim vTitle2 As Title = Nothing
      Dim vTitleName1 As String = ""
      Dim vTitleName2 As String = ""
      Dim vInitials1 As String = ""
      Dim vInitials2 As String = ""
      Dim vForename1 As String = ""
      Dim vForename2 As String = ""
      Dim vSurname1 As String = ""
      Dim vSurname2 As String = ""
      Dim vHonorifics1 As String = ""
      Dim vHonorifics2 As String = ""

      If pTitles Is Nothing And pTitle Is Nothing Then
        pTitles = GetTitles(mvEnv)
      End If
      If IsJointContact(pTitles, pTitle) Then
        mvClassFields(ContactFields.ContactType).Value = "J"
        TitleName = JointItem(TitleName, vTitleName1, vTitleName2, TitleName, pTitles, pTitle)
        JointItem(Initials, vInitials1, vInitials2, TitleName, pTitles, pTitle)
        JointItem(Forenames, vForename1, vForename2, TitleName, pTitles, pTitle)
        JointItem(Surname, vSurname1, vSurname2, TitleName, pTitles, pTitle)
        JointItem(Honorifics, vHonorifics1, vHonorifics2, TitleName, pTitles, pTitle)

        If Len(vTitleName1) > 0 Then
          If pTitles Is Nothing Then
            vTitle1 = New Title
            vTitle1.Init(mvEnv, vTitleName1)
          Else
            vTitle1 = pTitles(vTitleName1)
          End If
          If vTitle1.Existing = False Then mvClassFields(ContactFields.Title).Value = ""
        End If
        mvJointContact1 = New Contact(mvEnv)
        mvJointContact1.InitJointContact(Me, vTitle1, vInitials1, vForename1, vSurname1, vHonorifics1)

        If Len(vTitleName2) > 0 Then
          If pTitles Is Nothing Then
            vTitle2 = New Title
            vTitle2.Init(mvEnv, vTitleName2)
          Else
            vTitle2 = pTitles(vTitleName2)
          End If
          If vTitle2.Existing = False Then mvClassFields(ContactFields.Title).Value = ""
        End If
        mvJointContact2 = New Contact(mvEnv)
        mvJointContact2.InitJointContact(Me, vTitle2, vInitials2, vForename2, If(Len(vSurname2) > 0, vSurname2, vSurname1), vHonorifics2)
      Else
        mvJointContact1 = Nothing
        mvJointContact2 = Nothing
        mvClassFields(ContactFields.ContactType).Value = "C"
      End If
    End Sub

    Public Sub InitJointContact(ByVal pContact As Contact, ByVal pTitle As Title, ByVal pInitials As String, ByVal pForenames As String, ByVal pSurname As String, ByVal pHonorifics As String)
      Init()
      With mvClassFields
        .Item(ContactFields.Initials).Value = pInitials
        .Item(ContactFields.Forenames).Value = pForenames
        .Item(ContactFields.Surname).Value = pSurname
        .Item(ContactFields.Honorifics).Value = pHonorifics
        If Not pTitle Is Nothing Then
          .Item(ContactFields.Title).Value = pTitle.TitleName
          SetTitle(pTitle)
        End If
      End With
      With pContact
        mvClassFields(ContactFields.Source).Value = .Source
        mvClassFields(ContactFields.SourceDate).Value = .SourceDate
        mvClassFields(ContactFields.Status).Value = .Status
        mvClassFields(ContactFields.StatusDate).Value = .StatusDate
        mvClassFields(ContactFields.StatusReason).Value = .StatusReason
        mvClassFields(ContactFields.Department).Value = .Department
        mvClassFields(ContactFields.DiallingCode).Value = .DiallingCode
        mvClassFields(ContactFields.StdCode).Value = .StdCode
        mvClassFields(ContactFields.Telephone).Value = .Telephone
        mvClassFields(ContactFields.ExDirectory).Bool = .ExDirectory
        mvClassFields(ContactFields.ContactGroup).Value = .ContactGroupCode
        mvClassFields(ContactFields.Department).Value = .Department
        mvClassFields(ContactFields.OwnershipGroup).Value = .OwnershipGroup
        mvClassFields(ContactFields.ContactVatCategory).Value = .VATCategory
      End With
      SetValid(ContactFields.PreferredForename)
      If LabelNameFormatCode.Length = 0 Then
        If (Not mvLabelNameValid) Then SetLabelName(False)
      Else
        If (Not mvLabelNameValid) Then SetLabelName(False, LabelNameFormatCode)
      End If
      Dim vSalutation As String = Salutation 'Force setting of salutation if blank
    End Sub

    Public Sub SetTitle(ByRef pTitle As Title)
      Dim vSalutation As String = ""
      Dim vSurname As String

      mvTitle = pTitle
      If mvTitle.Sex.Length > 0 Then mvClassFields.Item(ContactFields.Sex).Value = mvTitle.Sex
      vSalutation = mvTitle.Salutation
      If PreferredForename.Length > 0 Then vSalutation = vSalutation.Replace("forename", PreferredForename)
      If SurnamePrefix.Length > 0 Then
        If PreferredForename.Length > 0 And InStr(mvTitle.Salutation, "forename") > 0 Then
          vSurname = Surname
        Else
          vSurname = SurnameCapitalisedPrefix
        End If
      Else
        vSurname = Surname
      End If
      vSalutation = vSalutation.Replace("surname", vSurname)
      vSalutation = vSalutation.Replace("nachname", FirstHyphenatedWord(vSurname))
      If Honorifics.Length > 0 Then
        vSalutation = vSalutation.Replace("honorifics", Honorifics)
      End If
      If vSalutation.Length = 0 Then vSalutation = vSurname
      Salutation = vSalutation
      If mvClassFields.Item(ContactFields.Sex).Value = "" Then Sex = ContactSex.cscUnknown
    End Sub

    Public ReadOnly Property SurnameCapitalisedPrefix() As String
      Get
        Dim vSurname As String = mvClassFields.Item(ContactFields.Surname).Value
        If mvClassFields.Item(ContactFields.SurnamePrefix).Value.Length > 0 Then vSurname = UCase(Mid(mvClassFields.Item(ContactFields.SurnamePrefix).Value, 1, 1)) & Mid(mvClassFields.Item(ContactFields.SurnamePrefix).Value, 2) & " " & vSurname
        Return vSurname
      End Get
    End Property

    Public ReadOnly Property ContactPositionNumber() As Integer
      Get
        GetPositionInfo()
        Return mvContactPositionNumber
      End Get
    End Property

    Public Sub SetAddress(ByVal pAddressNumber As Integer, Optional ByVal pUpdateDefaultAddress As Boolean = False, Optional ByVal pAddress As Address = Nothing)
      If pAddress Is Nothing Then
        mvCurrentAddress = Nothing
        mvCurrentAddress = New Address(mvEnv)
        mvCurrentAddress.Init(pAddressNumber)
        SetPositionInvalid()
        If pUpdateDefaultAddress Then AddressNumber = pAddressNumber
      Else
        'Currently used in DataImportCMT
        mvCurrentAddress = pAddress
        If pUpdateDefaultAddress Then AddressNumber = pAddress.AddressNumber
      End If
    End Sub

    Private Sub SetPositionInvalid()
      mvPositionValid = False
      mvPosition = ""
      mvPositionLocation = ""
      mvOrganisationNumber = 0
      mvContactPositionNumber = 0
      mvOrganisationName = ""
      mvOrgContactNumber = 0
    End Sub

    Public ReadOnly Property JointContact1() As Contact
      Get
        Return mvJointContact1
      End Get
    End Property

    Public ReadOnly Property JointContact2() As Contact
      Get
        Return mvJointContact2
      End Get
    End Property

    Public ReadOnly Property NameAndAddress(Optional ByVal pAddressLine As Boolean = False) As String
      Get
        Dim vNameAndAddress As String = LabelName
        If ContactNumber <> OrganisationNumber Then 'Not dummy contact
          If Position.Length > 0 Then vNameAndAddress = vNameAndAddress & vbCrLf & Position
          If OrganisationName.Length > 0 Then vNameAndAddress = vNameAndAddress & vbCrLf & OrganisationName
        End If
        If pAddressLine Then
          NameAndAddress = vNameAndAddress & vbCrLf & Address.AddressLine
        Else
          NameAndAddress = vNameAndAddress & vbCrLf & Address.AddressMultiLine
        End If
      End Get
    End Property

    Public ReadOnly Property HasValidGiftAidDeclaration(Optional ByVal pMethod As GiftAidDeclaration.GiftAidDeclarationMethods = GiftAidDeclaration.GiftAidDeclarationMethods.gadmAny) As Boolean
      Get
        For Each vGAD As GiftAidDeclaration In GiftAidDeclarations
          If vGAD.IsValid Then
            If pMethod <> GiftAidDeclaration.GiftAidDeclarationMethods.gadmAny Then
              If vGAD.Method = pMethod Then Return True
            Else
              Return True
            End If
          End If
        Next vGAD
      End Get
    End Property

    Public WriteOnly Property NewTitleName() As String
      Set(ByVal Value As String)
        mvClassFields.Item(ContactFields.Title).Value = Value
        SetLabelName(False)
        SetSalutation()
      End Set
    End Property

    Private Sub SetSalutation()
      Dim vTitle As New Title
      vTitle.Init(mvEnv)
      If TitleName.Length > 0 Then vTitle.Init(mvEnv, TitleName)
      SetTitle(vTitle)
    End Sub

    Public Function GetContactImageFileName() As String
      Dim vDocumentName As String = ""
      Dim vCommsNumber As Integer = IntegerValue(mvEnv.Connection.GetValue("SELECT communications_log_number FROM communications_log WHERE contact_number = " & ContactNumber & " AND document_type = '" & mvEnv.GetConfig("cd_contact_image_document_type") & "'"))
      If vCommsNumber > 0 Then
        Dim vRS As CDBRecordSet = New SQLStatement(mvEnv.Connection, "document", "communications_log", New CDBField("communications_log_number", vCommsNumber)).GetRecordSet(CDBConnection.RecordSetOptions.NoDataTable)
        If vRS.Fetch = True Then
          vDocumentName = vRS.Fields(1).Value
        End If
      End If
      Return vDocumentName
    End Function

    Public Sub UnformatedName(ByVal pName As String, ByVal pTitlesDict As Hashtable, Optional ByVal pSurnameFirst As Boolean = False)
      'Takes a name and uses it to populate the fields
      Dim vRS As CDBRecordSet
      Dim vTitle As Title
      Dim vTitlesDict As New Hashtable

      Dim vChar As String
      Dim vCount As Integer
      Dim vIndex As Integer
      Dim vIndex2 As Integer
      Dim vLen As Integer
      Dim vSurnameFirstSet As Boolean
      Dim vTitleFound As Boolean
      Dim vTitleWords() As String
      Dim vWordCount As Integer

      Dim vWords(20) As String
      Dim vType(20) As Integer 'Hold the contact field for the word

      For vIndex = 0 To 19
        vWords(vIndex) = ""
      Next

      '-----------------------------------------
      ' First, initialise the titles Dictionary
      '-----------------------------------------
      If pTitlesDict Is Nothing Then
        vTitle = New Title
        vTitle.Init(mvEnv)
        vRS = mvEnv.Connection.GetRecordSet("SELECT DISTINCT " & vTitle.GetRecordSetFields(Title.TitleRecordSetTypes.trtAll) & " FROM titles ORDER BY title")
        With vRS
          While .Fetch()
            vTitleWords = Split(.Fields(1).Value, " ")
            For vIndex = LBound(vTitleWords) To UBound(vTitleWords)
              If Not vTitlesDict.ContainsKey(UCase(vTitleWords(vIndex))) Then
                vTitlesDict.Add(UCase(vTitleWords(vIndex)), UCase(vTitleWords(vIndex)))
              End If
            Next
          End While
          .CloseRecordSet()
        End With
      Else
        vTitlesDict = pTitlesDict
      End If

      '-----------------------------------
      ' Split the pName string into words
      '-----------------------------------
      vLen = Len(pName)
      vWordCount = 0
      For vIndex = 1 To vLen
        vChar = Mid(pName, vIndex, 1)
        If vChar = " " Or vChar = "." Then
          If vWords(vWordCount).Length > 0 Then vWordCount = vWordCount + 1
        ElseIf vChar = "(" Or vChar = ")" Then
          'Ignore
        Else
          vWords(vWordCount) = vWords(vWordCount) & vChar
        End If
      Next

      '----------------------------------------------------------
      ' Go through the words array assigning a type to each word
      '----------------------------------------------------------
      vIndex = 0
      vTitleFound = False
      While vIndex <= vWordCount
        If Len(vWords(vIndex)) > 1 Then
          '-------------------------------------------------------------------------
          ' More then one character in a word means not initials - probably surname
          '-------------------------------------------------------------------------
          If vTitlesDict.ContainsKey(UCase(vWords(vIndex))) Then
            If UCase(vWords(vIndex + 1)) = "OF" Then
              vWords(vIndex) = vWords(vIndex) & " of"
              vIndex2 = vIndex + 1 'Set up to get rid of it
              While vIndex2 < vWordCount
                vWords(vIndex2) = vWords(vIndex2 + 1)
                vIndex2 = vIndex2 + 1
              End While
              vWords(vIndex2) = ""
              vWordCount = vWordCount - 1
            End If

            If vTitleFound Then
              'We already have a Title so maybe this is not really a Title
              If pSurnameFirst Then
                vIndex2 = vIndex - 1
                If vType(vIndex2) = ContactFields.Title Then
                  If vIndex2 > 0 Then
                    vIndex2 = vIndex2 - 1
                    If vType(vIndex2) = ContactFields.Surname Then
                      'We have surname/title/?title - assume surname/forename/title
                      vType(vIndex2 + 1) = ContactFields.Forenames
                      vType(vIndex) = ContactFields.Title
                    End If
                  End If
                End If
              End If
              If vType(vIndex) = 0 Then
                'We have not been able to allocate this
                vCount = 0
                For vIndex2 = 0 To vIndex
                  If vType(vIndex2) = ContactFields.Forenames Then
                    vCount = vCount + 1
                  End If
                Next
                If vCount = 0 Then
                  'Forename not set, assume Forename
                  vType(vIndex) = ContactFields.Forenames
                Else
                  'Forename set, look for Surname
                  vCount = 0
                  For vIndex2 = 0 To vIndex
                    If vType(vIndex2) = ContactFields.Surname Then
                      vCount = vCount + 1
                    End If
                  Next
                  If vCount = 0 Then
                    'Forename set but Surname not set
                    vType(vIndex) = ContactFields.Surname
                  Else
                    'Forename & Surname set, assume Forename
                    vType(vIndex) = ContactFields.Forenames
                  End If
                End If
              End If

            Else
              If pSurnameFirst And vIndex = 0 Then
                'First word so assume surname
                vType(vIndex) = ContactFields.Surname
                vSurnameFirstSet = True
              Else
                vType(vIndex) = ContactFields.Title
                vTitleFound = True
              End If
            End If
            vIndex = vIndex + 1
          Else 'Word is not in titles table
            Select Case UCase(vWords(vIndex))
              Case "OF", "DE"
                vType(vIndex) = ContactFields.Surname
                vIndex = vIndex + 1
              Case "THE"
                vIndex2 = vIndex 'Get rid of it
                While vIndex2 < vWordCount
                  vWords(vIndex2) = vWords(vIndex2 + 1)
                  vIndex2 = vIndex2 + 1
                End While
                vWords(vIndex2) = ""
                vWordCount = vWordCount - 1
              Case Else
                If vIndex > 0 Then
                  vIndex2 = vIndex - 1
                  Do
                    If vType(vIndex2) = ContactFields.Surname Then
                      Select Case UCase(vWords(vIndex2))
                        Case "MC", "LE", "LA", "DER"
                          'Leave
                        Case "VAN"
                          If vType(vIndex2 + 1) = ContactFields.Forenames Then vType(vIndex2 + 1) = ContactFields.Surname
                        Case "OF", "DE"
                          If vType(vIndex2 + 1) = ContactFields.Forenames Then vType(vIndex2 + 1) = ContactFields.Surname
                          If vIndex2 > 0 Then vIndex2 = vIndex2 - 1 'If OF or DE then step over previous
                        Case Else
                          If vIndex2 > 0 Then
                            If vType(vIndex2 - 1) <> ContactFields.Initials Then
                              If pSurnameFirst Then
                                vSurnameFirstSet = True
                              End If
                            End If
                          Else
                            If vType(vIndex2) <> ContactFields.Initials Then
                              If pSurnameFirst Then
                                vSurnameFirstSet = True
                              End If
                            End If
                          End If
                          If vSurnameFirstSet Then vType(vIndex) = ContactFields.Forenames 'If Surname then make Forename
                      End Select
                    End If
                    vIndex2 = vIndex2 - 1
                  Loop While vIndex2 >= 0
                End If
                If pSurnameFirst Then
                  For vIndex2 = 0 To UBound(vType)
                    If vType(vIndex2) = ContactFields.Surname Then
                      vCount = vCount + 1
                    End If
                  Next
                  If vSurnameFirstSet Or vCount >= 2 Then
                    vType(vIndex) = ContactFields.Forenames
                  Else
                    vType(vIndex) = ContactFields.Surname
                  End If
                Else
                  'Set to Surname
                  vType(vIndex) = ContactFields.Surname
                  If vIndex = 1 And (vIndex <> vWordCount) Then
                    'But if we have just set Title then set this as Forenames instead
                    If vType(vIndex - 1) = ContactFields.Title Then vType(vIndex) = ContactFields.Forenames
                  End If
                End If
                vSurnameFirstSet = False
                vIndex = vIndex + 1
            End Select
          End If
        Else
          '----------------------------------------------------------------------------
          ' Only 1 character means initials - Concat with previous initials and remove
          '----------------------------------------------------------------------------
          If vIndex > 0 Then
            If vType(vIndex - 1) = ContactFields.Initials Then
              vWords(vIndex - 1) = vWords(vIndex - 1) & " " & vWords(vIndex)
              vIndex2 = vIndex
              While vIndex2 < vWordCount
                vWords(vIndex2) = vWords(vIndex2 + 1)
                vIndex2 = vIndex2 + 1
              End While
              vWords(vIndex2) = ""
              vWordCount = vWordCount - 1
            Else
              vType(vIndex) = ContactFields.Initials
              vIndex = vIndex + 1
            End If
          Else
            vType(vIndex) = ContactFields.Initials
            vIndex = vIndex + 1
          End If
        End If
      End While

      If Not (pSurnameFirst) Then
        If vType(0) = ContactFields.Surname Then
          If Len(TitleName) = 0 Then
            vType(0) = ContactFields.Forenames
          End If
        End If
      End If

      '-----------------------------------------------
      ' Now assign the words to the correct locations
      '-----------------------------------------------
      vIndex = 0
      While vIndex <= vWordCount
        Select Case vType(vIndex)
          Case ContactFields.Title
            If TitleName.Length > 0 Then
              If TitleName <> vWords(vIndex) Then TitleName = TitleName & " " & vWords(vIndex)
            Else
              TitleName = vWords(vIndex)
            End If
          Case ContactFields.Initials
            If Len(Initials) = 0 Then
              Initials = vWords(vIndex)
            Else
              Select Case vType(vIndex - 1)
                Case ContactFields.Forenames
                  Forenames = Forenames & " " & vWords(vIndex)
                Case ContactFields.Surname
                  Surname = Surname & " " & vWords(vIndex)
              End Select
            End If
          Case ContactFields.Forenames
            If Len(Forenames) > 0 Then
              Forenames = Forenames & " " & vWords(vIndex)
            Else
              Forenames = vWords(vIndex)
            End If
          Case ContactFields.Surname
            If Surname.Length > 0 Then
              Surname = Surname & " " & vWords(vIndex)
            Else
              Surname = vWords(vIndex)
            End If
        End Select
        vIndex = vIndex + 1
      End While

      If Not pSurnameFirst Then
        'If the Surname was not the first field and we have not set the Surname
        'then set the Forenames to be the Surname
        If Len(Surname) = 0 And Len(Forenames) > 0 Then
          Surname = Forenames
          Forenames = ""
        End If
      End If
    End Sub

    Public WriteOnly Property NewForenames() As String
      Set(ByVal pNewValue As String)
        Dim vPos As Integer
        Dim vLen As Integer

        If pNewValue <> Forenames Then
          pNewValue = CapitaliseWords(pNewValue, False, True)
          vLen = Len(pNewValue)
          vPos = InStr(pNewValue, "Mac")
          If vPos <> 0 And vLen > vPos + 3 Then
            vPos = vPos + 3
            Mid$(pNewValue, vPos, 1) = LCase$(Mid$(pNewValue, vPos, 1))
          End If
        End If
        mvClassFields.Item(ContactFields.Forenames).Value = pNewValue
        mvClassFields.Item(ContactFields.PreferredForename).Value = FirstWord(pNewValue)
        mvClassFields.Item(ContactFields.Initials).Value = SpacePadInitials(pNewValue)
        SetLabelName(False)
        SetSalutation()
      End Set
    End Property

    Public Function ForenameAndOrInitialsMatch(ByVal pRecordSet As CDBRecordSet) As Boolean
      'Used as part of deduplication process - class will already contain contacts details
      'Copied from DataImport.ContactOrOrg
      Dim vForeLen As Integer = Forenames.Length
      Dim vInitLen As Integer = Initials.Length
      Dim vForenames As String = pRecordSet.Fields("forenames").Value
      Dim vInitials As String = pRecordSet.Fields("initials").Value
      If UCase(Left(vForenames, vForeLen)) = UCase(Forenames) And UCase(Left(vInitials, vInitLen)) = UCase(Initials) Then
        Return True
      Else
        If UCase(Left(vInitials, vInitLen)) = UCase(Initials) Then
          Return True
        Else
          Return False
        End If
      End If
    End Function

    Public WriteOnly Property ImportContactNumber() As Integer
      Set(ByVal value As Integer)
        mvClassFields(ContactFields.ContactNumber).IntegerValue = value
      End Set
    End Property

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      mvClassFields.Item(ContactFields.Sex).Value = "U"
      mvClassFields.Item(ContactFields.ExDirectory).Value = "N"
      mvClassFields.Item(ContactFields.ContactType).Value = "C"
      mvClassFields.Item(ContactFields.DobEstimated).Value = "N"
      mvClassFields.Item(ContactFields.Department).Value = mvEnv.User.Department
    End Sub
    ''' <summary>
    ''' If changing a duplicate contacts address to the target address would result in an attempt to create a duplicate contact address record, then delete the duplicate.
    ''' </summary>
    ''' <param name="pJob"></param>
    ''' <param name="pConn"></param>
    ''' <param name="pAddTo">Target Address</param>
    ''' <param name="pAddFrom">Duplicate Address</param>
    ''' <remarks>address_number and contact_number are a unique combination on contact_addresses BR15454</remarks>
    Private Sub DeleteDuplicateContactAddress(ByRef pJob As JobSchedule, ByVal pConn As CDBConnection, ByVal pAddTo As Integer, ByVal pAddFrom As Integer)

      Dim vAddressNumbers As List(Of Integer) = New List(Of Integer) ' List of primary and duplicate contact numbers
      Dim vAddressLinksToDelete As List(Of Integer) = New List(Of Integer)

      Dim vSQLStatement As SQLStatement
      Dim vFields As String = "address_link_number,address_number,contact_number"
      Dim vTable As String = "contact_addresses"
      Dim vWhereFields As New CDBFields
      Dim vWhereFieldsDelete As New CDBFields
      Dim vDataTable As DataTable
      Dim vDuplicateAddresses() As DataRow

      vAddressNumbers.Add(pAddTo)
      vAddressNumbers.Add(pAddFrom)
      vWhereFields = New CDBFields(New CDBField("address_number", vAddressNumbers)) ' Create an IN Clause from the list
      ' Which records will be updated
      vSQLStatement = New SQLStatement(pConn, vFields, vTable, vWhereFields)
      vDataTable = vSQLStatement.GetDataTable
      vDuplicateAddresses = vDataTable.Select("address_number = " & pAddFrom.ToString)
      For Each vdr As DataRow In vDuplicateAddresses
        'check to see if changing the duplicates address number to pAdd would create a record that already exists on the contacts address table
        Dim vdrFindUpdatedDuplicateRow() As DataRow = vDataTable.Select("address_number=" & pAddTo.ToString & " AND contact_number=" & CInt(vdr("contact_number")).ToString)
        If vdrFindUpdatedDuplicateRow.Length = 1 Then
          ' Change would attempt to create a record that already exists so delete the duplicate record, there is an entry already present for this contact after a merge
          vAddressLinksToDelete.Add(CInt(vdrFindUpdatedDuplicateRow(0)("address_link_number")))
        End If
      Next
      If vAddressLinksToDelete.Count > 0 Then
        vWhereFieldsDelete = New CDBFields(New CDBField("address_link_number", vAddressLinksToDelete)) ' Create an IN Clause from the list
        pConn.DeleteRecords(vTable, vWhereFieldsDelete)
      End If
    End Sub
  End Class
End Namespace
