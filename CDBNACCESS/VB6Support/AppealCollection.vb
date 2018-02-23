Namespace Access

  Public Class AppealCollection

    Public Enum AppealCollectionType
      actHouseToHouse
      actUnmanned
      actManned
    End Enum

    Public Enum AppealCollectionCopyCriteria
      acccAlwaysInclude = 0
      acccCriteriaBased
      acccAlwaysExclude
    End Enum

    Public Enum AppealCollectionRecordSetTypes 'These are bit values
      apcrtAll = &H10S
      'ADD additional recordset types here
      apcrtHouseToHouse = &H100S
      apcrtUnmanned = &H200S
      apcrtManned = &H400S
      apcrtCollectionWithRegion = &H800S
      apcrtReconciliation = &H1000S
    End Enum

    Private mvH2HCollection As H2hCollection
    Private mvUnMannedCollection As UnmannedCollection
    Private mvMannedCollection As MannedCollection
    Private mvSource As Source
    Private mvSourceValuesSet As Boolean
    Private mvCollectionRegions As Collection

    Public Sub CreateMannedCollection(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      If pParams.Exists("Coordinator") Then
        pParams.Add("ContactNumber", CDBField.FieldTypes.cftInteger, pParams("Coordinator").Value)
      End If
      pParams.Add("CollectionType", CDBField.FieldTypes.cftCharacter, "M")
      Create(pParams)
      mvMannedCollection = New MannedCollection
      With mvMannedCollection
        .Init(mvEnv)
        .Create(pEnv, pParams)
      End With
      SetSourceValues(pParams)
    End Sub

    Public Sub CreateH2HCollection(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      pParams.Add("CollectionType", CDBField.FieldTypes.cftCharacter, "H")
      pParams.Add("ReadyForAcknowledgement", CDBField.FieldTypes.cftCharacter, "N")
      Create(pParams)
      mvH2HCollection = New H2hCollection
      With mvH2HCollection
        .Init(mvEnv)
        .Create(pEnv, pParams)
      End With
      SetSourceValues(pParams)
    End Sub

    Public Sub CreateUnMannedCollection(ByVal pEnv As CDBEnvironment, ByRef pParams As CDBParameters)
      pParams.Add("CollectionType", CDBField.FieldTypes.cftCharacter, "U")
      Dim vContactNumber As Integer
      'Save the value of contact number and assign the value of coordinator to contact number
      'to save it to the appeal collections table
      If pParams.Exists("Coordinator") And pParams.Exists("ContactNumber") Then
        vContactNumber = pParams("ContactNumber").IntegerValue
        pParams("ContactNumber").Value = pParams("Coordinator").Value
      End If

      Create(pParams)

      'Now assign the value back to contact number to save the value to unmanned collection table
      If pParams.Exists("Coordinator") And pParams.Exists("ContactNumber") Then pParams("ContactNumber").Value = vContactNumber.ToString

      mvUnMannedCollection = New UnmannedCollection
      With mvUnMannedCollection
        .Init(mvEnv)
        .Create(pEnv, pParams)
      End With
      SetSourceValues(pParams)
    End Sub

    Public ReadOnly Property H2hCollection() As H2hCollection
      Get
        If mvH2HCollection Is Nothing Then
          mvH2HCollection = New H2hCollection
          mvH2HCollection.Init(mvEnv, CollectionNumber)
        End If
        H2hCollection = mvH2HCollection
      End Get
    End Property

    Public ReadOnly Property MannedCollection() As MannedCollection
      Get
        If mvMannedCollection Is Nothing Then
          mvMannedCollection = New MannedCollection
          mvMannedCollection.Init(mvEnv, CollectionNumber)
        End If
        MannedCollection = mvMannedCollection
      End Get
    End Property

    Public ReadOnly Property UnmannedCollection() As UnmannedCollection
      Get
        If mvUnMannedCollection Is Nothing Then
          mvUnMannedCollection = New UnmannedCollection
          mvUnMannedCollection.Init(mvEnv, CollectionNumber)
        End If
        UnmannedCollection = mvUnMannedCollection
      End Get
    End Property

    Public Function DeriveCollectionSource(Optional ByVal pCampaign As String = "", Optional ByVal pAppeal As String = "", Optional ByVal pCollection As String = "") As String
      Dim vSource As String
      Dim vSum As Integer
      Dim vCounter As Integer
      Dim vTimesThru As Integer
      Dim vChar As String
      Dim vValue As Integer
      Dim vRemainder As Integer
      'should put the check digit bit in a common place for segments and collections

      If Len(pCampaign) = 0 Then pCampaign = Campaign
      If Len(pAppeal) = 0 Then pAppeal = Appeal
      If Len(pCollection) = 0 Then pCollection = Collection

      If Len(pCampaign) = 3 And Len(pAppeal) = 3 And Len(pCollection) = 3 Then
        vSource = pCampaign & pAppeal & pCollection
        If mvEnv.GetConfigOption("ma_check_digit_on_source", True) Then
          vSum = 0
          vCounter = 1
          vTimesThru = 1
          While vCounter <= Len(vSource)
            vChar = Mid(vSource, vCounter, 1)
            If vChar Like "[A-Z]" Then
              'Convert chars to numbers A=10, B=11, etc.  (Asc("A") = 65)
              vValue = Asc(vChar) - 55
            Else
              vValue = IntegerValue(vChar)
            End If
            Select Case vTimesThru
              Case 1
                vSum = vSum + (vValue * 1)
              Case 2
                vSum = vSum + (vValue * 3)
              Case 3
                vSum = vSum + (vValue * 5)
              Case 4
                vSum = vSum + (vValue * 7)
                vTimesThru = 0
            End Select
            vCounter = vCounter + 1
            vTimesThru = vTimesThru + 1
          End While
          vRemainder = vSum Mod 10
          vSource = vSource & vRemainder.ToString("0")
        End If
      Else
        vSource = ""
      End If
      Return vSource
    End Function

    Public Overloads Function GetRecordSetFields(ByVal pRSType As AppealCollectionRecordSetTypes) As String
      Dim vFields As String = ""
      Dim vCollRegion As CollectionRegion

      'Modify below to add each recordset type as required
      If mvClassFields Is Nothing Then InitClassFields()
      If (pRSType And AppealCollectionRecordSetTypes.apcrtAll) <> 0 Then
        vFields = mvClassFields.FieldNames(mvEnv, "ac")
        vFields = Replace(vFields, "ac.contact_number", "ac.contact_number AS coordinator")
      End If
      If (pRSType And AppealCollectionRecordSetTypes.apcrtHouseToHouse) <> 0 Then
        If mvH2HCollection Is Nothing Then
          mvH2HCollection = New H2hCollection
          mvH2HCollection.Init(mvEnv)
        End If
        vFields = vFields & "," & mvH2HCollection.GetRecordSetFields(H2hCollection.H2hCollectionRecordSetTypes.hcrtAll)
      End If
      If (pRSType And AppealCollectionRecordSetTypes.apcrtUnmanned) = AppealCollectionRecordSetTypes.apcrtUnmanned Then
        If mvUnMannedCollection Is Nothing Then
          mvUnMannedCollection = New UnmannedCollection
          mvUnMannedCollection.Init(mvEnv)
        End If
        vFields = vFields & "," & mvUnMannedCollection.GetRecordSetFields(UnmannedCollection.UnmannedCollectionRecordSetTypes.ucrtAll)
      End If
      If (pRSType And AppealCollectionRecordSetTypes.apcrtManned) = AppealCollectionRecordSetTypes.apcrtManned Then
        If mvMannedCollection Is Nothing Then
          mvMannedCollection = New MannedCollection
          mvMannedCollection.Init(mvEnv)
        End If
        vFields = vFields & "," & mvMannedCollection.GetRecordSetFields(MannedCollection.MannedCollectionRecordSetTypes.mcrtAll)
      End If
      If (pRSType And AppealCollectionRecordSetTypes.apcrtCollectionWithRegion) <> 0 Then
        vCollRegion = New CollectionRegion
        vCollRegion.Init(mvEnv)
        vFields = vFields & ", " & Replace(vCollRegion.GetRecordSetFields(CollectionRegion.CollectionRegionRecordSetTypes.crertAll Or CollectionRegion.CollectionRegionRecordSetTypes.crertAllPlusPoints), "cr.collection_number,", "")
      End If

      If (pRSType And AppealCollectionRecordSetTypes.apcrtReconciliation) <> 0 Then
        If Len(vFields) > 0 Then vFields = vFields & ","
        vFields = vFields & "ac.collection_number,collection_type,ac.source,ac.product,ac.rate, ac.contact_number AS coordinator"
      End If
      GetRecordSetFields = vFields
    End Function

    Public Overrides Sub Init(ByVal pPrimaryKeyValue As Integer)
      MyBase.Init(pPrimaryKeyValue)
      Select Case CollectionType
        Case AppealCollectionType.actHouseToHouse
          mvH2HCollection = New H2hCollection
          mvH2HCollection.Init(mvEnv, CollectionNumber)
        Case AppealCollectionType.actUnmanned
          mvUnMannedCollection = New UnmannedCollection
          mvUnMannedCollection.Init(mvEnv, CollectionNumber)
        Case AppealCollectionType.actManned
          mvMannedCollection = New MannedCollection
          mvMannedCollection.Init(mvEnv, CollectionNumber)
      End Select
    End Sub

    Public Overloads Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As AppealCollectionRecordSetTypes)
      Dim vFields As CDBFields
      Dim vCollRegion As CollectionRegion

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(AppealCollectionFields.CollectionNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And AppealCollectionRecordSetTypes.apcrtAll) = AppealCollectionRecordSetTypes.apcrtAll Then
          .SetItem(AppealCollectionFields.Campaign, vFields)
          .SetItem(AppealCollectionFields.Appeal, vFields)
          .SetItem(AppealCollectionFields.Collection, vFields)
          .SetItem(AppealCollectionFields.CollectionDesc, vFields)
          .SetItem(AppealCollectionFields.CollectionType, vFields)
          '.Item(acfContactNumber).Value = vFields("coordinator").Value
          .Item(AppealCollectionFields.ContactNumber).SetValue = vFields("coordinator").Value
          '      .SetItem acfContactNumber, vFields
          .SetItem(AppealCollectionFields.TargetCollectors, vFields)
          .SetItem(AppealCollectionFields.TargetIncome, vFields)
          .SetItem(AppealCollectionFields.ActualCollectors, vFields)
          .SetItem(AppealCollectionFields.ActualIncome, vFields)
          .SetItem(AppealCollectionFields.Source, vFields)
          .SetItem(AppealCollectionFields.Product, vFields)
          .SetItem(AppealCollectionFields.Rate, vFields)
          .SetItem(AppealCollectionFields.BankAccount, vFields)
          .SetItem(AppealCollectionFields.CopyCriteria, vFields)
          .SetItem(AppealCollectionFields.ReadyForConfirmation, vFields)
          .SetItem(AppealCollectionFields.ReadyForLabels, vFields)
          .SetItem(AppealCollectionFields.ReadyForAcknowledgement, vFields)
          .SetItem(AppealCollectionFields.ConfirmationProducedOn, vFields)
          .SetItem(AppealCollectionFields.LabelsProducedOn, vFields)
          .SetItem(AppealCollectionFields.ResourcesProducedOn, vFields)
          .SetItem(AppealCollectionFields.EndOfCollectionProducedOn, vFields)
          .SetItem(AppealCollectionFields.AcknowledgementProducedOn, vFields)
          .SetItem(AppealCollectionFields.ReminderProducedOn, vFields)
          .SetItem(AppealCollectionFields.ActualCollectorsDate, vFields)
          .SetItem(AppealCollectionFields.ActualIncomeDate, vFields)
          .SetItem(AppealCollectionFields.Notes, vFields)
          .SetItem(AppealCollectionFields.AmendedBy, vFields)
          .SetItem(AppealCollectionFields.AmendedOn, vFields)
          .SetOptionalItem(AppealCollectionFields.TotalItemisedCost, vFields)
        End If

        If (pRSType And AppealCollectionRecordSetTypes.apcrtHouseToHouse) = AppealCollectionRecordSetTypes.apcrtHouseToHouse Then
          mvH2HCollection = New H2hCollection
          mvH2HCollection.InitFromRecordSet(mvEnv, pRecordSet, H2hCollection.H2hCollectionRecordSetTypes.hcrtAll)
        End If
        If (pRSType And AppealCollectionRecordSetTypes.apcrtManned) = AppealCollectionRecordSetTypes.apcrtManned Then
          mvMannedCollection = New MannedCollection
          mvMannedCollection.InitFromRecordSet(mvEnv, pRecordSet, MannedCollection.MannedCollectionRecordSetTypes.mcrtAll)
        End If
        If (pRSType And AppealCollectionRecordSetTypes.apcrtUnmanned) = AppealCollectionRecordSetTypes.apcrtUnmanned Then
          mvUnMannedCollection = New UnmannedCollection
          mvUnMannedCollection.InitFromRecordSet(mvEnv, pRecordSet, UnmannedCollection.UnmannedCollectionRecordSetTypes.ucrtAll)
        End If
        If (pRSType And AppealCollectionRecordSetTypes.apcrtCollectionWithRegion) = AppealCollectionRecordSetTypes.apcrtCollectionWithRegion Then
          'This will contain multiple records
          mvCollectionRegions = New Collection
          While (mvClassFields.Item(AppealCollectionFields.CollectionNumber).IntegerValue = pRecordSet.Fields("collection_number").IntegerValue) And pRecordSet.Status() = True
            If pRecordSet.Fields("collection_region_number").IntegerValue > 0 Then
              vCollRegion = New CollectionRegion
              vCollRegion.InitFromRecordSet(mvEnv, pRecordSet, CollectionRegion.CollectionRegionRecordSetTypes.crertAll Or CollectionRegion.CollectionRegionRecordSetTypes.crertAllPlusPoints)
              mvCollectionRegions.Add(vCollRegion, CStr(vCollRegion.CollectionRegionNumber))
            Else
              pRecordSet.Fetch()
            End If
          End While
        End If
        If (pRSType And AppealCollectionRecordSetTypes.apcrtReconciliation) > 0 Then
          .SetItem(AppealCollectionFields.CollectionType, vFields)
          .Item(AppealCollectionFields.ContactNumber).SetValue = vFields("coordinator").Value
          .SetItem(AppealCollectionFields.Source, vFields)
          .SetItem(AppealCollectionFields.Product, vFields)
          .SetItem(AppealCollectionFields.Rate, vFields)
        End If
      End With
    End Sub

    Public Overloads Sub Clone(ByVal pEnv As CDBEnvironment, ByVal pSrcAppealCollection As AppealCollection, ByVal pTgtCampaignCode As String, ByVal pTgtAppealCode As String, ByVal pTgtCollectionCode As String, Optional ByRef pTgtCollectionDesc As String = "")
      Dim vCollPoint As CollectionPoint
      Dim vSrcCollRegion As CollectionRegion
      Dim vTgtCollRegion As CollectionRegion
      Dim vTrans As Boolean

      Init()
      'Clone the AppealCollection
      With mvClassFields
        'Mandatory
        .Item(AppealCollectionFields.Campaign).Value = pTgtCampaignCode
        .Item(AppealCollectionFields.Appeal).Value = pTgtAppealCode
        .Item(AppealCollectionFields.Collection).Value = pTgtCollectionCode
        If Len(pTgtCollectionDesc) > 0 Then
          .Item(AppealCollectionFields.CollectionDesc).Value = pTgtCollectionDesc
        Else
          .Item(AppealCollectionFields.CollectionDesc).Value = pSrcAppealCollection.CollectionDesc
        End If
        .Item(AppealCollectionFields.CollectionType).Value = pSrcAppealCollection.CollectionTypeCode
        .Item(AppealCollectionFields.TargetIncome).Value = CStr(0)
        If pSrcAppealCollection.CollectionType = AppealCollectionType.actHouseToHouse Or pSrcAppealCollection.CollectionType = AppealCollectionType.actManned Then
          '.Item(acfActualCollectors).Value = 0
        Else
          If pSrcAppealCollection.ActualCollectors > 0 Then .Item(AppealCollectionFields.ActualCollectors).Value = CStr(pSrcAppealCollection.ActualCollectors)
        End If
        .Item(AppealCollectionFields.ActualIncome).Value = CStr(0)
        .Item(AppealCollectionFields.Source).Value = pSrcAppealCollection.Source
        .Item(AppealCollectionFields.Product).Value = pSrcAppealCollection.ProductCode
        .Item(AppealCollectionFields.Rate).Value = pSrcAppealCollection.RateCode
        .Item(AppealCollectionFields.BankAccount).Value = pSrcAppealCollection.BankAccount
        .Item(AppealCollectionFields.CopyCriteria).Value = pSrcAppealCollection.CopyCriteriaCode
        .Item(AppealCollectionFields.ReadyForConfirmation).Bool = True
        .Item(AppealCollectionFields.ReadyForLabels).Bool = True
        .Item(AppealCollectionFields.ReadyForAcknowledgement).Bool = True
        .Item(AppealCollectionFields.Notes).Value = pSrcAppealCollection.Notes
        'Optional
        If pSrcAppealCollection.ContactNumber > 0 Then .Item(AppealCollectionFields.ContactNumber).Value = CStr(pSrcAppealCollection.ContactNumber)
        If pSrcAppealCollection.TargetCollectors > 0 Then .Item(AppealCollectionFields.TargetCollectors).Value = CStr(pSrcAppealCollection.TargetCollectors)
        'If pSrcAppealCollection.TotalItemisedCost > 0 Then .Item(acfTotalItemisedCost).Value = pSrcAppealCollection.TotalItemisedCost
      End With

      'Clone the specific Collection
      Select Case pSrcAppealCollection.CollectionType
        Case AppealCollectionType.actHouseToHouse
          mvH2HCollection = New H2hCollection
          mvH2HCollection.Clone(pEnv, pSrcAppealCollection.H2hCollection)
        Case AppealCollectionType.actManned
          mvMannedCollection = New MannedCollection
          mvMannedCollection.Clone(pEnv, pSrcAppealCollection.MannedCollection)
        Case AppealCollectionType.actUnmanned
          mvUnMannedCollection = New UnmannedCollection
          mvUnMannedCollection.Clone(pEnv, pSrcAppealCollection.UnmannedCollection)
      End Select
      SetValid()
      'Clone the CollectionRegions & CollectionPoints
      If pSrcAppealCollection.CollectionType = AppealCollectionType.actHouseToHouse Or pSrcAppealCollection.CollectionType = AppealCollectionType.actManned Then
        mvCollectionRegions = New Collection
        For Each vSrcCollRegion In pSrcAppealCollection.CollectionRegions
          vTgtCollRegion = New CollectionRegion
          vTgtCollRegion.Clone(pEnv, vSrcCollRegion, mvClassFields.Item(AppealCollectionFields.CollectionNumber).IntegerValue)
          mvCollectionRegions.Add(vTgtCollRegion)
        Next vSrcCollRegion
      End If
      'Clone the CollectionResources
      If pSrcAppealCollection.CollectionType = AppealCollectionType.actManned Or pSrcAppealCollection.CollectionType = AppealCollectionType.actUnmanned Then
        'THIS IS NOT BEING IMPLEMENTED YET BUT MAY BE DONE AT A LATER DATE
      End If

      'Now go and save everything
      If Not pEnv.Connection.InTransaction Then
        pEnv.Connection.StartTransaction()
        vTrans = True
      End If

      'Save the AppealCollection and Save the specific Collection
      Save()

      'Save CollectionRegions & CollectionPoints
      If pSrcAppealCollection.CollectionType = AppealCollectionType.actHouseToHouse Or pSrcAppealCollection.CollectionType = AppealCollectionType.actManned Then
        For Each vTgtCollRegion In mvCollectionRegions
          vTgtCollRegion.Save()
          For Each vCollPoint In vTgtCollRegion.CollectionPoints
            vCollPoint.Save()
          Next vCollPoint
        Next vTgtCollRegion
      End If
      'Save CollectionResources
      If pSrcAppealCollection.CollectionType = AppealCollectionType.actManned Or pSrcAppealCollection.CollectionType = AppealCollectionType.actUnmanned Then
        'THIS IS NOT BEING IMPLEMENTED YET BUT MAY BE DONE AT A LATER DATE
      End If
      If vTrans Then mvEnv.Connection.CommitTransaction()
    End Sub

    Public ReadOnly Property CollectionRegions() As Collection
      Get
        CollectionRegions = mvCollectionRegions
      End Get
    End Property

    Public ReadOnly Property CopyCriteria() As AppealCollectionCopyCriteria
      Get
        Select Case CopyCriteriaCode
          Case "C"
            Return AppealCollectionCopyCriteria.acccCriteriaBased
          Case "E"
            Return AppealCollectionCopyCriteria.acccAlwaysExclude
          Case Else
            Return AppealCollectionCopyCriteria.acccAlwaysInclude
        End Select
      End Get
    End Property

    Private Sub SetSourceValues(ByRef pParams As CDBParameters)

      With mvSource
        If Not mvEnv.GetConfigOption("default_analysis_from_source") Then
          'Source field is mandatory and can not be null
          If Len(pParams.ParameterExists("Source").Value) = 0 Then RaiseError(DataAccessErrors.daeParameterNotFound, "Source")
        End If
        If Len(pParams.ParameterExists("Source").Value) = 0 Then
          If pParams.Exists("Source") Then
            pParams("Source").Value = DeriveCollectionSource()
          Else
            pParams.Add("Source", CDBField.FieldTypes.cftCharacter, DeriveCollectionSource)
          End If
          mvClassFields(AppealCollectionFields.Source).Value = pParams("Source").Value
          If pParams.Exists("SourceDesc") = False Then pParams.Add("SourceDesc")
          If Len(pParams("SourceDesc").Value) = 0 Then pParams("SourceDesc").Value = Left(CollectionDesc, 30)
        Else
          mvClassFields(AppealCollectionFields.Source).Value = pParams("Source").Value
        End If
        If mvSource Is Nothing Then mvSource = New Source
        mvSource.Init(mvEnv, (pParams("Source").Value))
        If mvSource.Existing Then
          mvSource.Update(pParams)
        Else
          mvSource.Create(mvEnv, pParams)
        End If
      End With
      mvSourceValuesSet = True
    End Sub

    Public Sub UpdateCollection(ByRef pParams As CDBParameters)
      Dim vContactNumber As Integer
      If pParams.Exists("Coordinator") And CollectionType = AppealCollectionType.actManned Then
        mvClassFields(AppealCollectionFields.ContactNumber).Value = pParams("Coordinator").Value
      End If

      'Save the value of contact number and assign the value of coordinator to contact number
      'to save it to the appeal collections table
      If pParams.Exists("Coordinator") And pParams.Exists("ContactNumber") And CollectionType = AppealCollectionType.actUnmanned Then
        vContactNumber = pParams("ContactNumber").IntegerValue
        pParams("ContactNumber").Value = pParams("Coordinator").Value
      End If

      Update(pParams)

      'Now assign the value back to contact number to save the value to unmanned collection table
      If pParams.Exists("Coordinator") And pParams.Exists("ContactNumber") And CollectionType = AppealCollectionType.actUnmanned Then pParams("ContactNumber").Value = vContactNumber.ToString

      Select Case CollectionType
        Case AppealCollectionType.actHouseToHouse
          mvH2HCollection.Update(pParams)
        Case AppealCollectionType.actUnmanned
          mvUnMannedCollection.Update(pParams)
        Case AppealCollectionType.actManned
          mvMannedCollection.Update(pParams)
      End Select
      SetSourceValues(pParams)
    End Sub
    Public Overrides ReadOnly Property DataTable() As CDBDataTable
      Get
        Dim vTable As New CDBDataTable
        Dim vLastColumn As Integer
        Dim vClassField As ClassField
        Dim vRow As CDBDataRow

        With vTable
          For Each vClassField In mvClassFields
            If vClassField.Name = "contact_number" Then
              .AddColumn("Coordinator", (vClassField.FieldType))
            Else
              .AddColumn(vClassField.ProperName, (vClassField.FieldType))
            End If
          Next vClassField
          vRow = .AddRow
          For vIndex As Integer = 0 To mvClassFields.Count - 1
            vRow.Item(vIndex + 1) = mvClassFields(vIndex + 1).Value
          Next
        End With
        vLastColumn = vTable.Columns.Count()
        With vTable.Rows.Item(0)
          Select Case CollectionType
            Case AppealCollectionType.actHouseToHouse
              vTable.AddColumnsFromList("StartDate,EndDate")
              .Item(vLastColumn + 1) = mvH2HCollection.StartDate
              .Item(vLastColumn + 2) = mvH2HCollection.EndDate
            Case AppealCollectionType.actUnmanned
              vTable.AddColumnsFromList("ContactNumber,OrganisationNumber,AddressNumber,StartDate,EndDate")
              .Item(vLastColumn + 1) = mvUnMannedCollection.ContactNumber.ToString
              .Item(vLastColumn + 2) = mvUnMannedCollection.OrganisationNumber.ToString
              .Item(vLastColumn + 3) = mvUnMannedCollection.AddressNumber.ToString
              .Item(vLastColumn + 4) = mvUnMannedCollection.StartDate
              .Item(vLastColumn + 5) = mvUnMannedCollection.EndDate
            Case AppealCollectionType.actManned
              vTable.AddColumnsFromList("OrganisationNumber,CollectionDate,StartTime,EndTime,MeetingPoint")
              .Item(vLastColumn + 1) = mvMannedCollection.OrganisationNumber.ToString
              .Item(vLastColumn + 2) = mvMannedCollection.CollectionDate
              .Item(vLastColumn + 3) = mvMannedCollection.StartTime
              .Item(vLastColumn + 4) = mvMannedCollection.EndTime
              .Item(vLastColumn + 5) = mvMannedCollection.MeetingPoint
          End Select
        End With
        Return vTable
      End Get
    End Property
    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      SetValid()
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
      Select Case CollectionType
        Case AppealCollectionType.actHouseToHouse
          If Not mvH2HCollection Is Nothing Then mvH2HCollection.Save(pAudit, CollectionNumber)
        Case AppealCollectionType.actUnmanned
          If Not mvUnMannedCollection Is Nothing Then mvUnMannedCollection.Save(pAudit, CollectionNumber)
        Case AppealCollectionType.actManned
          If Not mvMannedCollection Is Nothing Then mvMannedCollection.Save(pAudit, CollectionNumber)
      End Select
      If Not mvSource Is Nothing Then mvSource.Save(pAmendedBy, pAudit)
    End Sub
  End Class

End Namespace