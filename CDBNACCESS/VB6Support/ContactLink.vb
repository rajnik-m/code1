Imports System.Linq
Imports Advanced.LanguageExtensions

Namespace Access

  Partial Public Class ContactLink

    Public Enum ContactLinkTypes
      cltContact = 1
      cltOrganisation
    End Enum

    Public Overloads Sub Init(ByVal pEnv As CDBEnvironment, ByVal pLinkType As ContactLinkTypes, Optional ByVal pNumber1 As Integer = 0, Optional ByVal pNumber2 As Integer = 0, Optional ByVal pRelationship As String = "", Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "", Optional ByVal pOldValidFrom As String = "", Optional ByVal pOldValidTo As String = "")
      mvEnv = pEnv
      InitClassFields()
      SetLinkType(pLinkType)
      If (pNumber1 > 0 And pNumber2 > 0 And pRelationship.Length > 0) Then
        Dim vWhereFields As New CDBFields
        vWhereFields.Add(mvClassFields.Item(ContactLinkFields.ContactNumber1).Name, pNumber1)
        vWhereFields.Add(mvClassFields.Item(ContactLinkFields.ContactNumber2).Name, pNumber2)
        vWhereFields.Add(mvClassFields.Item(ContactLinkFields.Relationship).Name, pRelationship)
        If pValidFrom.Length > 0 Then vWhereFields.Add(mvClassFields.Item(ContactLinkFields.ValidFrom).Name, CDBField.FieldTypes.cftDate, pValidFrom)
        If pValidTo.Length > 0 Then vWhereFields.Add(mvClassFields.Item(ContactLinkFields.ValidTo).Name, CDBField.FieldTypes.cftDate, pValidTo)
        MyBase.InitWithPrimaryKey(vWhereFields)
      Else
        Init()
      End If
      If pOldValidFrom.Length > 0 Then mvOldValidFrom = pOldValidFrom
      If pOldValidTo.Length > 0 Then mvOldValidTo = pOldValidTo
    End Sub

    Private Sub SetLinkType(ByVal pLinkType As ContactLinkTypes)
      If pLinkType = ContactLinkTypes.cltOrganisation Then
        mvClassFields(ContactLinkFields.ContactNumber1).SetName("organisation_number_1")
        mvClassFields(ContactLinkFields.ContactNumber2).SetName("organisation_number_2")
        mvClassFields(ContactLinkFields.ContactLinkNumber).SetName("organisation_link_number")
        mvClassFields.DatabaseTableName = "organisation_links"
        Me.LinkType = pLinkType
      End If
    End Sub

    Public Sub InitNew(ByVal pEnv As CDBEnvironment, ByVal pLinkType As ContactLinkTypes, ByVal pNumber1 As Integer, ByVal pNumber2 As Integer, ByVal pRelationship As String, Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "", Optional ByVal pNotes As String = "", Optional ByVal pRelationShipStatus As String = "")
      mvEnv = pEnv
      SetLinkType(pLinkType)
      InitClassFields()
      mvClassFields.Item(ContactLinkFields.ContactNumber1).IntegerValue = pNumber1
      mvClassFields.Item(ContactLinkFields.ContactNumber2).IntegerValue = pNumber2
      mvClassFields.Item(ContactLinkFields.Relationship).Value = pRelationship
      mvClassFields.Item(ContactLinkFields.ValidFrom).Value = pValidFrom
      mvClassFields.Item(ContactLinkFields.ValidTo).Value = pValidTo
      mvClassFields.Item(ContactLinkFields.Historical).Bool = False
      If IsDate(pValidTo) Then
        If CDate(pValidTo) < Today Then mvClassFields.Item(ContactLinkFields.Historical).Bool = True
      End If
      mvClassFields.Item(ContactLinkFields.Notes).Value = pNotes
      mvClassFields.Item(ContactLinkFields.RelationshipStatus).Value = pRelationShipStatus
    End Sub

    Public Sub ChangeRelationship(ByRef pRelationshipCode As String)
      If pRelationshipCode <> RelationshipCode Then
        mvClassFields.Item(ContactLinkFields.Relationship).Value = pRelationshipCode
        mvClassFields.VerifyUnique(mvEnv.Connection)
      End If
    End Sub

    Public Overloads Sub Update(ByVal pNumber1 As Integer, ByVal pNumber2 As Integer, ByVal pRelationship As String, Optional ByVal pValidFrom As String = "", Optional ByVal pValidTo As String = "", Optional ByVal pNotes As String = "", Optional ByVal pRelationshipStatus As String = "")
      mvClassFields.Item(ContactLinkFields.ContactNumber1).IntegerValue = pNumber1
      mvClassFields.Item(ContactLinkFields.ContactNumber2).IntegerValue = pNumber2
      mvClassFields.Item(ContactLinkFields.Relationship).Value = pRelationship
      mvClassFields.Item(ContactLinkFields.ValidFrom).Value = pValidFrom
      mvClassFields.Item(ContactLinkFields.ValidTo).Value = pValidTo
      mvClassFields.Item(ContactLinkFields.Historical).Bool = False
      If IsDate(pValidTo) Then
        If CDate(pValidTo) < Today Then mvClassFields.Item(ContactLinkFields.Historical).Bool = True
      End If
      mvClassFields.Item(ContactLinkFields.Notes).Value = pNotes
      mvClassFields.Item(ContactLinkFields.RelationshipStatus).Value = pRelationshipStatus
    End Sub

    Private mvRelationship As Relationship
    Private mvLinkType As ContactLinkTypes = ContactLinkTypes.cltContact
    Private mvCompLinkType As ContactLinkTypes = ContactLinkTypes.cltContact
    Private mvCompLinkTypeSet As Boolean

    Protected Overrides Sub ClearFields()
      mvRelationship = Nothing
      mvCompLinkTypeSet = False
    End Sub

    Public ReadOnly Property Relationship As Relationship
      Get
        If mvRelationship Is Nothing AndAlso Me.RelationshipCode.HasValue Then
          mvRelationship = Me.GetRelatedInstance(Of Relationship)({ContactLinkFields.Relationship})
        End If
        Return mvRelationship
      End Get
    End Property

    Public Overloads Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pCheckSynch As Boolean)
      Dim vTransactionStarted As Boolean

      If pCheckSynch = True AndAlso mvEnv.Connection.InTransaction = False Then
        mvEnv.Connection.StartTransaction()
        vTransactionStarted = True
      End If
      MyBase.Delete(pAmendedBy, pAudit, 0)
      mvExisting = False 'bug in CARERecord - this needs to move to CARERecord. Needs to be fixed immediately after a code-split to maximise regression test.
      If pCheckSynch = True AndAlso
          Me.Relationship.AutoCreateComplementaryFlag = True AndAlso
              Me.ComplementaryLink IsNot Nothing AndAlso
                  Me.ComplementaryLink.Existing Then
        'Synchronised relationship needs to be deleted too.
        Me.ComplementaryLink.Delete(pAmendedBy, pAudit, False)
      End If
      If vTransactionStarted = True Then
        mvEnv.Connection.CommitTransaction()
      End If
    End Sub

    Public Overrides Sub Delete(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      Delete(pAmendedBy, pAudit, True)
    End Sub

    Public Overloads Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pCheckSynch As Boolean)
      SaveInternal(pAmendedBy, pAudit, 0, pCheckSynch, False)
    End Sub

    Public Overloads Sub Save(ByVal pAmendedBy As String, pAudit As Boolean)
      SaveInternal(pAmendedBy, pAudit, 0, True, Nothing)
    End Sub

    Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
      SaveInternal(pAmendedBy, pAudit, pJournalNumber, True, Nothing)
    End Sub

    Public Overrides Sub Save(pAmendedBy As String, pAudit As Boolean, pJournalNumber As Integer, pForceAmendmentHistory As Boolean)
      SaveInternal(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory, True)
    End Sub

    ''' <summary>
    ''' Call the base class Save method and checks if a complementary relationship needs to be created
    ''' </summary>
    ''' <param name="pAmendedBy"></param>
    ''' <param name="pAudit"></param>
    ''' <param name="pJournalNumber"></param>
    Private Sub SaveInternal(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer, pCheckSynchronisation As Boolean, pForceAmendmentHistory As Boolean?)
      Dim vCreateTransaction As Boolean = Not Me.Environment.Connection.InTransaction

      If vCreateTransaction Then
        Me.Environment.Connection.StartTransaction()
      End If

      CheckOverlaps() 'ensure there isn't already a Link record with the same dates.

      'We need to check synchronisation with complementary before the save, because we must only synchronise if the complementary was synchronised with the old values
      Dim vRequiresSynchronisation As Boolean = pCheckSynchronisation AndAlso RequiresComplementarySync() 'nb pCheckSynchronisation is set to false when the method is saving the complementary link

      'Base functionality here.  Overloaded methods in the base class are completely separate instead of one calling the other. Don't ask
      If pForceAmendmentHistory.HasValue Then
        MyBase.Save(pAmendedBy, pAudit, pJournalNumber, pForceAmendmentHistory.Value)
      Else
        MyBase.Save(pAmendedBy, pAudit, pJournalNumber)
      End If

      If vRequiresSynchronisation Then
        DoSyncComplementaryRelationship() 'sync dates and status between links.  Will work in add as well as update mode.
        If Me.ComplementaryLink IsNot Nothing AndAlso Me.ComplementaryLink.IsDirty Then
          Me.ComplementaryLink.SaveInternal(pAmendedBy, pAudit, pJournalNumber, False, False)
        End If
      End If

      If vCreateTransaction Then
        Me.Environment.Connection.CommitTransaction()
      End If
    End Sub

    Private Function RequiresComplementarySync() As Boolean
      Dim vRequiresSync As Boolean = False

      'Condition 1: safeguard against a potential null evaluation
      'Condition 2: Both current and complementary must be in the same CRUD state
      'Condition 3: sync is always in update mode, but it is only allowed in Add mode if it is set to automatically synch
      Dim vAllConditionsMet As Boolean =
          Me.ComplementaryLink IsNot Nothing AndAlso
              Me.Existing = Me.ComplementaryLink.Existing AndAlso
                  (Me.Existing = True OrElse Me.Relationship.CanAutoCreateComplementaryRelationship)

      If vAllConditionsMet Then
        'the only fields that get synchronised are the valid from, to and relationship status.
        'The records will be synched if any of the values in one record have changed and the values in the other record were previously synchronised
        'or the current record is new.
        Dim vSynchedValuesChanged As Boolean =
            Me.Existing = False OrElse
                mvClassFields(ContactLink.ContactLinkFields.ValidFrom).ValueChanged OrElse
                    mvClassFields(ContactLink.ContactLinkFields.ValidTo).ValueChanged OrElse
                        mvClassFields(ContactLink.ContactLinkFields.RelationshipStatus).ValueChanged

        'We'll only synch if the complementary record was synched prior to the save
        Dim vWasSynched As Boolean = False
        If Me.Existing Then
          'Check if complementary's current values are the same as current record's previous values (i.e. the SetValue property)
          vWasSynched =
            Me.ComplementaryLink.ValidFrom = Me.ClassFields(ContactLink.ContactLinkFields.ValidFrom).SetValue AndAlso
                 Me.ComplementaryLink.ValidTo = Me.ClassFields(ContactLink.ContactLinkFields.ValidTo).SetValue AndAlso
                    Me.ComplementaryLink.RelationshipStatus = Me.ClassFields(ContactLink.ContactLinkFields.RelationshipStatus).SetValue
        Else
          'For new records, we don't need to check the complementary's dates but we must check that the complementary's relationship has an equivalent status
          vWasSynched =
            Me.Relationship.CanAutoCreateComplementaryRelationship AndAlso
            (
              Me.RelationshipStatus.IsNullOrWhitespace OrElse
                Me.ComplementaryLink.Relationship.RelationshipStatuses.Any(Function(vStatus) vStatus.Field("relationship_status").Value.Equals(Me.RelationshipStatus))
             )
        End If
        vRequiresSync = vSynchedValuesChanged AndAlso vWasSynched
      End If
      Return vRequiresSync
    End Function

    Private Sub DoSyncComplementaryRelationship()
      If Me.ComplementaryLink IsNot Nothing Then
        Me.ComplementaryLink.ValidFrom = Me.ValidFrom
        Me.ComplementaryLink.ValidTo = Me.ValidTo
        'Before we set the complementary's status, we need to make sure that our status is valid for the complementary
        If Me.RelationshipStatus.IsNullOrWhitespace OrElse
          Me.ComplementaryLink.Relationship.RelationshipStatuses.Any(Function(vStatus) vStatus.Field("relationship_status").Value.Equals(Me.RelationshipStatus)) Then
          Me.ComplementaryLink.RelationshipStatus = Me.RelationshipStatus
        End If
      End If
    End Sub

    Public Property LinkType As ContactLinkTypes
      Get
        Return mvLinkType
      End Get
      Private Set(value As ContactLinkTypes)
        mvLinkType = value
      End Set
    End Property
  End Class

End Namespace
