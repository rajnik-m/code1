

Namespace Access
  Public Class StickyNote

    Public Enum StickyNoteRecordSetTypes 'These are bit values
      snrtAll = &HFFFFS
      'ADD additional recordset types here
    End Enum

    'Keep the enum items in the same order as in the InitClassFields function
    Private Enum StickyNoteFields
      snfAll = 0
      snfNoteNumber
      snfUniqueId
      snfRecordType
      snfNotes
      snfCreatedOn
      snfPermanent
      snfAmendedBy
      snfAmendedOn
    End Enum

    'Standard Class Setup
    Private mvEnv As CDBEnvironment
    Private mvClassFields As ClassFields
    Private mvExisting As Boolean

    '-----------------------------------------------------------
    ' PRIVATE PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Private Sub InitClassFields()
      If mvClassFields Is Nothing Then
        mvClassFields = New ClassFields
        With mvClassFields
          .DatabaseTableName = "sticky_notes"
          'There should be an entry here for each field in the table
          'Keep these in the same order as the Fields enum
          .Add("note_number", CDBField.FieldTypes.cftLong)
          .Add("unique_id", CDBField.FieldTypes.cftLong)
          .Add("record_type")
          .Add("notes")
          .Add("created_on", CDBField.FieldTypes.cftTime)
          .Add("permanent")
          .Add("amended_by")
          .Add("amended_on", CDBField.FieldTypes.cftTime)

          .Item(StickyNoteFields.snfNoteNumber).SetPrimaryKeyOnly()

          .Item(StickyNoteFields.snfPermanent).SpecialColumn = True
        End With

      Else
        mvClassFields.ClearItems()
      End If
      mvExisting = False
    End Sub

    Private Sub SetDefaults()
      'Add code here to initialise the class with default values for a new record
    End Sub

    Private Sub SetValid(ByVal pField As StickyNoteFields)
      'Add code here to ensure all values are valid before saving
      If NoteNumber <= 0 Then mvClassFields.Item(StickyNoteFields.snfNoteNumber).IntegerValue = mvEnv.GetControlNumber("SN")
      If Len(CreatedOn) = 0 Then mvClassFields(StickyNoteFields.snfCreatedOn).Value = TodaysDateAndTime()
      mvClassFields.Item(StickyNoteFields.snfAmendedOn).Value = TodaysDateAndTime()
      mvClassFields.Item(StickyNoteFields.snfAmendedBy).Value = mvEnv.User.UserID
    End Sub

    '-----------------------------------------------------------
    ' PUBLIC PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public Function GetRecordSetFields(ByVal pRSType As StickyNoteRecordSetTypes) As String
      Dim vFields As String = ""
      'Modify below to add each recordset type as required
      If pRSType = StickyNoteRecordSetTypes.snrtAll Then
        If mvClassFields Is Nothing Then InitClassFields()
        vFields = mvClassFields.FieldNames(mvEnv, "sn")
      End If
      GetRecordSetFields = vFields
    End Function

    Public Sub Init(ByVal pEnv As CDBEnvironment, Optional ByRef pNoteNumber As Integer = 0)
      Dim vRecordSet As CDBRecordSet

      mvEnv = pEnv
      If pNoteNumber > 0 Then
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & GetRecordSetFields(StickyNoteRecordSetTypes.snrtAll) & " FROM sticky_notes sn WHERE note_number = " & pNoteNumber)
        If vRecordSet.Fetch() = True Then
          InitFromRecordSet(pEnv, vRecordSet, StickyNoteRecordSetTypes.snrtAll)
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

    Public Sub InitFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pRSType As StickyNoteRecordSetTypes)
      Dim vFields As CDBFields

      mvEnv = pEnv
      InitClassFields()
      vFields = pRecordSet.Fields
      mvExisting = True
      With mvClassFields
        'Always include the primary key attributes
        .SetItem(StickyNoteFields.snfNoteNumber, vFields)
        'Modify below to handle each recordset type as required
        If (pRSType And StickyNoteRecordSetTypes.snrtAll) = StickyNoteRecordSetTypes.snrtAll Then
          .SetItem(StickyNoteFields.snfUniqueId, vFields)
          .SetItem(StickyNoteFields.snfRecordType, vFields)
          .SetItem(StickyNoteFields.snfNotes, vFields)
          .SetItem(StickyNoteFields.snfCreatedOn, vFields)
          .SetItem(StickyNoteFields.snfPermanent, vFields)
          .SetItem(StickyNoteFields.snfAmendedBy, vFields)
          .SetItem(StickyNoteFields.snfAmendedOn, vFields)
        End If
      End With
    End Sub

    Public Sub CreateWithType(ByRef pType As String, ByRef pUniqueID As Integer, ByRef pNotes As String, ByRef pPermanent As Boolean)
      mvClassFields(StickyNoteFields.snfUniqueId).IntegerValue = pUniqueID
      mvClassFields(StickyNoteFields.snfRecordType).Value = pType
      mvClassFields(StickyNoteFields.snfNotes).Value = pNotes
      mvClassFields(StickyNoteFields.snfPermanent).Bool = pPermanent
    End Sub

    Public Sub Create(ByRef pContact As Contact, ByRef pNotes As String, ByRef pPermanent As Boolean)
      mvClassFields(StickyNoteFields.snfUniqueId).IntegerValue = pContact.ContactNumber
      If pContact.ContactType = Contact.ContactTypes.ctcOrganisation Then
        mvClassFields(StickyNoteFields.snfRecordType).Value = "O"
      Else
        mvClassFields(StickyNoteFields.snfRecordType).Value = "C"
      End If
      mvClassFields(StickyNoteFields.snfNotes).Value = pNotes
      mvClassFields(StickyNoteFields.snfPermanent).Bool = pPermanent
    End Sub

    Public Sub Update(ByRef pNotes As String, ByRef pPermanent As Boolean)
      mvClassFields(StickyNoteFields.snfNotes).Value = pNotes
      mvClassFields(StickyNoteFields.snfPermanent).Bool = pPermanent
    End Sub

    Public Sub Delete()
      mvClassFields.Delete(mvEnv.Connection)
    End Sub

    Public Sub Save(Optional ByRef pAmendedBy As String = "", Optional ByRef pAudit As Boolean = False)
      SetValid(StickyNoteFields.snfAll)
      mvClassFields.Save(mvEnv, mvExisting, pAmendedBy, pAudit)
    End Sub

    '-----------------------------------------------------------
    ' PROPERTY PROCEDURES FOLLOW
    '-----------------------------------------------------------
    Public ReadOnly Property Existing() As Boolean
      Get
        Existing = mvExisting
      End Get
    End Property

    Public ReadOnly Property AmendedBy() As String
      Get
        AmendedBy = mvClassFields.Item(StickyNoteFields.snfAmendedBy).Value
      End Get
    End Property

    Public ReadOnly Property AmendedOn() As String
      Get
        AmendedOn = mvClassFields.Item(StickyNoteFields.snfAmendedOn).Value
      End Get
    End Property

    Public ReadOnly Property CreatedOn() As String
      Get
        CreatedOn = mvClassFields.Item(StickyNoteFields.snfCreatedOn).Value
      End Get
    End Property

    Public ReadOnly Property NoteNumber() As Integer
      Get
        NoteNumber = mvClassFields.Item(StickyNoteFields.snfNoteNumber).IntegerValue
      End Get
    End Property

    Public ReadOnly Property Notes() As String
      Get
        Notes = mvClassFields.Item(StickyNoteFields.snfNotes).MultiLineValue
      End Get
    End Property

    Public ReadOnly Property Permanent() As Boolean
      Get
        Permanent = mvClassFields.Item(StickyNoteFields.snfPermanent).Bool
      End Get
    End Property

    Public ReadOnly Property RecordType() As String
      Get
        RecordType = mvClassFields.Item(StickyNoteFields.snfRecordType).Value
      End Get
    End Property

    Public ReadOnly Property UniqueId() As Integer
      Get
        UniqueId = mvClassFields.Item(StickyNoteFields.snfUniqueId).IntegerValue
      End Get
    End Property
  End Class
End Namespace
