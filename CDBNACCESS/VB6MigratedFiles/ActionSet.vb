Namespace Access
  Public Class ActionSet
    Implements System.Collections.IEnumerable

    Private mvCol As New Collection

    Public Sub Init(ByVal pEnv As CDBEnvironment, ByRef pMasterAction As Integer)
      Dim vAction As New Action(pEnv)
      Dim vRecordSet As CDBRecordSet

      vAction.Init()
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT " & vAction.GetRecordSetFields() & " FROM actions ac WHERE master_action = " & pMasterAction & " ORDER BY sequence_number")
      While vRecordSet.Fetch() = True
        vAction = New Action(pEnv)
        vAction.InitFromRecordSet(vRecordSet)
        mvCol.Add(vAction, CStr(vAction.ActionNumber))
      End While
      vRecordSet.CloseRecordSet()
    End Sub

    Public Function CreateStatusChangeAction(ByVal pEnv As CDBEnvironment, ByVal pGroup As String, ByVal pNumber As Integer, ByVal pStatus1 As String, ByVal pStatus2 As String) As Integer
      Dim vRecordSet As CDBRecordSet
      Dim vWhereFields As New CDBFields
      Dim vActionNumber As Integer
      Dim vNewActionSet As New ActionSet
      Dim vEntityGroup As EntityGroup

      If pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbDataCustomFinderTab) Then
        vWhereFields.Add("contact_group", CDBField.FieldTypes.cftCharacter, pGroup)
        vWhereFields.Add("status_1", CDBField.FieldTypes.cftCharacter, pStatus1)
        vWhereFields.Add("status_2", CDBField.FieldTypes.cftCharacter, pStatus2)
        vRecordSet = pEnv.Connection.GetRecordSet("SELECT action_number FROM status_transitions WHERE " & pEnv.Connection.WhereClause(vWhereFields))
        If vRecordSet.Fetch() = True Then vActionNumber = vRecordSet.Fields(1).IntegerValue
        vRecordSet.CloseRecordSet()
        If vActionNumber > 0 Then
          Init(pEnv, vActionNumber)
          If Count = 0 Then RaiseError(DataAccessErrors.daeInvalidProFormaAction, CStr(vActionNumber))
          If Not Item(1).ActionStatus = Action.ActionStatuses.astProForma Then RaiseError(DataAccessErrors.daeInvalidProFormaAction, CStr(vActionNumber))
          'Set up a new action set from it
          vEntityGroup = pEnv.EntityGroups(pGroup)
          Dim vActionLinkObjectType As IActionLink.ActionLinkObjectTypes = IActionLink.ActionLinkObjectTypes.alotContact
          If vEntityGroup.EntityGroupType = EntityGroup.EntityGroupTypes.egtOrganisation Then vActionLinkObjectType = IActionLink.ActionLinkObjectTypes.alotOrganisation
          vNewActionSet.CreateFromProForma(pEnv, Me, 0, vActionLinkObjectType, pNumber)
          CreateStatusChangeAction = vNewActionSet.Item(1).MasterAction
        End If
      End If
    End Function

    ''' <summary>Create a new set of Actions for each <see cref="Access.Action">Action</see> in the specified Proforma (Template) <see cref="ActionSet">ActionSet</see>.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pProFormaSet">The <see cref="ActionSet">ActionSet</see> of Proforma Actions to be used as the basis for the new Actions.</param>
    ''' <param name="pNewActionNumber">The new Action number for the first Action in the set.</param>
    ''' <param name="pRelatedType">The link type for a related link to be created for each Action.</param>
    ''' <param name="pRelatedNumber">The link number for a related link to be created for each Action.</param>
    Public Sub CreateFromProForma(ByVal pEnv As CDBEnvironment, ByVal pProFormaSet As ActionSet, ByVal pNewActionNumber As Integer, ByVal pRelatedType As IActionLink.ActionLinkObjectTypes, ByVal pRelatedNumber As Integer)
      CreateFromProForma(pEnv, pProFormaSet, pNewActionNumber, pRelatedType, pRelatedNumber, IActionLink.ActionLinkObjectTypes.alotContact, 0, 0, 0, Date.Today, Date.Today, True)
    End Sub
    ''' <summary>Create a new set of Actions for each <see cref="Access.Action">Action</see> in the specified Proforma (Template) <see cref="ActionSet">ActionSet</see>.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pProFormaSet">The <see cref="ActionSet">ActionSet</see> of Proforma Actions to be used as the basis for the new Actions.</param>
    ''' <param name="pNewActionNumber">The new Action number for the first Action in the set.</param>
    ''' <param name="pRelatedType">The type of object for a related link to be created for each Action.</param>
    ''' <param name="pRelatedNumber">The link number for a related link to be created for each Action.</param>
    ''' <param name="pActionerType">The type of object for an Actioner link to be created for each Action.</param>
    ''' <param name="pActionerNumber">The number of the Actioner to be created for each Action.</param>
    Public Sub CreateFromProForma(ByVal pEnv As CDBEnvironment, ByVal pProFormaSet As ActionSet, ByVal pNewActionNumber As Integer, ByVal pRelatedType As IActionLink.ActionLinkObjectTypes, ByVal pRelatedNumber As Integer, ByVal pActionerType As IActionLink.ActionLinkObjectTypes, ByVal pActionerNumber As Integer)
      CreateFromProForma(pEnv, pProFormaSet, pNewActionNumber, pRelatedType, pRelatedNumber, pActionerType, pActionerNumber, 0, 0, Date.Today, Date.Today, True)
    End Sub
    ''' <summary>Create a new set of Actions for each <see cref="Access.Action">Action</see> in the specified Proforma (Template) <see cref="ActionSet">ActionSet</see>.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pProFormaSet">The <see cref="ActionSet">ActionSet</see> of Proforma Actions to be used as the basis for the new Actions.</param>
    ''' <param name="pNewActionNumber">The new Action number for the first Action in the set.</param>
    ''' <param name="pRelatedType">The type of object for a related link to be created for each Action.</param>
    ''' <param name="pRelatedNumber">The link number for a related link to be created for each Action.</param>
    ''' <param name="pActionerType">The type of object for an Actioner link to be created for each Action.</param>
    ''' <param name="pActionerNumber">The number of the Actioner to be created for each Action.</param>
    ''' <param name="pRelatedDocument">The number of a Document to be linked to each Action.</param>
    ''' <param name="pRelatedExamCentreId">The number of an Exam Centre to be linked to each Action.</param>
    Public Sub CreateFromProForma(ByVal pEnv As CDBEnvironment, ByVal pProFormaSet As ActionSet, ByVal pNewActionNumber As Integer, ByVal pRelatedType As IActionLink.ActionLinkObjectTypes, ByVal pRelatedNumber As Integer, ByVal pActionerType As IActionLink.ActionLinkObjectTypes, ByVal pActionerNumber As Integer, ByVal pRelatedDocument As Integer, ByVal pRelatedExamCentreId As Integer)
      CreateFromProForma(pEnv, pProFormaSet, pNewActionNumber, pRelatedType, pRelatedNumber, pActionerType, pActionerNumber, pRelatedDocument, pRelatedExamCentreId, Date.Today, Date.Today, True)
    End Sub
    ''' <summary>Create a new set of Actions for each <see cref="Access.Action">Action</see> in the specified Proforma (Template) <see cref="ActionSet">ActionSet</see>.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pProFormaSet">The <see cref="ActionSet">ActionSet</see> of Proforma Actions to be used as the basis for the new Actions.</param>
    ''' <param name="pNewActionNumber">The new Action number for the first Action in the set.</param>
    ''' <param name="pRelatedType">The type of object for a related link to be created for each Action.</param>
    ''' <param name="pRelatedNumber">The link number for a related link to be created for each Action.</param>
    ''' <param name="pActionerType">The type of object for an Actioner link to be created for each Action.</param>
    ''' <param name="pActionerNumber">The number of the Actioner to be created for each Action.</param>
    ''' <param name="pRelatedDocument">The number of a Document to be linked to each Action.</param>
    ''' <param name="pRelatedExamCentreId">The number of an Exam Centre to be linked to each Action.</param>
    ''' <param name="pStartProcessingDate">The base date to use for calculating the new Action dates when the Proforma Action uses negative offsets or is the main (first) Action in the set.</param>
    ''' <param name="pEndProcessingDate">The base date to use for calculating the new Action dates when the Proforma Action does not use negative offsets.</param>
    ''' <param name="pCreateActions">True to save the new Actions and related data in the database, otherwise False to set the Action data without saving.</param>
    Public Sub CreateFromProForma(ByVal pEnv As CDBEnvironment, ByVal pProFormaSet As ActionSet, ByVal pNewActionNumber As Integer, ByVal pRelatedType As IActionLink.ActionLinkObjectTypes, ByVal pRelatedNumber As Integer, ByVal pActionerType As IActionLink.ActionLinkObjectTypes, ByVal pActionerNumber As Integer, ByVal pRelatedDocument As Integer, ByVal pRelatedExamCentreId As Integer, ByVal pStartProcessingDate As Date, ByVal pEndProcessingDate As Date, ByVal pCreateActions As Boolean)
      Dim vProForma As Action
      Dim vAction As Action
      Dim vMasterNumber As Integer
      Dim vActionNumber As Integer

      Dim vCount As Integer
      For Each vProForma In pProFormaSet
        vCount = vProForma.PriorActions.Count() 'Force a read outside of any transaction
        vCount = vProForma.Subjects.Count() 'Force a read outside of any transaction
        vCount = vProForma.Links.Count() 'Force a read outside of any transaction
      Next vProForma

      If pNewActionNumber > 0 Then
        vActionNumber = pNewActionNumber
        vMasterNumber = pNewActionNumber
      End If

      pEnv.Connection.StartTransaction()

      '(1) Process the first Action before any others
      vProForma = pProFormaSet.Item(1)
      If vActionNumber = 0 Then vActionNumber = If(pCreateActions = True, pEnv.GetControlNumber("AC"), vProForma.ActionNumber * 10000)
      If vMasterNumber = 0 Then vMasterNumber = vActionNumber
      vAction = New Action(pEnv)
      vAction.CreateFromProForma(pEnv, vProForma, vMasterNumber, vActionNumber, pRelatedType, pRelatedNumber, pActionerType, pActionerNumber, pRelatedDocument, pRelatedExamCentreId, pStartProcessingDate, pCreateActions)
      vActionNumber = 0
      mvCol.Add(vAction, vAction.ActionNumber.ToString)

      '(2) Process all Actions in the Proforma that are using negative off-sets
      Dim vFirst As Boolean = True
      If pProFormaSet.Count > 1 Then
        For Each vProForma In pProFormaSet
          If vFirst = False AndAlso vProForma.UseNegativeOffsets = True Then
            If vActionNumber = 0 Then vActionNumber = If(pCreateActions = True, pEnv.GetControlNumber("AC"), vProForma.ActionNumber * 10000)
            'If vMasterNumber = 0 Then vMasterNumber = vActionNumber
            vAction = New Action(pEnv)
            vAction.CreateFromProForma(pEnv, vProForma, vMasterNumber, vActionNumber, pRelatedType, pRelatedNumber, pActionerType, pActionerNumber, pRelatedDocument, pRelatedExamCentreId, pStartProcessingDate, pCreateActions)
            vActionNumber = 0
            mvCol.Add(vAction, vAction.ActionNumber.ToString)
          End If
          vFirst = False
        Next
      End If

      '(3) Process the remaining Actions
      If pProFormaSet.Count > 1 Then
        vFirst = True
        For Each vProForma In pProFormaSet
          If vFirst = False AndAlso vProForma.UseNegativeOffsets = False Then
            If vActionNumber = 0 Then vActionNumber = If(pCreateActions = True, pEnv.GetControlNumber("AC"), vProForma.ActionNumber * 10000)
            'If vMasterNumber = 0 Then vMasterNumber = vActionNumber
            vAction = New Action(pEnv)
            vAction.CreateFromProForma(pEnv, vProForma, vMasterNumber, vActionNumber, pRelatedType, pRelatedNumber, pActionerType, pActionerNumber, pRelatedDocument, pRelatedExamCentreId, pEndProcessingDate, pCreateActions)
            vActionNumber = 0
            mvCol.Add(vAction, vAction.ActionNumber.ToString)
          End If
          vFirst = False
        Next
      End If

      'Now set up dependencies for the new actions where the pro-formas had dependencies
      Dim vFields As New CDBFields(New CDBField("action_number", CDBField.FieldTypes.cftInteger))
      vFields.Add("prior_action", CDBField.FieldTypes.cftInteger)
      vFields.AddAmendedOnBy(pEnv.User.UserID)
      Dim vPriorNumber As Integer = 0
      For Each vProForma In pProFormaSet
        If vProForma.PriorActions.Count() > 0 Then 'If there were dependencies
          For vIndex As Integer = 1 To vProForma.PriorActions.Count() 'For each dependency
            vPriorNumber = CInt(vProForma.PriorActions.Item(vIndex))    'PriorActions is a collection of Action Numbers
            vFields("action_number").Value = GetNumberByProForma(vProForma.ActionNumber).ToString
            'Then get the action whose proforma was the prior we are processing
            vFields("prior_action").Value = GetNumberByProForma(vPriorNumber).ToString
            pEnv.Connection.InsertRecord("action_dependencies", vFields)
          Next
        End If
      Next vProForma
      pEnv.Connection.CommitTransaction()

      If pCreateActions = False Then
        'When not creating actions, perform an additional step to populate priority & status descriptions
        Dim vCodeList As New List(Of String)
        For Each vAction In mvCol
          If vCodeList.Contains(vAction.ActionPriority) = False Then vCodeList.Add(vAction.ActionPriority)
        Next
        Dim vValues As String = "'" & vCodeList.AsCommaSeperated.Replace(",", "','") & "'"
        Dim vWhereFields As New CDBFields(New CDBField("action_priority", CDBField.FieldTypes.cftCharacter, vValues, CDBField.FieldWhereOperators.fwoIn))
        Dim vSQLStatement As New SQLStatement(pEnv.Connection, "action_priority, action_priority_desc", "action_priorities", vWhereFields)
        Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          For Each vAction In mvCol
            If vAction.ActionPriority.Equals(vRS.Fields(1).Value, StringComparison.InvariantCulture) Then vAction.ActionPriorityDescription = vRS.Fields(2).Value
          Next
        End While
        vRS.CloseRecordSet()
        vCodeList = New List(Of String)
        For Each vAction In mvCol
          If vCodeList.Contains(vAction.ActionStatusCode) = False Then vCodeList.Add(vAction.ActionStatusCode)
        Next
        vValues = "'" & vCodeList.AsCommaSeperated.Replace(",", "','") & "'"
        vWhereFields = New CDBFields(New CDBField("action_status", CDBField.FieldTypes.cftCharacter, vValues, CDBField.FieldWhereOperators.fwoIn))
        vSQLStatement = New SQLStatement(pEnv.Connection, "action_status, action_status_desc", "action_statuses", vWhereFields)
        vRS = vSQLStatement.GetRecordSet()
        While vRS.Fetch
          For Each vAction In mvCol
            If vAction.ActionStatusCode.Equals(vRS.Fields(1).Value, StringComparison.InvariantCulture) Then vAction.ActionStatusDescription = vRS.Fields(2).Value
          Next
        End While
        vRS.CloseRecordSet()
      End If


    End Sub

    Private Function GetNumberByProForma(ByVal pActionNumber As Integer) As Integer
      For Each vAction As Action In mvCol
        If vAction.ActionTemplateNumber = pActionNumber Then
          Return vAction.ActionNumber
        End If
      Next vAction
    End Function

    Public ReadOnly Property Item(ByVal pIndexKey As String) As Action
      Get
        Item = DirectCast(mvCol.Item(pIndexKey), Action)
      End Get
    End Property

    Public ReadOnly Property Item(ByVal pIndexKey As Integer) As Action
      Get
        Item = DirectCast(mvCol.Item(pIndexKey), Action)
      End Get
    End Property

    Public ReadOnly Property Count() As Integer
      Get
        Count = mvCol.Count()
      End Get
    End Property

    Public Function GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
      GetEnumerator = mvCol.GetEnumerator
    End Function
    Public Function Exists(ByVal pIndexKey As String) As Boolean
      Return mvCol.Contains(pIndexKey)
    End Function

    Public Sub Remove(ByRef pIndexKey As String)
      mvCol.Remove(pIndexKey)
    End Sub

    ''' <summary>Create or update the set of Actions for each <see cref="Access.Action">Action</see> in the specified Proforma (Template) <see cref="ActionSet">ActionSet</see>.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pProFormaSet">The <see cref="ActionSet">ActionSet</see> of Proforma Actions to be used as the basis for the new Actions.</param>
    ''' <param name="pWhereClause">A <see cref="CDBFields">CDBFields Collection</see> containing the where clause to find existing records.</param>
    ''' <param name="pLinkTableJoin">An <see cref="AnsiJoin">AnsiJoin</see> record containing a join from the actions table to he link table.</param>
    ''' <param name="pStartProcessingDate">The base date to use for calculating the new Action dates when the Proforma Action uses negative offsets or is the main (first) Action in the set.</param>
    ''' <param name="pEndProcessingDate">The base date to use for calculating the new Action dates when the Proforma Action does not use negative offsets.</param>
    ''' <remarks></remarks>
    Public Sub UpdateFromProforma(ByVal pEnv As CDBEnvironment, ByVal pProFormaSet As ActionSet, ByVal pWhereClause As CDBFields, ByVal pLinkTableJoin As AnsiJoin, ByVal pStartProcessingDate As Date, ByVal pEndProcessingDate As Date)
      Dim vExistingActions As New CollectionList(Of Action)

      '(1) Select all existing Actions linked to the specified ActionSet
      Dim vTemplateNumbers As String = String.Empty
      Dim vAction As Action
      For Each vAction In pProFormaSet
        If vTemplateNumbers.Length > 0 Then vTemplateNumbers &= ","
        vTemplateNumbers &= vAction.ActionNumber.ToString
      Next
      pWhereClause.Add("action_template_number", CDBField.FieldTypes.cftInteger, vTemplateNumbers, CDBField.FieldWhereOperators.fwoInOrEqual)
      Dim vAnsiJoins As New AnsiJoins({pLinkTableJoin})
      vAction = New Action(pEnv)
      vAction.Init()
      Dim vSQLStatement As New SQLStatement(pEnv.Connection, vAction.GetRecordSetFields, "actions ac", pWhereClause, "master_action,sequence_number", vAnsiJoins)
      Dim vRS As CDBRecordSet = vSQLStatement.GetRecordSet()
      While vRS.Fetch
        vAction = New Action(pEnv)
        vAction.InitFromRecordSet(vRS)
        vExistingActions.Add(vAction.ActionTemplateNumber.ToString, vAction)
      End While
      vRS.CloseRecordSet()

      '(2) Select all related data for existing Proforma
      Dim vCount As Integer
      Dim vProforma As Action
      For Each vProforma In pProFormaSet
        vCount = vProforma.PriorActions.Count() 'Force a read outside of any transaction
        vCount = vProforma.Subjects.Count() 'Force a read outside of any transaction
        vCount = vProforma.Links.Count() 'Force a read outside of any transaction
      Next vProforma

      pEnv.Connection.StartTransaction()

      '(3) Process the first Action before any others
      Dim vActionNumber As Integer = 0
      Dim vMasterNumber As Integer = 0
      vProforma = pProFormaSet.Item(1)
      If vExistingActions.ContainsKey(vProforma.ActionNumber.ToString) Then
        vAction = vExistingActions.Item(vProforma.ActionNumber.ToString)
        vActionNumber = vAction.ActionNumber
        vMasterNumber = vAction.MasterAction
        vAction.UpdateDatesFromProforma(pEnv, vProforma, pStartProcessingDate)
      Else
        If vActionNumber = 0 Then vActionNumber = vProforma.ActionNumber * 10000
        If vMasterNumber = 0 Then vMasterNumber = vActionNumber
        vAction.CreateFromProForma(pEnv, vProforma, vMasterNumber, vActionNumber, IActionLink.ActionLinkObjectTypes.alotContact, 0, IActionLink.ActionLinkObjectTypes.alotContact, 0, 0, 0, pStartProcessingDate, False)
      End If
      vActionNumber = 0
      mvCol.Add(vAction, vAction.ActionNumber.ToString)

      '(4) Process all Actions in the Proforma that are using negative off-sets
      Dim vFirst As Boolean = True
      If pProFormaSet.Count > 1 Then
        For Each vProforma In pProFormaSet
          If vFirst = False AndAlso vProforma.UseNegativeOffsets = True Then
            If vExistingActions.ContainsKey(vProforma.ActionNumber.ToString) Then
              vAction = vExistingActions.Item(vProforma.ActionNumber.ToString)
              vActionNumber = vAction.ActionNumber
              vAction.UpdateDatesFromProforma(pEnv, vProforma, pStartProcessingDate)
            Else
              vActionNumber = vProforma.ActionNumber * 10000
              vAction = New Action(pEnv)
              vAction.CreateFromProForma(pEnv, vProforma, vMasterNumber, vActionNumber, IActionLink.ActionLinkObjectTypes.alotContact, 0, IActionLink.ActionLinkObjectTypes.alotContact, 0, 0, 0, pStartProcessingDate, False)
            End If
            vActionNumber = 0
            mvCol.Add(vAction, vAction.ActionNumber.ToString)
          End If
          vFirst = False
        Next
      End If

      '(5) Process the remaining Actions
      If pProFormaSet.Count > 1 Then
        vFirst = True
        For Each vProforma In pProFormaSet
          If vFirst = False AndAlso vProforma.UseNegativeOffsets = False Then
            If vExistingActions.ContainsKey(vProforma.ActionNumber.ToString) Then
              vAction = vExistingActions.Item(vProforma.ActionNumber.ToString)
              vActionNumber = vAction.ActionNumber
              vAction.UpdateDatesFromProforma(pEnv, vProforma, pEndProcessingDate)
            Else
              vActionNumber = vProforma.ActionNumber * 10000
              vAction = New Action(pEnv)
              vAction.CreateFromProForma(pEnv, vProforma, vMasterNumber, vActionNumber, IActionLink.ActionLinkObjectTypes.alotContact, 0, IActionLink.ActionLinkObjectTypes.alotContact, 0, 0, 0, pEndProcessingDate, False)
            End If
            vActionNumber = 0
            mvCol.Add(vAction, vAction.ActionNumber.ToString)
          End If
          vFirst = False
        Next
      End If

      'Now set up dependencies for the new actions where the pro-formas had dependencies
      Dim vFields As New CDBFields(New CDBField("action_number", CDBField.FieldTypes.cftInteger))
      vFields.Add("prior_action", CDBField.FieldTypes.cftInteger)
      vFields.AddAmendedOnBy(pEnv.User.UserID)
      Dim vPriorNumber As Integer = 0
      For Each vProforma In pProFormaSet
        If vProforma.PriorActions.Count() > 0 Then 'If there were dependencies
          For vIndex As Integer = 1 To vProforma.PriorActions.Count() 'For each dependency
            vPriorNumber = CInt(vProforma.PriorActions.Item(vIndex))    'PriorActions is a collection of Action Numbers
            vFields("action_number").Value = GetNumberByProForma(vProforma.ActionNumber).ToString
            'Then get the action whose proforma was the prior we are processing
            vFields("prior_action").Value = GetNumberByProForma(vPriorNumber).ToString
            pEnv.Connection.DeleteRecords("action_dependencies", vFields, False)    'Delete any existing records first then re-create
            pEnv.Connection.InsertRecord("action_dependencies", vFields)
          Next
        End If
      Next vProforma

      pEnv.Connection.CommitTransaction()

    End Sub

    ''' <summary>Create <see cref="Action">Actions</see> from Proforma (Template) when those Action objects have already been set up.</summary>
    ''' <param name="pEnv"></param>
    ''' <param name="pActionsColl">A collection containing the Actions to be created that are from this ActionSet.</param>
    ''' <param name="pWorkstreamId">The number of the Workstream to be linked to each Action.</param>
    Public Sub CreateActionsFromProvisional(ByVal pEnv As CDBEnvironment, ByVal pActionsColl As CollectionList(Of Action), ByVal pWorkstreamId As Integer)

      Dim vTemplate As Action = Nothing
      Dim vNewAction As Action = Nothing

      Dim vCount As Integer
      For Each vTemplate In mvCol
        vCount = vTemplate.PriorActions.Count() 'Force a read outside of any transaction
        vCount = vTemplate.Subjects.Count() 'Force a read outside of any transaction
        vCount = vTemplate.Links.Count() 'Force a read outside of any transaction
      Next

      pEnv.Connection.StartTransaction()

      Dim vMasterActionNumber As Integer = 0
      For Each vTemplate In mvCol
        If pActionsColl.ContainsKey(vTemplate.ActionNumber.ToString) Then
          vNewAction = pActionsColl.Item(vTemplate.ActionNumber.ToString)
          vNewAction.CreateFromProvisionalProforma(vTemplate, vMasterActionNumber, pWorkstreamId)
          vMasterActionNumber = vNewAction.MasterAction
        End If
      Next

      pEnv.Connection.CommitTransaction()

    End Sub



  End Class
End Namespace
