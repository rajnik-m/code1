Namespace Access

    Public Class ExamUnitLink
        Inherits CARERecord
        Implements IRecordCreate

#Region "AutoGenerated Code"

        '--------------------------------------------------
        'Enum defining all the fields in the table
        '--------------------------------------------------
        Private Enum ExamUnitLinkFields
            AllFields = 0
            ExamUnitLinkId
            ExamUnitId1
            ExamUnitId2
            CreatedBy
            CreatedOn
            LongDescription
            ParentUnitLinkId
            BaseUnitLinkId
            AccreditationStatus
            AccreditationValidfrom
            AccreditationValidTo
            AmendedBy
            AmendedOn
        End Enum

        '--------------------------------------------------
        'Required overrides for the class
        '--------------------------------------------------
        Protected Overrides Sub AddFields()
            With mvClassFields
                .Add("exam_unit_link_id", CDBField.FieldTypes.cftInteger)
                .Add("exam_unit_id_1", CDBField.FieldTypes.cftInteger)
                .Add("exam_unit_id_2", CDBField.FieldTypes.cftInteger)
                .Add("created_by")
                .Add("created_on", CDBField.FieldTypes.cftDate)
                .Add("long_description")
                .Add("parent_unit_link_id", CDBField.FieldTypes.cftInteger)
                .Add("base_unit_link_id", CDBField.FieldTypes.cftInteger)
                .Add("accreditation_status")
                .Add("accreditation_valid_from", CDBField.FieldTypes.cftDate)
                .Add("accreditation_valid_to", CDBField.FieldTypes.cftDate)

                '.Item(ExamUnitLinkFields.ExamUnitId1).PrimaryKey = True
                .Item(ExamUnitLinkFields.ExamUnitId1).PrefixRequired = True


                '.Item(ExamUnitLinkFields.ExamUnitId2).PrimaryKey = True
                .Item(ExamUnitLinkFields.ExamUnitId2).PrefixRequired = True

                .Item(ExamUnitLinkFields.CreatedBy).PrefixRequired = True
                .Item(ExamUnitLinkFields.CreatedOn).PrefixRequired = False

                If mvEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbExamUnitLinkLongDescription) Then
                    .Item(ExamUnitLinkFields.ExamUnitLinkId).InDatabase = True
                    .Item(ExamUnitLinkFields.ExamUnitLinkId).PrimaryKey = True
                    .Item(ExamUnitLinkFields.ExamUnitLinkId).PrefixRequired = True
                    .SetControlNumberField(ExamUnitLinkFields.ExamUnitLinkId, "EUL")
                    .Item(ExamUnitLinkFields.AccreditationStatus).InDatabase = True
                    .Item(ExamUnitLinkFields.AccreditationValidfrom).InDatabase = True
                    .Item(ExamUnitLinkFields.AccreditationValidTo).InDatabase = True
                    .Item(ExamUnitLinkFields.ParentUnitLinkId).InDatabase = True
                    .Item(ExamUnitLinkFields.ParentUnitLinkId).PrefixRequired = True
                    .Item(ExamUnitLinkFields.BaseUnitLinkId).InDatabase = True
                    .Item(ExamUnitLinkFields.BaseUnitLinkId).PrefixRequired = True
                    .Item(ExamUnitLinkFields.LongDescription).InDatabase = True
                    .Item(ExamUnitLinkFields.LongDescription).PrefixRequired = True
                    .Item(ExamUnitLinkFields.AccreditationStatus).PrefixRequired = True
                    .Item(ExamUnitLinkFields.AccreditationValidfrom).PrefixRequired = True
                    .Item(ExamUnitLinkFields.AccreditationValidTo).PrefixRequired = True
                End If
            End With
        End Sub

        Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
            Get
                Return True
            End Get
        End Property
        Protected Overrides ReadOnly Property TableAlias() As String
            Get
                Return "eul"
            End Get
        End Property
        Protected Overrides ReadOnly Property DatabaseTableName() As String
            Get
                Return "exam_unit_links"
            End Get
        End Property

        '--------------------------------------------------
        'Default constructor
        '--------------------------------------------------
        Public Sub New(ByVal pEnv As CDBEnvironment)
            MyBase.New(pEnv)
        End Sub

        '--------------------------------------------------
        'IRecordCreate
        '--------------------------------------------------
        Public Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
            Return New ExamUnitLink(mvEnv)
        End Function
        '--------------------------------------------------
        'Public property procedures
        '--------------------------------------------------
        Public ReadOnly Property ExamUnitId1() As Integer
            Get
                Return mvClassFields(ExamUnitLinkFields.ExamUnitId1).IntegerValue
            End Get
        End Property
        Public ReadOnly Property ExamUnitId2() As Integer
            Get
                Return mvClassFields(ExamUnitLinkFields.ExamUnitId2).IntegerValue
            End Get
        End Property
        Public ReadOnly Property CreatedBy() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.CreatedBy).Value
            End Get
        End Property
        Public ReadOnly Property CreatedOn() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.CreatedOn).Value
            End Get
        End Property
        Public ReadOnly Property AmendedBy() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.AmendedBy).Value
            End Get
        End Property
        Public ReadOnly Property AmendedOn() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.AmendedOn).Value
            End Get
        End Property
        Public ReadOnly Property LongDescription() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.LongDescription).Value
            End Get
        End Property
        Public ReadOnly Property ExamUnitLinkId() As Integer
            Get
                Return mvClassFields(ExamUnitLinkFields.ExamUnitLinkId).IntegerValue
            End Get
        End Property
        Public ReadOnly Property ParentUnitLinkId() As Integer
            Get
                Return mvClassFields(ExamUnitLinkFields.ParentUnitLinkId).IntegerValue
            End Get
        End Property
        Public ReadOnly Property BaseUnitLinkId() As Integer
            Get
                Return mvClassFields(ExamUnitLinkFields.BaseUnitLinkId).IntegerValue
            End Get
        End Property
        Public ReadOnly Property AccreditationStatus() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.AccreditationStatus).Value
            End Get
        End Property
        Public ReadOnly Property AccreditationValidFrom() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.AccreditationValidfrom).Value
            End Get
        End Property
        Public ReadOnly Property AccreditationValidTo() As String
            Get
                Return mvClassFields(ExamUnitLinkFields.AccreditationValidTo).Value
            End Get
        End Property

        '--------------------------------------------------
        'AddDeleteCheckItems
        '--------------------------------------------------
        Public Overrides Sub AddDeleteCheckItems()
            AddCascadeDeleteItem("category_links", "exam_unit_link_id")
        End Sub

        Public Overrides Sub Save(ByVal pAmendedBy As String, ByVal pAudit As Boolean, ByVal pJournalNumber As Integer)
            SaveAccreditationHistory()
            MyBase.Save(pAmendedBy, pAudit, pJournalNumber, True)
        End Sub

        Private Sub SaveAccreditationHistory()
            If (mvClassFields(ExamUnitLinkFields.AccreditationStatus).ValueChanged And Me.Existing) Or
      (mvClassFields(ExamUnitLinkFields.AccreditationValidfrom).ValueChanged And Me.Existing) Or
      (mvClassFields(ExamUnitLinkFields.AccreditationValidTo).ValueChanged And Me.Existing) Then
                Dim vAccreditationHistoryRecord As New ExamAccreditationHistory(mvEnv)
                Dim vParams As New CDBParameters()
                If mvClassFields(ExamUnitLinkFields.AccreditationStatus).SetValue.Length > 0 Then
                    vParams.Add("AccreditationStatus", mvClassFields(ExamUnitLinkFields.AccreditationStatus).SetValue)
                    vParams.Add("AccreditationValidFrom", mvClassFields(ExamUnitLinkFields.AccreditationValidfrom).SetValue)
                    vParams.Add("AccreditationValidTo", mvClassFields(ExamUnitLinkFields.AccreditationValidTo).SetValue)
                    vAccreditationHistoryRecord.Create(vParams)
                    vAccreditationHistoryRecord.Save()
                    If vAccreditationHistoryRecord.AccreditationId > 0 Then
                        Dim vAccreditationHistoryLink As New ExamAccreditationHistLink(mvEnv)
                        Dim vLinkParams As New CDBParameters()
                        vLinkParams.Add("ExamUnitLinkId", mvClassFields(ExamUnitLinkFields.ExamUnitLinkId).Value.ToString)
                        vLinkParams.Add("AccreditationId", vAccreditationHistoryRecord.AccreditationId)
                        vAccreditationHistoryLink.Create(vLinkParams)
                        vAccreditationHistoryLink.Save()
                    End If
                End If
            End If
        End Sub
#End Region
        ''' <summary>
        ''' Returns the exam unit link Id that matches the current record's ExamUnitId1 and ExamUnitID2's Session 0 records
        ''' </summary>
        ''' <param name="pEnv">You must pass an existing Environment to this method as this object is not database-aware</param>
        ''' <returns></returns>
        ''' <remarks>The BasedUnitLinkId is the ExamUnitLink record whose ExamUnitId1 and ExamUnitId2 are the BaseExamUnitIds for the current ExamUnitLink's Parent (Unit1) and Child (Unit2) Units </remarks>
        Public Shared Function GetBaseUnitLinkId(pEnv As CDBEnvironment, pExamUnitLinkId As Integer) As Integer

            Dim vRtn = 0

            Dim vJoins As New AnsiJoins
            Dim vWhereFields As New CDBFields
            vJoins.Add("exam_units eu2", "eu2.exam_unit_id", "eul.exam_unit_id_2")
            vJoins.Add("exam_units beu2", "eu2.exam_base_unit_id", "beu2.exam_unit_id")

            vJoins.AddLeftOuterJoin("exam_units eu1", "eu1.exam_unit_id", "eul.exam_unit_id_1")
            vJoins.AddLeftOuterJoin("exam_units beu1", "beu1.exam_unit_id", "eu1.exam_base_unit_id") 'base unit Id of the parent
            vJoins.Add("exam_unit_links beul", "beul.exam_unit_id_1", pEnv.Connection.DBIsNull("beu1.exam_unit_id", "0"), "beul.exam_unit_id_2", "eu2.exam_base_unit_id") 'base unit link id
            vWhereFields.Add("beu1.exam_session_id")

            vWhereFields.Add("beu2.exam_session_id")
            vWhereFields.Add("eul.exam_unit_link_id", pExamUnitLinkId.ToString)

            Dim vSQL = New SQLStatement(pEnv.Connection, "beul.exam_unit_link_id", "exam_unit_links eul", vWhereFields, "", vJoins)

            Dim vRecordset As CDBRecordSet = pEnv.Connection.GetRecordSet(vSQL.SQL)
            If vRecordset.Fetch Then vRtn = vRecordset.Fields(1).IntegerValue

            Return vRtn

        End Function

        Public Shared Function GetUnitIdFromLinkId(pEnv As CDBEnvironment, pExamUnitLinkId As Integer) As Integer
            Dim vRtn = 0

            Dim vJoins As New AnsiJoins
            Dim vWhereFields As New CDBFields
            vWhereFields.Add("exam_unit_link_id", pExamUnitLinkId.ToString)

            Dim vSQL = New SQLStatement(pEnv.Connection, "exam_unit_id_2", "exam_unit_links eul", vWhereFields, "", vJoins)

            Dim vRecordset As CDBRecordSet = pEnv.Connection.GetRecordSet(vSQL.SQL)
            If vRecordset.Fetch Then vRtn = vRecordset.Fields(1).IntegerValue

            Return vRtn
        End Function

        Public Function GetBaseUnitLinkId(pEnv As CDBEnvironment) As Integer
            Return ExamUnitLink.GetBaseUnitLinkId(pEnv, ExamUnitLinkId)
        End Function

        Public Function GetChildExamUnitLinks() As IEnumerable(Of ExamUnitLink)
            Dim vField As CDBField = Me.ClassFields(ExamUnitLinkFields.ParentUnitLinkId)
            vField.Value = ExamUnitLinkId.ToString()

            Dim vResult As IEnumerable(Of ExamUnitLink) = CARERecordFactory.SelectList(Of ExamUnitLink)(Me.Environment, New CDBFields(vField))
            Return vResult

        End Function

    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the create methods
      PreValidateCreateSelfReference(pParameterList)
    End Sub

        Private Sub PreValidateCreateSelfReference(ByVal pParameterList As CDBParameters)
            If pParameterList("ExamUnitId1").Value.Equals(pParameterList("ExamUnitId2").Value) Then
                RaiseError(DataAccessErrors.daeParameterValueInvalid, String.Format("ExamUnitId1(={0})", pParameterList("ExamUnitId1").Value), String.Format("ExamUnitId2(={0})", pParameterList("ExamUnitId2").Value))
            End If
        End Sub
    End Class
End Namespace
