Namespace Access

  Public Class Survey
    Inherits CARERecord
    Implements IRecordCreate

#Region "AutoGenerated Code"

'--------------------------------------------------
'Enum defining all the fields in the table
'--------------------------------------------------
    Private Enum SurveyFields
      AllFields = 0
      CreatedBy
      CreatedOn
      LongDescription
      Notes
      RespondedActivity
      RespondedActivityValue
      SentActivity
      SentActivityValue
      SurveyName
      SurveyNumber
      AmendedBy
      AmendedOn
    End Enum

'--------------------------------------------------
'Required overrides for the class
'--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)
        .Add("long_description", CDBField.FieldTypes.cftMemo)
        .Add("notes", CDBField.FieldTypes.cftMemo)
        .Add("responded_activity")
        .Add("responded_activity_value")
        .Add("sent_activity")
        .Add("sent_activity_value")
        .Add("survey_name")
        .Add("survey_number", CDBField.FieldTypes.cftLong)

        .Item(SurveyFields.SurveyNumber).PrimaryKey = True
        .SetControlNumberField(SurveyFields.SurveyNumber, "SU")
      End With
    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "s"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "surveys"
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
    Public ReadOnly Property AmendedBy() As String
      Get
        Return mvClassFields(SurveyFields.AmendedBy).Value
      End Get
    End Property
    Public ReadOnly Property AmendedOn() As String
      Get
        Return mvClassFields(SurveyFields.AmendedOn).Value
      End Get
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(SurveyFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(SurveyFields.CreatedOn).Value
      End Get
    End Property
    Public ReadOnly Property LongDescription() As String
      Get
        Return mvClassFields(SurveyFields.LongDescription).Value
      End Get
    End Property
    Public ReadOnly Property Notes() As String
      Get
        Return mvClassFields(SurveyFields.Notes).Value
      End Get
    End Property
    Public ReadOnly Property RespondedActivity() As String
      Get
        Return mvClassFields(SurveyFields.RespondedActivity).Value
      End Get
    End Property
    Public ReadOnly Property RespondedActivityValue() As String
      Get
        Return mvClassFields(SurveyFields.RespondedActivityValue).Value
      End Get
    End Property
    Public ReadOnly Property SentActivity() As String
      Get
        Return mvClassFields(SurveyFields.SentActivity).Value
      End Get
    End Property
    Public ReadOnly Property SentActivityValue() As String
      Get
        Return mvClassFields(SurveyFields.SentActivityValue).Value
      End Get
    End Property
    Public ReadOnly Property SurveyName() As String
      Get
        Return mvClassFields(SurveyFields.SurveyName).Value
      End Get
    End Property
    Public ReadOnly Property SurveyNumber() As Integer
      Get
        Return mvClassFields(SurveyFields.SurveyNumber).IntegerValue
      End Get
    End Property
#End Region

    Public Function CreateInstance(ByVal pEnv As CDBEnvironment) As CARERecord Implements IRecordCreate.CreateInstance
      Return New Survey(pEnv)
    End Function

    Public Shared Function CreateInstance(ByVal pEnv As CDBEnvironment, ByVal pParameterList As CDBParameters) As Survey
      Dim vSurvey As New Survey(pEnv)
      vSurvey.ValidateParameters(pParameterList)
      vSurvey.Init(pParameterList)
      Return vSurvey
    End Function

    Public Function DuplicateSurvey(ByVal pEnv As CDBEnvironment, ByVal pNewSurveyName As String) As Integer
      Try
        'Checking whether the new survey name is already existing
        Dim vNewSurveyWhereFields As New CDBFields
        Dim vNewSurveyNumber As Integer
        Dim vCount As Integer
        vNewSurveyWhereFields.Add("survey_name", pNewSurveyName)
        vCount = pEnv.Connection.GetCount("surveys", vNewSurveyWhereFields)
        If vCount = 0 Then
          Dim vOldSurveyQuestion As New SurveyQuestion(pEnv)
          Dim vNewSurveyQuestion As SurveyQuestion
          Dim vOldSurveyVersion As New SurveyVersion(pEnv)
          Dim vNewSurveyVersion As SurveyVersion
          Dim vOldSurveyContactGroup As New SurveyContactGroup(pEnv)
          Dim vNewSurveyContactGroup As SurveyContactGroup
          Dim vOldSurveyAnswer As New SurveyAnswer(pEnv)
          Dim vNewSurveyAnswer As SurveyAnswer

          'Retrieve all existing data before we start duplicating the Survey
          Dim vWhereFields As New CDBFields(New CDBField("survey_number", SurveyNumber))
          Dim vSurveyVersionList As List(Of SurveyVersion) = vOldSurveyVersion.GetList(Of SurveyVersion)(vOldSurveyVersion, vWhereFields)
          Dim vSurveyGroupList As List(Of SurveyContactGroup) = vOldSurveyContactGroup.GetList(Of SurveyContactGroup)(vOldSurveyContactGroup, vWhereFields)
          Dim vSurveyQuestionList As List(Of SurveyQuestion) = vOldSurveyQuestion.GetList(Of SurveyQuestion)(vOldSurveyQuestion, vWhereFields)
          For Each vSurveyQuestion As SurveyQuestion In vSurveyQuestionList
            vSurveyQuestion.GetSurveyAnswers()
          Next

          'Start a transaction
          Dim vTrans As Boolean = pEnv.Connection.StartTransaction()

          'Duplicating Survey
          Dim vNewSurvey As New Survey(pEnv)
          vNewSurvey.Clone(Me)
          Dim vNewSurveyList As New CDBParameters
          vNewSurveyList.Add("SurveyName", pNewSurveyName)
          vNewSurvey.Update(vNewSurveyList)
          vNewSurvey.Save()
          vNewSurveyNumber = vNewSurvey.SurveyNumber

          vNewSurveyList = New CDBParameters
          vNewSurveyList.Add("SurveyNumber", vNewSurveyNumber)

          'Duplicating All Survey Versions
          For Each vSurveyVersion As SurveyVersion In vSurveyVersionList
            'Clone the survey questions
            vNewSurveyVersion = New SurveyVersion(pEnv)
            vNewSurveyVersion.Clone(vSurveyVersion)
            vNewSurveyVersion.Update(vNewSurveyList)
            vNewSurveyVersion.Save()
          Next

          'Duplicating All Survey Contact Groups
          For Each vSurveyGroup As SurveyContactGroup In vSurveyGroupList
            'Clone the survey contact groups
            vNewSurveyContactGroup = New SurveyContactGroup(pEnv)
            Dim vParams As CDBParameters = vSurveyGroup.GetAddParameters()
            vParams("SurveyNumber").Value = vNewSurveyList("SurveyNumber").Value
            vNewSurveyContactGroup.Create(vParams)
            vNewSurveyContactGroup.Save()
          Next

          'Duplicating All Survey Questions
          Dim vSurveyQuestionParams As CDBParameters
          For Each vSurveyQuestion As SurveyQuestion In vSurveyQuestionList
            Dim vSurveyAnswerList As List(Of SurveyAnswer) = vSurveyQuestion.GetSurveyAnswers()
            'Clone the survey questions
            vNewSurveyQuestion = New SurveyQuestion(pEnv)
            vNewSurveyQuestion.Clone(vSurveyQuestion)
            vNewSurveyQuestion.Update(vNewSurveyList)
            vNewSurveyQuestion.Save()
            vSurveyQuestionParams = New CDBParameters
            vSurveyQuestionParams.Add("SurveyQuestionNumber", vNewSurveyQuestion.SurveyQuestionNumber.ToString)
            'Duplicating All Survey Answers
            For Each vSurveyAnswer As SurveyAnswer In vSurveyAnswerList
              'Clone the survey answers
              vNewSurveyAnswer = New SurveyAnswer(pEnv)
              vNewSurveyAnswer.Clone(vSurveyAnswer)
              vNewSurveyAnswer.Update(vSurveyQuestionParams)
              vNewSurveyAnswer.Save()
            Next
          Next

          If vTrans Then pEnv.Connection.CommitTransaction()

          Return vNewSurveyNumber
        Else
          RaiseError(DataAccessErrors.daeRecordExists, "New Survey Name")
        End If
      Catch vEX As Exception
        Throw vEX
      End Try
    End Function
    ''' <summary>
    ''' Validate Parameters
    ''' </summary>
    ''' <param name="pParameterList"></param>
    ''' <remarks></remarks>
    Public Sub ValidateParameters(ByVal pParameterList As CDBParameters)
      ValidateSurveyNumber(pParameterList)
    End Sub
    ''' <summary>
    ''' Validate the Survey Number if it is present.
    ''' </summary>
    ''' <param name="pParameterList"></param>
    ''' <remarks>Not present when creating.</remarks>
    Public Sub ValidateSurveyNumber(ByVal pParameterList As CDBParameters)
      If pParameterList.Exists("SurveyNumber") Then
        Dim vInteger As Integer
        If pParameterList("SurveyNumber").Value.Length = 0 Or (pParameterList("SurveyNumber").Value.Length > 0 AndAlso Not Integer.TryParse(pParameterList("SurveyNumber").Value, vInteger)) Then
          RaiseError(DataAccessErrors.daeSurveyNumberInvalid)
        End If
      End If
    End Sub

    Protected Overrides Sub PreValidateCreateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the create methods
      MyBase.PreValidateCreateParameters(pParameterList)
      ValidateParameters(pParameterList)
    End Sub

    Protected Overrides Sub PreValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the update methods
      MyBase.PreValidateUpdateParameters(pParameterList)
      ValidateParameters(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateCreateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the create methods
      MyBase.PostValidateCreateParameters(pParameterList)
    End Sub

    Protected Overrides Sub PostValidateUpdateParameters(ByVal pParameterList As CDBParameters)
      'Add code here to validate parameters passed to the update methods
      MyBase.PostValidateUpdateParameters(pParameterList)
    End Sub

  End Class
End Namespace
