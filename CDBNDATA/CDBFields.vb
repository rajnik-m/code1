Namespace Data

  Partial Public Class CDBFields
    Inherits CollectionList(Of CDBField)

    Private mvTableAlias As String

    Public Sub New()
      MyBase.New(1)
    End Sub

    Public Sub New(ByVal pField As CDBField)
      MyBase.New(1)
      MyBase.Add(pField.Name, pField)
    End Sub

    ''' <summary>
    ''' Create a new collection of fields from the an array of fields.
    ''' </summary>
    ''' <param name="pFields">An array containing the <see cref="CDBField"/> objects that should form the
    ''' initial content of the collection.</param>
    Public Sub New(ByVal pFields As CDBField())
      MyBase.New(1)
      For Each vField As CDBField In pFields
        Add(vField)
      Next vField
    End Sub

    Public Overloads Function Add(ByVal pField As CDBField) As CDBField
      MyBase.Add(pField.Name, pField)
      Return pField
    End Function

    Public Overloads Function Add(ByVal pName As String) As CDBField
      Dim vField As New CDBField(pName)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pType As CDBField.FieldTypes) As CDBField
      Dim vField As New CDBField(pName, pType)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pType As CDBField.FieldTypes, ByVal pValue As String) As CDBField
      Dim vField As CDBField = New CDBField(pName, pType, pValue)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pType As CDBField.FieldTypes, ByVal pValue As Integer) As CDBField
      Dim vField As CDBField = New CDBField(pName, pType, pValue.ToString)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As String) As CDBField
      Dim vField As CDBField = New CDBField(pName, pValue)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As String, ByVal pWhereOperator As CDBField.FieldWhereOperators) As CDBField
      Dim vField As CDBField = New CDBField(pName, pValue, pWhereOperator)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As Integer) As CDBField
      Dim vField As CDBField = New CDBField(pName, pValue)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As Date) As CDBField
      Dim vField As CDBField = New CDBField(pName, CDBField.FieldTypes.cftDate, pValue.ToString(CAREDateFormat))
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As Date, ByVal pFWO As CDBField.FieldWhereOperators) As CDBField
      Dim vField As CDBField = New CDBField(pName, CDBField.FieldTypes.cftDate, pValue.ToString(CAREDateFormat), pFWO)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pValue As Integer, ByVal pWhereOperator As CDBField.FieldWhereOperators) As CDBField
      Dim vField As CDBField = New CDBField(pName, pValue, pWhereOperator)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function Add(ByVal pName As String, ByVal pType As CDBField.FieldTypes, ByVal pValue As String, ByVal pWhereOperator As CDBField.FieldWhereOperators) As CDBField
      Dim vField As CDBField = New CDBField(pName, pType, pValue, pWhereOperator)
      MyBase.Add(pName, vField)
      Return vField
    End Function

    Public Overloads Function AddJoin(ByVal pField1 As String, ByVal pField2 As String) As CDBField
      Dim vField As CDBField = New CDBField(pField1, CDBField.FieldTypes.cftInteger, pField2)
      MyBase.Add(pField1, vField)
      Return vField
    End Function

    Public Sub AddClientDeptLogname(ByVal pClientCode As String, ByVal pDepartment As String, ByVal pLogname As String)
      If pClientCode.Length = 0 Then
        Add("client")
      Else
        Add("client", pClientCode, CDBField.FieldWhereOperators.fwoOpenBracket)
        Add("client#2", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      If pDepartment.Length = 0 Then
        Add("department")
      Else
        Add("department", pDepartment, CDBField.FieldWhereOperators.fwoOpenBracket)
        Add("department#2", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
      If pLogname.Length = 0 Then
        Add("logname")
      Else
        Add("logname", pLogname, CDBField.FieldWhereOperators.fwoOpenBracket)
        Add("logname#2", "", CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
      End If
    End Sub

    Public Property TableAlias() As String
      Get
        TableAlias = mvTableAlias
      End Get
      Set(ByVal pValue As String)
        mvTableAlias = pValue
      End Set
    End Property

    Public Sub AddAmendedOnBy(ByVal pAmendedBy As String, Optional ByVal pAmendedOn As String = "")
      Add("amended_by", CDBField.FieldTypes.cftCharacter, pAmendedBy)
      If IsDate(pAmendedOn) Then
        Add("amended_on", CDBField.FieldTypes.cftDate, pAmendedOn)
      Else
        Add("amended_on", CDBField.FieldTypes.cftDate, TodaysDate)
      End If
    End Sub

    Public Function Exists(ByVal pName As String) As Boolean
      Return Me.ContainsKey(pName)
    End Function

    Public Function ExistsInsensitive(ByVal pName As String) As Boolean
      Return Me.ContainsKey(pName) OrElse Me.ContainsKey(pName.ToLower) OrElse Me.ContainsKey(pName.ToUpper)
    End Function

    Public Function FieldExists(ByVal pName As String) As CDBField
      If Me.ContainsKey(pName) Then
        Return Me.Item(pName)
      Else
        Return New CDBField("", CDBField.FieldTypes.cftCharacter)
      End If
    End Function

    Public Function FieldExistsInsensitive(ByVal pName As String) As CDBField
      If ExistsInsensitive(pName) Then
        Return ItemInsesitive(pName)
      Else
        Return New CDBField("", CDBField.FieldTypes.cftCharacter)
      End If
    End Function

    Default Public Overrides ReadOnly Property Item(ByVal pKey As String) As CDBField
      Get
        'Dim vPos As Integer = pKey.IndexOf(".")
        'If vPos >= 0 Then pKey = pKey.Substring(vPos + 1)
        Return MyBase.Item(pKey)
      End Get
    End Property

    Public ReadOnly Property ItemInsesitive(ByVal pKey As String) As CDBField
      Get
        Dim vResult As CDBField = Nothing
        If Me.ContainsKey(pKey) Then
          vResult = MyBase.Item(pKey)
        ElseIf Me.ContainsKey(pKey.ToLower) Then
          vResult = MyBase.Item(pKey.ToLower)
        ElseIf Me.ContainsKey(pKey.ToUpper) Then
          vResult = MyBase.Item(pKey.ToUpper)
        Else
          vResult = MyBase.Item(pKey)
        End If
        Return vResult
      End Get
    End Property

    Public Sub WriteToMessageQueue()
      If Me.ContainsKey("message_queue_name") Then
        'Get the message queue name then build the XML
        Dim vMessageQueueName As String = Me("message_queue_name").Value
        Dim vStream As New IO.MemoryStream
        Dim vSettings As New Xml.XmlWriterSettings()
        vSettings.NewLineHandling = Xml.NewLineHandling.None
        Dim vWriter As Xml.XmlWriter = Xml.XmlWriter.Create(vStream, vSettings)
        vWriter.WriteProcessingInstruction("xml", "version=""1.0""")
        vWriter.WriteWhitespace(Environment.NewLine)
        vWriter.WriteStartElement("AmendmentHistory")
        For Each vField As CDBField In Me
          Select Case vField.Name
            Case "message_queue_name"
              'Ignore
            Case "data_values"
              Dim vEndSection As Boolean = False
              Dim vItems() As String = Split(vField.Value, Chr(22))
              For vIndex As Integer = 0 To vItems.Length - 1
                If vItems(vIndex).Trim = "OLD" Then
                  If vEndSection Then vWriter.WriteEndElement()
                  vWriter.WriteStartElement("OldValues")
                  vEndSection = True
                ElseIf vItems(vIndex).Trim = "NEW" Then
                  If vEndSection Then vWriter.WriteEndElement()
                  vWriter.WriteStartElement("NewValues")
                  vEndSection = True
                ElseIf Mid(vItems(vIndex), 3) = "NEW" Then
                  If vEndSection Then vWriter.WriteEndElement()
                  vWriter.WriteStartElement("NewValues")
                  vEndSection = True
                Else
                  Dim vValues() As String = Split(vItems(vIndex), ":")
                  If vValues.Length > 1 Then
                    vWriter.WriteAttributeString(ProperName(vValues(0)), "", vValues(1))
                  End If
                End If
              Next
              If vEndSection Then vWriter.WriteEndElement()
            Case Else
              vWriter.WriteElementString(ProperName(vField.Name), "", vField.Value)
          End Select
        Next
        vWriter.WriteEndElement()
        vWriter.Flush()
        Dim vReader As New System.IO.StreamReader(vStream)
        vStream.Position = 0
        Dim vXML As String = vReader.ReadToEnd()
        Dim vMessage As New System.Messaging.Message(vXML)
        vMessage.Label = "AmendmentHistory"

        Dim vMessageQueues() As String = vMessageQueueName.Split("|".ToCharArray, StringSplitOptions.RemoveEmptyEntries)
        For Each vQueueName As String In vMessageQueues
          Dim vMessageQueue As System.Messaging.MessageQueue = New System.Messaging.MessageQueue(vQueueName)
          vMessageQueue.Send(vMessage)
        Next
      End If
    End Sub

  End Class

End Namespace