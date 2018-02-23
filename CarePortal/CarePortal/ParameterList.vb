Imports System.Xml
Imports System.Web.Security
Imports System.IO

Public Class ParameterList
  Inherits SortedList

  Sub New()
    MyBase.New()
  End Sub

  'ParameterList is passed an XML string containing a result set
  'from the care business objects layer. It is assumed to contain a single record
  'The string is then parsed and a list of the names and values is returned

  Sub New(ByVal pXMLString As String)
    'If the string has some data in it we assume it is a single record
    MyBase.new()
    FillFromXMLString(pXMLString)
  End Sub

  Sub New(ByVal pContext As HttpContext)
    MyBase.New()
    AddConectionData(pContext)
  End Sub

  Private Sub AddIdentity(ByVal pUserLogname As String, ByVal pDatabase As String)
    MyBase.Add("UserLogname", pUserLogname)
    MyBase.Add("Database", pDatabase)
  End Sub

  Private Sub AddIdentity(ByVal pIdentity As FormsIdentity)
    MyBase.Add("UserLogname", pIdentity.Name)
    If pIdentity.Ticket.UserData.Length > 0 Then
      Dim vString As String = pIdentity.Ticket.UserData
      Dim vItems As String() = vString.Split("|"c)
      MyBase.Add("Database", vItems(0))
      If vItems.Length > 3 AndAlso vItems(3).Length > 0 Then
        'Got a department so must be a CARE user
      Else
        MyBase.Add("UserID", vItems(1))
      End If
    End If
  End Sub

  Public Sub FillFromValueList(ByVal pValues As String)
    If pValues.Length > 0 Then
      Dim vItems() As String = pValues.Split(","c)
      For Each vItem As String In vItems
        Dim vValues() As String = vItem.Split("="c)
        If vValues.Length = 2 Then
          Me.Item(vValues(0)) = vValues(1)
        End If
      Next
    End If
  End Sub

  Public Sub FillFromXMLString(ByVal pXMLString As String)
    Dim vStartIndex As Integer

    If Len(pXMLString) > 0 Then
      Dim vDoc As New System.Xml.XmlDocument
      vDoc.LoadXml(pXMLString)
      Dim vXMLRoot As XmlNode
      Do
        vXMLRoot = vDoc.ChildNodes(vStartIndex)           'Parameters
        vStartIndex = vStartIndex + 1
      Loop While vXMLRoot.NodeType = XmlNodeType.XmlDeclaration

      If vXMLRoot.ChildNodes.Count > 0 Then
        Dim vNode As XmlNode
        Dim vError As Boolean = False
        Dim vErrorNumber As Long
        Dim vErrorSource As String = ""
        Dim vErrorMethod As String = ""
        Dim vErrorModule As String = ""

        If vXMLRoot.Name = "Result" Then
          For Each vNode In vXMLRoot.ChildNodes
            With vNode
              MyBase.Add(.Name, Trim(.InnerText))
              If .Name = "ErrorMessage" Then vError = True
              If vError Then
                Select Case .Name
                  Case "ErrorNumber"
                    vErrorNumber = CInt(.InnerText)
                  Case "Source"
                    vErrorSource = .InnerText
                  Case "Method"
                    vErrorMethod = .InnerText
                  Case "Module"
                    vErrorModule = .InnerText
                End Select
              End If
            End With
          Next
          If vError = True Then Throw New CareException(MyBase.Item("ErrorMessage").ToString, vErrorNumber, vErrorSource, vErrorModule, vErrorMethod)
        Else
          Dim vList As XmlNodeList
          vList = vDoc.GetElementsByTagName("DataRow")
          For Each vNode In vList(0).ChildNodes
            MyBase.Add(vNode.Name, vNode.InnerText)
          Next
        End If
      End If
    End If
  End Sub

  Public Function XMLResultString() As String
    Dim vStream As New IO.MemoryStream
    Dim vWriter As New XmlTextWriter(vStream, Nothing)
    Dim vItem As DictionaryEntry

    With vWriter
      .WriteStartElement("Result")
      For Each vItem In Me
        .WriteElementString(vItem.Key.ToString, vItem.Value.ToString)
      Next
      .WriteEndElement()
      .Flush()
    End With
    Dim vReader As New System.IO.StreamReader(vStream)
    vStream.Position = 0
    Return vReader.ReadToEnd()
  End Function

  Public Function XMLParameterString() As String
    Dim vStream As New IO.MemoryStream
    Dim vWriter As New XmlTextWriter(vStream, Nothing)
    Dim vItem As DictionaryEntry

    Debug.Assert(Me.ContainsKey("Database"))
    With vWriter
      .WriteStartElement("Parameters")
      For Each vItem In Me
        .WriteElementString(vItem.Key.ToString, vItem.Value.ToString)
      Next
      .WriteEndElement()
      .Flush()
    End With
    Dim vReader As New System.IO.StreamReader(vStream)
    vStream.Position = 0
    XMLParameterString = vReader.ReadToEnd()
  End Function

  Public ReadOnly Property OptionalValue(ByVal pName As String) As String
    Get
      If MyBase.ContainsKey(pName) Then
        Return MyBase.Item(pName).ToString
      Else
        Return ""
      End If
    End Get
  End Property

  Public Function ToCSVFile() As String
    Dim vFileName As String = My.Computer.FileSystem.GetTempFileName()
    Dim vStreamWriter As StreamWriter = New StreamWriter(vFileName, False)
    Dim vAddSeparator As Boolean
    For Each vItem As DictionaryEntry In Me
      If vAddSeparator Then vStreamWriter.Write(",")
      vStreamWriter.Write(vItem.Key)
      vAddSeparator = True
    Next
    vAddSeparator = False
    vStreamWriter.WriteLine()
    For Each vItem As DictionaryEntry In Me
      If vAddSeparator Then vStreamWriter.Write(",")
      vStreamWriter.Write("""")
      vStreamWriter.Write(vItem.Value)
      vStreamWriter.Write("""")
      vAddSeparator = True
    Next
    vStreamWriter.WriteLine()
    vStreamWriter.Close()
    Return vFileName
  End Function

  Public Sub AddConectionData(ByVal pContext As HttpContext)
    If pContext.User.Identity.IsAuthenticated Then
      If TypeOf (pContext.User.Identity) Is System.Security.Principal.WindowsIdentity Then
        AddIdentity(pContext.Session("UserLogname").ToString, pContext.Session("Database").ToString)
      Else
        AddIdentity(CType(pContext.User.Identity, FormsIdentity))
      End If
    Else
      MyBase.Add("Database", DataHelper.Database)
    End If
  End Sub

End Class

''' <summary>
''' Portal exceptions are not recorded in the error_logs table and the message is shown to the end user
''' </summary>
''' <remarks></remarks>
Friend Class PortalAccessException
  Inherits CareException
End Class

Friend Class PortalAccessOrganisationException
  Inherits CareException
End Class

Friend Class CareException
  Inherits System.Exception

  Private mvErrorNumber As ErrorNumbers

  Public Enum ErrorNumbers
    enDuplicateRecord = 1048
    enInsufficientRights = 1049
    enParameterInvalidValue = 1050
    enNoSelectionData = 1051
    enNoContactItemFound = 1052
    enInsufficientStock = 1053
    enCannotSetAsDefault = 1054
    enGiftAidDeclarationsOverlap = 1055
    enALBACSVerify = 1056
    enAppointmentConflict = 1120
    enPasswordPreviouslyUsedInHistory = 1555
    enUserNameAlreadyInUse = 10248
    enPositionDatesExceedSiteDates = 10282
    enSoleMembership = 10368
    enSingleMembershipOnly = 10434
    enUserDoesNotExist = 10454
    enInvalidEmailAddress = 10455
  End Enum

  Sub New()
    MyBase.New() ' pass control to the parent constructor of the same signature
  End Sub

  Sub New(ByVal message As String, ByVal pErrorNumber As Long, ByVal pSource As String, ByVal pModule As String, ByVal pMethod As String)
    MyBase.New(message) ' pass control to the parent constructor of the same signature
    Select Case pErrorNumber
      Case 10246
        mvErrorNumber = ErrorNumbers.enParameterInvalidValue
      Case 10250, 1226
        mvErrorNumber = ErrorNumbers.enDuplicateRecord
      Case 10253
        mvErrorNumber = ErrorNumbers.enInsufficientRights
      Case 1104
        mvErrorNumber = ErrorNumbers.enNoSelectionData
      Case 10266
        mvErrorNumber = ErrorNumbers.enInsufficientStock
      Case 1158
        mvErrorNumber = ErrorNumbers.enCannotSetAsDefault
      Case 10319      '
        mvErrorNumber = ErrorNumbers.enGiftAidDeclarationsOverlap
      Case 1107
        mvErrorNumber = ErrorNumbers.enALBACSVerify
      Case 10282
        mvErrorNumber = ErrorNumbers.enPositionDatesExceedSiteDates
      Case 10368
        mvErrorNumber = ErrorNumbers.enSoleMembership
      Case 10434
        mvErrorNumber = ErrorNumbers.enSingleMembershipOnly
      Case 1555
        mvErrorNumber = ErrorNumbers.enPasswordPreviouslyUsedInHistory
      Case 3
        mvErrorNumber = ErrorNumbers.enInvalidEmailAddress
      Case Else
        mvErrorNumber = CType(pErrorNumber, CarePortal.CareException.ErrorNumbers)
    End Select
    MyBase.Source &= pSource & "." & pModule & "." & pMethod
  End Sub

  Sub New(ByVal message As String, ByVal ErrorNumber As ErrorNumbers)
    MyBase.New(message) ' pass control to the parent constructor
    mvErrorNumber = ErrorNumber
  End Sub

  Sub New(ByVal message As String)
    MyBase.New(message) ' pass control to the parent constructor of the same signature
  End Sub

  Sub New(ByVal message As String, ByVal inner As Exception)
    MyBase.New(message, inner) ' pass control to the parent constructor of the same signature
  End Sub

  Overrides ReadOnly Property Message() As String
    Get
      Select Case mvErrorNumber
        Case ErrorNumbers.enDuplicateRecord
          Message = "A Record Already Exists with these Values"
        Case ErrorNumbers.enNoSelectionData
          Message = "At least one field must contain information to base the selection on"
        Case ErrorNumbers.enNoContactItemFound
          Message = "Contact Data Could not be retrieved for Editing"
        Case Else
          Message = MyBase.Message
      End Select
    End Get
  End Property

  ReadOnly Property ErrorNumber() As CareException.ErrorNumbers
    Get
      ErrorNumber = mvErrorNumber
    End Get
  End Property
End Class