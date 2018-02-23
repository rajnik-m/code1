Imports System.Text.RegularExpressions
Imports System.Net.Mail

Public Module Utilities
  Friend Const PasswordEncrypted As String = "encrypted"
  Public Function GetCustomPageElement(ByVal pItem As String) As String
    'pItem will be header or footer
    Dim vItemContent As String = ""
    Try
      If HttpContext.Current.Cache.Item(pItem) Is Nothing Then
        Dim vCustomConfig As System.Xml.XmlDocument = GetCustomConfig()
        Dim vNode As System.Xml.XmlNode = vCustomConfig.SelectSingleNode("CustomConfiguration/CustomURL")
        If vNode IsNot Nothing Then
          Dim vWebClient As Net.WebClient = New Net.WebClient()
          ' Read web page HTML to byte array
          Dim vURL As String = vNode.InnerText
          Dim vContactNumber As Integer
          Dim vQueryString As New StringBuilder
          Dim vFriendlyUrl As String = String.Empty

          If HttpContext.Current.Session("UserContactNumber") IsNot Nothing Then vContactNumber = IntegerValue(HttpContext.Current.Session("UserContactNumber").ToString)

          'Add all the values in the query string pass it across as custom URL to be used by thirdparty website 
          Dim vKeyCount As Integer
          If HttpContext.Current.Request.QueryString.Count > 0 Then
            For Each vKey As String In HttpContext.Current.Request.QueryString.AllKeys
              vQueryString.Append(vKey + "=" + HttpContext.Current.Request.QueryString(vKey))
              If vQueryString.Length > 0 AndAlso vKeyCount < HttpContext.Current.Request.QueryString.AllKeys.GetUpperBound(0) Then vQueryString.Append("+")
              vKeyCount += 1
            Next
          End If

          If HttpContext.Current.Session("FriendlyUrl") IsNot Nothing Then vFriendlyUrl = HttpContext.Current.Session("FriendlyUrl").ToString
          If vURL.Contains("{3}") Then
            vURL = String.Format(vURL, pItem, vContactNumber.ToString, vQueryString.ToString, vFriendlyUrl)
          ElseIf vURL.Contains("{2}") Then
            vURL = String.Format(vURL, pItem, vContactNumber.ToString, vQueryString.ToString)
          ElseIf vURL.Contains("{1}") Then
            vURL = String.Format(vURL, pItem, vContactNumber.ToString)
          Else
            vURL = String.Format(vNode.InnerText, pItem)
          End If
          Dim vPageHTMLBytes() As Byte = vWebClient.DownloadData(vURL)
          ' Convert result from byte array to string
          Dim vUTF8 As UTF8Encoding = New UTF8Encoding()
          vItemContent = vUTF8.GetString(vPageHTMLBytes)
        Else
          Dim vPath As String = HttpContext.Current.Server.MapPath(String.Format("custom\{0}.htm", pItem))
          If My.Computer.FileSystem.FileExists(vPath) Then
            vItemContent = My.Computer.FileSystem.ReadAllText(vPath)
          Else
            vItemContent = ""
          End If
          HttpContext.Current.Cache.Insert(pItem, vItemContent, New CacheDependency(vPath))
        End If
      Else
        vItemContent = CStr(HttpContext.Current.Cache.Item(pItem))
      End If
    Catch vEx As Exception
      Dim vList As New ParameterList
      vList("ErrorNumber") = 0
      vList("ErrorSource") = "GetCustomPageElement"
      vList("WebPageNumber") = IIf(HttpContext.Current.Request.QueryString("PN") IsNot Nothing, HttpContext.Current.Request.QueryString("PN"), 0)
      vList("ErrorMessage") = vEx.Message
      vList("StackTrace") = vEx.StackTrace
      'Record Error in Database.
      Dim vResult As DataTable = DataHelper.AddErrorLog(vList)
    End Try
    Return vItemContent
  End Function


  Public Sub LogException(ByVal pException As Exception)
    Try
      Dim vEventSource As String = "NGPortal"
      If Not EventLog.SourceExists(vEventSource) Then
        EventLog.CreateEventSource(vEventSource, vEventSource)
      End If
      If EventLog.SourceExists(vEventSource) Then
        EventLog.WriteEntry(vEventSource, pException.Message & vbCrLf & "Source: " & pException.Source & "Stack Trace: " & pException.StackTrace, EventLogEntryType.Error)
      End If
    Catch ex As Exception
      Debug.Print("No rights to the event log")
    End Try
  End Sub

  Public Function GetDataTable(ByVal pResult As String, Optional ByVal pGetData As Boolean = False) As DataTable
    Dim vStream As System.IO.StringReader = New System.IO.StringReader(pResult)
    Dim vDataSet As New DataSet
    Dim vTable As DataTable = Nothing
    vDataSet.ReadXml(vStream, XmlReadMode.Auto)
    If vDataSet.Tables.Count > 0 Then
      If vDataSet.Tables(0).Columns.Contains("ErrorMessage") Then
        Throw New CareException(vDataSet.Tables(0).Rows(0).Item("ErrorMessage").ToString)
      Else
        If pGetData Then
          If vDataSet.Tables.Count > 1 AndAlso vDataSet.Tables(1) IsNot Nothing Then
            vTable = vDataSet.Tables(1)
          Else
            vTable = Nothing
          End If
        Else
          vTable = vDataSet.Tables(0)
        End If
        End If
    End If
    Return vTable
  End Function

  Public Function IntegerValue(ByVal pString As String) As Integer
    If pString.Length > 0 Then
      Return CInt(pString)
    Else
      Return 0
    End If
  End Function

  Public Function DoubleValue(ByVal pString As String) As Double
    If pString.Length = 0 Then
      Return 0
    Else
      Return CDbl(pString)
    End If
  End Function

  Public Function FixTwoPlaces(ByVal pValue As String) As Double
    Dim vDouble As Double
    Double.TryParse(pValue, vDouble)
    Return CDbl(vDouble.ToString("F"))
  End Function

  Public Function StripHTML(ByVal pSource As String) As String
    Dim vResult As String
    ' Remove HTML Development formatting 
    ' Replace line breaks with space 
    ' because browsers inserts space 
    vResult = pSource.Replace(vbCr, " ")
    ' Replace line breaks with space 
    ' because browsers inserts space 
    vResult = vResult.Replace(vbLf, " ")
    ' Remove step-formatting 
    vResult = vResult.Replace(vbTab, String.Empty)
    ' Remove repeating spaces because browsers ignore them 
    vResult = Regex.Replace(vResult, "( )+", " ")

    ' Remove the header (prepare first by clearing attributes) 
    vResult = Regex.Replace(vResult, "<( )*head([^>])*>", "<head>", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(<( )*(/)( )*head( )*>)", "</head>", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(<head>).*(</head>)", String.Empty, RegexOptions.IgnoreCase)

    ' remove all scripts (prepare first by clearing attributes) 
    vResult = Regex.Replace(vResult, "<( )*script([^>])*>", "<script>", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(<( )*(/)( )*script( )*>)", "</script>", RegexOptions.IgnoreCase)
    'result = Regex.Replace(result, 
    ' @"(<script>)([^(<script>\.</script>)])*(</script>)", 
    ' string.Empty, 
    ' RegexOptions.IgnoreCase); 
    vResult = Regex.Replace(vResult, "(<script>).*(</script>)", String.Empty, RegexOptions.IgnoreCase)

    ' remove all styles (prepare first by clearing attributes) 
    vResult = Regex.Replace(vResult, "<( )*style([^>])*>", "<style>", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(<( )*(/)( )*style( )*>)", "</style>", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(<style>).*(</style>)", String.Empty, RegexOptions.IgnoreCase)

    ' insert tabs in spaces of <td> tags 
    vResult = Regex.Replace(vResult, "<( )*td([^>])*>", vbTab, RegexOptions.IgnoreCase)

    ' insert line breaks in places of <BR> and <LI> tags 
    vResult = Regex.Replace(vResult, "<( )*br( )*>", vbCr, RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "<( )*li( )*>", vbCr, RegexOptions.IgnoreCase)

    ' insert line paragraphs (double line breaks) in place 
    ' if <P>, <DIV> and <TR> tags 
    vResult = Regex.Replace(vResult, "<( )*div([^>])*>", vbCr & vbCr, RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "<( )*tr([^>])*>", vbCr & vbCr, RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "<( )*p([^>])*>", vbCr & vbCr, RegexOptions.IgnoreCase)

    ' Remove remaining tags like <a>, links, images, 
    ' comments etc - anything that's enclosed inside < > 
    vResult = Regex.Replace(vResult, "<[^>]*>", String.Empty, RegexOptions.IgnoreCase)

    ' replace special characters: 
    vResult = Regex.Replace(vResult, " ", " ", RegexOptions.IgnoreCase)

    vResult = Regex.Replace(vResult, "&bull;", " * ", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&lsaquo;", "<", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&rsaquo;", ">", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&trade;", "(tm)", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&frasl;", "/", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&lt;", "<", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&gt;", ">", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&copy;", "(c)", RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "&reg;", "(r)", RegexOptions.IgnoreCase)
    ' Remove all others. More can be added, see 
    ' http://hotwired.lycos.com/webmonkey/reference/special_characters/ 
    vResult = Regex.Replace(vResult, "&(.{2,6});", String.Empty, RegexOptions.IgnoreCase)

    ' make line breaking consistent 
    vResult = vResult.Replace(vbLf, vbCr)

    ' Remove extra line breaks and tabs: 
    ' replace over 2 breaks with 2 and over 4 tabs with 4. 
    ' Prepare first to remove any whitespaces in between 
    ' the escaped characters and remove redundant tabs in between line breaks 
    vResult = Regex.Replace(vResult, "(" & vbCr & ")( )+(" & vbCr & ")", vbCr & vbCr, RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(" & vbTab & ")( )+(" & vbTab & ")", vbTab & vbTab, RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(" & vbTab & ")( )+(" & vbCr & ")", vbTab & vbCr, RegexOptions.IgnoreCase)
    vResult = Regex.Replace(vResult, "(" & vbCr & ")( )+(" & vbTab & ")", vbCr & vbTab, RegexOptions.IgnoreCase)
    ' Remove redundant tabs 
    vResult = Regex.Replace(vResult, "(" & vbCr & ")(" & vbTab & ")+(" & vbCr & ")", vbCr & vbCr, RegexOptions.IgnoreCase)
    ' Remove multiple tabs following a line break with just one tab 
    vResult = Regex.Replace(vResult, "(" & vbCr & ")(" & vbTab & ")+", vbCr & vbTab, RegexOptions.IgnoreCase)

    While vResult.StartsWith(vbCr)
      vResult = vResult.Substring(1)
    End While
    Return vResult
  End Function

  ''' <summary>
  ''' Checks if the requested url exists in the WhiteList node of the custom config file. 
  ''' if no custom config, no WhiteList node or an empty WhiteList node is present then no check is performed.
  ''' Redirect to any local address is allowed
  ''' </summary>
  ''' <param name="pURL"></param>
  ''' <remarks></remarks>
  Public Sub RedirectViaWhiteList(ByVal pURL As String)
    Dim vValid As Boolean = True
    'Any redirect to http://localhost should be accepted without checking the white list
    'urls that dont start with http are relative paths...skip them
    If pURL.StartsWith("http") AndAlso ((Not pURL.StartsWith("http://localhost")) OrElse (Not pURL.StartsWith("https://localhost"))) Then
      Dim vCustomConfig As System.Xml.XmlDocument = GetCustomConfig()
      Dim vNode As System.Xml.XmlNode = vCustomConfig.SelectSingleNode("CustomConfiguration/WhiteList")
      If vNode IsNot Nothing Then
        'if list has not values,then no checking performed
        If vNode.ChildNodes.Count > 0 Then
          vValid = False
          For Each vChildNode As System.Xml.XmlNode In vNode.ChildNodes
            If vChildNode.InnerText.StartsWith(pURL) Then
              vValid = True
              Exit For
            End If
          Next
        End If
      End If
    End If
    If vValid Then
      HttpContext.Current.Response.Redirect(pURL, True)
    Else
      Throw New CareException(String.Format("Invalid redirection attempted to {0}", pURL))
    End If
  End Sub

  Public Sub ValidateEmailAddress(ByVal pEmail As String)
    If Not String.IsNullOrWhiteSpace(pEmail) AndAlso
      BooleanValue(DataHelper.ControlValue(DataHelper.ControlTables.email_controls, DataHelper.ControlValues.force_smtp_address)) Then
      Try
        Dim mailAddress As New MailAddress(pEmail)
      Catch ex As Exception
        RaiseError(DataAccessErrors.daeInvalidEmailAddress, pEmail)
      End Try
    End If
  End Sub

  Public Function ValidateTelephoneNumber(ByVal vPhoneNumber As String) As Boolean
    Dim vValid As Boolean = True
    If Not String.IsNullOrWhiteSpace(vPhoneNumber) Then
      If vPhoneNumber.StartsWith("+") Then 'Remove leading + and any following space
        vPhoneNumber = vPhoneNumber.Substring(1).TrimStart
      End If
      If vPhoneNumber.Contains("(") Then  'Remove any Brackets
        vPhoneNumber = vPhoneNumber.Replace("(", "").Trim
        vPhoneNumber = vPhoneNumber.Replace(")", "").Trim
      End If
      vPhoneNumber = vPhoneNumber.Replace(" ", "")
      If Not IsNumeric(vPhoneNumber) Then
        vValid = False
      Else
        If vPhoneNumber.Contains(",") OrElse vPhoneNumber.StartsWith("-") OrElse vPhoneNumber.EndsWith("-") OrElse vPhoneNumber.StartsWith("+") OrElse vPhoneNumber.EndsWith("+") OrElse vPhoneNumber.Contains(".") Then
          vValid = False
        End If
      End If
    End If
    Return vValid
  End Function

  ''' <summary>
  ''' Reads the custom config file into the application cache. 
  ''' Subsequent calls will return the config stored in the cache as long as the config doesn't change
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetCustomConfig() As System.Xml.XmlDocument
    Dim vCustomConfig As System.Xml.XmlDocument = Nothing
    Dim vCache As Cache = HttpContext.Current.Cache
    If vCache.Item("CustomConfig") Is Nothing Then
      vCustomConfig = New System.Xml.XmlDocument
      Dim vConfigPath As String = HttpContext.Current.Server.MapPath("custom\custom.config")
      If My.Computer.FileSystem.FileExists(vConfigPath) Then
        vCustomConfig.Load(vConfigPath)
        vCache.Insert("CustomConfig", vCustomConfig, New CacheDependency(vConfigPath))
      Else
        vCache.Insert("CustomConfig", vCustomConfig)
      End If
    End If
    vCustomConfig = CType(vCache.Item("CustomConfig"), System.Xml.XmlDocument)
    Return vCustomConfig
  End Function

  Public Function GetCustomConfigItem(ByVal pItem As String) As String
    Return GetCustomConfigItem(pItem, False)
  End Function
  Public Function GetCustomConfigItem(ByVal pItem As String, ByVal pGetInnerText As Boolean) As String
    Dim vCustomConfig As System.Xml.XmlDocument = GetCustomConfig()
    Dim vNode As System.Xml.XmlNode = vCustomConfig.SelectSingleNode(pItem)
    If vNode IsNot Nothing Then
      If pGetInnerText Then
        Return vNode.InnerText
      Else
        Return vNode.InnerXml
      End If
    Else
      Return ""
    End If
  End Function

  ''' <summary>
  ''' Generates the Random Password with at least 6 characters or the length set by portal_password_min_length if this is more than 6.
  ''' The password has one upper case alpha character (no vowels or y), one number character (1-9) and the remaining characters are lower case alpha (no vowels or y)
  ''' </summary>
  ''' <returns>String</returns>
  ''' <remarks></remarks>
  Public Function GeneratePassword() As String
    Dim vDefaultPasswordLen As Integer = 6
    Dim vUpperCaseCharacter() As Char = "BCDFGHJKLMNPQRSTWXZ".ToCharArray
    Dim vLowerCaseCharacter() As Char = "bcdfgjkmnpqrstwxz".ToCharArray
    Dim vNumeric() As Char = "123456789".ToCharArray
    Dim vRand As New Random()
    Dim vCounter As Integer
    Dim vPosition As Integer
    Dim vHoldChar As Char
    Dim vGenPass As String
    Dim vPassGen As New System.Text.StringBuilder
    Dim vRandomPass As New System.Text.StringBuilder

    'Get the portal_passsword_min_length 
    Dim vResult As String = DataHelper.ConfigurationValue(DataHelper.ConfigurationValues.portal_password_min_length)
    'If greater than 6 set it as defaultPasswordLen
    If IntegerValue(vResult) > vDefaultPasswordLen Then
      vDefaultPasswordLen = IntegerValue(vResult)
    End If
    'one upper case alpha character - no vowels or 'y' to prevent offensive words being generated
    vPassGen.Append(vUpperCaseCharacter(vRand.Next(vUpperCaseCharacter.Length - 1)))
    'one number character (1-9)
    vPassGen.Append(vNumeric(vRand.Next(vNumeric.Length - 1)))
    'remaining lower case alpha character - no vowels or 'y' (as above)
    For value As Integer = 1 To vDefaultPasswordLen - 2
      vPassGen.Append(vLowerCaseCharacter(vRand.Next(vLowerCaseCharacter.Length - 1)))
    Next
    'Randomizing Generate Password
    Dim vChars() As Char = CType(vPassGen.ToString, Char())
    For vCounter = 0 To vChars.Length - 1
      vPosition = vRand.Next Mod vChars.Length
      vHoldChar = vChars(vCounter)
      vChars(vCounter) = vChars(vPosition)
      vChars(vPosition) = vHoldChar
    Next vCounter
    vGenPass = New String(vChars)
    Return vGenPass
  End Function
  Public Function GetImage(ByVal pImagePath As String, ByVal pAlternateText As String, ByVal pDefaultImagePath As String, ByVal pCssClass As String) As String
    Try
      If Not FileIO.FileSystem.FileExists(HttpContext.Current.Server.MapPath(pImagePath)) Then
        pImagePath = pDefaultImagePath
      End If
      Return String.Format("<img src ='{0}'alt='{1}'class='{2}'", pImagePath, pAlternateText, pCssClass)
    Catch ex As Exception
      If pAlternateText.Contains(">") Then
        pAlternateText = pAlternateText.Replace(">", "")
        pImagePath = pDefaultImagePath
      End If
      Return String.Format("<img src ='{0}'alt='{1}'class='{2}'", pImagePath, pAlternateText, pCssClass)
    End Try
  End Function

  Public Function GetDataGridItemIndex(ByVal pDataGrid As DataGrid, ByVal pColumnName As String) As Integer

    Dim vIndex As Integer
    For Each vCol As DataGridColumn In pDataGrid.Columns
      If TypeOf (vCol) Is BoundColumn Then
        Dim vBoundColumn As BoundColumn = DirectCast(vCol, BoundColumn)
        If vBoundColumn.DataField = pColumnName Then
          Return vIndex
        End If
      ElseIf TypeOf (vCol) Is TemplateColumn Then
        Dim vTemplateColumn As TemplateColumn = DirectCast(vCol, TemplateColumn)
        If TypeOf (vTemplateColumn.ItemTemplate) Is DisplayTemplate AndAlso
          DirectCast(vTemplateColumn.ItemTemplate, DisplayTemplate).DataItem = pColumnName Then
          Return vIndex
        End If
      End If
      vIndex += 1
    Next
  End Function

  Public Function GetBaseDataListItem(ByVal pBaseDataList As BaseDataList, ByVal pIndex As Integer, ByVal pColumnName As String) As String
    'Gets the Value of the Column Name within an Item of the list
    'This supports both DataGrid (row format) and DataList (column format)
    If TypeOf (pBaseDataList) Is DataGrid Then
      Dim vDataGrid As DataGrid = DirectCast(pBaseDataList, DataGrid)
      Return vDataGrid.Items(pIndex).Cells(GetDataGridItemIndex(vDataGrid, pColumnName)).Text
    Else
      'DataKeyField must be set for DataList
      Dim vDataList As DataList = DirectCast(pBaseDataList, DataList)
      Dim vDataKeyFieldValue As String = vDataList.DataKeys(pIndex).ToString
      If vDataList.DataKeyField = pColumnName Then
        Return vDataKeyFieldValue
      Else
        Return CType(vDataList.DataSource, DataSet).Tables("DataRow").Select(vDataList.DataKeyField & " = '" & vDataKeyFieldValue & "'")(0)(pColumnName).ToString
      End If
    End If
  End Function
  Public Function GetBaseDataListCheckBox(ByVal pBaseDataList As BaseDataList, ByVal pIndex As Integer) As CheckBox
    'Currently used to return the first control in the list which is a CheckBox
    If TypeOf (pBaseDataList) Is DataGrid Then
      Dim vDataGrid As DataGrid = DirectCast(pBaseDataList, DataGrid)
      For vCellIndex As Integer = 0 To vDataGrid.Items(pIndex).Cells.Count - 1
        If vDataGrid.Items(pIndex).Cells(vCellIndex).Controls.Count > 0 AndAlso TryCast(vDataGrid.Items(pIndex).Cells(vCellIndex).Controls(0), CheckBox) IsNot Nothing Then
          Return TryCast(vDataGrid.Items(pIndex).Cells(vCellIndex).Controls(0), CheckBox)
        End If
      Next
    Else
      Dim vDataList As DataGrid = DirectCast(pBaseDataList, DataGrid)
      Return TryCast(vDataList.Items(pIndex).Controls(0), CheckBox)
    End If
    Return Nothing
  End Function

  Public Enum DataAccessErrors
    daeUniservError
    daeInvalidContactDataSelectionFilter
    daeSessionValueNotSet
    daeInvalidEmailAddress
  End Enum

  Public Sub RaiseError(ByVal pError As DataAccessErrors, ByVal pParm1 As String)
    Dim vDesc As String = ""
    Select Case pError
      Case DataAccessErrors.daeUniservError
        vDesc = "%1"
      Case DataAccessErrors.daeInvalidContactDataSelectionFilter
        vDesc = "Could not set Data Selection filter '%1' due to invalid column name or filter format"
      Case DataAccessErrors.daeSessionValueNotSet
        vDesc = "Session value is not set for %1"
      Case DataAccessErrors.daeInvalidEmailAddress
        vDesc = "'%1' is not a valid email address"
      Case Else
        vDesc = "Undefined Error"
    End Select
    vDesc = vDesc.Replace("%1", pParm1)
    Throw New CareException(vDesc, CInt(pError), "CARE.Access", String.Empty, String.Empty)
  End Sub

  Public Function BooleanValue(ByVal pString As String) As Boolean
    If pString.Length = 0 Then
      Return False
    ElseIf pString.Length = 1 Then
      Return (pString = "Y")
    Else
      Return (pString.Substring(0, 1) = "Y")
    End If
  End Function

  Public Function BooleanString(ByVal pValue As Boolean) As String
    If pValue Then
      Return "Y"
    Else
      Return "N"
    End If
  End Function

  Public Sub PreserveStackTrace(ByVal pEx As Exception)
    Dim preserveStackTrace As System.Reflection.MethodInfo = GetType(Exception).GetMethod("InternalPreserveStackTrace", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)
    preserveStackTrace.Invoke(pEx, Nothing)
  End Sub

  Public Function Substring(ByVal pString As String, ByVal pStartIndex As Integer) As String
    Return Substring(pString, pStartIndex, pString.Length)
  End Function
  Public Function Substring(ByVal pString As String, ByVal pStartIndex As Integer, ByVal pLength As Integer) As String
    If pStartIndex < 0 Then pStartIndex = 0
    If pLength < 0 Then pLength = 0
    If pString.Length <= pStartIndex Then
      pStartIndex = 0
      pLength = 0
    End If
    If pLength + pStartIndex > pString.Length Then
      pLength = pString.Length - pStartIndex
    End If
    Return pString.Substring(pStartIndex, pLength)
  End Function

  Private mvDateFormat As String
  Public Function CAREDateFormat() As String
    If String.IsNullOrWhiteSpace(mvDateFormat) Then
      Dim vDate As String
      Dim vFormat As New StringBuilder

      mvDateFormat = ""
      vDate = New Date(1998, 11, 20).ToShortDateString()
      While vDate.Length > 0
        If vDate.StartsWith("20") Then
          vFormat.Append("dd")          'NoTranslate
          vDate = vDate.Substring(2)
        ElseIf vDate.StartsWith("11") Then
          vFormat.Append("MM")          'NoTranslate
          vDate = vDate.Substring(2)
        ElseIf vDate.StartsWith("98") Then
          vFormat.Append("yyyy")
          vDate = vDate.Substring(2)
        ElseIf vDate.StartsWith("1998") Then
          vFormat.Append("yyyy")
          vDate = vDate.Substring(4)
        Else
          If mvDateFormat.Length < 10 Then
            vFormat.Append(vDate.Substring(0, 1))
            vDate = vDate.Substring(1)
          End If
        End If
      End While
      mvDateFormat = vFormat.ToString
    End If
    Return mvDateFormat
  End Function

  Public Function CAREDateTimeFormat() As String
    Return CAREDateFormat() & " HH:mm:ss"
  End Function

  Public Function TodaysDate() As String
    Return Date.Today.ToString(CAREDateFormat)
  End Function

  Public Enum FieldTypes
    cftCharacter = 1
    cftMemo = 2
    cftInteger = 3
    cftLong = 4
    cftNumeric = 5
    cftDate = 6
    cftTime = 7
    cftBulk = 8
    cftFile = 9
  End Enum

  Public Function GetFieldType(ByVal pType As String, Optional ByVal pTableName As String = "") As FieldTypes
    Dim vType As FieldTypes
    Select Case pType
      Case "D"
        vType = FieldTypes.cftDate
      Case "T"
        vType = FieldTypes.cftTime
      Case "N"
        vType = FieldTypes.cftNumeric
      Case "I", "L"
        vType = FieldTypes.cftLong
      Case "M"
        vType = FieldTypes.cftMemo
      Case "C", "", "U" 'Added 'U' to support the new Unicode fields
        vType = FieldTypes.cftCharacter
      Case Else
        vType = CType(CInt(pType), FieldTypes)
        'Temporary fix until changed on the thick client
        If vType = FieldTypes.cftDate AndAlso pTableName = "contact_appointments" Then vType = FieldTypes.cftTime
    End Select
    Return vType
  End Function

End Module
