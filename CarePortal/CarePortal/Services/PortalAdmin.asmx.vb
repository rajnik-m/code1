Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Xml
Imports System.IO
Imports System.Security.Principal
Imports System.Security.Permissions
Imports System.web.Configuration
Imports System.Configuration

<Assembly: SecurityPermissionAttribute(SecurityAction.RequestMinimum, UnmanagedCode:=True), _
 Assembly: PermissionSetAttribute(SecurityAction.RequestMinimum, Name:="FullTrust")> 

<System.Web.Services.WebService(Namespace:="http://care.co.uk/webservices/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class PortalAdmin
  Inherits System.Web.Services.WebService

  Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As [String], _
       ByVal lpszDomain As [String], ByVal lpszPassword As [String], _
       ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, _
       ByRef phToken As IntPtr) As Boolean


  Private Declare Auto Function DuplicateToken Lib "advapi32.dll" (ByVal ExistingTokenHandle As IntPtr, _
          ByVal SECURITY_IMPERSONATION_LEVEL As Integer, _
          ByRef DuplicateTokenHandle As IntPtr) As Boolean


  Dim mvWI As WindowsIdentity
  Dim mvWIC As WindowsImpersonationContext

  Public Enum LogonType As Integer
    'This logon type is intended for users who will be interactively using the computer, such as a user being logged on 
    'by a terminal server, remote shell, or similar process.
    'This logon type has the additional expense of caching logon information for disconnected operations 
    'therefore, it is inappropriate for some client/server applications,
    'such as a mail server.
    LOGON32_LOGON_INTERACTIVE = 2

    'This logon type is intended for high performance servers to authenticate plaintext passwords.
    'The LogonUser function does not cache credentials for this logon type.
    LOGON32_LOGON_NETWORK = 3

    'This logon type is intended for batch servers, where processes may be executing on behalf of a user without 
    'their direct intervention. This type is also for higher performance servers that process many plaintext
    'authentication attempts at a time, such as mail or Web servers. 
    'The LogonUser function does not cache credentials for this logon type.
    LOGON32_LOGON_BATCH = 4

    'Indicates a service-type logon. The account provided must have the service privilege enabled. 
    LOGON32_LOGON_SERVICE = 5

    'This logon type is for GINA DLLs that log on users who will be interactively using the computer. 
    'This logon type can generate a unique audit record that shows when the workstation was unlocked. 
    LOGON32_LOGON_UNLOCK = 7

    'This logon type preserves the name and password in the authentication package, which allows the server to make 
    'connections to other network servers while impersonating the client. A server can accept plaintext credentials 
    'from a client, call LogonUser, verify that the user can access the system across the network, and still 
    'communicate with other servers.
    'NOTE: Windows NT:  This value is not supported. 
    LOGON32_LOGON_NETWORK_CLEARTEXT = 8

    'This logon type allows the caller to clone its current token and specify new credentials for outbound connections.
    'The new logon session has the same local identifier but uses different credentials for other network connections. 
    'NOTE: This logon type is supported only by the LOGON32_PROVIDER_WINNT50 logon provider.
    'NOTE: Windows NT:  This value is not supported. 
    LOGON32_LOGON_NEW_CREDENTIALS = 9
  End Enum

  Public Enum LogonProvider As Integer
    'Use the standard logon provider for the system. 
    'The default security provider is negotiate, unless you pass NULL for the domain name and the user name 
    'is not in UPN format. In this case, the default provider is NTLM. 
    'NOTE: Windows 2000/NT:   The default security provider is NTLM.
    LOGON32_PROVIDER_DEFAULT = 0
  End Enum

  <WebMethod(Description:="Updates the Web Name and Web Number for this Application")> _
  Public Function SetWebInfo(ByVal pWebNumber As Integer, ByVal pWebName As String) As String
    Dim vTable As DataTable = Nothing
    Dim vWebName As String = ""
    Try
      Dim vList As New ParameterList()
      vList("WebNumber") = pWebNumber
      vTable = DataHelper.GetWebInfo(vList)
      Dim vReturnList As New ParameterList
      If vTable IsNot Nothing Then vWebName = vTable.Rows(0).Item("WebName").ToString
      If vTable IsNot Nothing AndAlso pWebName = vWebName Then
        Dim vChanged As Boolean
        Dim vCfg As Configuration = WebConfigurationManager.OpenWebConfiguration("~")
        Dim vSetting As KeyValueConfigurationElement
        vSetting = CType(vCfg.AppSettings.Settings("WebNumber"), KeyValueConfigurationElement)
        If Not vSetting Is Nothing Then
          If vSetting.Value <> pWebNumber.ToString Then
            vSetting.Value = pWebNumber.ToString
            vChanged = True
          End If
        End If
        vSetting = CType(vCfg.AppSettings.Settings("WebName"), KeyValueConfigurationElement)
        If Not vSetting Is Nothing Then
          If vSetting.Value <> pWebName.ToString Then
            vSetting.Value = pWebName
            vChanged = True
          End If
        End If
        If vChanged Then vCfg.Save()
        vReturnList.Add("Result", "OK")
        Return vReturnList.XMLResultString
      Else
        vReturnList("ErrorMessge") = "The given Web Information does not match the Database data"
        Return vReturnList.XMLResultString
      End If
    Catch vException As Exception
      Return GetErrorReturn(vException)
    End Try
  End Function

  <WebMethod(Description:="Returns a result set of all the document names in the documents sub-directory")> _
  Public Function GetDocuments() As String
    Return GetFiles(GetDocumentPath, "*.doc", "*.pdf", "*.txt")
  End Function

  <WebMethod(Description:="Returns a result set of all the image names in the images sub-directory")> _
  Public Function GetImages() As String
    Return GetFiles(GetImagePath, "*.jpg", "*.gif", "*.bmp", "*.png")
  End Function

  Private Function GetFiles(ByVal pPathName As String, ByVal ParamArray pWildcards() As String) As String
    Dim vStream As New IO.MemoryStream
    Dim vWriter As New XmlTextWriter(vStream, Nothing)
    Dim vFileInfo As FileInfo
    Try
      With vWriter
        .WriteStartElement("Results")
        For Each vImage As String In My.Computer.FileSystem.GetFiles(pPathName, FileIO.SearchOption.SearchTopLevelOnly, pWildcards)
          vFileInfo = My.Computer.FileSystem.GetFileInfo(vImage)
          .WriteElementString("Image", vFileInfo.Name)
        Next
        .WriteEndElement()
        .Flush()
      End With
      Dim vReader As New System.IO.StreamReader(vStream)
      vStream.Position = 0
      Return vReader.ReadToEnd()
    Catch vEx As Exception
      Return GetErrorReturn(vEx)
    End Try
  End Function

  <WebMethod(Description:="Deletes an existing document from the documents sub-directory")> _
  Public Function DeleteDocument(ByVal pFileName As String) As String
    Return DeleteFile(GetDocumentPath, pFileName)
  End Function

  <WebMethod(Description:="Deletes an existing image from the images sub-directory")> _
  Public Function DeleteImage(ByVal pFileName As String) As String
    Return DeleteFile(GetImagePath, pFileName)
  End Function

  Private Function DeleteFile(ByVal pPathname As String, ByVal pFileName As String) As String
    Try
      Dim vList As New ParameterList
      If Impersonate() Then
        Dim vFileInfo As New FileInfo(pPathname & "\" & pFileName)
        If vFileInfo.Exists Then
          vFileInfo.Attributes = FileAttributes.Normal
          vFileInfo.Delete()
          If mvWIC IsNot Nothing Then mvWIC.Undo()
          vList.Add("Result", "OK")
          Return vList.XMLResultString
        Else
          If mvWIC IsNot Nothing Then mvWIC.Undo()
          vList.Add("ErrorMessage", "File Not Found")
          Return vList.XMLResultString
        End If
      Else
        vList.Add("ErrorMessage", "Invalid user credentials")
        Return vList.XMLResultString
      End If
    Catch vEx As Exception
      Return GetErrorReturn(vEx)
    End Try
  End Function

  <WebMethod(Description:="Uploads a new document from a byte stream into the documents sub-directory")> _
  Public Function UploadDocument(ByVal pFileName As String, ByVal pByte As Byte()) As String
    Return UploadFile(GetDocumentPath, pFileName, pByte)
  End Function

  <WebMethod(Description:="Uploads a new image from a byte stream into the images sub-directory")> _
  Public Function UploadImage(ByVal pFileName As String, ByVal pByte As Byte()) As String
    Return UploadFile(GetImagePath, pFileName, pByte)
  End Function

  Private Function UploadFile(ByVal pPathname As String, ByVal pFileName As String, ByVal pByte As Byte()) As String
    Dim vList As New ParameterList
    Dim vFS As FileStream = Nothing
    Try
      vFS = New FileStream(pPathname & "\" & pFileName, FileMode.Create)
      vFS.Write(pByte, 0, pByte.Length)
      vFS.Close()
    Catch vException As Exception
      If Not vFS Is Nothing Then vFS.Close()
      GetErrorReturn(vException)
    End Try
    vList.Add("Result", "OK")
    Return vList.XMLResultString
  End Function

  Private Function GetDocumentPath() As String
    Return Server.MapPath(Context.Request.ApplicationPath) & "/documents"
  End Function

  Private Function GetImagePath() As String
    Return Server.MapPath(Context.Request.ApplicationPath) & "/images"
  End Function

  Private Function GetErrorReturn(ByVal pException As Exception) As String
    Dim vList As New ParameterList
    vList("ErrorMessge") = pException.Message
    vList("Source") = pException.Source
    Return vList.XMLResultString
  End Function

  <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
  Private Function Impersonate() As Boolean
    Dim vToken As New IntPtr(0)
    Dim vDuplicateToken As New IntPtr(0)

    Return True

    'vToken = IntPtr.Zero
    'Dim vReturn As Boolean = LogonUser("simon", "carebusinesssol", "sdt1206", LogonType.LOGON32_LOGON_INTERACTIVE, LogonProvider.LOGON32_PROVIDER_DEFAULT, vToken)
    'If vReturn <> False Then
    '  Debug.Print("Before Impersonation " & WindowsIdentity.GetCurrent.Name)
    '  mvWI = New System.Security.Principal.WindowsIdentity(vToken)
    '  mvWIC = mvWI.Impersonate()
    '  Debug.Print("After Impersonation " & WindowsIdentity.GetCurrent.Name)
    '  Return True
    'End If
  End Function

End Class