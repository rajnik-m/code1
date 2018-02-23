Public Class SagePayNotification
  Inherits System.Web.UI.Page

  Private Const OK As String = "OK"
  Private Const INVALID As String = "INVALID"

  Private Enum PageType As Integer
    None = 0
    ProcessPayment
    AddMemberCC
  End Enum


  Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    Dim vResponse As String = String.Empty

    If Request.QueryString("Redirect") Is Nothing Then
      Response.Clear()
      Response.ContentType = "text/plain"

      If String.Compare(Request.Form("Status"), OK) = 0 Then
        vResponse = If(CheckMD5Hash(), OK, INVALID)
      Else
        vResponse = Request.Form("Status")
      End If

      If Request.Form.Keys.Count > 0 AndAlso String.Compare(vResponse, OK, True) = 0 Then
        Response.Write("Status=" & OK & Chr(13))
        If Cache("ProcessPaymentPageUrl") IsNot Nothing OrElse Cache("AddMemberCCUrl") IsNot Nothing Then
          Response.Write(GetQueryStringParameters(OK, If(Cache("ProcessPaymentPageUrl") IsNot Nothing, PageType.ProcessPayment, PageType.AddMemberCC)))
        Else
          Response.Write(GetQueryStringParameters(OK, PageType.None))
        End If
      Else
        Response.Write("Status=" & INVALID & Chr(13))

        If Cache("ProcessPaymentPageUrl") IsNot Nothing OrElse Cache("AddMemberCCUrl") IsNot Nothing Then
          Response.Write(GetQueryStringParameters(INVALID, If(Cache("ProcessPaymentPageUrl") IsNot Nothing, PageType.ProcessPayment, PageType.AddMemberCC)))
        Else
          Response.Write(GetQueryStringParameters(INVALID, PageType.None))
        End If
      End If
    End If

      'Response.Redirect("http://localhost:1234/default.aspx?pn=3000028&&Status=Ok&VendorTxCode=1000837&Last4Digit=0006&Token=&CardType=VISA&StatusDetail=0000 : The Authorisation was Successful.&ExpiryDate=0115")

  End Sub

  ''' <summary>
  ''' Create a query string based on the response from Sage Pay. 
  ''' </summary>
  ''' <param name="pStatus">Status from sage pay</param>
  ''' <param name="pPageType">Page Type to decide where the call was made from Portal or Trader</param>
  ''' <returns></returns>
  Private Function GetQueryStringParameters(pStatus As String, pPageType As PageType) As String
    Dim vResult As String

    If pPageType = PageType.None Then
      'Trader
      Dim vRedirectUrl As String = Request.Url.AbsoluteUri

      If String.Compare(pStatus, OK, StringComparison.CurrentCultureIgnoreCase) = 0 Then
        'vRedirectUrl.Replace(Request.Url.AbsoluteUri.Substring(Request.Url.AbsoluteUri.LastIndexOf("/")), "/Default.aspx"
        vResult = String.Format("RedirectURL={0}?Status={1}&VendorTxCode={2}&Last4Digit={3}&Token={4}&CardType={5}&StatusDetail={6}&ExpiryDate={7}&Redirect={8}",
                                 Request.Url.AbsoluteUri, OK, Server.UrlDecode(Request.Form("VendorTxCode")), Server.UrlDecode(Request.Form("Last4Digits")),
                              If(Request.Form("Token") IsNot Nothing, Server.UrlDecode(Request.Form("Token")), String.Empty),
                              If(Request.Form("CardType") IsNot Nothing, Server.UrlDecode(Request.Form("CardType")), String.Empty),
                              If(Request.Form("StatusDetail") IsNot Nothing, Server.UrlDecode(Request.Form("StatusDetail")), String.Empty),
                              If(Request.Form("ExpiryDate") IsNot Nothing, Server.UrlDecode(Request.Form("ExpiryDate")), String.Empty),
                              "Y")
      Else
        vResult = String.Format("RedirectURL={0}?Status={1}&VendorTxCode={2}&StatusDetail={3}&Redirect={4}",
                                 Request.Url.AbsoluteUri,
                                 Server.UrlDecode(Request.Form("Status")),
                                 Server.UrlDecode(Request.Form("VendorTxCode")),
                                 If(Request.Form("StatusDetail") IsNot Nothing, Server.UrlDecode(Request.Form("StatusDetail")), String.Empty),
                                 "Y")
      End If

    Else
      'Portal 
      Dim vPageLink As String = If(pPageType = PageType.ProcessPayment, Cache("ProcessPaymentPageUrl").ToString, Cache("AddMemberCCUrl").ToString)

      If String.Compare(pStatus, OK, StringComparison.CurrentCultureIgnoreCase) = 0 Then
        vResult = String.Format("RedirectURL={0}&Status={1}&VendorTxCode={2}&Last4Digit={3}&Token={4}&CardType={5}&StatusDetail={6}&ExpiryDate={7}",
                                 vPageLink, OK,
                                 Server.UrlDecode(Request.Form("VendorTxCode")),
                                 Server.UrlDecode(Request.Form("Last4Digits")),
                                 If(Request.Form("Token") IsNot Nothing, Server.UrlDecode(Request.Form("Token")), String.Empty),
                                 If(Request.Form("CardType") IsNot Nothing, Server.UrlDecode(Request.Form("CardType")), String.Empty),
                                 If(Request.Form("StatusDetail") IsNot Nothing, Server.UrlDecode(Request.Form("StatusDetail")), String.Empty),
                                 If(Request.Form("ExpiryDate") IsNot Nothing, Server.UrlDecode(Request.Form("ExpiryDate")), String.Empty))
      Else
        vResult = String.Format("RedirectURL={0}&Status={1}&VendorTxCode={2}&StatusDetail={3}",
                                 vPageLink,
                                 Server.UrlDecode(Request.Form("Status")),
                                 Server.UrlDecode(Request.Form("VendorTxCode")),
                                 Server.UrlDecode(Request.Form("StatusDetail")))
      End If
    End If
    Return vResult
  End Function

  Private Function CheckMD5Hash() As Boolean
    Dim vResponseData As New StringBuilder()
    Dim vQueue As New Queue(Of String)

    'VPSTxId + VendorTxCode + Status + TxAuthNo + VendorName+ AVSCV2 + SecurityKey + AddressResult + PostCodeResult + CV2Result + GiftAid + 3DSecureStatus + CAVV + AddressStatus + PayerStatus + CardType + Last4Digits + DeclineCode + ExpiryDate + FraudResponse + BankAuthCode
    Dim vKeyArray() As String = {"VPSTxId", "VendorTxCode", "Status", "TxAuthNo", "VendorName",
                                 "AVSCV2", "SecurityKey", "AddressResult", "PostCodeResult",
                                 "CV2Result", "GiftAid", "3DSecureStatus", "CAVV", "AddressStatus", "PayerStatus",
                                 "CardType", "Last4Digits", "DeclineCode", "ExpiryDate", "FraudResponse", "BankAuthCode"}


    If Cache("VendorName") Is Nothing Then Return True

    Dim vSecurityCode As String = GetSecurityCode(CInt(Request.Form("VendorTxCode")))
    If vSecurityCode.Length = 0 Then Return False

    If Request.Form("VPSSignature") IsNot Nothing Then
      For Each item As String In vKeyArray
        vQueue.Enqueue(item)
      Next

      Do While vQueue.Count <> 0
        Dim vItem As String = vQueue.Dequeue()
        Select vItem
          Case "SecurityKey"
            vResponseData.Append(vSecurityCode)
          Case "VendorName"
            vResponseData.Append(Cache("VendorName"))
          Case Else
            If Not String.IsNullOrWhiteSpace(Request.Form(vItem)) Then vResponseData.Append(Request.Form(vItem))
        End Select
      Loop

      Dim vMD5 As System.Security.Cryptography.MD5
      vMD5 = System.Security.Cryptography.MD5.Create()
      vMD5.ComputeHash(Encoding.UTF8.GetBytes(vResponseData.ToString))

      Dim vHashValue As New StringBuilder
      For vIndex As Integer = 0 To vMD5.Hash.Length - 1
        vHashValue.Append(vMD5.Hash(vIndex).ToString("X2")) 'Create 32 characters hexadecimal-formatted hash string
      Next
      Return If(String.Compare(vHashValue.ToString, Request.Form("VPSSignature")) = 0, True, False)
    Else
      Return False
    End If
  End Function


  Private Function GetSecurityCode(pAuthorisationNumber As Integer) As String
    Dim vSecurityCode As String = String.Empty
    Dim vDataTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtCreditCardAuthorisations, (Function(pNumber As Integer)
                                                                                                                               Dim vParams As New ParameterList(HttpContext.Current)
                                                                                                                               vParams.Add("AuthorisationNumber", pNumber)
                                                                                                                               Return vParams
                                                                                                                             End Function)(pAuthorisationNumber))

    If vDataTable IsNot Nothing AndAlso vDataTable.Rows.Count > 0 Then
      vSecurityCode = vDataTable.Rows(0).Item("AuthorisationCode").ToString()
    End If
    Return vSecurityCode
  End Function

End Class
