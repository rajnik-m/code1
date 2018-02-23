Imports CDBNServices.QASProOnDemand
Imports CARE.Access.PostcodeValidation
Imports CARE.Access.Interfaces
Imports CARE.Config

Namespace Access

  Public Class Postcoder

    'Private Const AFD_SERVER = "http://localhost:81"

    Private Enum GBResponses
      GBResponseNone = 0
      GBErrorResponse = 2
      GBErrorSynch = 3
      GBResponseOK = -1
    End Enum

    Public Enum ValidatePostcodeStatuses
      vpsNone
      vpsAddressPostcoded
      vpsAddressNotPostcoded
      vpsBuildingValidated
      vpsBuildingNotValidated
      vpsPostcodeValidated
      vpsPostcodeNotValidated
      vpsAddressRePostcoded
      vpsQASBatchReportCode
      vpsError
      vpsUpdateError
    End Enum

    Public Enum PostcoderTypes
      pctNone
      pctQAS
      pctGBMailing
      pctAFD
      pctEDEN
    End Enum

    Private Const DEF_UDP_TIMEOUT As Integer = 10 'Default Timeout

    Private Const MAX_RESPONSE As Integer = 550

    Private mvPostcoderType As PostcoderTypes
    Private mvCountries As String
    Private mvResponse As String
    Private mvEnv As CDBEnvironment
    'Private mvWinSock As Object
    Private mvRemoteHost As String
    Private mvUDPTimeout As Integer
    Private mvLastError As String
    Private mvInitialised As Boolean

    Private mvCheckAddress As Boolean
    Private mvAddress As String
    Private mvTown As String
    Private mvCounty As String
    Private mvPostCode As String
    Private mvOrgName As String
    Private mvBuildingNumber As String
    Private mvEasting As String
    Private mvNorthing As String
    Private mvDPS As String
    Private mvLEACode As String
    Private mvLEAName As String

    Private mvQASGBRLines As Integer
    Private mvQASDataSet As String

    Private mvThorofares As CollectionList(Of AddressInformation)
    Private mvBuildings As CollectionList(Of AddressInformation)

    Private mvGBIMTK As Boolean
    Private mvGBIMTKINI As String
    Private mvGBIMTKConfig As String
    Private mvGBHandle As String
    Private mvNoRecords As Boolean

    Private QASODPostcodeValidator As QASProOnDemandPostcodeValidator
    Private mvQASODInitialised As Boolean
    Private mvQASODLicencedCountries As String

    Private Enum IMTKErrors
      IMTKELInvalidRecordNumber = 1704
      IMTKELicenceWillExpireSoon = 1715
    End Enum

    Private Declare Function GBIMAPI Lib "GBNRTI32.DLL" Alias "_GBIMAPI@16" (ByVal Query As String, ByVal Buffr As String, ByVal Reply As String, ByVal Handl As String) As Short

    '----------------------------------------------------------------------------------
    'Quick Address stuff after this
    '----------------------------------------------------------------------------------

    Public Enum QuickAddressTypes
      qatClientServer
      qatNonClientServer
      qatProV4orAbove
      qatProOnDemand
    End Enum

    Public Enum QAProGetItemInfoTypes
      qapStepInfo = 0
    End Enum

    Public Enum QAProGetItemResults As Integer
      qgirNoStepIn
      qgirStepInRequired
      qgirInformation
      qgirWarning
    End Enum

    Private mvQAType As QuickAddressTypes
    Private mvQASInitialised As Boolean
    Private mvQASINIFile As String
    Private mvQAHandle As Integer
    Private mvQASOpened As Boolean
    Private mvQASBatchInitialised As Boolean
    Private mvQASBatchINIFile As String
    Private mvQASBatchHandle As Integer
    Private mvQASBatchOpened As Boolean
    Private mvQASBatchReportCode As String
    Private mvQASDPS As Boolean
    Private mvQASGRD As Boolean
    Private mvQASLEA As Boolean
    Private mvQASDPSIndex As Integer
    Private mvQASGRDEastIndex As Integer
    Private mvQASGRDNorthIndex As Integer
    Private mvQASLEACodeIndex As Integer
    Private mvQASLEANameIndex As Integer

    '--------------------------------------------------------------------------------
    'DECLARATIONS FOR THE CLIENT SERVER VERSION OF QUICK ADDRESS
    '--------------------------------------------------------------------------------
    Private Declare Sub CQAInitialise Lib "QAPUIEN.DLL" Alias "QAInitialise" (ByVal vi1 As Integer)
    Private Declare Sub CQAErrorMessage Lib "QAPUIEN.DLL" Alias "QAErrorMessage" (ByVal vi1 As Integer, ByVal rs2 As String, ByVal vi3 As Integer)
    Private Declare Function CQAPro_Open Lib "QAPUIEN.DLL" Alias "QAPro_Open" (ByVal vs1 As String, ByVal vs2 As String) As Integer
    Private Declare Sub CQAPro_Close Lib "QAPUIEN.DLL" Alias "QAPro_Close" ()
    Private Declare Function CQAPro_Count Lib "QAPUIEN.DLL" Alias "QAPro_Count" () As Integer
    Private Declare Sub CQAPro_EndSearch Lib "QAPUIEN.DLL" Alias "QAPro_EndSearch" ()
    Private Declare Function CQAPro_First Lib "QAPUIEN.DLL" Alias "QAPro_First" (ByVal vs1 As String, ByVal vi2 As Integer, ByRef rl3 As Integer) As Integer
    Private Declare Function CQAPro_FormatLine Lib "QAPUIEN.DLL" Alias "QAPro_FormatLine" (ByVal vl1 As Integer, ByVal vi2 As Integer, ByVal vs3 As String, ByVal rs4 As String, ByVal vi5 As Integer) As Integer
    Private Declare Function CQAPro_GetItemInfo Lib "QAPUIEN.DLL" Alias "QAPro_GetItemInfo" (ByVal vl1 As Integer, ByVal vi2 As Integer, ByRef rl3 As Integer) As Integer
    Private Declare Function CQAPro_ListItem Lib "QAPUIEN.DLL" Alias "QAPro_ListItem" (ByVal vl1 As Integer, ByVal rs2 As String, ByVal vi3 As Integer, ByVal vi4 As Integer) As Integer
    Private Declare Function CQAPro_Search Lib "QAPUIEN.DLL" Alias "QAPro_Search" (ByVal vs1 As String) As Integer
    Private Declare Function CQAPro_StepIn Lib "QAPUIEN.DLL" Alias "QAPro_StepIn" (ByVal vl1 As Integer) As Integer
    Private Declare Function CQAPro_StepOut Lib "QAPUIEN.DLL" Alias "QAPro_StepOut" () As Integer
    '--------------------------------------------------------------------------------
    'DECLARATIONS FOR THE NON-CLIENT SERVER VERSION OF QUICK ADDRESS
    '--------------------------------------------------------------------------------
    Private Declare Sub NQAInitialise Lib "QAPUIEB.DLL" Alias "QAInitialise" (ByVal vi1 As Integer)
    Private Declare Sub NQAErrorMessage Lib "QAPUIEB.DLL" Alias "QAErrorMessage" (ByVal vi1 As Integer, ByVal rs2 As String, ByVal vi3 As Integer)
    Private Declare Function NQAPro_Open Lib "QAPUIEB.DLL" Alias "QAPro_Open" (ByVal vs1 As String, ByVal vs2 As String) As Integer
    Private Declare Sub NQAPro_Close Lib "QAPUIEB.DLL" Alias "QAPro_Close" ()
    Private Declare Function NQAPro_Count Lib "QAPUIEB.DLL" Alias "QAPro_Count" () As Integer
    Private Declare Sub NQAPro_EndSearch Lib "QAPUIEB.DLL" Alias "QAPro_EndSearch" ()
    Private Declare Function NQAPro_First Lib "QAPUIEB.DLL" Alias "QAPro_First" (ByVal vs1 As String, ByVal vi2 As Integer, ByRef rl3 As Integer) As Integer
    Private Declare Function NQAPro_FormatLine Lib "QAPUIEB.DLL" Alias "QAPro_FormatLine" (ByVal vl1 As Integer, ByVal vi2 As Integer, ByVal vs3 As String, ByVal rs4 As String, ByVal vi5 As Integer) As Integer
    Private Declare Function NQAPro_GetItemInfo Lib "QAPUIEB.DLL" Alias "QAPro_GetItemInfo" (ByVal vl1 As Integer, ByVal vi2 As Integer, ByRef rl3 As Integer) As Integer
    Private Declare Function NQAPro_ListItem Lib "QAPUIEB.DLL" Alias "QAPro_ListItem" (ByVal vl1 As Integer, ByVal rs2 As String, ByVal vi3 As Integer, ByVal vi4 As Integer) As Integer
    Private Declare Function NQAPro_Search Lib "QAPUIEB.DLL" Alias "QAPro_Search" (ByVal vs1 As String) As Integer
    Private Declare Function NQAPro_StepIn Lib "QAPUIEB.DLL" Alias "QAPro_StepIn" (ByVal vl1 As Integer) As Integer
    Private Declare Function NQAPro_StepOut Lib "QAPUIEB.DLL" Alias "QAPro_StepOut" () As Integer
    '--------------------------------------------------------------------------------
    'DECLARATIONS FOR THE PRO V4 VERSION OF QUICK ADDRESS
    '--------------------------------------------------------------------------------
    'Private Declare Function QA_SetLibraryFlags Lib "QAUPIED.DLL" (ByVal vl1 As Long) As Long
    Private Declare Function QA_Open Lib "QAUPIED.DLL" (ByVal vs1 As String, ByVal vs2 As String, ByRef ri3 As Integer) As Integer
    Private Declare Function QA_Close Lib "QAUPIED.DLL" (ByVal vi1 As Integer) As Integer
    Private Declare Sub QA_Shutdown Lib "QAUPIED.DLL" ()
    'Private Declare Function QA_SetEngine Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long) As Long
    'Private Declare Function QA_GetEngine Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
    Private Declare Function QA_SetEngineOption Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer, ByVal vl3 As Integer) As Integer
    'Private Declare Function QA_GetEngineOption Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, rl3 As Long) As Long
    'Private Declare Function QA_GetPromptStatus Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ri3 As Long, ByVal rs4 As String, ByVal vi5 As Long) As Long
    'Private Declare Function QA_GetPrompt Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ri5 As Long, ByVal rs6 As String, ByVal vi7 As Long) As Long
    Private Declare Function QA_Search Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vs2 As String) As Integer
    'Private Declare Function QA_CancelSearch Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vl2 As Long) As Long
    Private Declare Function QA_EndSearch Lib "QAUPIED.DLL" (ByVal vi1 As Integer) As Integer
    Private Declare Function QA_GetSearchStatus Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByRef ri2 As Integer, ByRef ri3 As Integer, ByRef rl4 As Integer) As Integer
    Private Declare Function QA_GetSearchStatusDetail Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer, ByRef rl3 As Integer, ByVal rs4 As String, ByVal vi5 As Integer) As Integer
    Private Declare Function QA_StepIn Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer) As Integer
    Private Declare Function QA_StepOut Lib "QAUPIED.DLL" (ByVal vi1 As Integer) As Integer
    Private Declare Function QA_GetResult Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer, ByVal rs3 As String, ByVal vi4 As Integer, ByRef ri5 As Integer, ByRef rl6 As Integer) As Integer
    Private Declare Function QA_GetResultDetail Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer, ByVal vi3 As Integer, ByRef rl4 As Integer, ByVal rs5 As String, ByVal vi6 As Integer) As Integer
    Private Declare Function QA_FormatResult Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer, ByVal vs3 As String, ByRef ri4 As Integer, ByRef rl5 As Integer) As Integer
    Private Declare Function QA_GetFormattedLine Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer, ByVal rs3 As String, ByVal vi4 As Integer, ByVal rs5 As String, ByVal vi6 As Integer, ByRef rl7 As Integer) As Integer
    'Private Declare Function QA_GetExampleCount Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
    'Private Declare Function QA_FormatExample Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ri5 As Long, rl6 As Long) As Long
    'Private Declare Function QA_GetLayoutCount Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
    'Private Declare Function QA_GetLayout Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
    'Private Declare Function QA_GetActiveLayout Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
    Private Declare Function QA_SetActiveLayout Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vs2 As String) As Integer
    Private Declare Function QA_GetDataCount Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByRef ri2 As Integer) As Integer
    'Private Declare Function QA_GetData Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
    'Private Declare Function QA_GetDataDetail Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vi3 As Long, rl4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
    'Private Declare Function QA_GetLicensingCount Lib "QAUPIED.DLL" (ByVal vi1 As Long, ri2 As Long, rl3 As Long) As Long
    'Private Declare Function QA_GetLicensingDetail Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal vi3 As Long, rl4 As Long, ByVal rs5 As String, ByVal vi6 As Long) As Long
    Private Declare Function QA_GetActiveData Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal rs2 As String, ByVal vi3 As Integer) As Integer
    Private Declare Function QA_SetActiveData Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal vs2 As String) As Integer
    'Private Declare Function QA_GenerateSystemInfo Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ri3 As Long) As Long
    'Private Declare Function QA_GetSystemInfo Lib "QAUPIED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
    Private Declare Function QA_ErrorMessage Lib "QAUPIED.DLL" (ByVal vi1 As Integer, ByVal rs2 As String, ByVal vi3 As Integer) As Integer

    '--------------------------------------------------------------------------------
    'DECLARATIONS FOR QUICK ADDRESS BATCH 4.6
    '--------------------------------------------------------------------------------
    'Private Declare Function QABeginInstance Lib "QABWVED.DLL" () As Long
    'Private Declare Sub QAEndInstance Lib "QABWVED.DLL" ()
    'Private Declare Sub QAInitialise Lib "QABWVED.DLL" (ByVal vi1 As Long)
    Private Declare Sub QAErrorMessage Lib "QABWVED.DLL" (ByVal vi1 As Integer, ByVal rs2 As String, ByVal vi3 As Integer)
    'Private Declare Function QAErrorLevel Lib "QABWVED.DLL" (ByVal vi1 As Long) As Long
    'Private Declare Function QAErrorHistory Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
    'Private Declare Sub QAVersionInfo Lib "QABWVED.DLL" (ByVal rs1 As String, ByVal vi2 As Long)
    'Private Declare Function QADataInfo Lib "QABWVED.DLL" (ByVal rs1 As String, ByVal vi2 As Long, ri3 As Long) As Long
    'Private Declare Function QASystemInfo Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
    'Private Declare Function QAUpdateKey Lib "QABWVED.DLL" (ByVal rs1 As String, ByVal vi2 As Long) As Long
    'Private Declare Function QAUpdateCode Lib "QABWVED.DLL" (ByVal vs1 As String) As Long
    'Private Declare Function QALicenseInfo Lib "QABWVED.DLL" (ri1 As Long, ri2 As Long, ri3 As Long) As Long
    'Private Declare Function QAAuthorise Lib "QABWVED.DLL" (ByVal vs1 As String, ByVal vl2 As Long) As Long
    'Private Declare Function QASetStrMode Lib "QABWVED.DLL" (ByVal vi1 As Long) As Long

    Private Declare Function QABatchWV_Startup Lib "QABWVED.DLL" (ByVal vl1 As Integer) As Integer
    Private Declare Function QABatchWV_Shutdown Lib "QABWVED.DLL" () As Integer
    'Private Declare Function QABatchWV_LayoutCount Lib "QABWVED.DLL" (ByVal vs1 As String, ri2 As Long) As Long
    'Private Declare Function QABatchWV_GetLayout Lib "QABWVED.DLL" (ByVal vs1 As String, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
    Private Declare Function QABatchWV_Open Lib "QABWVED.DLL" (ByVal vs1 As String, ByVal vs2 As String, ByVal vl3 As Integer, ByRef ri4 As Integer) As Integer
    Private Declare Function QABatchWV_Close Lib "QABWVED.DLL" (ByVal vi1 As Integer) As Integer
    'Private Declare Function QABatchWV_Search Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ri3 As Long, ByVal rs4 As String, ByVal vi5 As Long, ByVal rs6 As String, ByVal rs7 As String, ByVal vi8 As Long) As Long
    Private Declare Function QABatchWV_Clean Lib "QABWVED.DLL" (ByVal vi1 As Integer, ByVal vs2 As String, ByRef ri3 As Integer, ByVal rs4 As String, ByVal vi5 As Integer, ByVal rs6 As String, ByVal rs7 As String, ByVal vi8 As Integer) As Integer
    'Private Declare Function QABatchWV_GetMatchInfo Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal rs3 As String, ri4 As Long, ri5 As Long, ri6 As Long, rl7 As Long, rl8 As Long, rl9 As Long) As Long
    Private Declare Function QABatchWV_EndSearch Lib "QABWVED.DLL" (ByVal vi1 As Integer) As Integer
    'Private Declare Function QABatchWV_LayoutLineCount Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ri3 As Long) As Long
    Private Declare Function QABatchWV_FormattedLineCount Lib "QABWVED.DLL" (ByVal vi1 As Integer, ByRef ri2 As Integer) As Integer
    'Private Declare Function QABatchWV_LayoutLineElements Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ByVal vi3 As Long, ByVal rs4 As String, ByVal vi5 As Long, rl6 As Long) As Long
    'Private Declare Function QABatchWV_GetLayoutLine Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
    Private Declare Function QABatchWV_GetFormattedLine Lib "QABWVED.DLL" (ByVal vi1 As Integer, ByVal vi2 As Integer, ByVal rs3 As String, ByVal vi4 As Integer) As Integer
    'Private Declare Function QABatchWV_UnusedLineCount Lib "QABWVED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
    'Private Declare Function QABatchWV_GetUnusedLine Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long) As Long
    'Private Declare Function QABatchWV_GetUnusedInput Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal vi4 As Long, rl5 As Long, rl6 As Long, rl7 As Long, ri8 As Long, ri9 As Long) As Long
    'Private Declare Function QABatchWV_CountryCount Lib "QABWVED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long
    'Private Declare Function QABatchWV_GetCountry Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vi2 As Long, ByVal rs3 As String, ByVal rs4 As String, ByVal vi5 As Long) As Long
    'Private Declare Function QABatchWV_DataInfo Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ri3 As Long, ByVal rs4 As String, ByVal vi5 As Long, ByVal rs6 As String, ByVal vi7 As Long) As Long
    'Private Declare Function QABatchWV_DataSetInfo Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ri3 As Long, ri4 As Long, ri5 As Long, ByVal rs6 As String, ByVal vi7 As Long, ByVal rs8 As String, ByVal vi9 As Long) As Long
    'Private Declare Function QABatchWV_LicenceInfoCount Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ri3 As Long) As Long
    'Private Declare Function QABatchWV_GetLicenceInfo Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String, ByVal vi3 As Long, ByVal rs4 As String, ByVal vi5 As Long) As Long
    'Private Declare Function QABatchWV_GetUSPS Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
    'Private Declare Function QABatchWV_GetDPBC Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
    'Private Declare Function QABatchWV_DPVGetCode Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal rs2 As String, ByVal vi3 As Long) As Long
    'Private Declare Function QABatchWV_DPVSetKey Lib "QABWVED.DLL" (ByVal vi1 As Long, ByVal vs2 As String) As Long
    'Private Declare Function QABatchWV_DPVState Lib "QABWVED.DLL" (ByVal vi1 As Long, ri2 As Long) As Long

    Private Const QAERR_FIELDTRUNCATED As Integer = -1302

    Private Const qassint_PICKLISTSIZE As Integer = 1
    Private Const qaresultint_ISCANSTEP As Integer = 15
    Private Const qaresultstr_DESCRIPTION As Integer = 2
    Private Const qaresultstr_PARTIALADDRESS As Integer = 3
    Private Const qaresult_CANSTEP As Integer = 4
    Private Const qaresult_INFORMATION As Integer = 1024
    Private Const qaresult_WARN_INFORMATION As Integer = 2048
    Private Const qaengopt_TIMEOUT As Integer = 7

#Region " General Initialisation & Properties "

    Public Sub Init(ByVal pEnv As CDBEnvironment)
      mvEnv = pEnv
      Dim vINIFile As New INIReader
      If Len(mvEnv.GetConfig("afd_everywhere_server")) > 0 And _
        (Len(mvEnv.GetConfig("afd_software")) = 0 Or mvEnv.GetConfig("afd_software") = "POSTCODE" Or mvEnv.GetConfig("afd_software") = "BOTH") Then
        mvPostcoderType = PostcoderTypes.pctAFD
        mvInitialised = True
      ElseIf mvEnv.GetConfigOption("qas_installed") Then
        mvPostcoderType = PostcoderTypes.pctQAS
      ElseIf Not String.IsNullOrWhiteSpace(mvEnv.GetConfig("qas_pro_ondemand_url")) Then
        mvPostcoderType = PostcoderTypes.pctQAS
      ElseIf vINIFile.ReadString("PREFERENCES", "QASInterface", "0") <> "0" Then
        mvPostcoderType = PostcoderTypes.pctQAS
      ElseIf mvEnv.GetConfigOption("option_gb_mailing", True) Then
        mvPostcoderType = PostcoderTypes.pctGBMailing
      Else
        mvPostcoderType = PostcoderTypes.pctNone
      End If
      If mvPostcoderType = PostcoderTypes.pctQAS Then mvCountries = mvEnv.GetConfig("qas_country_codes")
      If (mvPostcoderType <> PostcoderTypes.pctNone) And Len(mvCountries) = 0 Then mvCountries = "UK"
      mvQASDataSet = "UK"
    End Sub

    Public Function UsePostcoder(ByVal pCountry As String, ByRef pUK As Boolean) As Boolean
      If pUK Then pCountry = "UK"
      If InStr(mvCountries, pCountry) > 0 Then
        UsePostcoder = True
      Else
        UsePostcoder = False
      End If
    End Function

    Public ReadOnly Property PostcoderType() As PostcoderTypes
      Get
        PostcoderType = mvPostcoderType
      End Get
    End Property

    Public ReadOnly Property QuickAddressType() As QuickAddressTypes
      Get
        QuickAddressType = mvQAType
      End Get
    End Property
    Public ReadOnly Property MaxRetryCount() As Integer
      Get
        If mvGBIMTK Or PostcoderType <> PostcoderTypes.pctGBMailing Then
          MaxRetryCount = 0
        Else
          MaxRetryCount = 10
        End If
      End Get
    End Property

    Public ReadOnly Property PostcodeValidationAvailable() As Boolean
      Get
        PostcodeValidationAvailable = mvInitialised Or mvQASInitialised Or mvQASODInitialised
      End Get
    End Property

    Public ReadOnly Property LastError() As String
      Get
        LastError = mvLastError
      End Get
    End Property

    Public ReadOnly Property CheckAddress() As Boolean
      Get
        CheckAddress = mvCheckAddress
      End Get
    End Property

    Public ReadOnly Property Address(Optional ByVal pGetOrgName As Boolean = False) As String
      Get
        If pGetOrgName And Left(mvAddress, Len(mvOrgName)) <> mvOrgName Then
          Address = mvOrgName & vbCrLf & mvAddress
        Else
          Address = mvAddress
        End If
      End Get
    End Property

    Public ReadOnly Property Town() As String
      Get
        Town = mvTown
      End Get
    End Property

    Public ReadOnly Property County() As String
      Get
        County = mvCounty
      End Get
    End Property

    Public ReadOnly Property Postcode() As String
      Get
        Postcode = mvPostCode
      End Get
    End Property

    Public ReadOnly Property BuildingNumber() As String
      Get
        BuildingNumber = mvBuildingNumber
      End Get
    End Property

    Public ReadOnly Property OrganisationName() As String
      Get
        OrganisationName = mvOrgName
      End Get
    End Property

    Public ReadOnly Property Easting() As String
      Get
        Easting = mvEasting
      End Get
    End Property

    Public ReadOnly Property Northing() As String
      Get
        Northing = mvNorthing
      End Get
    End Property

    Public ReadOnly Property DPS() As String
      Get
        DPS = mvDPS
      End Get
    End Property

    Public ReadOnly Property LEACode() As String
      Get
        LEACode = mvLEACode
      End Get
    End Property

    Public ReadOnly Property LEAName() As String
      Get
        LEAName = mvLEAName
      End Get
    End Property

#End Region

#Region " Postcoding - All "

    Public Function IsCountrySupported(ByRef pCountry As String) As Boolean
      Dim vCountryToCheck As String

      vCountryToCheck = GetBaseCountry(pCountry)
      Select Case mvPostcoderType
        Case PostcoderTypes.pctQAS
          IsCountrySupported = InStr(mvCountries, vCountryToCheck) >= 0 'QAS only supported if in the countries list
        Case PostcoderTypes.pctNone
          IsCountrySupported = False 'No postcoder so not supported
        Case Else
          IsCountrySupported = vCountryToCheck = "UK" 'Other postcoders support UK only
      End Select
    End Function

    Public Function CanSearchByJustPostcode(ByRef pCountry As String) As Boolean
      Select Case GetBaseCountry(pCountry)
        Case "UK", "NL"
          CanSearchByJustPostcode = True
        Case Else
          CanSearchByJustPostcode = False
      End Select
    End Function

    Public Function SupportsBuildingNumber(ByRef pCountry As String) As Boolean
      Select Case GetBaseCountry(pCountry)
        Case "NL"
          SupportsBuildingNumber = True
        Case Else
          SupportsBuildingNumber = False
      End Select
    End Function

    Public Function PostcodeIfNoPostcodeGiven(ByRef pCountry As String) As Boolean
      Select Case GetBaseCountry(pCountry)
        Case "UK"
          PostcodeIfNoPostcodeGiven = True
        Case Else
          PostcodeIfNoPostcodeGiven = False
      End Select
    End Function

    Private Function GetBaseCountry(ByRef pCountry As String) As String
      If mvEnv.IsCountryUK(pCountry) Then
        GetBaseCountry = "UK"
      Else
        GetBaseCountry = pCountry
      End If
    End Function
    Public Function PostcodeAddress(ByVal Address As IAddress, ByRef pCountry As String, Optional ByVal pDataTable As CDBDataTable = Nothing) As ValidatePostcodeStatuses
      If IsCountrySupported(pCountry) Then
        If PostcoderType = PostcoderTypes.pctQAS AndAlso mvQAType = QuickAddressTypes.qatProOnDemand Then
          Try
            Dim results As IEnumerable(Of IAddress)
            results = QASODPostcodeValidator.PostcodeAddress(Address)
            PostcoderAddress.ConvertIAddressToDataTable(results, pDataTable)
          Catch vEx As Exception
            RaiseError(DataAccessErrors.daeQASProOnDemandGeneralError, vEx.Message)
          End Try
        End If
      Else
        Return ValidatePostcodeStatuses.vpsNone
      End If
    End Function
    Public Function PostcodeAddress(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String, ByRef pStrict As Boolean, Optional ByVal pDataTable As CDBDataTable = Nothing) As ValidatePostcodeStatuses
      'pDataTable is currently used by XMLLookupData.GetPAFPostcode to get all the available postcodes and not just the first matching one
      If IsCountrySupported(pCountry) Then
        Select Case PostcoderType
          Case PostcoderTypes.pctGBMailing
            Return PostcodeAddressGB(pAddress, pTown, pCounty, pPostCode, pDataTable)
          Case PostcoderTypes.pctQAS
            If mvQASBatchInitialised Then
              Return ValidateBuildingQASBatch(pAddress, pTown, pCounty, pPostCode)
            Else
              Return PostcodeAddressQAS(pAddress, pTown, pCounty, pPostCode, pStrict, pDataTable)
            End If
          Case PostcoderTypes.pctAFD
            Return ValidateBuildingAFD(False, "", pAddress, pTown, pCounty, pPostCode, pCountry, False)
        End Select
      Else
        Return ValidatePostcodeStatuses.vpsNone
      End If
    End Function

    Private Sub AddAddressToTable(ByVal pDataTable As CDBDataTable)
      Dim vRow As CDBDataRow = pDataTable.AddRow
      vRow.Item("AddressLine") = Replace(Address, vbCrLf, ", ")
      vRow.Item("Town") = Town
      vRow.Item("OrganisationName") = OrganisationName
      vRow.Item("County") = County
      vRow.Item("Postcode") = Postcode
      vRow.Item("Address") = Address
      If PostcoderType = Postcoder.PostcoderTypes.pctQAS Then
        vRow.Item("DeliveryPointSuffix") = DPS
        vRow.Item("Easting") = Easting
        vRow.Item("Northing") = Northing
        vRow.Item("LeaCode") = LEACode
        vRow.Item("LeaName") = LEAName
      End If
    End Sub
    Private Sub ClearAdditionalItems()
      mvEasting = ""
      mvNorthing = ""
      mvDPS = ""
      mvLEACode = ""
      mvLEAName = ""
    End Sub

    Public Function ValidateBuilding(ByRef pBuildingNumber As Boolean, ByRef pBuilding As String, ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String) As ValidatePostcodeStatuses
      If IsCountrySupported(pCountry) Then
        Select Case PostcoderType
          Case PostcoderTypes.pctGBMailing
            ValidateBuilding = ValidateBuildingGB(pBuildingNumber, pBuilding, pAddress, pTown, pCounty, pPostCode, pCountry)
          Case PostcoderTypes.pctQAS
            If mvQASBatchInitialised Then
              ValidateBuilding = ValidateBuildingQASBatch(pAddress, pTown, pCounty, pPostCode)
            ElseIf mvQAType = QuickAddressTypes.qatProOnDemand Then
              ValidateBuilding = ValidateBuildingQASProOnDemand(pBuildingNumber, pBuilding, pAddress, pTown, pCounty, pPostCode, pCountry)
            Else
              ValidateBuilding = ValidateBuildingQAS(pBuildingNumber, pBuilding, pAddress, pTown, pCounty, pPostCode, pCountry)
            End If
          Case PostcoderTypes.pctAFD
            ValidateBuilding = ValidateBuildingAFD(pBuildingNumber, pBuilding, pAddress, pTown, pCounty, pPostCode, pCountry, False)
        End Select
      Else
        ValidateBuilding = ValidatePostcodeStatuses.vpsNone
      End If
    End Function
    Private Function ValidateBuildingQASProOnDemand(ByRef pBuildingNumber As Boolean, ByRef pBuilding As String, ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String) As ValidatePostcodeStatuses
      Try
        Dim addressLine As String = String.Empty
        addressLine = GetAddressLine(pAddress) 'Address Elements need to be delimited for Search
        Dim iso3CountryCode As String
        If Not String.IsNullOrWhiteSpace(pCountry) Then
          iso3CountryCode = pCountry
        Else
          iso3CountryCode = mvEnv.DefaultCountry
        End If
        If Not String.IsNullOrWhiteSpace(iso3CountryCode) Then ConvertCountryCodeToISO3(iso3CountryCode)
        Dim buildingAddress As New PostcoderAddress With {.Address = addressLine,
                                                          .Town = pTown,
                                                          .County = pCounty,
                                                          .Postcode = pPostCode,
                                                          .Iso3166Alpha3CountryCode = iso3CountryCode}
        If pBuildingNumber Then buildingAddress.BuildingNumber = pBuilding
        ValidateBuildingQASProOnDemand = QASODPostcodeValidator.ValidateBuilding(buildingAddress)
        If buildingAddress.AddressVerified Then
          mvCheckAddress = True
          pAddress = Split(buildingAddress.Address, vbCrLf)
          If pBuildingNumber Then pBuilding = buildingAddress.BuildingNumber
          pTown = buildingAddress.Town
          pCounty = buildingAddress.County
          pPostCode = buildingAddress.Postcode
          mvAddress = buildingAddress.Address
          mvTown = buildingAddress.Town
          mvCounty = buildingAddress.County
          mvPostCode = buildingAddress.Postcode
          mvDPS = buildingAddress.DPS
          mvEasting = buildingAddress.Easting.ToString
          mvNorthing = buildingAddress.Northing.ToString
          mvLEACode = buildingAddress.LEACode
          mvLEAName = buildingAddress.LEAName
        End If
      Catch vEx As Exception
        RaiseError(DataAccessErrors.daeQASProOnDemandGeneralError, vEx.Message)
      End Try
    End Function
    Public Function GetAddressLine(ByVal pAddress() As String) As String
      Dim addressLine As String = String.Empty
      Dim index As Integer = 0
      Do
        If Not String.IsNullOrWhiteSpace(pAddress(index)) Then
          If String.IsNullOrWhiteSpace(addressLine) Then
            addressLine = pAddress(index)
          Else
            addressLine = addressLine + "|" + pAddress(index)
          End If
        End If
        index = index + 1
      Loop While index < 4

      Return addressLine
    End Function
    Public Function ValidatePostcode(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String) As ValidatePostcodeStatuses
      If IsCountrySupported(pCountry) Then
        Select Case PostcoderType
          Case PostcoderTypes.pctGBMailing
            ValidatePostcode = ValidatePostcodeGB(pAddress, pTown, pCounty, pPostCode, pCountry)
          Case PostcoderTypes.pctQAS
            If mvQASBatchInitialised OrElse mvQAType = QuickAddressTypes.qatProOnDemand Then
              'Don't do this?
              ValidatePostcode = ValidatePostcodeStatuses.vpsNone
            Else
              ValidatePostcode = ValidatePostcodeQAS(pAddress, pTown, pCounty, pPostCode)
            End If
          Case PostcoderTypes.pctAFD
            ValidatePostcode = ValidateBuildingAFD(False, "", pAddress, pTown, pCounty, pPostCode, pCountry, True)
        End Select
      Else
        ValidatePostcode = ValidatePostcodeStatuses.vpsNone
      End If
    End Function

#End Region

#Region " GB Mailing "

    Public Function InitGBMailing(ByRef pWinSock As Object) As Boolean
      Dim vGBIMTKPath As String
      Dim vSection As String

      vGBIMTKPath = mvEnv.GetConfig("gb_imtk_path")
      mvGBIMTK = False
      If Len(vGBIMTKPath) > 0 Then
        If Right(vGBIMTKPath, 1) = "\" Then vGBIMTKPath = Left(vGBIMTKPath, Len(vGBIMTKPath) - 1)
        mvGBIMTKINI = vGBIMTKPath & "\imtkcfg.ini"
        mvGBIMTK = True
        If CallGBIMAPI("INIT('" & vGBIMTKPath & "')") Then
          mvInitialised = True
          vSection = "Config:CDB Config"
          Dim vINIFile As New INIReader(mvGBIMTKINI)
          If vINIFile.ReadString(vSection, "Line18", "").Length = 0 Then
            'If InStr(vINIFile.ReadString(vSection, "Exclude", ""), "ORGN") <= 0 Then
            'If Len(vINIFile.ReadString(vSection, "Line14", "")) = 0 Then
            With vINIFile
              .Write(vSection, "Description", "Care Contacts Database Configuration V2")
              .Write(vSection, "PadFields", "0")
              .Write(vSection, "PadChar", "32")
              .Write(vSection, "CountyAlways", "0")
              .Write(vSection, "AbbrCounty", "0")
              .Write(vSection, "NoOfLines", "18")
              .Write(vSection, "Exclude", "TOWN|CNTY|PCOD|ORGN")
              .Write(vSection, "Line1", "35,ADDR,DEFAULT,,")
              .Write(vSection, "Line2", "35,ADDR,DEFAULT,,")
              .Write(vSection, "Line3", "35,ADDR,DEFAULT,,")
              .Write(vSection, "Line4", "35,ADDR,DEFAULT,,")
              .Write(vSection, "Line5", "35,TOWN,DEFAULT|COMPONENT,,")
              .Write(vSection, "Line6", "35,CNTY,DEFAULT|COMPONENT,,")
              .Write(vSection, "Line7", "10,PCOD,DEFAULT,,")
              .Write(vSection, "Line8", "35,SUBB,DEFAULT,,")
              .Write(vSection, "Line9", "35,BULD,DEFAULT,,")
              .Write(vSection, "Line10", "35,ORGN,DEFAULT,,")
              .Write(vSection, "Line11", "35,THOR,DEFAULT,,")
              .Write(vSection, "Line12", "35,DDLO,DEFAULT,,")
              .Write(vSection, "Line13", "35,DPLO,DEFAULT,,")
              .Write(vSection, "Line14", "35,BNAM,DEFAULT,,")
              .Write(vSection, "Line15", "35,BNUM,DEFAULT,,")
              .Write(vSection, "Line16", "5,DPTS,DEFAULT,,")
              .Write(vSection, "Line17", "6,EAST,DEFAULT,,")
              .Write(vSection, "Line18", "6,NRTH,DEFAULT,,")
              .Write(vSection, "Format", "LALIGN|LEFT|MIXED")
            End With
            AddGBErrorMessages()
          End If
          mvGBIMTKConfig = "LOADFCONFIG('" & mvGBIMTKINI & "','" & vSection & "')"
          If CallGBIMAPI(mvGBIMTKConfig) Then
            InitGBMailing = True
          End If
        End If
      Else
        If pWinSock Is Nothing Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
        'mvCheckAddress = False
        'mvWinSock = pWinSock
        'vHostName = mvEnv.GetConfig("cbs_server host name")
        'vUDPAddress = mvEnv.GetConfig("cbs_server address")
        'vUDPPort = mvEnv.GetConfig("cbs_server port")
        'If Val(vUDPPort) < 1 Then RaiseError(DataAccessErrors.daeInvalidConfig, "cbs_server_port")
        'mvUDPTimeout = IntegerValue(mvEnv.GetConfig("cbs_server timeout"))
        'If mvUDPTimeout < 1 Then mvUDPTimeout = DEF_UDP_TIMEOUT

        ''UPGRADE_WARNING: Couldn't resolve default property of object mvWinSock.RemotePort. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'mvWinSock.RemotePort = vUDPPort
        ''wsk.LocalPort = vUDPPort DO NOT USE LOCAL PORT AS THIS FAILS WITH MULTIPLE INSTANCES

        ''Get the IP Address if available
        ''UPGRADE_WARNING: Couldn't resolve default property of object mvWinSock.RemoteHost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'mvWinSock.RemoteHost = vUDPAddress
        ''UPGRADE_WARNING: Couldn't resolve default property of object mvWinSock.RemoteHost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'mvRemoteHost = mvWinSock.RemoteHost
        ''Check if a valid address (needs more)
        'If InStr(vUDPAddress, ".") <= 0 Then
        '  'Not valid so check the host name
        '  If Len(vHostName) Then
        '    'UPGRADE_WARNING: Couldn't resolve default property of object mvWinSock.RemoteHost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '    mvWinSock.RemoteHost = vHostName
        '  Else
        '    RaiseError(DataAccessErrors.daeInvalidConfig, "cbs_server")
        '  End If
        '  'gvGBMailingInfo.error = LoadString(24203)    'Failed to connect to GB Mailing Server
        '  'gvGBMailingInfo.error = LoadStringP1(24204, wsk.ToAddress)   'Invalid GB Mailing Server IP Address %s
        'End If

        'If GBSend("GBOP") = GBResponses.GBResponseOK Then
        '  vGBResponse = GBSend("CBSV")
        '  If vGBResponse = GBResponses.GBResponseOK Then
        '    If Left(mvResponse, 6) = "CBSVOK" And Val(Mid(mvResponse, 7)) >= 1.04 Then
        '      mvInitialised = True
        '      InitGBMailing = True
        '    Else
        '      If Left(mvResponse, 6) = "CBSVOK" Then
        '        vVersion = Mid(mvResponse, 7)
        '      Else
        '        vVersion = "< 1.04"
        '      End If
        '      mvLastError = gvSystem.LoadStringP1(24205, vVersion) 'The CBS_SERVER running on the host is the wrong version %s\r\n\r\nPlease install a newer version
        '    End If
        '  Else
        '    If vGBResponse = GBResponses.GBErrorResponse Then
        '      mvLastError = mvLastError & gvSystem.LoadString(24208)
        '    End If
        '  End If
        'End If
      End If
    End Function
    Public Sub Close()
      Select Case mvEnv.Postcoder.PostcoderType
        Case Postcoder.PostcoderTypes.pctGBMailing
          CloseGBMailing()
        Case Postcoder.PostcoderTypes.pctQAS
          QASClose()
      End Select
    End Sub
    Public Sub CloseGBMailing()
      If mvInitialised Then
        If mvGBIMTK Then
          CallGBIMAPI("TERM")
        Else
          'GBSend("GBCL")
        End If
        mvInitialised = False
      End If
    End Sub

    'Public Sub ProcessDataArrival(ByRef pWinSock As Object)
    '  Dim vString As String
    '  Dim vRemoteHostIP As String
    '  Dim vLocalIP As String

    '  With pWinSock
    '    'UPGRADE_WARNING: Couldn't resolve default property of object pWinSock.GetData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    .GetData(vString, VariantType.String)
    '    'UPGRADE_WARNING: Couldn't resolve default property of object pWinSock.RemoteHostIP. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    vRemoteHostIP = .RemoteHostIP
    '    'UPGRADE_WARNING: Couldn't resolve default property of object pWinSock.LocalIP. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    vLocalIP = .LocalIP
    '  End With
    '  Debug.Print("Data Arrived " & vString & " RemoteHostIP = " & vRemoteHostIP & " LocalIP = " & vLocalIP)
    '  If (vRemoteHostIP = mvRemoteHost) Then
    '    If Mid(vString, 5, Len(vLocalIP)) = vLocalIP Then
    '      mvResponse = Left(vString, 4) & Mid(vString, 5 + Len(vLocalIP))
    '    End If
    '  End If
    'End Sub

    Public Function GetFormattedAddress(ByVal pPostCode As String, ByVal pThorofareNumber As Integer, ByVal pBuildingNumber As Integer, Optional ByRef pAllowOrg As Boolean = False) As Boolean
      Dim vGBSelect As String
      Dim vAddressInfo As New AddressInformation

      If Not mvInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      mvLastError = ""

      If mvGBIMTK Then
        CallGBIMAPI(mvGBIMTKConfig)
        vGBSelect = "SELECT CONFIGURE FROM NR WHERE WALK = '" & pPostCode & "'"
        If Not mvBuildings Is Nothing Then
          vAddressInfo = mvBuildings.Item(pBuildingNumber)
          AddWhereClause(vGBSelect, "THOR", (vAddressInfo.Thorofare))
          AddWhereClause(vGBSelect, "DPLO", (vAddressInfo.DependantLocality))
          AddWhereClause(vGBSelect, "DDLO", (vAddressInfo.DoubleDependantLocality))
          AddWhereClause(vGBSelect, "ORGN", (vAddressInfo.Organisation))
          AddWhereClause(vGBSelect, "SUBB", (vAddressInfo.SubBuilding))
          AddWhereClause(vGBSelect, "BNAM", (vAddressInfo.BuildingName))
          AddWhereClause(vGBSelect, "BNUM", (vAddressInfo.BuildingNumber))
        End If
        If CallGBIMAPI(vGBSelect, True) Then
          If mvNoRecords = False Then
            If CallGBIMAPI("GETNEXT") Then
              GBExtractAddress(1, vbCrLf)
              CheckForCounty()
              If pAllowOrg Then
                If Len(vAddressInfo.Organisation) > 0 Then
                  If Len(vAddressInfo.SubBuilding) + Len(vAddressInfo.BuildingName) + Len(vAddressInfo.BuildingNumber) = 0 Then
                    If InStr(mvAddress, vAddressInfo.Organisation) <= 0 Then
                      If mvAddress.Length > 0 Then
                        mvAddress = vAddressInfo.Organisation & vbCrLf & mvAddress
                      Else
                        mvAddress = vAddressInfo.Organisation
                      End If
                    End If
                  End If
                End If
              Else
                If Len(vAddressInfo.Organisation) > 0 Then
                  mvAddress = Replace(mvAddress, vAddressInfo.Organisation, "")
                  If Left(mvAddress, 2) = vbCrLf Then mvAddress = Mid(mvAddress, 3)
                End If
              End If
              GetFormattedAddress = True
            End If
          Else
            ClearResponse()
            GBExtractAddress(1, vbCrLf)
          End If
        End If
      Else
        'If GBSend("GBBA " & Cstr(pThorofareNumber, "0000") & " " & Cstr(pBuildingNumber, "0000") & " " & pPostCode) = True Then
        '  GBExtractAddress(18, vbCrLf)
        '  GetFormattedAddress = True
        'End If
      End If
    End Function

    Private Sub AddWhereClause(ByRef pGBSelect As String, ByRef pCode As String, ByRef pValue As String)
      If pValue.Length > 0 Then
        pValue = pValue.Replace("'", "''")
        If InStr(pValue, ",") > 0 Then
          pValue = Trim(Replace(pValue, ",", " "))
        End If
        If Len(pValue) >= 34 Then
          pGBSelect = pGBSelect & " AND " & pCode & " LIKE '" & pValue & "*'"
        Else
          pGBSelect = pGBSelect & " AND " & pCode & " = '" & pValue & "'"
        End If
      End If
    End Sub

    Public Function GetThorofares(ByVal pPostCode As String) As CDBCollection
      Dim vThorofareNumber As Integer
      Dim vGBSelect As String
      Dim vAddressInfo As AddressInformation
      Dim vCollection As New CDBCollection

      If Not mvInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      mvLastError = ""
      vThorofareNumber = 1
      mvThorofares = New CollectionList(Of AddressInformation)(1)

      If mvGBIMTK Then
        If CallGBIMAPI("SETDATALEVEL('THOROFARE')") Then
          vGBSelect = "SELECT CONFIGURE FROM NR WHERE WALK = '" & pPostCode & "'"
          If CallGBIMAPI(vGBSelect, True) Then
            While mvNoRecords = False
              If CallGBIMAPI("GETNEXT") Then
                If mvNoRecords = False Then
                  vAddressInfo = New AddressInformation
                  vAddressInfo.ExtractData(mvResponse)
                  mvThorofares.Add("T" & vThorofareNumber, vAddressInfo)
                  vCollection.Add((vAddressInfo.ThorofareString), "T" & vThorofareNumber)
                  vThorofareNumber = vThorofareNumber + 1
                End If
              End If
            End While
          End If
        End If
      Else
        'The GBTF call looks like this GBTF 0001 POSTCODE
        'ST MOD - Do While GBSend(Me, "GBTF " & Format$(vThorofareNumber, "0000") & " " & pPostcode) = True
        'Do While GBSend("GBTF " & Cstr(vThorofareNumber, "0000") & " " & pPostCode) = True
        '  'The response looks like this GBTF   0001 address_data
        '  GBExtractAddress(12, ", ")
        '  vThorofare = mvAddress
        '  If mvTown.Length > 0 Then
        '    If vThorofare.Length > 0 Then vThorofare = vThorofare & ", "
        '      vThorofare = vThorofare & mvTown
        '    End If
        '    If mvCounty.Length > 0 Then
        '    If vThorofare.Length > 0 Then vThorofare = vThorofare & ", "
        '      vThorofare = vThorofare & mvCounty
        '    End If
        '    vCollection.Add(vThorofare, "T" & vThorofareNumber)
        '    vThorofareNumber = vThorofareNumber + 1
        '    If Left(mvResponse, 6) = "GBTFOK" Then Exit Do
        'Loop
      End If
      GetThorofares = vCollection
    End Function

    Public Function GetBuildings(ByVal pPostCode As String, ByVal pThorofareNumber As Integer) As CDBCollection
      Dim vBuildingNumber As Integer
      Dim vCollection As New CDBCollection
      Dim vGBSelect As String
      Dim vAddressInfo As AddressInformation
      Dim vNewAddressInfo As AddressInformation

      If Not mvInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      mvLastError = ""
      vBuildingNumber = 1
      mvBuildings = New CollectionList(Of AddressInformation)(1)

      If mvGBIMTK Then
        If CallGBIMAPI("SETDATALEVEL('BUILDING')") Then
          vGBSelect = "SELECT CONFIGURE FROM NR WHERE WALK = '" & pPostCode & "'"
          vAddressInfo = mvThorofares.Item(pThorofareNumber)
          AddWhereClause(vGBSelect, "THOR", (vAddressInfo.Thorofare))
          AddWhereClause(vGBSelect, "DPLO", (vAddressInfo.DependantLocality))
          AddWhereClause(vGBSelect, "DDLO", (vAddressInfo.DoubleDependantLocality))
          If CallGBIMAPI(vGBSelect, True) Then
            While mvNoRecords = False
              If CallGBIMAPI("GETNEXT") Then
                If mvNoRecords = False Then
                  vNewAddressInfo = New AddressInformation
                  vNewAddressInfo.ExtractData(mvResponse)
                  If vNewAddressInfo.Thorofare = vAddressInfo.Thorofare Then
                    mvBuildings.Add("B" & vBuildingNumber, vNewAddressInfo)
                    vCollection.Add((vNewAddressInfo.BuildingString), "B" & vBuildingNumber)
                    vBuildingNumber = vBuildingNumber + 1
                  End If
                End If
              End If
            End While
          End If
        End If
      Else
        'The GBBU call looks like this GBBU 0001 0001 POSTCODE
        'Do While GBSend("GBBU " & Cstr(pThorofareNumber, "0000") & " " & VB6.Format(vBuildingNumber, "0000") & " " & pPostCode) = True
        '  'The response is GBBU 0001 0001 building_data
        '  If Left(mvResponse, 6) = "GBBUOK" Then Exit Do
        '  vBuilding = Trim(Mid(mvResponse, 41 + 15, 40))
        '  vSubBuilding = Trim(Mid(mvResponse, 81 + 15, 40))
        '  vBuildingName = Trim(Mid(mvResponse, 121 + 15, 40))
        '  If vSubBuilding.Length > 0 Then
        '    If vBuilding.Length > 0 Then vBuilding = vBuilding & ", "
        '      vBuilding = vBuilding & vSubBuilding
        '    End If
        '    If vBuildingName.Length > 0 Then
        '      If vBuilding.Length > 0 Then vBuilding = vBuilding & ", "
        '      vBuilding = vBuilding & vBuildingName
        '    End If
        '    vCollection.Add(vBuilding, "B" & vBuildingNumber)
        '    vBuildingNumber = vBuildingNumber + 1
        'Loop
      End If
      GetBuildings = vCollection
    End Function

    Private Function PostcodeAddressGB(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, Optional ByVal pDataTable As CDBDataTable = Nothing) As ValidatePostcodeStatuses
      Dim vGBAddress As String
      Dim vBuildingNo As String = ""
      Dim vBuildingName As String = ""
      Dim vThorofare As String = ""


      If Not mvInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      ClearAdditionalItems()
      mvCheckAddress = False
      If mvGBIMTK Then
        GetBuildingInfo(pAddress, vBuildingNo, vBuildingName, vThorofare)
        vGBAddress = "SELECT CONFIGURE FROM NR WHERE "
        If vBuildingNo.Length > 0 Then
          AddWhereClause(vGBAddress, "BNUM", vBuildingNo)
        Else
          AddWhereClause(vGBAddress, "BNAM", vBuildingName)
        End If
        vGBAddress = Replace(vGBAddress, " AND ", "") 'Remove and from the start
        AddWhereClause(vGBAddress, "THOR", vThorofare)
        AddWhereClause(vGBAddress, "TOWN", pTown)
        AddWhereClause(vGBAddress, "CNTY", pCounty)
        If CallGBIMAPI(vGBAddress, True) Then
          If mvNoRecords Then
            PostcodeAddressGB = ValidatePostcodeStatuses.vpsAddressNotPostcoded
          Else
            If vBuildingNo.Length > 0 AndAlso vThorofare.Length = 0 Then
              PostcodeAddressGB = ValidatePostcodeStatuses.vpsAddressNotPostcoded 'Report cannot get postcode
            ElseIf CallGBIMAPI("GETNEXT") Then
              If mvNoRecords Then
                PostcodeAddressGB = ValidatePostcodeStatuses.vpsAddressNotPostcoded 'Report cannot get postcode
              Else
                Dim vCount As Integer = 0
                Do
                  Dim vContinue As Boolean = True
                  'BR13924: Changes to resolve the issue where an incorrect postcode was used when importing an address having a Dot (.) e.g. imported address = 34 St. Marks Road W10 6JN but the system makes it 34 St Marks Road W7 2PW
                  If pDataTable Is Nothing AndAlso Me.Address IsNot Nothing Then  'We have found one address previously which did not match with imported address for some reasons
                    Dim vNewAddInfo As New AddressInformation()
                    vNewAddInfo.ExtractData(mvResponse) 'Just extract the new data. Do not update Postcoder values yet.
                    If vNewAddInfo.AddressMatch(vBuildingNo, Me.Address) Then 'If the address for next record is same as the address of the last record then we do not need to re-postcode. We will use the last address.
                      mvCheckAddress = True 'Set this to true so that the address changes are reported correctly
                      vContinue = False
                    End If
                  End If
                  If vContinue Then
                    PostcodeAddressGB = ValidatePostcodeStatuses.vpsAddressPostcoded
                    GBExtractAddress(1, vbCrLf)
                    CheckForCounty()
                    If pDataTable IsNot Nothing Then
                      AddAddressToTable(pDataTable)
                      CallGBIMAPI("GETNEXT")
                      vCount += 1
                    End If
                  End If
                Loop While pDataTable IsNot Nothing AndAlso Not mvNoRecords AndAlso vCount <= 2000
              End If
            Else
              PostcodeAddressGB = ValidatePostcodeStatuses.vpsError
            End If
          End If
        Else
          PostcodeAddressGB = ValidatePostcodeStatuses.vpsError
        End If
      Else
        'Dim vGBResponse As GBResponses
        ''No postcode so make GBPC call eg. Address1|Address2|Address3|Address4|Town|County|Postcode
        'vGBAddress = pAddress(0) & "|" & pAddress(1) & "|" & pAddress(2) & "|" & pAddress(3) & "|" & pTown & "|" & pCounty & "|" & pPostCode
        'vGBResponse = GBSend("GBPC " & vGBAddress)
        'If vGBResponse = GBResponses.GBResponseOK Then
        '  'Check for OK or Full postcode applied warnings
        '  If Left(mvResponse, 6) = "GBPCOK" Or Left(mvResponse, 7) = "GBPCN32" Or Left(mvResponse, 7) = "GBPCN33" Then
        '    'If OK extract the address save the postcode and set the paf status to VB
        '    GBExtractAddress(8)
        '    PostcodeAddressGB = ValidatePostcodeStatuses.vpsAddressPostcoded
        '  Else
        '    PostcodeAddressGB = ValidatePostcodeStatuses.vpsAddressNotPostcoded 'Report cannot get postcode
        '  End If
        'ElseIf vGBResponse = GBResponses.GBResponseNone Then
        '  PostcodeAddressGB = ValidatePostcodeStatuses.vpsAddressNotPostcoded 'Report cannot get postcode
        'Else
        '  PostcodeAddressGB = ValidatePostcodeStatuses.vpsError
        'End If
      End If
    End Function
    Private Function ValidateBuildingAFD(ByRef pBuildingNumber As Boolean, ByRef pBuilding As String, ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String, ByVal pUsePostcode As Boolean) As ValidatePostcodeStatuses
      Dim postcodeStatus As ValidatePostcodeStatuses
      Dim datatable As New CDBDataTable
      Dim addressString As String
      If Not mvInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      addressString = Join(pAddress, ",")
      addressString = Replace(addressString.ToUpper, pTown.ToUpper, "")
      addressString = addressString.TrimEnd(",".ToCharArray)
      postcodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
      mvCheckAddress = False

      datatable.AddColumnsFromList("AddressLine,Town,OrganisationName,County,Postcode,Address,DeliveryPointSuffix,Easting,Northing,LeaCode,LeaName,PostcoderID")
      If pUsePostcode Then
        AFDGetPostcode(addressString, pTown, pCounty, datatable, pPostCode)
      Else
        AFDGetPostcode(addressString, pTown, pCounty, datatable)
      End If

      If datatable.Rows.Count > 0 Then
        Dim searchAddress As New PostcoderAddress With {.Address = addressString, .BuildingNumber = pBuilding, .County = pCounty, .Postcode = pPostCode, .Town = pTown}

        For Each vRow In datatable.Rows
          Dim buildingNumber As String = ""
          If datatable.Columns.ContainsKey("BuildingNumber") Then
            buildingNumber = vRow.Item("BuildingNumber")
          End If
          Dim resultAddress As New PostcoderAddress With {.OrganisationName = vRow.Item("OrganisationName"), .Address = Replace(vRow.Item("AddressLine"), ", ", ","), .County = vRow.Item("County"), .Postcode = vRow.Item("Postcode"), .Town = vRow.Item("Town")}
          If PostcoderAddress.MatchAddress(resultAddress, searchAddress) Then
            mvAddress = String.Empty
            mvTown = String.Empty
            mvCounty = String.Empty
            mvPostCode = String.Empty
            mvOrgName = String.Empty
            mvBuildingNumber = ""
            ClearAdditionalItems()
            postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingValidated
            pAddress = Split(resultAddress.Address, ",")
            If pBuildingNumber Then pBuilding = resultAddress.BuildingNumber
            pTown = resultAddress.Town
            pCounty = resultAddress.County
            pPostCode = resultAddress.Postcode
            mvAddress = Replace(resultAddress.Address, ",", vbLf)
            mvTown = resultAddress.Town
            mvCounty = resultAddress.County
            mvPostCode = resultAddress.Postcode
            mvDPS = resultAddress.DPS
            mvEasting = resultAddress.Easting.ToString
            mvNorthing = resultAddress.Northing.ToString
            mvLEACode = resultAddress.LEACode
            mvLEAName = resultAddress.LEAName
            Exit For
          End If
        Next

        If postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingValidated Then
          If String.IsNullOrWhiteSpace(searchAddress.Postcode) AndAlso Not String.IsNullOrWhiteSpace(mvPostCode) Then
            'No Search Postcode so Address was postcoded
            postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsAddressPostcoded
          ElseIf Not String.IsNullOrWhiteSpace(searchAddress.Postcode) Then
            If String.Compare(searchAddress.Postcode, mvPostCode, StringComparison.InvariantCultureIgnoreCase) <> 0 Then
              postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsAddressRePostcoded
            End If
          End If
          mvCheckAddress = True
        Else
          postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingNotValidated
        End If
      Else
        postcodeStatus = Postcoder.ValidatePostcodeStatuses.vpsBuildingNotValidated
      End If
      Return postcodeStatus
    End Function
    
    Private Function ValidateBuildingGB(ByRef pBuildingNumber As Boolean, ByRef pBuilding As String, ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String) As ValidatePostcodeStatuses
      Dim vVPS As ValidatePostcodeStatuses
      Dim vGBSelect As String
      Dim vNewAddressInfo As AddressInformation
      Dim vGotNext As Boolean
      Dim vContinue As Boolean
      Dim vPostCodeStatus As ValidatePostcodeStatuses

      If Not mvInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      mvCheckAddress = False
      ClearAdditionalItems()
      mvOrgName = ""

      If mvGBIMTK Then
        vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
        If CallGBIMAPI("SETDATALEVEL('BUILDING')") Then
          vGBSelect = "SELECT CONFIGURE FROM NR WHERE WALK = '" & pPostCode & "'"
          If pBuildingNumber And Len(pBuilding) > 0 Then vGBSelect = vGBSelect & " AND (BNUM = '" & pBuilding & "' OR SUBB = '" & pBuilding & "')"
          If CallGBIMAPI(vGBSelect, True) Then
            Do While mvNoRecords = False
              If vGotNext Then
                vGotNext = False
                vContinue = True
              Else
                vContinue = CallGBIMAPI("GETNEXT")
              End If
              If vContinue Then
                If mvNoRecords = False Then
                  vNewAddressInfo = New AddressInformation
                  vNewAddressInfo.ExtractData(mvResponse)
                  If vNewAddressInfo.ValidBuilding(pBuildingNumber, pBuilding) Then
                    vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated
                    GBExtractAddress(1, vbCrLf)
                    If Len(vNewAddressInfo.Organisation) > 0 Then
                      If Len(vNewAddressInfo.SubBuilding) + Len(vNewAddressInfo.BuildingName) + Len(vNewAddressInfo.BuildingNumber) = 0 Then
                        mvOrgName = vNewAddressInfo.Organisation
                      End If
                    End If
                    If Len(pCounty) = 0 Then
                      CheckForCounty()
                    End If
                    If pBuildingNumber Then
                      If Not vNewAddressInfo.AddressMatch(pBuilding, pAddress(0)) Then
                        'Validate building worked but produced a different first word of the address
                        'Go and get the next record
                        If CallGBIMAPI("GETNEXT") Then
                          If mvNoRecords = False Then
                            'Since there is a next record, go back to the top of the loop to re-validate the building against this record
                            vGotNext = True
                          Else
                            'In this case the postcode is probably wrong and we should try to remove it and re-postcode the address
                            vVPS = PostcodeAddress(pAddress, pTown, pCounty, "", pCountry, False)
                            Select Case vVPS
                              Case ValidatePostcodeStatuses.vpsAddressPostcoded
                                vPostCodeStatus = ValidatePostcodeStatuses.vpsAddressRePostcoded
                              Case ValidatePostcodeStatuses.vpsAddressNotPostcoded
                                vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
                              Case ValidatePostcodeStatuses.vpsError
                                vPostCodeStatus = ValidatePostcodeStatuses.vpsError
                            End Select
                          End If
                        End If
                      End If
                    End If
                    If Not vGotNext Then Exit Do
                  ElseIf vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated Then
                    'we've been around the loop at least once at had the building match
                    'In this case the postcode is probably wrong and we should try to remove it and re-postcode the address
                    vVPS = PostcodeAddress(pAddress, pTown, pCounty, "", pCountry, False)
                    Select Case vVPS
                      Case ValidatePostcodeStatuses.vpsAddressPostcoded
                        vPostCodeStatus = ValidatePostcodeStatuses.vpsAddressRePostcoded
                      Case ValidatePostcodeStatuses.vpsAddressNotPostcoded
                        vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
                      Case ValidatePostcodeStatuses.vpsError
                        vPostCodeStatus = ValidatePostcodeStatuses.vpsError
                    End Select
                    Exit Do
                  End If
                End If
              End If
            Loop
          End If
        End If
        If vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated Then
          'Debug.Print
          'Debug.Print "Could NOT Validate: " & pAddress(0), pTown, pPostCode
          'GetPostcodeAddresses mvEnv, pPostCode
          'Debug.Print
        End If
      Else
        'Do
        '  'Try a VB call to see if the building matches - If OK then set paf status to VB
        '  vGBResponse = GBSend("GBVB 0001 " & Left(pPostCode & "         ", 9) & pBuilding)
        '  vRetry = False
        '  If vGBResponse = GBResponses.GBResponseOK Then
        '    GBExtractAddress(8)
        '    vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated
        '    If pBuildingNumber Then
        '      'Should do a soundex comparison here
        '      If LCase(FirstWord(Trim(Mid(pAddress(0), Len(pBuilding) + 1)))) <> LCase(FirstWord(Trim(Mid(mvResponse, 8 + Len(pBuilding), 35)))) Then
        '        'Validate building worked but produced a different first word of the address
        '        'In this case the postcode is probably wrong and we should try to remove it and re-postcode the address
        '        vVPS = PostcodeAddress(pAddress, pTown, pCounty, "", pCountry, False)
        '        Select Case vVPS
        '          Case ValidatePostcodeStatuses.vpsAddressPostcoded
        '            vPostCodeStatus = ValidatePostcodeStatuses.vpsAddressRePostcoded
        '          Case ValidatePostcodeStatuses.vpsAddressNotPostcoded
        '            vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
        '          Case ValidatePostcodeStatuses.vpsError
        '            vPostCodeStatus = ValidatePostcodeStatuses.vpsError
        '        End Select
        '      End If
        '    End If
        '  ElseIf vGBResponse = GBResponses.GBResponseNone Then
        '    vPos = InStr(pBuilding, " ")
        '    If vPos > 0 Then
        '      vPos = Len(pBuilding)
        '      While (Mid(pBuilding, vPos, 1) <> " ") And (vPos > 1)
        '        vPos = vPos - 1
        '      End While
        '      pBuilding = RTrim(Left(pBuilding, vPos))
        '      vRetry = True
        '    Else
        '      vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
        '    End If
        '  Else
        '    vPostCodeStatus = ValidatePostcodeStatuses.vpsError
        '  End If
        'Loop While vRetry
      End If
      ValidateBuildingGB = vPostCodeStatus
    End Function

    Private Function ValidatePostcodeGB(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String) As ValidatePostcodeStatuses
      Dim vVPS As ValidatePostcodeStatuses

      If Not mvInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      mvCheckAddress = False
      ClearAdditionalItems()
      If mvGBIMTK Then
        ValidatePostcodeGB = ValidatePostcodeStatuses.vpsPostcodeNotValidated
        vVPS = PostcodeAddress(pAddress, pTown, pCounty, "", pCountry, False)
        Select Case vVPS
          Case ValidatePostcodeStatuses.vpsAddressPostcoded
            If mvPostCode = pPostCode Then
              ValidatePostcodeGB = ValidatePostcodeStatuses.vpsPostcodeValidated
            Else
              ValidatePostcodeGB = ValidatePostcodeStatuses.vpsAddressRePostcoded
            End If
        End Select
      Else
        'Dim vGBAddress As String
        'Dim vGBResponse As GBResponses
        'vGBAddress = pAddress(0) & "|" & pAddress(1) & "|" & pAddress(2) & "|" & pAddress(3) & "|" & pTown & "|" & pCounty & "|" & pPostCode
        'vGBResponse = GBSend("GBVP " & vGBAddress)
        'If vGBResponse = GBResponses.GBResponseOK Then
        '  ValidatePostcodeGB = ValidatePostcodeStatuses.vpsPostcodeValidated
        'ElseIf vGBResponse = GBResponses.GBResponseNone Then
        '  'We have a postcode but cannot validate it - Let's try to re-postcode the address
        '  vVPS = PostcodeAddress(pAddress, pTown, pCounty, "", pCountry, False)
        '  Select Case vVPS
        '    Case ValidatePostcodeStatuses.vpsAddressPostcoded
        '      ValidatePostcodeGB = ValidatePostcodeStatuses.vpsAddressRePostcoded
        '    Case ValidatePostcodeStatuses.vpsAddressNotPostcoded
        '      ValidatePostcodeGB = ValidatePostcodeStatuses.vpsPostcodeNotValidated
        '    Case ValidatePostcodeStatuses.vpsError
        '      ValidatePostcodeGB = ValidatePostcodeStatuses.vpsError
        '  End Select
        'Else
        '  ValidatePostcodeGB = ValidatePostcodeStatuses.vpsError
        'End If
      End If
    End Function

    '-----------------------------------------------------------------------------------
    ' PRIVATE METHODS
    '-----------------------------------------------------------------------------------

    '  Private Function GBSend(ByVal pCommand As String) As GBResponses
    '    mvResponse = ""
    '    'Debug.Print "Sending Data " & Left$(pCommand, 4) & mvWinSock.LocalIP & Mid$(pCommand, 5)
    '    On Error GoTo WinsockError
    '    'UPGRADE_WARNING: Couldn't resolve default property of object mvWinSock.LocalIP. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    'UPGRADE_WARNING: Couldn't resolve default property of object mvWinSock.SendData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
    '    mvWinSock.SendData(Left(pCommand, 4) & mvWinSock.LocalIP & Mid(pCommand, 5))
    '    GBSend = GBGetResponse(pCommand)
    '    Exit Function

    'WinsockError:
    '    If LCase(Err.Source) = "winsock" Then
    '      RaiseError(DataAccessErrors.daeWinSockError, Err.Description)
    '    Else
    '      Err.Raise(Err.Number, Err.Source, Err.Description)
    '    End If
    '  End Function

    'Private Function GBGetResponse(ByVal pCommand As String) As GBResponses
    '  Dim vStartTime As Date
    '  Dim vResponse As Boolean
    '  Dim vGBResponse As GBResponses
    '  Dim vError As String
    '  Dim vIgnoreError As Boolean
    '  Dim vWaitTime As Integer

    '  vStartTime = Now
    '  vResponse = False
    '  vGBResponse = GBResponses.GBResponseNone
    '  Do
    '    System.Windows.Forms.Application.DoEvents()
    '    If mvResponse <> "" Then
    '      vResponse = True
    '      Select Case Left(pCommand, 4)
    '        Case "GBOP"
    '          If mvResponse = "GBOPOK" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          ElseIf mvResponse = "GBOPE81" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          Else
    '            If mvResponse = "GBOPE11" Or mvResponse = "GBOPE12" Then
    '              GBGetError(Mid(mvResponse, 5, 3))
    '              vGBResponse = GBResponses.GBResponseOK
    '            End If
    '          End If

    '        Case "GBCL"
    '          If mvResponse = "GBCLOK" Then vGBResponse = GBResponses.GBResponseOK

    '        Case "CBSV" 'Get the version number
    '          vGBResponse = GBResponses.GBResponseOK 'Any response is OK since we will check it later

    '        Case "GBTF"
    '          If Left(mvResponse, 6) = "GBTFOK" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          ElseIf Left(mvResponse, 6) = "GBTF  " Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          End If

    '        Case "GBBU"
    '          If Left(mvResponse, 6) = "GBBUOK" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          ElseIf Left(mvResponse, 6) = "GBBU  " Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          End If

    '        Case "GBBA"
    '          If Left(mvResponse, 6) = "GBBAOK" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          End If

    '        Case "GBVP"
    '          If Left(mvResponse, 6) = "GBVPOK" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          Else
    '            If Left(mvResponse, 7) = "GBVPN51" Or Left(mvResponse, 7) = "GBVPN52" Then
    '              'Debug.Print "Ignoring error: " & mvResponse
    '              vIgnoreError = True
    '            End If
    '          End If

    '        Case "GBVB"
    '          If Left(mvResponse, 6) = "GBVBOK" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          ElseIf Left(mvResponse, 6) = "GBVBN0" Then
    '            'Debug.Print "Ignoring error: " & mvResponse
    '            vIgnoreError = True
    '          End If

    '        Case "GBPC"
    '          If Left(mvResponse, 6) = "GBPCOK" Then
    '            vGBResponse = GBResponses.GBResponseOK
    '          ElseIf Left(mvResponse, 6) = "GBPCN3" Then
    '            vError = Mid(mvResponse, 6, 1)
    '            If vError = "1" Or vError = "2" Or vError = "3" Or vError = "4" Then
    '              vGBResponse = GBResponses.GBResponseOK
    '            End If
    '          End If
    '      End Select

    '      If vGBResponse = GBResponses.GBResponseNone Then
    '        If Left(mvResponse, 4) <> Left(pCommand, 4) Then
    '          mvLastError = gvSystem.LoadStringP2(16551, Left(pCommand, 4), Left(mvResponse, 4)) 'GB Mailing synchronisation error: %s %s
    '          vGBResponse = GBResponses.GBErrorSynch
    '        Else
    '          If Not vIgnoreError Then
    '            GBGetError(Mid(mvResponse, 5, 3))
    '            vGBResponse = GBResponses.GBErrorSynch
    '          End If
    '        End If
    '      End If
    '    Else
    '      'No response
    '      'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
    '      vWaitTime = DateDiff(Microsoft.VisualBasic.DateInterval.Second, vStartTime, Now)
    '      'If vWaitTime > 1 Then Debug.Print "W: " & vWaitTime
    '    End If
    '  Loop While vResponse = False And vWaitTime < mvUDPTimeout
    '  If vResponse = False Then
    '    mvLastError = gvSystem.LoadString(16552) 'GB Mailing server failed to respond
    '    vGBResponse = GBResponses.GBErrorResponse
    '  End If
    '  'Debug.Print vWaitTime
    '  GBGetResponse = vGBResponse
    'End Function

    Private Sub GBGetError(ByVal pCode As String)
      Dim vRecordSet As CDBRecordSet

      If pCode = "" Then
        mvLastError = ProjectText.String16547 'GB Mailing error - No Error Code
      Else
        vRecordSet = mvEnv.Connection.GetRecordSet("SELECT message, error_type FROM gb_message WHERE error_code = '" & pCode & "'")
        If vRecordSet.Fetch() Then
          If vRecordSet.Fields(2).Value = "W" Then
            MsgBox(String.Format(ProjectText.String16518, vRecordSet.Fields(1).Value), MsgBoxStyle.Information) 'GB Mailing Warning: %s
          Else
            mvLastError = vRecordSet.Fields(1).Value
          End If
        Else
          mvLastError = String.Format(ProjectText.String16548, pCode)
        End If
        vRecordSet.CloseRecordSet()
      End If
    End Sub

    Private Sub GBExtractAddress(ByVal pStartPos As Integer, Optional ByRef pSeparator As String = vbLf)
      Dim vAdd2 As String
      Dim vAdd3 As String
      Dim vAdd4 As String
      Dim vEastNorth As String

      mvAddress = Trim(Mid(mvResponse, pStartPos, 35))
      vAdd2 = Trim(Mid(mvResponse, 35 + pStartPos, 35))
      vAdd3 = Trim(Mid(mvResponse, 70 + pStartPos, 35))
      vAdd4 = Trim(Mid(mvResponse, 105 + pStartPos, 35))
      mvTown = Trim(Mid(mvResponse, 140 + pStartPos, 35))
      mvCounty = Trim(Mid(mvResponse, 175 + pStartPos, 35))
      mvPostCode = Trim(Mid(mvResponse, 210 + pStartPos, 8))
      If mvGBIMTK Then
        vEastNorth = RTrim(Mid(mvResponse, 504 + pStartPos))
      Else
        vEastNorth = RTrim(Mid(mvResponse, 218 + pStartPos))
      End If
      mvEasting = Mid(vEastNorth, 2, 6)
      mvNorthing = Mid(vEastNorth, 8, 6)
      If vAdd2.Length > 0 Then
        If mvAddress.Length > 0 Then mvAddress = mvAddress & pSeparator
        mvAddress = mvAddress & vAdd2
      End If
      If vAdd3.Length > 0 Then
        If mvAddress.Length > 0 Then mvAddress = mvAddress & pSeparator
        mvAddress = mvAddress & vAdd3
      End If
      If vAdd4.Length > 0 Then
        If mvAddress.Length > 0 Then mvAddress = mvAddress & pSeparator
        mvAddress = mvAddress & vAdd4
      End If
      mvCheckAddress = True
    End Sub

    Private Sub CheckForCounty()
      Dim vAddressParts() As String
      Dim vCount As Integer
      Dim vPos As Integer
      Dim vRemovedCounty As Boolean

      If Len(mvCounty) > 0 Then
        If InStr(mvAddress, vbCrLf & mvCounty) > 0 Then
          ''If the address returned from GB Mailing contains the county on a separate line, remove the county from the address
          'mvAddress = Replace(mvAddress, vbCrLf & mvCounty, "")
          vAddressParts = Split(mvAddress, vbCrLf)
          For vCount = 0 To UBound(vAddressParts)
            If InStr(1, vAddressParts(vCount), mvCounty) > 0 Then
              If vCount = UBound(vAddressParts) Then
                ' It's the last part of the address. Ensure that nothing is after it (in this string)
                ' If so, remove it
                If Len(vAddressParts(vCount)) = Len(mvCounty) Then
                  ' Remove it
                  vAddressParts(vCount) = ""
                  vRemovedCounty = True
                End If
              End If
            End If
          Next vCount
          ' Now rebuild the address string
          mvAddress = ""
          For vCount = 0 To UBound(vAddressParts)
            mvAddress = mvAddress & If(Len(vAddressParts(vCount)) > 0, vAddressParts(vCount) & vbCrLf, "")
          Next vCount
          If Right(mvAddress, 2) = vbCrLf Then mvAddress = Left(mvAddress, Len(mvAddress) - 2)
        End If

        ' Now do other checks
        If Right(mvAddress, Len(mvCounty) + 1) = " " & mvCounty Then
          vPos = InStr(1, mvAddress, " " & mvCounty)
          If vPos > 2 Then
            If UCase(Mid(mvAddress, vPos - 2, 2)) <> "OF" Then
              'If the address returned from GB Mailing ends with the county, remove the county from the address
              mvAddress = Mid(mvAddress, 1, Len(mvAddress) - (Len(mvCounty) + 1))
              vRemovedCounty = True
            Else
              ' Special case hack - Need to avoid chopping off the 'of' for such places as 'Bow of Fife'
              ' Add more as necessary (if we wish to go down this route!)
              ' Ignore the 'of' and leave it be
            End If
          End If
        End If

        If Not vRemovedCounty Then
          'If the county does not appear in the address returned from GB Mailing then it's not necessary to store it on the address, so clear it
          mvCounty = ""
        End If
      End If
    End Sub

    Private Sub ClearResponse()
      mvResponse = Space(MAX_RESPONSE)
    End Sub

    Private Function CallGBIMAPI(ByRef pQuery As String, Optional ByRef pCheckRecordsFound As Boolean = False) As Boolean
      Dim vReply As String
      Dim vResult As Integer

      ClearResponse()
      vReply = pQuery.Length.ToString & "     " 'Place length of query into reply
      If Len(mvGBHandle) = 0 Then
        mvGBHandle = "0000 "
      Else
        mvGBHandle = Right("0000" & CStr(Val(mvGBHandle)), 4) & " "
      End If
      ' Debug.Print pQuery
      vResult = GBIMAPI(pQuery, mvResponse, vReply, mvGBHandle)
      Select Case vResult
        Case 0
          CallGBIMAPI = True 'All OK
          If pCheckRecordsFound Then
            If Left(mvResponse, 1) = "0" Then
              mvNoRecords = True
            Else
              mvNoRecords = False
            End If
          Else
            mvNoRecords = False
          End If
        Case IMTKErrors.IMTKELInvalidRecordNumber
          mvNoRecords = True
          CallGBIMAPI = True 'OK but no records
        Case IMTKErrors.IMTKELicenceWillExpireSoon
          CallGBIMAPI = True 'OK for now
        Case Else
          GBGetError(CStr(vResult))
          mvNoRecords = True
      End Select
    End Function

    Private Sub GetBuildingInfo(ByRef pAddress() As String, ByRef pBuildingNo As String, ByRef pBuildingName As String, ByRef pThorofare As String)
      Dim vAddress As String
      Dim vBuildingNo As Integer

      vAddress = pAddress(0)
      If Val(vAddress) < 1 Or Val(vAddress) > 999999 Then
        pBuildingName = pAddress(0)
        pThorofare = pAddress(1)
      Else
        vBuildingNo = CInt(Val(vAddress))
        pBuildingNo = CStr(vBuildingNo)
        pThorofare = Trim(Mid(pAddress(0), Len(pBuildingNo) + 1))
      End If
    End Sub

    Friend Sub AddGBErrorMessages()
      Dim vWhereFields As New CDBFields

      vWhereFields.Add("error_code", CDBField.FieldTypes.cftCharacter, "101")
      If mvEnv.Connection.GetCount("gb_message", vWhereFields) = 0 Then
        AddGBError("101", "No handles available in reformatter")
        AddGBError("102", "Insufficient memory available for reformatter operation")
        AddGBError("103", "Invalid handle supplied to reformatter")
        AddGBError("104", "Could not locate configuration file for reformatter")
        AddGBError("105", "Could not find configurations for reformatter")
        AddGBError("106", "Configuration name to long")
        AddGBError("107", "Null pointer passed as address in reformatter")
        AddGBError("108", "Invalid line number supplied to reformatter")
        AddGBError("109", "Mutex control failure in reformatter")
        AddGBError("300", "No free handles")
        AddGBError("301", "Insuffiecient memory for an NRT operation")
        AddGBError("303", "Unable to initialise the internal postcode index")
        AddGBError("304", "Insufficient file handles")
        AddGBError("305", "Unable to open a required file")
        AddGBError("306", "GBPATH could not be set correctly")
        AddGBError("307", "Required file not found")
        AddGBError("309", "Error seeking in required file")
        AddGBError("310", "Error reading in required file")
        AddGBError("326", "Unable to access control index")
        AddGBError("327", "Mismatched data files")

        AddGBError("337", "One or more NR files missing or GBPATH incorrect")
        AddGBError("342", "cifgen.ini is in the wrong case (should be lower case)")
        AddGBError("344", "GBACCESS>CIF is in the wrong case (should be upper) or GBPATH set incorrectly")
        AddGBError("346", "Could not find GOBO.DAT on GBPATH")
        AddGBError("347", "Checksum error in GOBO.DAT possibly a corrupt file")
        AddGBError("348", "GOBO access has not been initialised")
        AddGBError("1321", "Error occured in batch engine during NRT filtering")
        AddGBError("1324", "Could not open query bitmap file in batch engine")
        AddGBError("1326", "Query bitmap is not valid for the current version of the NRF")
        AddGBError("1327", "Batch engine could not read from bitmap file")
        AddGBError("1340", "Operator not valid in 16 bit environments")
        AddGBError("1701", "Memory allocation error")
        AddGBError("1702", "No free handles available")
        AddGBError("1703", "Invalid handle supplied")
        AddGBError("1704", "Invalid record number")
        AddGBError("1707", "No configuration loaded")
        AddGBError("1711", "Process timed out")
        AddGBError("1714", "Invalid IMQL syntax specified")
        AddGBError("1715", "Accelerator Developer Toolkit license will expire soon")
        AddGBError("1716", "Invalid dataset")
        AddGBError("1717", "Fatal security error")
        AddGBError("1718", "Mutex error")
        AddGBError("1719", "Data encryption/Decryption error")
        AddGBError("1720", "Invalid field number")

        AddGBError("1721", "Invalid context")
        AddGBError("1722", "Authentication is disabled")
        AddGBError("1723", "Could not find Authentication DLL")
        AddGBError("1724", "Could not load Authentication function")
        AddGBError("1725", "Could not load Authentication version function")
        AddGBError("1726", "Could not load Authentication get version function")
        AddGBError("1727", "Invalid data code specified")
        AddGBError("1728", "Invalid reply code specified")
        AddGBError("1729", "Invalid Authenticate function")
        AddGBError("1730", "Cannot proceed. Authentication logic not loaded")
      End If
    End Sub

    Friend Sub AddGBError(ByRef pErrorCode As String, ByRef pErrorMsg As String)
      Dim vFields As New CDBFields

      vFields.Add("error_code", CDBField.FieldTypes.cftCharacter, pErrorCode)
      vFields.Add("message", CDBField.FieldTypes.cftCharacter, pErrorMsg)
      vFields.Add("error_type", CDBField.FieldTypes.cftCharacter, "F")
      mvEnv.Connection.InsertRecord("gb_message", vFields)
    End Sub

    Public Function GetAddressInfo(ByVal pIndex As Integer) As AddressInformation
      Return mvBuildings(pIndex)
    End Function

#End Region

#Region " QAS "

    Public ReadOnly Property QASBatchReportCode() As String
      Get
        QASBatchReportCode = mvQASBatchReportCode
      End Get
    End Property

    Public ReadOnly Property QAProType() As QuickAddressTypes
      Get
        QAProType = mvQAType
      End Get
    End Property

    Private Function PostcodeAddressQAS(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pStrict As Boolean, Optional ByVal pDataTable As CDBDataTable = Nothing) As ValidatePostcodeStatuses
      'pDataTable is currently used in PostcodeAddress to get all the available postcodes and not just the first matching one
      If Not mvQASInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      QAProOpen()
      Dim vSearch As String = Join(pAddress, vbCrLf)
      If vSearch.Length > 0 Then vSearch = vSearch & ","
      vSearch = vSearch & pTown & ","
      If Len(pCounty) > 0 Then vSearch = vSearch & pCounty & "@C,"
      If Len(pPostCode) > 0 Then vSearch = vSearch & pPostCode
      If Right(vSearch, 1) = "," Then vSearch = Left(vSearch, Len(vSearch) - 1)
      vSearch = Replace(vSearch, vbCrLf, ",")
      Dim vQASError As Integer = QAProSearch(vSearch)
      Dim vStatus As ValidatePostcodeStatuses = ValidatePostcodeStatuses.vpsNone
      If vQASError < 0 Then
        QAProEndSearch()
        RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
      ElseIf vQASError > 0 Then
        QAProEndSearch()
        RaiseError(DataAccessErrors.daeInvalidParameter, "Postcode")
      Else
        vStatus = PostcodeAddressesQAS(pAddress, pTown, pCounty, pPostCode, pStrict, pDataTable)
      End If
      QAProEndSearch()
      If Not pStrict Then QAProClose()
      PostcodeAddressQAS = vStatus
    End Function

    Private Function PostcodeAddressesQAS(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pStrict As Boolean, ByVal pDataTable As CDBDataTable, Optional ByRef pStepOutCount As Integer = 0) As ValidatePostcodeStatuses
      'Called from PostcodeAddressQAS. Use PostcodeAddressQAS in all cases.
      Dim vResult As QAProGetItemResults
      Dim vQASError As Integer
      Dim vConfidence As Integer
      Dim vQASCount As Integer

      Dim vStatus As ValidatePostcodeStatuses = ValidatePostcodeStatuses.vpsNone
      Do
        vQASCount = QAProCount()
        If vQASCount < 0 Then
          QAProEndSearch()
          RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASCount))
        ElseIf vQASCount = 1 Then
          'We have only one possible return
          vQASError = QAProGetItemInfo(0, QAProGetItemInfoTypes.qapStepInfo, vResult, vConfidence)
          If vQASError < 0 Then
            QAProEndSearch()
            RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
          Else
            If vResult = QAProGetItemResults.qgirNoStepIn Then
              'We don't need to step in so just return the data
              If QAGetAddress(0, False) >= 0 Then
                'Debug.Print "Found: ", gvSystem.FirstLine(Address), Town, County, Postcode
                If pStrict Then
                  If QASAddressMatch(pAddress, pTown, pPostCode, "", False) Then
                    vStatus = ValidatePostcodeStatuses.vpsAddressPostcoded 'We are done
                  End If
                Else
                  If pDataTable IsNot Nothing Then AddAddressToTable(pDataTable)
                  vStatus = ValidatePostcodeStatuses.vpsAddressPostcoded 'We are done
                End If
              End If
            ElseIf vResult = QAProGetItemResults.qgirStepInRequired Then
              vQASError = QAProStepIn(0)
              If vQASError < 0 Then
                QAProEndSearch()
                RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
              ElseIf pStepOutCount > 0 Then
                pStepOutCount += 1
              End If
            End If
          End If
        Else
          'There is more than one possible item so iterate looking for the first matching address
          For vBIndex As Integer = 0 To vQASCount - 1
            If pDataTable IsNot Nothing Then  'Only for web service
              vQASError = QAProGetItemInfo(vBIndex, QAProGetItemInfoTypes.qapStepInfo, vResult, vConfidence)
              If vQASError < 0 Then
                QAProEndSearch()
                RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
              Else
                If vResult = QAProGetItemResults.qgirStepInRequired Then
                  vQASError = QAProStepIn(vBIndex)
                  If vQASError < 0 Then
                    QAProEndSearch()
                    RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
                  Else
                    Dim vStepOutCount As Integer = 1
                    vStatus = PostcodeAddressesQAS(pAddress, pTown, pCounty, pPostCode, pStrict, pDataTable, vStepOutCount)
                    While vStepOutCount > 0
                      vQASError = QAProStepOut()
                      vStepOutCount -= 1
                    End While
                    If vQASError < 0 Then
                      QAProEndSearch()
                      RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
                    Else
                      Continue For
                    End If
                  End If
                End If
              End If
            End If

            If QAGetAddress(vBIndex, False) >= 0 Then
              'Debug.Print "Multiple Entries:", gvSystem.FirstLine(Address), Town, County, Postcode
              If pStrict Then
                If QASAddressMatch(pAddress, pTown, pPostCode, "", False) Then
                  vStatus = ValidatePostcodeStatuses.vpsAddressPostcoded 'We are done
                  Exit For
                End If
              Else
                If ((Val(Address) > 0) AndAlso (Val(Address) = Val(pAddress(0)))) OrElse (StrComp(FirstLine(Address), pAddress(0), CompareMethod.Text) = 0 OrElse pDataTable IsNot Nothing) Then
                  vStatus = ValidatePostcodeStatuses.vpsAddressPostcoded 'We are done
                  If pDataTable IsNot Nothing Then
                    AddAddressToTable(pDataTable)
                  Else
                    Exit For
                  End If
                End If
              End If
            End If
          Next
        End If
      Loop While vQASCount = 1 And vQASError >= 0 And vResult = QAProGetItemResults.qgirStepInRequired
      Return vStatus
    End Function

    Private Function ValidateBuildingQASBatch(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String) As ValidatePostcodeStatuses
      Dim vSearch As String = ""
      Dim vQASError As Integer
      Dim vPostCodeStatus As ValidatePostcodeStatuses
      Dim vSearchHandle As Integer
      Dim vPostcode As String
      Dim vISOCode As String
      Dim vReturnCode As String
      Dim vLineCount As Integer
      Dim vIndex As Integer
      Dim vAddrLine As String

      If Not mvQASBatchInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      mvCheckAddress = False
      Debug.Print("ValidateBuildingQASBatch: " & Join(pAddress, ",") & " " & pTown & " " & pPostCode)

      vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
      mvAddress = ""
      mvTown = ""
      mvCounty = ""
      mvPostCode = ""
      mvOrgName = ""
      mvBuildingNumber = ""
      ClearAdditionalItems()
      QABatchOpen()
      If Len(pAddress(0)) > 0 Then vSearch = Join(pAddress, ",") & ","
      If Len(pTown) > 0 Then vSearch = vSearch & pTown & ","
      If Len(pPostCode) > 0 Then vSearch = vSearch & pPostCode
      If Right(vSearch, 1) = "," Then vSearch = Left(vSearch, Len(vSearch) - 1)
      vSearch = vSearch & Chr(0)
      vISOCode = Space(10)
      vPostcode = Space(20)
      vReturnCode = Space(30)
      vQASError = QABatchWV_Clean(mvQASBatchHandle, vSearch, vSearchHandle, vPostcode, 20, vISOCode, vReturnCode, 30)
      If vQASError < 0 Then
        QABatchWV_EndSearch(vSearchHandle)
        RaiseError(DataAccessErrors.daeQuickAddressError, QASBatchErrorMessage(vQASError))
      Else
        If Left(vReturnCode, 2) = "R9" Or Left(vReturnCode, 2) = "P9" Then
          vQASError = QABatchWV_FormattedLineCount(vSearchHandle, vLineCount)
          For vIndex = 0 To vLineCount - 1
            vAddrLine = Space(81)
            vQASError = QABatchWV_GetFormattedLine(vSearchHandle, vIndex, vAddrLine, 50)
            If vQASError < 0 Then
              QABatchWV_EndSearch(vSearchHandle)
              RaiseError(DataAccessErrors.daeQuickAddressError, QASBatchErrorMessage(vQASError))
            Else
              SetStringLen(vAddrLine)
              If Len(vAddrLine) > 0 Or vIndex = 8 Then
                'If vQASError = QAERR_FIELDTRUNCATED Then Debug.Print vAddrLine
                Select Case vIndex
                  Case 0
                    mvOrgName = vAddrLine
                  Case 1
                    mvAddress = vAddrLine & Chr(13) & Chr(10)
                  Case 2
                    mvAddress = mvAddress & vAddrLine & Chr(13) & Chr(10)
                  Case 3
                    mvAddress = mvAddress & vAddrLine & Chr(13) & Chr(10)
                  Case 4
                    mvAddress = mvAddress & vAddrLine & Chr(13) & Chr(10)
                  Case 5
                    mvTown = UCase(vAddrLine)
                  Case 6
                    mvCounty = vAddrLine
                  Case 7
                    mvPostCode = UCase(vAddrLine)
                  Case 8
                    If False And Len(vAddrLine) = 0 Then
                      If Len(mvAddress) > 0 Then
                        mvAddress = mvOrgName & vbCrLf & mvAddress
                      Else
                        mvAddress = mvOrgName
                      End If
                    End If
                  Case 9
                    mvEasting = CStr(Val(vAddrLine) * 10)
                  Case 10
                    mvNorthing = CStr(Val(vAddrLine) * 10)
                End Select
              End If
            End If
          Next
          If Right(mvAddress, 2) = Chr(13) & Chr(10) Then mvAddress = Left(mvAddress, Len(mvAddress) - 2)
          If Left(vReturnCode, 2) = "R9" Then
            vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated
            mvCheckAddress = True
          ElseIf Left(vReturnCode, 2) = "P9" Then
            If Len(pAddress(0)) > 0 And LCase(pAddress(0)) = LCase(FirstLine(Address)) Then
              vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated
              mvCheckAddress = True
            Else
              If pPostCode = mvPostCode Then
                vPostCodeStatus = ValidatePostcodeStatuses.vpsPostcodeValidated
              Else
                mvQASBatchReportCode = vReturnCode
                vPostCodeStatus = ValidatePostcodeStatuses.vpsQASBatchReportCode
              End If
            End If
          End If
          If vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated Then
            Debug.Print("Building Validated: " & Left(vReturnCode, 2) & " " & Replace(mvAddress, vbCrLf, ",") & " " & mvTown & " " & mvPostCode)
            If pPostCode <> mvPostCode Then
              If Len(pPostCode) = 0 Then
                vPostCodeStatus = ValidatePostcodeStatuses.vpsAddressPostcoded
              Else
                vPostCodeStatus = ValidatePostcodeStatuses.vpsAddressRePostcoded
              End If
            End If
            'If mvTown <> pTown Or mvPostCode <> pPostCode Or gvSystem.FirstLine(mvAddress) <> pAddress(0) Then Stop
          End If
        Else
          mvQASBatchReportCode = vReturnCode
          vPostCodeStatus = ValidatePostcodeStatuses.vpsQASBatchReportCode
        End If
      End If
      QABatchWV_EndSearch(vSearchHandle)
      ValidateBuildingQASBatch = vPostCodeStatus
    End Function

    Private Function ValidateBuildingQAS(ByRef pBuildingNumber As Boolean, ByRef pBuilding As String, ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String, ByRef pCountry As String) As ValidatePostcodeStatuses
      Dim vSearch As String = ""
      Dim vQASError As Integer
      Dim vBIndex As Integer
      Dim vQASCount As Integer
      Dim vResult As QAProGetItemResults
      Dim vPostCodeStatus As ValidatePostcodeStatuses
      Dim vConfidence As Integer
      Dim vUseBuildingNumber As Boolean

      If Not mvQASInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      mvCheckAddress = False
      'Debug.Print "ValidateBuildingQAS: ", pAddress(0), pTown, pPostCode

      vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingNotValidated
      QAProOpen()
      If Len(pCountry) > 0 Then
        QASetCountry(pCountry)
        If pCountry = "NL" And pBuildingNumber Then vUseBuildingNumber = True
      End If
      If Len(pAddress(0)) > 0 Then vSearch = Join(pAddress, ",") & ","
      If vUseBuildingNumber Then vSearch = vSearch & pBuilding & ","
      If Len(pTown) > 0 Then vSearch = vSearch & pTown & "@T,"
      If Len(pPostCode) > 0 Then vSearch = vSearch & pPostCode & "@X"
      If Right(vSearch, 1) = "," Then vSearch = Left(vSearch, Len(vSearch) - 1)
      vQASError = QAProSearch(vSearch)
      If vQASError < 0 Then
        QAProEndSearch()
        RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
      ElseIf vQASError > 0 Then
        QAProEndSearch()
        RaiseError(DataAccessErrors.daeParameterNotFound, "Postcode")
      Else
        Do
          vQASCount = QAProCount()
          If vQASCount < 0 Then
            QAProEndSearch()
            RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASCount))
          ElseIf vQASCount = 1 Then
            'We have only one possible return
            vQASError = QAProGetItemInfo(0, QAProGetItemInfoTypes.qapStepInfo, vResult, vConfidence)
            If vQASError < 0 Then
              QAProEndSearch()
              RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
            Else
              If vResult = QAProGetItemResults.qgirNoStepIn Then
                'We don't need to step in so just return the data
                For vInt As Integer = 1 To 2
                  'Loop twice- First check building numbers without removing dashes and spaces
                  'If no match do standard building number check (removing dashes and spaces)
                  If vPostCodeStatus <> ValidatePostcodeStatuses.vpsBuildingValidated Then
                    Dim vCheckBuildingNumberExact As Boolean = (vInt = 1)
                    If QAGetAddress(0, False) >= 0 Then
                      'Debug.Print "Validate: " & gvSystem.FirstLine(Address), Town, County, Postcode
                      If QASAddressMatch(pAddress, pTown, pPostCode, pBuilding, vUseBuildingNumber, vCheckBuildingNumberExact) Then
                        vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated 'Done
                      End If
                    End If
                  End If
                Next
              ElseIf vResult = QAProGetItemResults.qgirStepInRequired Then
                vQASError = QAProStepIn(0)
                If vQASError < 0 Then
                  QAProEndSearch()
                  RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
                End If
              End If
            End If
          Else
            'We looked up using the building and postcode but got more than one result
            'Don't reckon we can call it VB if there was more than one
            For vInt As Integer = 1 To 2
              'Loop twice- First check building numbers without removing dashes and spaces
              'If no match do standard building number check (removing dashes and spaces)
              If vPostCodeStatus <> ValidatePostcodeStatuses.vpsBuildingValidated Then
                Dim vCheckBuildingNumberExact As Boolean = (vInt = 1)
                For vBIndex = 0 To vQASCount - 1
                  If QAGetAddress(vBIndex, False) >= 0 Then
                    'Debug.Print "Multiple Entries:", gvSystem.FirstLine(Address), Town, County, Postcode
                    If QASAddressMatch(pAddress, pTown, pPostCode, pBuilding, vUseBuildingNumber, vCheckBuildingNumberExact) Then
                      vPostCodeStatus = ValidatePostcodeStatuses.vpsBuildingValidated
                      vResult = 0
                      Exit For
                    End If
                  End If
                Next
              End If
            Next
          End If
        Loop While vQASCount = 1 And vQASError >= 0 And vResult = QAProGetItemResults.qgirStepInRequired
      End If
      QAProEndSearch()
      'QAProClose
      ValidateBuildingQAS = vPostCodeStatus
    End Function

    Private Function QASAddressMatch(ByRef pAddress() As String, ByRef pTown As String, ByRef pPostCode As String, ByRef pBuildingNumber As String, ByRef pCheckBuildingNumber As Boolean) As Boolean
      Return QASAddressMatch(pAddress, pTown, pPostCode, pBuildingNumber, pCheckBuildingNumber, False)
    End Function

    Private Function QASAddressMatch(ByRef pAddress() As String, ByRef pTown As String, ByRef pPostCode As String, ByRef pBuildingNumber As String, ByRef pCheckBuildingNumber As Boolean, ByVal pCheckBuildingNumberExact As Boolean) As Boolean
      Dim vGivenBno As String = ""
      Dim vReturnedBno As String = ""
      Dim vGivenAddressLine As String
      Dim vReturnedAddressLine As String

      If Len(pPostCode) > 0 And Replace(pPostCode, " ", "") <> Replace(Postcode, " ", "") Then Exit Function
      If Len(pTown) > 0 And pTown <> Town And mvQASDataSet <> "NL" Then Exit Function

      If Len(pBuildingNumber) > 0 Then
        vGivenBno = UCase(pBuildingNumber)
        If pCheckBuildingNumberExact = False Then
          vGivenBno = UCase(Replace(Replace(pBuildingNumber, " ", ""), "-", ""))
        End If
      End If
      If Len(BuildingNumber) > 0 Then
        vReturnedBno = UCase(BuildingNumber)
        If pCheckBuildingNumberExact = False Then
          vReturnedBno = UCase(Replace(Replace(BuildingNumber, " ", ""), "-", ""))
        End If
      End If
      vGivenAddressLine = UCase(Replace(Replace(pAddress(0), "-", ""), " ", ""))
      vReturnedAddressLine = UCase(Replace(Replace(FirstLine(Address), "-", ""), " ", ""))

      If pCheckBuildingNumber Then
        If vGivenBno = vReturnedBno Then
          QASAddressMatch = True
          mvCheckAddress = True
          'Debug.Print "Address Validated"
        ElseIf Len(vGivenBno) > 0 And Len(vReturnedBno) = 0 Then
          If Right(vReturnedAddressLine, Len(vGivenBno)) = vGivenBno Then
            QASAddressMatch = True
            mvCheckAddress = True
            'Debug.Print "Address Validated"
          End If
        End If
      Else
        If Len(vReturnedBno) > 0 Then
          'We got a building number back from QAS
          If Right(vGivenAddressLine, Len(vReturnedBno)) = vReturnedBno Then
            'Here we found the building number at the end of the given address so we should assume that this is a match
            QASAddressMatch = True
            mvCheckAddress = True
            'Debug.Print "Address Validated  " & pAddress(0) & " " & Address & " (" & BuildingNumber & ")"
          Else
            'Debug.Print "Address NOT Validated  " & pAddress(0) & " " & Address & " (" & BuildingNumber & ")"
          End If
        Else
          'No building number from QAS so compare the first line of the addresses
          If vGivenAddressLine.Length > 0 Then
            'BR13924: Special case for addresses having the word 'Flat' as first line e.g. importing 19 Redhall Close where the actual address is Flat vbCrLf 19 Redhall Close
            If vReturnedAddressLine.ToUpper = "FLAT" Then
              Dim vPos As Integer = Address.IndexOfAny(vbCrLf.ToCharArray)
              'Try to get the second line
              If vPos >= 0 Then vReturnedAddressLine = UCase(Replace(Replace(FirstLine(Address.Substring(vPos + 2)), "-", ""), " ", ""))
            End If
            If vGivenAddressLine = vReturnedAddressLine Then
              QASAddressMatch = True
              mvCheckAddress = True
              'Debug.Print "Address Validated"
            End If
          End If
        End If
      End If
    End Function

    Private Function ValidatePostcodeQAS(ByRef pAddress() As String, ByRef pTown As String, ByRef pCounty As String, ByRef pPostCode As String) As ValidatePostcodeStatuses
      Dim vVPS As ValidatePostcodeStatuses

      If Not mvQASInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)
      ClearAdditionalItems()
      mvCheckAddress = False
      ValidatePostcodeQAS = ValidatePostcodeStatuses.vpsPostcodeNotValidated
      vVPS = PostcodeAddressQAS(pAddress, pTown, pCounty, "", True)
      Select Case vVPS
        Case ValidatePostcodeStatuses.vpsAddressPostcoded
          If mvPostCode = pPostCode And Len(pPostCode) > 0 Then
            ValidatePostcodeQAS = ValidatePostcodeStatuses.vpsPostcodeValidated
          ElseIf Len(mvPostCode) > 0 Then
            ValidatePostcodeQAS = ValidatePostcodeStatuses.vpsAddressRePostcoded
          Else
            ValidatePostcodeQAS = ValidatePostcodeStatuses.vpsPostcodeNotValidated
          End If
      End Select
    End Function
    Public Sub GetAddresses(ByVal postcode As String, ByVal buildingNumber As String, ByVal iso3Countrycode As String, ByVal postcoderID As String, Optional ByRef pDT As CDBDataTable = Nothing, Optional ByRef pHeadings As String = Nothing, Optional ByRef pWidths As String = Nothing)
      If Not mvQASODInitialised Then RaiseError(DataAccessErrors.daePostCoderNotInitialised)

      Try
        If mvQAType = QuickAddressTypes.qatProOnDemand Then
          Dim results As IEnumerable(Of IAddress)
          Dim verifyLevel As Boolean = False
          results = QASODPostcodeValidator.GetAddresses({New PostcoderAddress With {.Postcode = postcode,
                                                                                    .BuildingNumber = buildingNumber,
                                                                                    .PostcoderID = postcoderID,
                                                                                    .Iso3166Alpha3CountryCode = iso3Countrycode}}, verifyLevel)

          If Not verifyLevel AndAlso pHeadings IsNot Nothing AndAlso pWidths IsNot Nothing Then ' Postcode has not been verified results will be Similar Matches
            'Adjust heading of datatable to show postcode first 
            If pDT.ColumnNames().Contains("BuildingNumber") Then
              pDT = New CDBDataTable
              pDT.AddColumnsFromList("Postcode,AddressLine,BuildingNumber,Town,OrganisationName,County,Address,DeliveryPointSuffix,Easting,Northing,LeaCode,LeaName,PostcoderID")
              pHeadings = "Postcode,Address,Building,Town,Organisation Name,County,Address"
              pWidths = "15,1200,900,1200,1200,1,1,1,1,1,1,1,1"
            Else
              pDT = New CDBDataTable
              pDT.AddColumnsFromList("Postcode,AddressLine,Town,County,OrganisationName,Address,DeliveryPointSuffix,Easting,Northing,LeaCode,LeaName,PostcoderID")
              pHeadings = "Postcode,Address,Town,County,Organisation Name,Address"
              pWidths = "15,1200,1200,1200,1200,1,1,1,1,1,1,1"
            End If
          End If
          PostcoderAddress.ConvertIAddressToDataTable(results, pDT)
        Else
          Throw New ArgumentException("Function not implemented for this Address Type", "QuickAddressTypes")
        End If
      Catch vEx As Exception
        RaiseError(DataAccessErrors.daeQASProOnDemandGeneralError, vEx.Message)
      End Try
    End Sub
    Public Function IsCountryLicenced(ByVal countryCodes As String) As Boolean
      Dim vPos As Integer = 0
      Dim vStart As Integer = 1
      Dim country As String = Nothing
      Dim isoCountryList As String = Nothing
      Dim vWhereFields As New CDBFields
      Do
        vPos = InStr(vStart, countryCodes, "|")
        If vPos > 0 Then
          country = Trim(Mid(countryCodes, vStart, vPos - vStart))
        Else
          country = Trim(Mid(countryCodes, vStart))
        End If

        ConvertCountryCodeToISO3(country)
        If String.IsNullOrWhiteSpace(country) Then
          Throw New ArgumentException("QAS Country Codes have not been set-up with ISO 3 Country Codes in the Countries table.", "qas_country_codes")
        Else
          If InStr(mvQASODLicencedCountries, country) = 0 Then Return False
        End If
        vStart = vPos + 1
      Loop While vPos > 0
      Return True
    End Function
    Public Function ConvertCountryCodeToISO3(ByRef countryCode As String) As Boolean
      Dim vCountry As New Country(mvEnv)
      vCountry.Init(countryCode)
      countryCode = vCountry.Iso3166Alpha3CountryCode
      Return True
    End Function
    Public Function InitQASProOnDemand(ByRef pUseBatch As Boolean) As Boolean
      Try
        Dim iso3DefaultCountryCode As String = mvEnv.DefaultCountry
        If Not String.IsNullOrWhiteSpace(iso3DefaultCountryCode) Then ConvertCountryCodeToISO3(iso3DefaultCountryCode)
        QASODPostcodeValidator = PostcodeValidatorFactory.GetValidator(mvEnv, iso3DefaultCountryCode)
        If mvPostcoderType = PostcoderTypes.pctQAS AndAlso QASODPostcodeValidator IsNot Nothing AndAlso Not mvQASODInitialised Then
          mvQASODLicencedCountries = QASODPostcodeValidator.GetLicencedCountries
          If IsCountryLicenced(mvEnv.DefaultCountry) AndAlso Not String.IsNullOrWhiteSpace(mvQASODLicencedCountries) Then 'Check Base Country
            If InStr(mvCountries, "|") > 0 Then 'More than one country in qas_country_codes
              If Not IsCountryLicenced(mvCountries) Then Exit Function
            End If
            mvQAType = QuickAddressTypes.qatProOnDemand
            mvQASODInitialised = True
            InitQASProOnDemand = mvQASODInitialised
            Exit Function
          End If
        End If
      Catch vEx As Exception
        RaiseError(DataAccessErrors.daeQASProOnDemandGeneralError, vEx.Message)
      End Try
    End Function
    Public Function InitQAS(ByRef pUseBatch As Boolean) As Boolean
      Dim vQASError As Integer
      Dim vSection As String = ""
      Dim vEntry As String = ""
      Dim vPos As Integer
      Dim vINIFile As String
      Dim vINIChecked As String
      Dim vMsg As String
      Dim vQASINIFile As String = ""
      Dim vDataPlusLines As Integer
      Dim vAddressLines As Integer
      Dim vUpdateQASINI As Boolean

      If Not String.IsNullOrWhiteSpace(mvEnv.GetConfig("qas_pro_ondemand_url")) Then
        InitQAS = InitQASProOnDemand(pUseBatch)
        Exit Function
      End If



      On Error GoTo QASInterfaceError

      If Not mvQASInitialised And mvPostcoderType = PostcoderTypes.pctQAS Then
        vSection = "CDBDefault"
        vINIFile = mvEnv.GetConfig("qas_ini_path")
        'If the ini path config is set then assume QAS V4 or above
        If Len(vINIFile) > 0 Then
          mvQAType = QuickAddressTypes.qatProV4orAbove
          mvQASGBRLines = 9
          If Right(vINIFile, 1) <> "\" Then vINIFile = vINIFile & "\"
          vQASINIFile = vINIFile & "QAWSERVE.INI"
          mvQASINIFile = vINIFile & QAINIFileName()
          vINIChecked = mvQASINIFile
          If My.Computer.FileSystem.FileExists(mvQASINIFile) Then vEntry = GetQAEntry("QADefault", "Language", "", mvQASINIFile)
        Else
          mvQAType = QuickAddressTypes.qatClientServer 'Try Client server first
          mvQASGBRLines = 8
          Dim vWinDirName As String = Environment.GetFolderPath(Environment.SpecialFolder.System)
          vWinDirName = IO.Path.Combine(vWinDirName, "..")
          mvQASINIFile = vWinDirName & QAINIFileName()
          vINIChecked = mvQASINIFile
          If My.Computer.FileSystem.FileExists(mvQASINIFile) Then
            vEntry = GetQAEntry("QADefault", "AddressLines", "")
          Else
            mvQAType = QuickAddressTypes.qatNonClientServer 'Try Non Client server next

            mvQASINIFile = vWinDirName & QAINIFileName()
            vINIChecked = vINIChecked & " or " & mvQASINIFile
            If My.Computer.FileSystem.FileExists(mvQASINIFile) Then
              vEntry = GetQAEntry("QADefault", "AddressLines", "")
            Else
              mvQAType = QuickAddressTypes.qatProV4orAbove
              mvQASINIFile = vWinDirName & QAINIFileName()
              vQASINIFile = vWinDirName & "QAWSERVE.INI"
              vINIChecked = vINIChecked & " or " & mvQASINIFile
              If My.Computer.FileSystem.FileExists(mvQASINIFile) Then vEntry = GetQAEntry("QADefault", "Language", "")
            End If
          End If
        End If
        If Len(vEntry) = 0 Then
          RaiseError(DataAccessErrors.daeQuickAddressCannotFind, vINIChecked) 'Failed to Initialise QuickAddress - Cannot find %1 or the process user may not have the correct permissions for this file
        Else
          If mvQAType = QuickAddressTypes.qatProV4orAbove Then
            'See if there is a local server file
            vINIFile = vQASINIFile
            If My.Computer.FileSystem.FileExists(vQASINIFile) Then
              mvQASDPS = mvEnv.GetConfigOption("qas_delivery_point_suffix")
              mvQASGRD = mvEnv.GetConfigOption("qas_grid_references")
              mvQASLEA = mvEnv.GetConfigOption("qas_lea_data")
              vAddressLines = 9
              If mvQASDPS Then
                mvQASDPSIndex = vAddressLines
                vAddressLines = vAddressLines + 1
              End If
              If mvQASGRD Then
                mvQASGRDEastIndex = vAddressLines
                mvQASGRDNorthIndex = vAddressLines + 1
                vAddressLines = vAddressLines + 2
                vDataPlusLines = vDataPlusLines + 2
              End If
              If mvQASLEA Then
                mvQASLEACodeIndex = vAddressLines
                mvQASLEANameIndex = vAddressLines + 1
                vAddressLines = vAddressLines + 2
                vDataPlusLines = vDataPlusLines + 2
              End If
              If Val(GetQAEntry(vSection, "GBRAddressLineCount", "", vQASINIFile)) <> vAddressLines Then vUpdateQASINI = True
              If vUpdateQASINI = False Then
                If Val(GetQAEntry(vSection, "GBRDataPlusLines", "", vQASINIFile)) <> vDataPlusLines Then vUpdateQASINI = True
              End If
              mvQASGBRLines = vAddressLines
              If vUpdateQASINI Then
                SetQAEntry(vSection, "GBRComment", "CARE Formatting", vQASINIFile)
                SetQAEntry(vSection, "GBRConfigured", "220", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine1", "W80,O11", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine2", "W35", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine3", "W35", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine4", "W35", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine5", "W35", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine6", "W35,L21", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine7", "W35,L11", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine8", "W8,C11", vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLine9", "W35,P22,P21,P11,P12", vQASINIFile)
                SetQAEntry(vSection, "GBRCapitaliseItem", "L21", vQASINIFile)
                vAddressLines = 9
                If mvQASDPS Then
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "GBRAddressLine" & vAddressLines, "W2,A11", vQASINIFile)
                End If
                If mvQASGRD Then
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "GBRAddressLine" & vAddressLines, "W10,GBRGRD.RawEast", vQASINIFile)
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "GBRAddressLine" & vAddressLines, "W10,GBRGRD.RawNorth", vQASINIFile)
                End If
                If mvQASLEA Then
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "GBRAddressLine" & vAddressLines, "W60,GBRGOV.LEACode", vQASINIFile)
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "GBRAddressLine" & vAddressLines, "W60,GBRGOV.LEAName", vQASINIFile)
                End If
                SetQAEntry(vSection, "GBRDataPlusLines", CStr(vDataPlusLines), vQASINIFile)
                SetQAEntry(vSection, "GBRAddressLineCount", CStr(vAddressLines), vQASINIFile)
              End If

              If InStr(mvCountries, "IRL") >= 0 Then
                vEntry = GetQAEntry(vSection, "IRLAddressLineCount", "", vQASINIFile)
                If Len(vEntry) = 0 Then
                  SetQAEntry(vSection, "IRLComment", "CARE Irish Formatting", vQASINIFile)
                  SetQAEntry(vSection, "IRLConfigured", "220", vQASINIFile)
                  SetQAEntry(vSection, "IRLDataPlusLines", "0", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLineCount", "7", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLine1", "W35", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLine2", "W35", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLine3", "W35", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLine4", "W35", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLine5", "W35,L21", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLine6", "W35,L11", vQASINIFile)
                  SetQAEntry(vSection, "IRLAddressLine7", "W60", vQASINIFile)
                  SetQAEntry(vSection, "IRLCapitaliseItem", "L22 L21 L23 L24", vQASINIFile)
                End If
              End If
              If InStr(mvCountries, "NL") >= 0 Then
                vEntry = GetQAEntry(vSection, "NLDAddressLineCount", "", vQASINIFile)
                If Len(vEntry) = 0 Then
                  SetQAEntry(vSection, "NLDComment", "CARE Dutch Formatting", vQASINIFile)
                  SetQAEntry(vSection, "NLDConfigured", "220", vQASINIFile)
                  SetQAEntry(vSection, "NLDDataPlusLines", "0", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLineCount", "7", vQASINIFile)
                  SetQAEntry(vSection, "NLDCDFVariation", "2", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLine1", "W35", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLine2", "W35", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLine3", "W35", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLine4", "W35", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLine5", "W35,L21", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLine6", "W35,P11", vQASINIFile)
                  SetQAEntry(vSection, "NLDAddressLine7", "W12,C11", vQASINIFile)
                  SetQAEntry(vSection, "NLDCapitaliseItem", "L21 S12 L22 L23 X11", vQASINIFile)
                  SetQAEntry(vSection, "NLDExcludeItem", "P11 C11 L21", vQASINIFile)
                  SetQAEntry(vSection, "NLDDataPlusLines", "0", vQASINIFile)
                End If
              End If
            End If
          Else
            vINIFile = mvQASINIFile
            vEntry = GetQAEntry(vSection, "AddressLines", "")
            If Len(vEntry) = 0 Then
              SetQAEntry(vSection, "AddressLines", "8")
              SetQAEntry(vSection, "DefaultLineWidth", "35")
              SetQAEntry(vSection, "PostcodeFormat", "S")
              SetQAEntry(vSection, "NoCounty", "IfOptional")
              SetQAEntry(vSection, "LinePad", "")
              SetQAEntry(vSection, "FieldEnd", "','")
              SetQAEntry(vSection, "LineEnd", "")
              SetQAEntry(vSection, "Abbreviate", "")
              SetQAEntry(vSection, "Capitalise", "Town")
              SetQAEntry(vSection, "Line1", "W80 ORGA")
              SetQAEntry(vSection, "Line2", "W35")
              SetQAEntry(vSection, "Line3", "W35")
              SetQAEntry(vSection, "Line4", "W35")
              SetQAEntry(vSection, "Line5", "W35")
              SetQAEntry(vSection, "Line6", "TOWN")
              SetQAEntry(vSection, "Line7", "COUN")
              SetQAEntry(vSection, "Line8", "W8 POST")
            End If
          End If
          QAInitialise(1)
          vQASError = QAPro_Open(vINIFile, vSection)
          If vQASError <> 0 Then
            RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError)) 'Failed to Initialise Quick Address ERROR: %1
          Else
            mvQASInitialised = True
          End If
          QAProClose()
          If Not mvQASInitialised Then QAInitialise(0)
        End If
      End If
      If mvQASInitialised = True And pUseBatch = True And mvQASBatchInitialised = False Then
        If mvEnv.GetConfigOption("qas_batch_installed") Then
          vQASError = QABatchWV_Startup(0)
          If vQASError <> 0 Then
            RaiseError(DataAccessErrors.daeQuickAddressError, QASBatchErrorMessage(vQASError))
          Else
            vINIFile = mvEnv.GetConfig("qas_batch_ini_path")
            If Len(vINIFile) > 0 Then
              If Right(vINIFile, 1) <> "\" Then vINIFile = vINIFile & "\"
            End If
            mvQASBatchINIFile = vINIFile & "QAWORLD.INI"
            vEntry = ""
            If My.Computer.FileSystem.FileExists(mvQASBatchINIFile) Then vEntry = GetQAEntry("QADefault", "Language", "", mvQASBatchINIFile)
            If Len(vEntry) = 0 Then
              RaiseError(DataAccessErrors.daeQuickAddressCannotFind, mvQASBatchINIFile) 'Failed to Initialise QuickAddress - Cannot find %1 or the process user may not have the correct permissions for this file
            Else
              vEntry = GetQAEntry(vSection, "AddressLineCount", "", mvQASBatchINIFile)
              If CDbl(vEntry) <> mvQASGBRLines Then
                SetQAEntry(vSection, "CountryBase", "GBR", mvQASBatchINIFile)
                SetQAEntry(vSection, "CleaningAction", "Address", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLineCount", CStr(mvQASGBRLines), mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine1", "W80,O11", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine2", "W35", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine3", "W35", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine4", "W35", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine5", "W35", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine6", "W35,L21", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine7", "W35,L11", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine8", "W8,C11", mvQASBatchINIFile)
                SetQAEntry(vSection, "AddressLine9", "W35,P22,P21,P11,P12", mvQASBatchINIFile)
                vAddressLines = 9
                If mvQASDPS Then
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "AddressLine" & vAddressLines, "W2,A11", mvQASBatchINIFile)
                End If
                If mvQASGRD Then
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "AddressLine" & vAddressLines, "W10,GBRGRD.RawEast", mvQASBatchINIFile)
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "AddressLine" & vAddressLines, "W10,GBRGRD.RawNorth", mvQASBatchINIFile)
                End If
                If mvQASLEA Then
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "AddressLine" & vAddressLines, "W60,GBRGOV.LAName", mvQASBatchINIFile)
                  vAddressLines = vAddressLines + 1
                  SetQAEntry(vSection, "AddressLine" & vAddressLines, "W60,GBRGOV.LEACode", mvQASBatchINIFile)
                End If
                SetQAEntry(vSection, "SeparateElements", "Yes", mvQASBatchINIFile)
                SetQAEntry(vSection, "ElementSeparator", "{, } C11{ ^ } P11{, ^ } P21{, ^ } X11{ ^ } A11{ ^ }", mvQASBatchINIFile)
                SetQAEntry(vSection, "CapitaliseItem", "L21", mvQASBatchINIFile)
              End If
            End If
            mvQASBatchInitialised = True
          End If
        End If
      End If

QASInterfaceEnd:
      InitQAS = mvQASInitialised
      Exit Function

QASInterfaceError:
      If Err.Number = 48 Or Err.Number = 53 Then
        If mvQASInitialised Then
          RaiseError(DataAccessErrors.daeQuickAddressCannotFind, "QABWVED.DLL") 'Failed to Initialise Quick Address - Cannot find QABWVED.DLL
        Else
          RaiseError(DataAccessErrors.daeQuickAddressCannotFind, QADLLName) 'Failed to Initialise Quick Address - Cannot find QAPUIEN.DLL
        End If
      Else
        Err.Raise(Err.Number, Err.Source, Err.Description)
      End If
    End Function

    Private Function QASBatchErrorMessage(ByRef pQASError As Integer) As String
      Dim vString As String
      vString = Space(500)
      QAErrorMessage(pQASError, vString, 500)
      QASBatchErrorMessage = vString
    End Function

    Public Sub QASClose()
      If mvQASBatchInitialised Then
        QABatchClose()
        QABatchWV_Shutdown()
        mvQASBatchInitialised = False
      End If
      If mvQASInitialised Then
        If mvQASOpened Then QAProClose()
        QAInitialise(0)
        mvQASInitialised = False
      End If
    End Sub

    Private Function GetQAEntry(ByVal pSection As String, ByVal pID As String, ByVal pDefault As String, Optional ByRef pServerINI As String = "") As String
      Dim vINIFile As String

      If Len(pServerINI) > 0 Then
        vINIFile = pServerINI
      Else
        vINIFile = QAINIFileName()
      End If
      Dim vINIReader As New INIReader(vINIFile)
      Return vINIReader.ReadString(pSection, pID, pDefault)
    End Function

    Private Sub SetQAEntry(ByVal pSection As String, ByVal pID As String, ByVal pValue As String, Optional ByRef pServerINI As String = "")
      Dim vINIFile As String

      If Len(pServerINI) > 0 Then
        vINIFile = pServerINI
      Else
        vINIFile = QAINIFileName()
      End If
      Dim vINIReader As New INIReader(vINIFile)
      vINIReader.Write(pSection, pID, pValue)
    End Sub

    Private Function QAINIFileName() As String
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          Return "QADDRESS.INI"
        Case QuickAddressTypes.qatNonClientServer
          Return "QAPRO.INI"
        Case QuickAddressTypes.qatProV4orAbove
          Return "QAWORLD.INI"
        Case Else                         'Default to V4 or above
          Return "QAWORLD.INI"
      End Select
    End Function

    Private Function QADLLName() As String
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          Return "QAPUIEN.DLL"
        Case QuickAddressTypes.qatNonClientServer
          Return "QAPUIEB.DLL"
        Case QuickAddressTypes.qatProV4orAbove
          Return "QAUPIED.DLL"
        Case Else
          Return "QAUPIED.DLL"
      End Select
    End Function

    Private Sub QAInitialise(ByVal p1 As Integer)
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          CQAInitialise(p1)
        Case QuickAddressTypes.qatNonClientServer
          NQAInitialise(p1)
        Case QuickAddressTypes.qatProV4orAbove
          'No action
      End Select
    End Sub

    Public Function QASetCountry(ByRef pCountry As String) As Integer
      Dim vCountry As String
      Select Case mvQAType
        Case QuickAddressTypes.qatProV4orAbove
          vCountry = GetBaseCountry(pCountry)
          mvQASDataSet = vCountry
          If vCountry = "UK" Then vCountry = "GBR"
          If vCountry = "NL" Then vCountry = "NLD"
          QASetCountry = QA_SetActiveData(mvQAHandle, vCountry)
      End Select
    End Function

    Public Function QASErrorMessage(ByVal p1 As Integer) As String
      Dim vErrorMsg As String

      vErrorMsg = Space(500)
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          CQAErrorMessage(p1, vErrorMsg, 500)
        Case QuickAddressTypes.qatNonClientServer
          NQAErrorMessage(p1, vErrorMsg, 500)
        Case QuickAddressTypes.qatProV4orAbove
          QA_ErrorMessage(p1, vErrorMsg, 500)
      End Select
      SetStringLen(vErrorMsg)
      QASErrorMessage = "Quick Address ERROR: " & p1 & " " & vErrorMsg
      Debug.Print("QASErrorMessage " & vErrorMsg)
    End Function

    Private Sub SetStringLen(ByRef pString As String)
      Dim vPos As Integer

      vPos = InStr(pString, Chr(0))
      If vPos > 0 Then pString = Trim(Left(pString, vPos - 1))
    End Sub

    Public Sub QABatchOpen()
      Dim vResult As Integer

      If Not mvQASBatchOpened Then
        mvQASBatchINIFile = ""
        vResult = QABatchWV_Open(mvQASBatchINIFile, "CDBDefault", 0, mvQASBatchHandle)
        mvQASBatchOpened = True
        If vResult <> 0 Then RaiseError(DataAccessErrors.daeQuickAddressError, QASBatchErrorMessage(vResult))
      End If
    End Sub

    Public Sub QABatchClose()
      If mvQASBatchOpened Then
        QABatchWV_Close(mvQASBatchHandle)
        mvQASBatchOpened = False
      End If
    End Sub

    Public Function QAProOpen() As Integer
      Dim vResult As Integer

      If Not mvQASOpened Then
        vResult = QAPro_Open(mvQASINIFile, "CDBDefault")
        If vResult = 0 Then
          If mvQAType = QuickAddressTypes.qatProV4orAbove Then vResult = QA_SetEngineOption(mvQAHandle, qaengopt_TIMEOUT, 15000) 'Timeout in 15 seconds
        End If
        If vResult <> 0 Then RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vResult))
      End If
      QAProOpen = vResult
    End Function

    Private Function QAPro_Open(ByRef p1 As String, ByRef p2 As String) As Integer

      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          QAPro_Open = CQAPro_Open(p1, p2)
          If QAPro_Open = 0 Then mvQASOpened = True
        Case QuickAddressTypes.qatNonClientServer
          QAPro_Open = NQAPro_Open(p1, p2)
          If QAPro_Open = 0 Then mvQASOpened = True
        Case QuickAddressTypes.qatProV4orAbove
          QAPro_Open = QA_Open(p1, p2, mvQAHandle)
          If QAPro_Open = 0 Then
            QAPro_Open = QA_SetActiveLayout(mvQAHandle, "CDBDefault")
            mvQASOpened = True
          End If
      End Select
    End Function

    Public Sub QAProClose()
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          CQAPro_Close()
        Case QuickAddressTypes.qatNonClientServer
          NQAPro_Close()
        Case QuickAddressTypes.qatProV4orAbove
          QA_Close(mvQAHandle)
          QA_Shutdown()
      End Select
      mvQASOpened = False
    End Sub

    Public Function QAProCount() As Integer
      Dim vString As String
      Dim vCount As Integer

      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          QAProCount = CQAPro_Count()
        Case QuickAddressTypes.qatNonClientServer
          QAProCount = NQAPro_Count()
        Case QuickAddressTypes.qatProV4orAbove
          vString = Space(80)
          If QA_GetSearchStatusDetail(mvQAHandle, qassint_PICKLISTSIZE, vCount, vString, 80) >= 0 Then QAProCount = vCount
      End Select
    End Function

    Public Sub QAProEndSearch()
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          CQAPro_EndSearch()
        Case QuickAddressTypes.qatNonClientServer
          NQAPro_EndSearch()
        Case QuickAddressTypes.qatProV4orAbove
          QA_EndSearch(mvQAHandle)
      End Select
    End Sub

    Public Function QAProFirst(ByVal p1 As String, ByVal p2 As Integer, ByRef p3 As Integer) As Integer
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          QAProFirst = CQAPro_First(p1, p2, p3)
        Case QuickAddressTypes.qatNonClientServer
          QAProFirst = NQAPro_First(p1, p2, p3)
      End Select
    End Function

    Public Function QAProFormatLine(ByVal p1 As Integer, ByVal p2 As Integer, ByVal p3 As String, ByRef p4 As String, ByVal p5 As Integer) As Integer
      Dim vString As String = ""
      Dim vLineCount As Integer
      Dim vInfo As Integer
      Dim vLabel As String = ""
      Dim vLabelLen As Integer
      Dim vContents As Integer
      Dim vResult As Integer

      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          QAProFormatLine = CQAPro_FormatLine(p1, p2, p3, p4, p5)
        Case QuickAddressTypes.qatNonClientServer
          QAProFormatLine = NQAPro_FormatLine(p1, p2, p3, p4, p5)
        Case QuickAddressTypes.qatProV4orAbove
          If p2 = 0 Then vResult = QA_FormatResult(mvQAHandle, p1, vString, vLineCount, vInfo)
          If vResult < 0 Then
            QAProFormatLine = vResult
          Else
            QAProFormatLine = QA_GetFormattedLine(mvQAHandle, p2, p4, p5, vLabel, vLabelLen, vContents)
          End If
      End Select
    End Function

    Public Function QAProGetItemInfo(ByVal p1 As Integer, ByVal p2 As QAProGetItemInfoTypes, ByRef p3 As QAProGetItemResults, ByRef pConfidence As Integer) As Integer
      Dim vString As String
      Dim vResult As Integer
      Dim vFlags As Integer

      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          vResult = p3
          QAProGetItemInfo = CQAPro_GetItemInfo(p1, p2, vResult)
          p3 = CType(vResult, QAProGetItemResults)
          If p3 = 0 Then
            p3 = QAProGetItemResults.qgirNoStepIn
          Else
            p3 = QAProGetItemResults.qgirStepInRequired
          End If
        Case QuickAddressTypes.qatNonClientServer
          vResult = p3
          QAProGetItemInfo = NQAPro_GetItemInfo(p1, p2, vResult)
          p3 = CType(vResult, QAProGetItemResults)
          If p3 = 0 Then
            p3 = QAProGetItemResults.qgirNoStepIn
          Else
            p3 = QAProGetItemResults.qgirStepInRequired
          End If
        Case QuickAddressTypes.qatProV4orAbove
          'Get Step info
          vString = Space(80)
          vResult = QA_GetResult(mvQAHandle, p1, vString, 80, pConfidence, vFlags)
          If (vFlags And qaresult_CANSTEP) > 0 Then
            p3 = QAProGetItemResults.qgirStepInRequired
          ElseIf (vFlags And qaresult_INFORMATION) > 0 Then
            p3 = QAProGetItemResults.qgirInformation
          ElseIf (vFlags And qaresult_WARN_INFORMATION) > 0 Then
            p3 = QAProGetItemResults.qgirWarning
          Else
            p3 = QAProGetItemResults.qgirNoStepIn
          End If
          'vResult = QA_GetResultDetail(mvQAHandle, p1, qaresultint_ISCANSTEP, p3, vString, 0)
          QAProGetItemInfo = vResult
      End Select
    End Function

    Public Function QAProListItem(ByVal p1 As Integer, ByRef p2 As String, ByVal p3 As Integer, ByVal p4 As Integer) As Integer
      Dim vConfidence As Integer
      Dim vFlags As Integer

      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          QAProListItem = CQAPro_ListItem(p1, p2, p3, p4)
        Case QuickAddressTypes.qatNonClientServer
          QAProListItem = NQAPro_ListItem(p1, p2, p3, p4)
        Case QuickAddressTypes.qatProV4orAbove
          QAProListItem = QA_GetResult(mvQAHandle, p1, p2, p4, vConfidence, vFlags)
          If (vFlags And qaresult_CANSTEP) > 0 Then p2 = "+ " & p2
      End Select
    End Function

    'Public Function QAProGetContext(ByRef pIndex As Integer, ByRef pResult As String, ByRef pResultLen As Integer) As Object
    '  Select Case mvQAType
    '    Case QuickAddressTypes.qatProV4orAbove
    '      QAProGetContext = QA_GetResultDetail(mvQAHandle, pIndex, qaresultstr_PARTIALADDRESS, 0, pResult, pResultLen)
    '  End Select
    'End Function

    Public Function QAProSearch(ByVal p1 As String) As Integer
#If VERBOSE_QAS_MESSAGES Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression VERBOSE_QAS_MESSAGES did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		Debug.Print "QAProSearch " & p1
#End If
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          If Left(p1, 3) = "@P," Then p1 = Mid(p1, 4)
          QAProSearch = CQAPro_Search(p1)
        Case QuickAddressTypes.qatNonClientServer
          If Left(p1, 3) = "@P," Then p1 = Mid(p1, 4)
          QAProSearch = NQAPro_Search(p1)
        Case QuickAddressTypes.qatProV4orAbove
          QAProSearch = QA_Search(mvQAHandle, p1)
      End Select
#If VERBOSE_QAS_MESSAGES Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression VERBOSE_QAS_MESSAGES did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		Debug.Print "QAProSearch  End"
#End If
    End Function

    Public Function QAProStepIn(ByVal p1 As Integer) As Integer
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          QAProStepIn = CQAPro_StepIn(p1)
        Case QuickAddressTypes.qatNonClientServer
          QAProStepIn = NQAPro_StepIn(p1)
        Case QuickAddressTypes.qatProV4orAbove
          QAProStepIn = QA_StepIn(mvQAHandle, p1)
      End Select
    End Function

    Public Function QAProStepOut() As Integer
      Select Case mvQAType
        Case QuickAddressTypes.qatClientServer
          QAProStepOut = CQAPro_StepOut
        Case QuickAddressTypes.qatNonClientServer
          QAProStepOut = NQAPro_StepOut
        Case QuickAddressTypes.qatProV4orAbove
          QAProStepOut = QA_StepOut(mvQAHandle)
      End Select
    End Function

    Public Function QAGetAddress(ByVal pIndex As Integer, ByVal pGetOrgName As Boolean) As Integer
      Dim vIndex As Integer
      Dim vAddrLine As String
      Dim vQASError As Integer
      Dim vLines As Integer
      Dim vOffset As Integer
      Dim vRemoveCounty As Boolean
      Dim vRemoveTown As Boolean
      Dim vLastLine As String = ""
      Dim vLastItem As String
      Dim vBuildingNumberNotCounty As Boolean
      Dim vCheckPostbus As Boolean

      mvAddress = ""
      mvTown = ""
      mvCounty = ""
      mvPostCode = ""
      mvOrgName = ""
      mvBuildingNumber = ""
      ClearAdditionalItems()
      Select Case mvQASDataSet
        Case "NL"
          vLines = 6
          vOffset = 1
          vBuildingNumberNotCounty = True
          vCheckPostbus = True
        Case "IRL"
          vLines = 6
          vOffset = 1
          vRemoveCounty = True
          vRemoveTown = True
        Case Else
          vLines = mvQASGBRLines - 1
          vOffset = 0
      End Select

      For vIndex = 0 To vLines
        vAddrLine = Space(81)
        vQASError = QAProFormatLine(pIndex, vIndex, "", vAddrLine, 50)
        If vQASError < 0 And vQASError <> QAERR_FIELDTRUNCATED Then
          QAGetAddress = vQASError
        Else
          SetStringLen(vAddrLine)
          If Len(vAddrLine) > 0 Or vIndex = 8 Then
            If vQASError = QAERR_FIELDTRUNCATED Then Debug.Print(vAddrLine)
            Select Case vIndex + vOffset
              Case 0
                mvOrgName = vAddrLine
              Case 1
                mvAddress = vAddrLine & Chr(13) & Chr(10)
              Case 2
                mvAddress = mvAddress & vAddrLine & Chr(13) & Chr(10)
              Case 3
                mvAddress = mvAddress & vAddrLine & Chr(13) & Chr(10)
              Case 4
                mvAddress = mvAddress & vAddrLine & Chr(13) & Chr(10)
                vLastLine = vAddrLine
              Case 5
                mvTown = UCase(vAddrLine)
              Case 6
                If vBuildingNumberNotCounty Then
                  mvBuildingNumber = vAddrLine
                Else
                  mvCounty = vAddrLine
                End If
              Case 7
                mvPostCode = UCase(vAddrLine)
              Case 8
                If pGetOrgName And Len(vAddrLine) = 0 Then
                  If Len(mvAddress) > 0 Then
                    mvAddress = mvOrgName & vbCrLf & mvAddress
                  Else
                    mvAddress = mvOrgName
                  End If
                End If
              Case Else
                If vIndex = mvQASDPSIndex Then
                  mvDPS = vAddrLine
                ElseIf vIndex = mvQASGRDEastIndex Then
                  mvEasting = CStr(Val(vAddrLine) * 10)
                ElseIf vIndex = mvQASGRDNorthIndex Then
                  mvNorthing = CStr(Val(vAddrLine) * 10)
                ElseIf vIndex = mvQASLEACodeIndex Then
                  mvLEACode = vAddrLine
                ElseIf vIndex = mvQASLEANameIndex Then
                  mvLEAName = vAddrLine
                End If
            End Select
          End If
        End If
      Next
      If Right(mvAddress, 2) = Chr(13) & Chr(10) Then mvAddress = Left(mvAddress, Len(mvAddress) - 2)
      If pGetOrgName And Len(mvAddress) = 0 Then
        mvAddress = mvOrgName
      End If
      If vRemoveCounty Then
        If Len(mvCounty) > 0 Then
          mvAddress = Replace(mvAddress, Chr(13) & Chr(10) & mvCounty, "")
        End If
      End If
      If vRemoveTown Then
        mvAddress = Replace(mvAddress, Chr(13) & Chr(10) & mvTown, "")
        'Now deal with the situation where the town has been concatenated to the end of the last line
        If InStr(vLastLine, ",") > 0 Then
          For vIndex = Len(vLastLine) - 1 To 1 Step -1
            If Mid(vLastLine, vIndex, 1) = "," Then
              vLastItem = Trim(Mid(vLastLine, vIndex + 1))
              If Left(mvTown, Len(vLastItem)) = vLastItem Then mvAddress = Replace(mvAddress, Mid(vLastLine, vIndex), "")
              Exit For
            End If
          Next
        End If
      End If
      If vCheckPostbus Then
        If Left(mvAddress, 8) = "Postbus " And mvBuildingNumber = "" Then
          mvAddress = FirstLine(mvAddress)
          mvBuildingNumber = Mid(mvAddress, 9)
          mvAddress = "Postbus"
        End If
      End If
    End Function

    Public Function QAGetFormattedAddress(ByVal pPostCode As String) As Boolean
      Dim vQASError As Integer
      Dim vPostcode As String

      If InitQAS(False) Then
        QAProOpen()
        vPostcode = UCase(pPostCode)
        vQASError = QAProSearch("@P," & vPostcode)
        If vQASError < 0 Then
          QAProEndSearch()
          RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
        ElseIf vQASError > 0 Then
          QAProEndSearch()
          RaiseError(DataAccessErrors.daeQuickAddressError, "Invalid Postcode")
        Else
          QAGetFormattedAddress = GetQASAddresses(vPostcode)
        End If
        QAProClose()
      End If
    End Function

    Private Function GetQASAddresses(ByRef pPostCode As String) As Boolean
      Dim vQASCount As Integer
      Dim vQASError As Integer
      Dim vResult As QAProGetItemResults
      Dim vIndex As Integer
      Dim vConfidence As Integer

      vQASCount = QAProCount()
      If vQASCount < 0 Then
        QAProEndSearch()
        RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASCount))
      Else
        For vIndex = 0 To vQASCount - 1 'For each item - see if we can step in
          vQASError = QAProGetItemInfo(vIndex, QAProGetItemInfoTypes.qapStepInfo, vResult, vConfidence)
          If vQASError < 0 Then
            QAProEndSearch()
            RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
          Else
            If vResult <> QAProGetItemResults.qgirNoStepIn Then
              vQASError = QAProStepIn(vIndex) 'Yes we can step in so do it
              If vQASError < 0 Then
                QAProEndSearch()
                RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
              Else
                If GetQASAddresses(pPostCode) Then
                  GetQASAddresses = True
                  Exit For
                End If
                vQASError = QAProStepOut() 'Step out now
                If vQASError < 0 Then
                  QAProEndSearch()
                  RaiseError(DataAccessErrors.daeQuickAddressError, QASErrorMessage(vQASError))
                End If
              End If
            Else 'No we cannot step in so get the record
              If QAGetAddress(vIndex, False) >= 0 Then
                If UCase(Postcode) = pPostCode Then
                  Debug.Print(Address & " " & Town & " " & County & " " & Postcode)
                  GetQASAddresses = True
                  Exit For
                End If
              End If
            End If
          End If
        Next
      End If
    End Function

#End Region

#Region " AFD "

    Public Sub AFDGetAddressesTable(ByRef pDT As CDBDataTable, ByRef pPostCode As String, Optional ByRef pBuilding As String = "")
      Dim vXMLDoc As Xml.XmlDocument
      Dim vRoot As Xml.XmlElement
      Dim vItemNodes As Xml.XmlNodeList
      Dim vIndex As Integer
      Dim vValid As Boolean

      vXMLDoc = CType(AFDGetXMLDoc(False, pPostCode, ""), Xml.XmlDocument)
      vRoot = vXMLDoc.DocumentElement
      vItemNodes = vRoot.SelectNodes("Item")
      If vItemNodes(0) IsNot Nothing Then
        SetAFDAddress(vItemNodes(0))
        For vIndex = 0 To vItemNodes.Count - 1
          SetAFDAddress(vItemNodes(vIndex))
          If Len(pBuilding) > 0 Then
            If Left(mvAddress, Len(pBuilding)) <> pBuilding Then
              vValid = False
            Else
              vValid = True
            End If
          Else
            vValid = True
          End If
          If vValid Then AddAddressToTable(pDT)
        Next
      End If
    End Sub

    Public Function AFDGetAddresses(ByRef pPostCode As String, Optional ByRef pAddress As String = "", Optional ByRef pTown As String = "") As CDBParameters
      Dim vXMLDoc As Xml.XmlDocument
      Dim vRoot As Xml.XmlElement
      Dim vItemNodes As Xml.XmlNodeList
      Dim vListNode As Xml.XmlNode
      Dim vKeyNode As Xml.XmlNode
      Dim vParams As New CDBParameters
      Dim vDataNode As Xml.XmlNode

      vXMLDoc = AFDGetXMLDoc(True, pPostCode, "", pAddress, pTown)
      vRoot = vXMLDoc.DocumentElement
      vItemNodes = vRoot.SelectNodes("Item")

      ' Display any changed postcode if applicable
      '  Set pcFromNode = itemNodes(0).selectSingleNode("PostcodeFrom")
      '  Set dataNode = itemNodes(0).selectSingleNode("Postcode")
      '  If Not (pcFromNode Is Nothing) And Not (dataNode Is Nothing) Then
      '    If pcFromNode.Text <> "" Then
      '      MsgBox "Postcode has changed from " + pcFromNode.Text + " to " + dataNode.Text
      '    End If
      '  End If
      For Each vDataNode In vItemNodes ' Get the data nodes
        vListNode = vDataNode.SelectSingleNode("List")
        vKeyNode = vDataNode.SelectSingleNode("Key")
        If Not (vListNode Is Nothing) And Not (vKeyNode Is Nothing) Then
          If vParams.Exists(vKeyNode.InnerText) = False Then
            vParams.Add(vKeyNode.InnerText, CDBField.FieldTypes.cftCharacter, vListNode.InnerText)
          End If
        End If
      Next vDataNode
      AFDGetAddresses = vParams
    End Function

    Public Function AFDGetPostcode(ByRef pAddress As String, ByRef pTown As String, ByRef pCounty As String, Optional ByVal pDataTable As CDBDataTable = Nothing, Optional ByVal pPostcode As String = "") As Boolean
      Dim vXMLDoc As Xml.XmlDocument
      Dim vRoot As Xml.XmlElement
      Dim vItemNodes As Xml.XmlNodeList
      Dim vListNode As Xml.XmlNode
      Dim vKeyNode As Xml.XmlNode
      Dim vParams As New CDBParameters

      vXMLDoc = AFDGetXMLDoc(True, pPostcode, "", pAddress, pTown)
      vRoot = vXMLDoc.DocumentElement
      vItemNodes = vRoot.SelectNodes("Item")
      If vItemNodes.Count = 1 OrElse (pDataTable IsNot Nothing AndAlso vItemNodes.Count > 0) Then
        Dim vIndex As Integer = 0
        Do
          vListNode = vItemNodes(vIndex).SelectSingleNode("List")
          vKeyNode = vItemNodes(vIndex).SelectSingleNode("Key")
          If Not (vListNode Is Nothing) And Not (vKeyNode Is Nothing) Then
            AFDGetFormattedAddresses(vKeyNode.InnerText)
            If pDataTable IsNot Nothing Then AddAddressToTable(pDataTable)
            AFDGetPostcode = True
          End If
          vIndex += 1
        Loop While pDataTable IsNot Nothing AndAlso vIndex < vItemNodes.Count
      End If
    End Function

    Public Sub AFDGetFormattedAddresses(ByRef pKey As String)
      Dim vXMLDoc As Xml.XmlDocument
      Dim vRoot As Xml.XmlElement
      Dim vItemNodes As Xml.XmlNodeList

      vXMLDoc = AFDGetXMLDoc(False, "", pKey)
      vRoot = vXMLDoc.DocumentElement
      vItemNodes = vRoot.SelectNodes("Item")
      SetAFDAddress(vItemNodes(0))
    End Sub

    Private Sub SetAFDAddress(ByRef pNode As Xml.XmlNode)
      Dim vDataNode As Xml.XmlNode
      Dim vProperty As String = ""
      Dim vStreet As String = ""
      Dim vLocality As String = ""

      vDataNode = pNode.SelectSingleNode("Organisation")
      If Not (vDataNode Is Nothing) Then mvOrgName = vDataNode.InnerText
      vDataNode = pNode.SelectSingleNode("Property")
      If Not (vDataNode Is Nothing) Then vProperty = vDataNode.InnerText
      vDataNode = pNode.SelectSingleNode("Street")
      If Not (vDataNode Is Nothing) Then vStreet = vDataNode.InnerText
      vDataNode = pNode.SelectSingleNode("Locality")
      If Not (vDataNode Is Nothing) Then vLocality = vDataNode.InnerText
      mvAddress = Replace(vProperty, ", ", vbCrLf)
      If Len(mvAddress) > 0 Then
        mvAddress = mvAddress & vbCrLf & vStreet
      Else
        mvAddress = vStreet
      End If
      If vLocality.Length > 0 Then mvAddress = mvAddress & vbCrLf & vLocality
      vDataNode = pNode.SelectSingleNode("Town")
      If Not (vDataNode Is Nothing) Then mvTown = UCase(vDataNode.InnerText)
      vDataNode = pNode.SelectSingleNode("PostalCounty")
      If Not (vDataNode Is Nothing) Then mvCounty = vDataNode.InnerText
      vDataNode = pNode.SelectSingleNode("Postcode")
      If Not (vDataNode Is Nothing) Then mvPostCode = UCase(vDataNode.InnerText)
    End Sub

    Private Function AFDGetXMLDoc(ByRef pList As Boolean, ByRef pPostCode As String, ByRef pKey As String, Optional ByRef pAddress As String = "", Optional ByRef pTown As String = "") As Xml.XmlDocument
      Dim vXMLDoc As Xml.XmlDocument
      Dim vXmlParams As String
      Dim vRoot As Xml.XmlElement
      Dim vDataNode As Xml.XmlNode
      Dim vItemNodes As Xml.XmlNodeList
      Dim vFields As String

      ' Initialise the Microsoft XML Document Object Model
      vXMLDoc = New Xml.XmlDocument
      If pList Then
        vFields = "List"
      Else
        vFields = "Standard"
      End If

      Dim vSerialNumber As String = String.Empty
      Dim vPassword As String = String.Empty
      If NfpConfigrationManager.QAAuthenticationValues IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(NfpConfigrationManager.QAAuthenticationValues.UsernameValue) AndAlso Not String.IsNullOrWhiteSpace(NfpConfigrationManager.QAAuthenticationValues.PasswordValue) Then
        vSerialNumber = NfpConfigrationManager.QAAuthenticationValues.UsernameValue
        vPassword = NfpConfigrationManager.QAAuthenticationValues.PasswordValue
      End If

      ' Build up the XML query string
      vXmlParams = mvEnv.GetConfig("afd_everywhere_server") & "/afddata.pce?"
      vXmlParams = vXmlParams & "Serial=" & vSerialNumber & "&"
      vXmlParams = vXmlParams & "Password=" & vPassword & "&"
      vXmlParams = vXmlParams & "UserID=" & "" & "&"
      If Len(pPostCode) > 0 Then
        vXmlParams = vXmlParams & "Data=Address&Task=FastFind&Fields=" & vFields
        If Not String.IsNullOrWhiteSpace(pAddress) Then
          vXmlParams = vXmlParams & "&Lookup=" & pAddress & ", " & pTown & ", " & pPostCode 'Postcode Validation Tasks
        Else
          vXmlParams = vXmlParams & "&Lookup=" & pPostCode
        End If
        vXmlParams = vXmlParams & "&Lookup=" & pPostCode
      ElseIf Len(pKey) > 0 Then
        vXmlParams = vXmlParams & "Data=Address&Task=Retrieve&Fields=" & vFields
        vXmlParams = vXmlParams & "&Key=" & pKey
      Else
        vXmlParams = vXmlParams & "Data=Address&Task=FastFind&Fields=" & vFields
        vXmlParams = vXmlParams & "&Lookup=" & Replace(pAddress, vbCrLf, ",") & ", " & pTown
      End If
      ' Set the maximum number of records to return
      vXmlParams = vXmlParams & "&MaxQuantity=100"
      ' Load the XML from the webserver with the query string
      Try
        vXMLDoc.Load(vXmlParams)
      Catch vEx As Xml.XmlException
        RaiseError(DataAccessErrors.daeUniservError, "Error: " & vEx.Message)
      End Try
      ' Check if PCE returned an error and if the document is valid
      vRoot = vXMLDoc.DocumentElement
      vDataNode = vRoot.SelectSingleNode("Result")
      vItemNodes = vRoot.SelectNodes("Item")
      If vDataNode Is Nothing Or vItemNodes Is Nothing Then
        RaiseError(DataAccessErrors.daeUniservError, "Invalid PCE XML Document")
      End If
      If Val(vDataNode.InnerText) < 1 Then
        Dim vErrorCode As String = vDataNode.InnerText
        vDataNode = vRoot.SelectSingleNode("ErrorText")
        If vDataNode Is Nothing Then
          RaiseError(DataAccessErrors.daeUniservError, "Invalid PCE XML Document")
        Else
          If vErrorCode <> "-2" Then RaiseError(DataAccessErrors.daeUniservError, "Postcoder Error - " + vDataNode.InnerText) ' Show the user the error
        End If
      End If
      Return vXMLDoc
    End Function

#End Region

  End Class

End Namespace
