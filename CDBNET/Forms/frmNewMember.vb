Public Class frmNewMember

  Private mvAddMember As Boolean
  Private mvPackToDonor As Boolean
  Private mvMemberAdded As Boolean
  Private mvContactNumber As Integer
  Private mvContactInfo As ContactInfo
  Private mvPaymentPlanMemberInfo As PaymentPlanMemberInfo
  Private mvMembersPerOrder As Integer
  Private mvCMDFileName As String

  Public ReadOnly Property PPNumber() As Integer
    Get
      Return mvPaymentPlanMemberInfo.PaymentPlanNumber
    End Get
  End Property

  Public Sub New(ByVal pPaymentPlanMemberInfo As PaymentPlanMemberInfo, ByVal pAddMember As Boolean)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.
    InitialiseControls(pPaymentPlanMemberInfo, pAddMember)
    'AddHandler epl.ValueChanged, AddressOf epl_ValueChanged
    mvCMDFileName = String.Empty
  End Sub

  Private Sub InitialiseControls(ByVal pPaymentPlanMemberInfo As PaymentPlanMemberInfo, ByVal pAddMember As Boolean)
    SetControlTheme()
    Me.MdiParent = MDIForm
    epl.Init(New EditPanelInfo(EditPanelInfo.OtherPanelTypes.optNewMember))
    mvAddMember = pAddMember
    mvPaymentPlanMemberInfo = pPaymentPlanMemberInfo
    Dim vMembershipTypeTextLookupBox As TextLookupBox = epl.FindTextLookupBox("MembershipType")
    epl.SetValue("Source", pPaymentPlanMemberInfo.Source)
    epl.SetValue("Branch", pPaymentPlanMemberInfo.BranchCode)
    'Set up the main membership type
    If mvAddMember Then
      epl.SetValue("MembershipType", mvPaymentPlanMemberInfo.PaymentPlanMembershipType)
    Else    'Replace
      epl.SetValue("MembershipType", pPaymentPlanMemberInfo.MembershipType)
    End If
    mvMembersPerOrder = vMembershipTypeTextLookupBox.GetDataRowInteger("MembersPerOrder")
    If mvMembersPerOrder = 0 Then
      epl.SetValue("Joined", AppValues.TodaysDate)
    Else
      epl.SetValue("Joined", pPaymentPlanMemberInfo.Joined.ToString(AppValues.DateFormat))
    End If
    If vMembershipTypeTextLookupBox.GetDataRowItem("BranchMembership") = "Y" Then epl.SetValue("BranchMember", "Y")
    epl.SetValue("Applied", AppValues.TodaysDate)
    epl.PanelInfo.PanelItems("ContactNumber").Mandatory = True
    epl.PanelInfo.PanelItems("Branch").Mandatory = True

    If Not mvAddMember Then         'Replace Member
      vMembershipTypeTextLookupBox.TextBox.Visible = False
      epl.EnableControl("MembershipType", False)
      mvPackToDonor = BooleanValue(AppValues.ConfigurationValues.me_gift_pack_to_donor_default.ToString())
      epl.SetValue("DistributionCode", pPaymentPlanMemberInfo.DistributionCode)
      If pPaymentPlanMemberInfo.GiftMembership Then
        epl.SetValue("GiftMembership", "Y")
        If pPaymentPlanMemberInfo.OneYearGift Then
          epl.SetValue("OneYearGift", "Y")
        End If
      Else
        epl.SetValue("GiftMembership", "N")
        epl.EnableControlList("OneYearGift,PackToDonor,GiftCardType_D,GiftCardType_R,GiftCardType_S", False)
      End If
      Dim vNewOrdersTable As DataTable
      Dim vPList As New ParameterList(True)
      vPList("PaymentPlanNumber") = pPaymentPlanMemberInfo.PaymentPlanNumber.ToString
      vPList("ExcludeFulfilled") = "Y"
      vNewOrdersTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtNewOrders, vPList)
      If vNewOrdersTable IsNot Nothing Then
        Select Case vNewOrdersTable.Rows(0).Item("GiftCardStatus").ToString
          Case "N"
            epl.SetValue("GiftCardType_D", "D")
          Case "B"
            epl.SetValue("GiftCardType_S", "S")
          Case "W"
            epl.SetValue("GiftCardType_R", "R")
        End Select
        If BooleanValue(vNewOrdersTable.Rows(0).Item("PackToDonor").ToString) Then
          epl.SetValue("PackToDonor", "Y")
        End If
      Else
        epl.SetValue("GiftCardType_D", "D")
      End If
      If pPaymentPlanMemberInfo.ContactNumber = mvPaymentPlanMemberInfo.PaymentPlanContactNumber Then
        'The payer is the member, so since we're replacing the member we must be creating a gift membership
        epl.SetValue("GiftMembership", "Y", True)
      Else
        epl.SetValue("GiftMembership", "Y", False)
      End If
      epl.PanelInfo.PanelItems("CancellationReason").Mandatory = True
      epl.PanelInfo.PanelItems("Branch").Mandatory = True
    Else    'New Member
      'First figure out what types of new member are allowed and filter the membership types accordingly
      Dim vMembershipType1 As String = ""
      Dim vMembershipType2 As String = ""
      Dim vDT As DataTable = DataHelper.GetPaymentPlanData(CareServices.XMLPaymentPlanDataSelectionTypes.xpdtPaymentPlanMemberMenu, mvPaymentPlanMemberInfo.PaymentPlanNumber, mvPaymentPlanMemberInfo.ContactNumber).Tables("DataRow")
      If vDT IsNot Nothing Then
        For Each vDR As DataRow In vDT.Rows
          Select Case vDR.Item("MenuItemOption").ToString
            Case "CanAddNewMainMember"
              If BooleanValue(vDR.Item("MenuItemAvailable").ToString) Then
                vMembershipType1 = mvPaymentPlanMemberInfo.PaymentPlanMembershipType
              End If
            Case "CanAddNewAssociateMember"
              If BooleanValue(vDR.Item("MenuItemAvailable").ToString) Then
                vMembershipType2 = epl.FindTextLookupBox("MembershipType").GetDataRowItem("AssociateMembershipType")
              End If
          End Select
        Next
      End If
      Dim vRestriction As String = "MembershipType IN("
      If vMembershipType1.Length > 0 Then
        vRestriction &= String.Format("'{0}'", vMembershipType1)
        If vMembershipType2.Length > 0 Then vRestriction &= ","
      End If
      If vMembershipType2.Length > 0 Then vRestriction &= String.Format("'{0}'", vMembershipType2)
      vRestriction &= ")"
      epl.FindTextLookupBox("MembershipType").SetFilter(vRestriction)
      'If associates are allowed default to the associate type
      If vMembershipType2.Length > 0 Then epl.SetValue("MembershipType", vMembershipType2)
      If Not vMembershipType1.Length > 0 AndAlso vMembershipType2.Length > 0 Then epl.EnableControl("MembershipType", False)

      epl.SetControlVisible("GiftMembership", False)
      epl.SetControlVisible("OneYearGift", False)
      epl.SetControlVisible("PackToDonor", False)
      epl.SetControlVisible("CancellationReason", False)
      epl.SetControlVisible("CancellationSource", False)
      epl.SetControlVisible("CardType", False)

      epl.PanelInfo.PanelItems("Joined").Mandatory = True
      epl.PanelInfo.PanelItems("Applied").Mandatory = True
      Me.Text = InformationMessages.ImAddMember
    End If
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub epl_ValueChanged(ByVal sender As System.Object, ByVal pParameterName As System.String, ByVal pValue As System.String) Handles epl.ValueChanged
    Try
      Select Case pParameterName
        Case "GiftMembership"
          Dim vGiftMembership As Boolean = BooleanValue(epl.GetValue("GiftMembership"))
          If vGiftMembership Then
            epl.SetValue("GiftCardType_S", "S")
            'The Pack To Donor checkbox may already be checked.
            'It could be that the Gift Membership checkbox has become checked BECAUSE Pack To Donor was checked.
            If epl.GetValue("PackToDonor") = "Y" And mvPackToDonor Then epl.SetValue("PackToDonor", "Y")
          Else
            epl.SetValue("GiftCardType_D", "D")
            epl.SetValue("OneYearGift", "N")
            epl.SetValue("PackToDonor", "N")
          End If
          epl.EnableControlList("OneYearGift,PackToDonor,GiftCardType_D,GiftCardType_S,GiftCardType_R", vGiftMembership)
        Case "ContactNumber"
          epl.SetValue("AddressLine", "")
          If epl.GetValue("ContactNumber").Length > 0 Then
            Dim vParams As New ParameterList(True)
            mvContactNumber = IntegerValue(epl.GetValue("ContactNumber"))
            Dim vAddressDataTable As DataTable = DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactAddresses, mvContactNumber)
            If Not vAddressDataTable Is Nothing Then
              Dim vDataRows() As DataRow = vAddressDataTable.Select("Default= 'Yes'")
              If vDataRows.Length > 0 Then
                epl.SetValue("AddressLine", vDataRows(0).Item("AddressLine").ToString)
                epl.SetValue("Branch", vDataRows(0).Item("Branch").ToString)
              End If
              mvContactInfo = New ContactInfo(mvContactNumber)
            End If
            If Not mvAddMember Then
              If mvContactNumber = mvPaymentPlanMemberInfo.PaymentPlanContactNumber Then
                epl.SetValue("GiftMembership", "N", True)
              Else
                epl.SetValue("GiftMembership", "Y", True)
              End If
            End If
          End If
        Case "AgeOverride"
          ValidateAgeOverride()
        Case "OneYearGift"
          If epl.GetValue("OneYearGift") = "Y" Then epl.SetValue("GiftMembership", "Y")
        Case "PackToDonor"
          If epl.GetValue("PackToDonor") = "Y" Then epl.SetValue("GiftMembership", "Y")
      End Select
    Catch vException As CareException
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      Dim vPPMemberList As New ParameterList(True)
      If epl.AddValuesToList(vPPMemberList, True) Then
        vPPMemberList.IntegerValue("OrderNumber") = mvPaymentPlanMemberInfo.PaymentPlanNumber
        vPPMemberList("BranchMember") = epl.GetValue("BranchMember")
        UpdateContact()
        If ValidateAgeOverride() Then
          If ShowQuestion(QuestionMessages.QmConfirmInsert, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
            If Not mvAddMember Then
              'Work out the gift card status selected
              If epl.GetValue("GiftCardType_D") = "D" Then
                vPPMemberList("GiftCardType") = "D"
              ElseIf epl.GetValue("GiftCardType_S") = "S" Then
                vPPMemberList("GiftCardType") = "S"
              Else
                vPPMemberList("GiftCardType") = "R"
              End If
            Else
              vPPMemberList("GiftCardType") = "D"
            End If
            Dim vReturnDS As DataSet = DataHelper.AddPaymentPlanMember(vPPMemberList)
            If vReturnDS IsNot Nothing Then
              Dim vRow As DataRow = DataHelper.GetRowFromDataSet(vReturnDS)
              If vRow IsNot Nothing Then
                Dim vWarningMessage As String = ""
                If vRow.Table.Columns.Contains("WarningMessage") Then
                  vWarningMessage = vRow("WarningMessage").ToString
                End If
                If vWarningMessage.Length > 0 Then ShowWarningMessage(vWarningMessage)
              End If
            End If
            If Not mvAddMember Then
              'Create Contact Mailing Document
              CreateCMD()
            End If
            Me.Close()
          End If
        End If
      End If
    Catch vEx As CareException
      If vEx.ErrorNumber = CareException.ErrorNumbers.enAlreadyMember Then
        ShowInformationMessage(vEx.Message)
      ElseIf vEx.ErrorNumber = CareException.ErrorNumbers.enPaymentScheduleNotCreatedRDInPast Then
        ShowInformationMessage(vEx.Message)
      ElseIf vEx.ErrorNumber = CareException.ErrorNumbers.enCannotAddMember Then
        ShowInformationMessage(vEx.Message)
      Else
        DataHelper.HandleException(vEx)
      End If
    End Try
  End Sub

  Private Sub CMDActionComplete(ByVal pAction As ExternalApplication.DocumentActions, ByVal pFileName As String)
    mvCMDFileName = pFileName
  End Sub

  Private Function ValidateAgeOverride() As Boolean
    Dim vMaxJuniorAge As Integer = epl.FindTextLookupBox("MembershipType").GetDataRowInteger("MaxJuniorAge")
    Dim vValidateAgeOverride As Boolean = True
    Dim vAgeOverride As Integer = IntegerValue(epl.GetValue("AgeOverride"))
    If vAgeOverride > 0 Then
      If vAgeOverride > vMaxJuniorAge Then
        ShowInformationMessage(InformationMessages.ImAgeOverrideInvalid, vMaxJuniorAge.ToString())
        epl.FindTextBox("AgeOverride").Focus()
        vValidateAgeOverride = False
      End If
    End If
    Return vValidateAgeOverride
  End Function

  Private Sub CreateCMD()
    Dim vSourceDataRow As DataRow
    Dim vSourceParams As New ParameterList(True)
    vSourceParams("Source") = epl.GetValue("Source")
    vSourceDataRow = DataHelper.GetLookupDataRow(CareNetServices.XMLLookupDataTypes.xldtSources, vSourceParams)
    If Not vSourceDataRow Is Nothing Then
      Dim vCMDList As New ParameterList(True)
      Dim vShowParagraphs As DialogResult = System.Windows.Forms.DialogResult.Yes
      vCMDList("Mailing") = vSourceDataRow.Item("ThankYouLetter").ToString()
      Dim vCheckMailingDocResult As ParameterList = DataHelper.CheckMailingDocuments(vCMDList)
      If vCheckMailingDocResult.ContainsKey("ContactWarningSuppressionsPrompt") AndAlso vCheckMailingDocResult.ContainsKey("WarningSuppressions") AndAlso vCheckMailingDocResult("ContactWarningSuppressionsPrompt") = "Y" AndAlso vCheckMailingDocResult("WarningSuppressions").Length > 0 Then
        vShowParagraphs = ShowQuestion(QuestionMessages.QmWarningSuppressions, MessageBoxButtons.YesNo, vCheckMailingDocResult("WarningSuppressions"))
      End If
      If vShowParagraphs = System.Windows.Forms.DialogResult.Yes Then
        If Not mvAddMember Then
          Dim vMailingCountList As New ParameterList(True)
          vMailingCountList("PaymentPlanNumber") = mvPaymentPlanMemberInfo.PaymentPlanNumber.ToString
          Dim vContactMailingData As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtContactMailingDocuments, vMailingCountList)
          If vContactMailingData IsNot Nothing Then
            If vContactMailingData.Rows.Count > 0 Then
              If ShowQuestion(QuestionMessages.QmUnfulfilledMailingDoc, MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.Yes Then
                DataHelper.DeleteContactMailingDocument(IntegerValue(vContactMailingData.Rows(0).Item("MailingDocumentNumber").ToString))
              End If
            End If
          End If
        End If

        'Retrieve the matching paragraphs
        vCMDList.Item("Mailing") = vSourceDataRow.Item("ThankYouLetter").ToString
        vCMDList.IntegerValue("ContactNumber") = IntegerValue(epl.GetValue("ContactNumber"))
        vCMDList.Item("ExistingTransaction") = "N"
        vCMDList.Item("NewPayerContact") = "Y"
        vCMDList.IntegerValue("PaymentPlanNumber") = mvPaymentPlanMemberInfo.PaymentPlanNumber
        vCMDList.Item("PaymentPlanCreated") = "Y"
        Dim vCount As Integer
        Dim vCMDDataSet As DataSet = DataHelper.GetMailingDocumentParagraphs(vCMDList)
        'Display the matching paragraphs
        Dim vParagraphsTable As DataTable = DataHelper.GetTableFromDataSet(vCMDDataSet)
        If vParagraphsTable IsNot Nothing AndAlso _
        vParagraphsTable.Columns.Contains("DisplayParagraphs") AndAlso vParagraphsTable.Rows(0).Item("DisplayParagraphs").ToString = "Y" Then
          DocumentApplication = New WordApplication
          AddHandler DocumentApplication.ActionComplete, AddressOf CMDActionComplete
          Dim vTransDocumentType As frmTransactionDocument.TransactionDocumentTypes = frmTransactionDocument.TransactionDocumentTypes.tdtTransaction
          If vCount > 0 Then vTransDocumentType = frmTransactionDocument.TransactionDocumentTypes.tdtPaymentPlan
          Dim vForm As frmTransactionDocument = New frmTransactionDocument(vTransDocumentType, vCMDDataSet, vCMDList)
          vForm.ShowDialog()
          'need to edit the document when Edit is pressed
          vCMDDataSet = vForm.DataSet
        End If
        'Create the mailing document
        If vCMDList.Contains("EarliestFulfilmentDate") Then
          If mvCMDFileName.Length = 0 Then
            Dim vSelectedParagraphs As New StringBuilder
            Dim vCMDTable As DataTable = DataHelper.GetTableFromDataSet(vCMDDataSet)
            If vCMDTable IsNot Nothing Then
              For Each vCMDRow As DataRow In vCMDTable.Rows
                If BooleanValue(vCMDRow.Item("Include").ToString) Then
                  If vSelectedParagraphs.Length > 0 Then vSelectedParagraphs.Append(",")
                  vSelectedParagraphs.Append(vCMDRow.Item("ParagraphNumber"))
                End If
              Next
            End If
            If vSelectedParagraphs.Length = 0 Then
              vCMDList("SelectedParagraphs") = "0"
            Else
              vCMDList("SelectedParagraphs") = vSelectedParagraphs.ToString
            End If
          End If
          vCMDDataSet = DataHelper.AddContactMailingDocument(vCMDList)
          If mvCMDFileName.Length > 0 Then
            Dim vResultRow As DataRow = vCMDDataSet.Tables("Result").Rows(0)
            DataHelper.UpdateContactMailingDocumentFile(IntegerValue(vResultRow.Item("MailingDocumentNumber").ToString), mvCMDFileName)
          End If
        End If
      End If
    End If
  End Sub

  Private Sub UpdateContact()
    Dim vParams As New ParameterList(True)
    Dim mvContactInfo As New ContactInfo(mvContactNumber)
    Dim vMembershipLevel As String = epl.FindTextLookupBox("MembershipType").GetDataRowItem("MembershipLevel")
    If mvContactInfo.DateOfBirth.Length = 0 AndAlso vMembershipLevel = "J" Then
      'Junior contact requires a date of birth
      vParams("LabelName") = mvContactInfo.ContactName
      vParams("DateOfBirth") = mvContactInfo.DateOfBirth
      vParams("DobEstimated") = CBoolYN(mvContactInfo.DOBEstimated)
      Dim vAPForm As New frmApplicationParameters(CareServices.FunctionParameterTypes.fptUpdateContact, vParams, Nothing)
      If vAPForm.ShowDialog = System.Windows.Forms.DialogResult.OK Then
        Dim vList As ParameterList = vAPForm.ReturnList
        vList("DOBEstimated") = vList("DobEstimated")
        vList.Remove("DobEstimated")
        vList.IntegerValue("ContactNumber") = mvContactInfo.ContactNumber
        DataHelper.UpdateItem(CareServices.XMLMaintenanceControlTypes.xmctContact, vList)
      End If
    End If
  End Sub

End Class