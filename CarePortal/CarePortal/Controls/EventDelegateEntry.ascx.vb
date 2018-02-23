Public Class EventDelegateEntry
    Inherits CareWebControl

    Private mvBookingNumber As Integer
    Private mvDelegateCount As Integer
    Private mvDelegates As New Dictionary(Of String, EventDelegate)
    Private mvFields() As String = {"SequenceNumber", "Title", "Forenames", "Surname", "Position", "OrganisationName", "EMailAddress"}
    Private mvFreeSeatNumbers As New List(Of Integer)

    'Adding of position at registered users organisation

    Public Sub New()
        mvNeedsAuthentication = True
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            InitialiseControls(CareNetServices.WebControlTypes.wctEventDelegateEntry, tblDataEntry)
            SelectEventDelegates()
            SetControlVisible("PageError", False)
            SetControlVisible("WarningMessage1", False)
            SetControlVisible("WarningMessage2", False)
            SetControlVisible("WarningMessage3", False)
            SetControlVisible("WarningMessage4", False)
        Catch vEx As ThreadAbortException
            Throw vEx
        Catch vException As Exception
            ProcessError(vException)
        End Try
    End Sub

    Private Sub RebuildDataEntryTable(pRowCount As Integer)
        Dim vLabels As New List(Of HtmlTableCell)
        Dim vItems As New List(Of Control)
        Dim vRows As New List(Of HtmlTableRow)
        Dim vItemCount As Integer

        For Each vExistingRow As HtmlTableRow In tblDataEntry.Rows
            If vExistingRow.Cells(0).Attributes("Class") = "DataEntryLabel" Then
                vLabels.Add(vExistingRow.Cells(0))
                vItems.Add(vExistingRow.Cells(1).Controls(0))
                vRows.Add(vExistingRow)
                vItemCount += 1
            Else
                If vExistingRow.Cells(0).Attributes("Class") = "DataMessage" Then
                    vExistingRow.Cells(0).ColSpan = vItemCount
                Else
                    vExistingRow.Cells(0).ColSpan = 1     'Need to set this to 1 to stop ASP.NET from setting it to the same value as above!!! Is this a bug...
                End If
            End If
        Next
        For Each vDeleteRow As HtmlTableRow In vRows
            tblDataEntry.Rows.Remove(vDeleteRow)
        Next
        Dim vRow As New HtmlTableRow
        For Each vLabel As HtmlTableCell In vLabels
            vLabel.ColSpan = 1
            vRow.Cells.Add(vLabel)
        Next
        tblDataEntry.Rows.Insert(0, vRow)
        For vIndex As Integer = 0 To pRowCount - 1
            vRow = New HtmlTableRow
            For Each vItem As Control In vItems
                Dim vCell As New HtmlTableCell

                If TryCast(vItem, TextBox) IsNot Nothing Then
                    Dim vItemTextBox As TextBox = DirectCast(vItem, TextBox)
                    If vIndex = 0 Then
                        vCell.Controls.Add(vItemTextBox)
                        If vItemTextBox.ID = "SequenceNumber" Then vItemTextBox.Text = CStr(vIndex + 1)
                    Else
                        Dim vTextBox As New TextBox
                        vTextBox.ID = vItemTextBox.ID & vIndex.ToString
                        vTextBox.CssClass = vItemTextBox.CssClass
                        vTextBox.MaxLength = vItemTextBox.MaxLength
                        vTextBox.Width = vItemTextBox.Width
                        If vItemTextBox.ID = "SequenceNumber" Then vTextBox.Text = CStr(vIndex + 1)
                        vCell.Controls.Add(vTextBox)
                    End If
                ElseIf TryCast(vItem, DropDownList) IsNot Nothing Then
                    Dim vItemDDL As DropDownList = DirectCast(vItem, DropDownList)
                    If vIndex = 0 Then
                        vCell.Controls.Add(vItem)
                    Else
                        Dim VDDL As New DropDownList
                        VDDL.ID = vItemDDL.ID & vIndex.ToString
                        VDDL.CssClass = vItemDDL.CssClass
                        VDDL.Width = vItemDDL.Width
                        VDDL.DataTextField = vItemDDL.DataTextField
                        VDDL.DataValueField = vItemDDL.DataValueField
                        Dim vDataTable As DataTable = TryCast(vItemDDL.DataSource, DataTable)
                        If vDataTable IsNot Nothing Then
                            Dim vNewTable As DataTable = vDataTable.Clone
                            For Each vOldRow As DataRow In vDataTable.Rows
                                Dim vNewRow As DataRow = vNewTable.NewRow
                                vNewRow.ItemArray = vOldRow.ItemArray
                                vNewTable.Rows.Add(vNewRow)
                            Next
                            VDDL.DataSource = vNewTable
                            VDDL.DataBind()
                        End If
                        vCell.Controls.Add(VDDL)
                    End If
                End If
                vRow.Cells.Add(vCell)
            Next
            tblDataEntry.Rows.Insert(vIndex + 1, vRow)
        Next
    End Sub

    Private Sub SetRequiredField(ByVal pHTMLCell As HtmlTableCell, ByVal pID As String)
        Dim vRFV As New RequiredFieldValidator
        With vRFV
            .CssClass = "DataValidator"
            .ID = "rfv" & pID
            .ControlToValidate = pID
            .Display = ValidatorDisplay.Dynamic
            .ErrorMessage = "*"
            .SetFocusOnError = True
        End With
        pHTMLCell.Controls.Add(vRFV)
    End Sub


    Private Sub SelectEventDelegates()
        Dim vList As New ParameterList(HttpContext.Current)
        If InitialParameters.ContainsKey("BookingNumber") Then
            vList("BookingNumber") = InitialParameters("BookingNumber")
            mvBookingNumber = IntegerValue(vList("BookingNumber").ToString)
        ElseIf Request.QueryString("BN") IsNot Nothing Then
            vList("BookingNumber") = Request.QueryString("BN")
            mvBookingNumber = IntegerValue(vList("BookingNumber").ToString)
        Else
            RebuildDataEntryTable(1)
            Exit Sub
        End If

        Dim vRow As DataRow = DataHelper.GetRowFromDataTable(DataHelper.GetDataTable(DataHelper.SelectEventData(CareNetServices.XMLEventDataSelectionTypes.xedtEventBookings, vList)))
        If vRow IsNot Nothing Then
            mvDelegateCount = IntegerValue(vRow.Item("Quantity").ToString)
            RebuildDataEntryTable(mvDelegateCount)
            Dim vTable As DataTable = DataHelper.GetDataTable(DataHelper.SelectEventData(CareNetServices.XMLEventDataSelectionTypes.xedtEventBookingDelegates, vList))
            If vTable IsNot Nothing Then

                'Determine list of Seat Numbers not used
                For vSeatNumber As Integer = 1 To mvDelegateCount
                    mvFreeSeatNumbers.Add(vSeatNumber)
                Next
                For Each vDelegateRow As DataRow In vTable.Rows
                    Dim vSequenceNumber As String
                    For Each vField As String In mvFields
                        If vField.StartsWith("SequenceNumber") Then
                            vSequenceNumber = vDelegateRow.Item(vField).ToString
                            If vSequenceNumber.Length > 0 Then
                                mvFreeSeatNumbers.Remove(IntegerValue(vSequenceNumber))
                            End If
                            Exit For
                        End If
                    Next
                Next

                Dim vIndex As Integer = 1
                For Each vDelegateRow As DataRow In vTable.Rows
                    'Find the row, i.e. sequence (seating) number, to display the delegate
                    Dim vSequenceNumber As String = (mvDelegateCount + 1).ToString 'default to a sequence number not displayed
                    For Each vField As String In mvFields
                        If vField.StartsWith("SequenceNumber") Then
                            vSequenceNumber = vDelegateRow.Item(vField).ToString
                            If vSequenceNumber.Length = 0 Then  'This delegate have no seat number so assign free seat number to this delegate 
                                vSequenceNumber = mvFreeSeatNumbers(0).ToString
                                mvFreeSeatNumbers.Remove(IntegerValue(vSequenceNumber))
                            Else
                                For Each vDelegate As EventDelegate In mvDelegates.Values
                                    If vDelegate.SequenceNumber.Equals(vSequenceNumber) Then
                                        'There is another delegate with the same seat/sequence number. Assign next free seat number to delegate
                                        vSequenceNumber = mvFreeSeatNumbers(0).ToString
                                        mvFreeSeatNumbers.Remove(IntegerValue(vSequenceNumber))
                                    End If
                                Next
                            End If
                            Exit For
                        End If
                    Next
                    Dim vDisplayRowIndex As Integer = IntegerValue(vSequenceNumber) - 1


                    If vIndex <= mvDelegateCount Then
                        Dim vEMailAddress As String = ""
                        Dim vContactNumber As Integer = IntegerValue(vDelegateRow.Item("ContactNumber").ToString)
                        Dim vEmailList As New ParameterList(HttpContext.Current)
                        vEmailList("ContactNumber") = vContactNumber
                        Dim vEmailTable As DataTable = DataHelper.GetDataTable(DataHelper.SelectContactData(CareNetServices.XMLContactDataSelectionTypes.xcdtContactEMailAddresses, vEmailList))
                        If vEmailTable IsNot Nothing AndAlso vEmailTable.Rows.Count > 0 Then
                            vEMailAddress = vEmailTable.Rows(0).Item("EMailAddress").ToString
                        End If
                        Dim vDelegate As New EventDelegate(vContactNumber.ToString, vDelegateRow.Item("EventDelegateNumber").ToString, vDelegateRow.Item("AddressNumber").ToString, vDelegateRow.Item("SequenceNumber").ToString, vDelegateRow.Item("Title").ToString, vDelegateRow.Item("Forenames").ToString, vDelegateRow.Item("Surname").ToString, vEMailAddress, vDelegateRow.Item("Position").ToString, vDelegateRow.Item("OrganisationName").ToString)
                        If String.IsNullOrEmpty(vDelegate.SequenceNumber) Then
                            vDelegate.SequenceNumber = vSequenceNumber
                        End If
                        mvDelegates.Add(vIndex.ToString, vDelegate)
                        For Each vField As String In mvFields
                            Dim vControlID As String = vField
                            If vDisplayRowIndex > 0 Then vControlID &= vDisplayRowIndex.ToString
                            If vField = "EMailAddress" Then
                                SetTextBoxText(vControlID, vEMailAddress)
                            ElseIf vField.StartsWith("Title") Then
                                SetDropDownText(vControlID, vDelegateRow.Item(vField).ToString)
                            ElseIf vField.StartsWith("SequenceNumber") Then              'The first delegate may not have a number - if not set it to 1
                                SetTextBoxText(vControlID, vSequenceNumber)
                            Else
                                SetTextBoxText(vControlID, vDelegateRow.Item(vField).ToString)
                            End If
                        Next
                    End If
                    vIndex += 1
                Next
            End If
        Else
            RebuildDataEntryTable(1)
        End If
    End Sub

    Protected Overrides Sub ButtonClickHandler(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If IsValid() Then
                RebuildDelegates()
                'Get the information for all the delegates
                For vIndex As Integer = 0 To mvDelegateCount - 1
                    Dim vDelegate As EventDelegate = mvDelegates((vIndex + 1).ToString)

                    Dim vControlIndex As Integer = vIndex
                    If Not String.IsNullOrEmpty(vDelegate.SequenceNumber) Then
                        vControlIndex = IntegerValue(vDelegate.SequenceNumber) - 1
                    End If

                    For Each vField As String In mvFields
                        Dim vControlID As String = vField
                        If vControlIndex > 0 Then
                            vControlID &= vControlIndex.ToString
                        End If
                        Dim vValue As String
                        If vField.StartsWith("Title") Then
                            vValue = GetDropDownValue(vControlID)
                            If vValue.Length > 0 Then vDelegate.AnyFieldsSet = True
                            Dim vDDL As DropDownList = TryCast(FindControlByName(tblDataEntry, vControlID), DropDownList)
                            If vDDL IsNot Nothing Then
                                If vDDL.CssClass = "DataEntryItemMandatory" AndAlso vValue.Length = 0 Then vDelegate.MandatoryFieldsNotSet = True
                            End If
                            vDelegate.Title = vValue
                        Else
                            vValue = GetTextBoxText(vControlID)
                            Select Case vField
                                Case "SequenceNumber"
                                    If Not String.IsNullOrEmpty(vValue) Then
                                        vDelegate.SequenceNumber = vValue
                                    End If
                                    If String.IsNullOrEmpty(vDelegate.SequenceNumber) AndAlso String.IsNullOrEmpty(vValue) Then
                                        'If Delegate Sequence Number not set and value not set or control not existing then assign Sequence Number from free seat numbers
                                        vDelegate.SequenceNumber = mvFreeSeatNumbers(0).ToString
                                        mvFreeSeatNumbers.RemoveAt(0)
                                    End If
                                Case Else
                                    If vValue.Length > 0 Then vDelegate.AnyFieldsSet = True
                                    Select Case vField
                                        Case "Forenames"
                                            vDelegate.Forenames = vValue
                                        Case "Position"
                                            vDelegate.Position = vValue
                                        Case "OrganisationName"
                                            vDelegate.OrganisationName = vValue
                                        Case "Surname"
                                            vDelegate.Surname = vValue
                                        Case "EMailAddress"
                                            Try
                                                vDelegate.Email = vValue
                                            Catch vEx As CareException
                                                If vEx.ErrorNumber = CareException.ErrorNumbers.enInvalidEmailAddress Then
                                                    SetControlVisible("PageError", True)
                                                    SetLabelText("PageError", String.Format("Row {0}: {1}", vIndex + 1, vEx.Message))
                                                    Exit Sub
                                                Else
                                                    ProcessError(vEx)
                                                End If
                                            End Try
                                    End Select
                            End Select
                            Dim vTextBox As TextBox = TryCast(FindControlByName(tblDataEntry, vControlID), TextBox)
                            If vTextBox IsNot Nothing Then
                                If vTextBox.CssClass = "DataEntryItemMandatory" AndAlso vValue.Length = 0 Then vDelegate.MandatoryFieldsNotSet = True
                            End If
                        End If
                    Next
                Next

                'Check if we have all the information
                For Each vDelegate As EventDelegate In mvDelegates.Values
                    If vDelegate.AnyFieldsSet OrElse vDelegate.Existing Then
                        If vDelegate.MandatoryFieldsNotSet Then
                            SetControlVisible("WarningMessage1", True)
                            Exit Sub
                        End If
                    ElseIf InitialParameters("AllDelegatesMandatory").ToString = "Y" Then
                        SetControlVisible("WarningMessage2", True)
                        Exit Sub
                    End If
                Next

                'Check for duplicate sequence numbers or duplicate contact details
                For Each vDelegate As EventDelegate In mvDelegates.Values
                    If vDelegate.AnyFieldsSet AndAlso vDelegate.SequenceNumber.Length > 0 Then
                        For Each vCheckDelegate As EventDelegate In mvDelegates.Values
                            If vCheckDelegate IsNot vDelegate AndAlso vCheckDelegate.AnyFieldsSet Then
                                If IntegerValue(vDelegate.SequenceNumber) = IntegerValue(vCheckDelegate.SequenceNumber) Then
                                    SetControlVisible("WarningMessage3", True)
                                    Exit Sub
                                End If
                                If vDelegate.Surname = vCheckDelegate.Surname AndAlso vDelegate.Email = vCheckDelegate.Email Then
                                    SetControlVisible("WarningMessage4", True)
                                    Exit Sub
                                End If
                            End If
                        Next
                    End If
                Next

                Dim vOrganisationNumber As Integer
                Dim vAddressNumber As Integer
                If InitialParameters("AddDelegatePosition").ToString = "Y" Then
                    Dim vList As New ParameterList(HttpContext.Current)
                    vList("ContactNumber") = UserContactNumber()
                    vList("AddressNumber") = UserAddressNumber()
                    vList("Current") = "Y"
                    Dim vDT As DataTable = DataHelper.GetContactDataTable(CareNetServices.XMLContactDataSelectionTypes.xcdtContactPositions, vList)
                    If vDT IsNot Nothing AndAlso vDT.Rows.Count > 0 Then
                        vOrganisationNumber = IntegerValue(vDT.Rows(0).Item("ContactNumber").ToString)
                        vAddressNumber = IntegerValue(vDT.Rows(0).Item("AddressNumber").ToString)
                    End If
                End If

                'Collate contact details for displayed row
                Dim vSequenceLists As New Dictionary(Of String, ParameterList) 'holds parameter list for each sequence/seat number
                Dim vRowIndex As Integer = 1
                For Each vDelegate As EventDelegate In mvDelegates.Values
                    If vDelegate.Existing Then
                        If vDelegate.UpdateRequired AndAlso vDelegate.NewDelegateRequired = False Then      'Position or OrganisationName changed
                            Dim vList As New ParameterList(HttpContext.Current)
                            vList("EventDelegateNumber") = vDelegate.EventDelegateNumber
                            vList("Position") = vDelegate.Position
                            vList("OrganisationName") = vDelegate.OrganisationName
                            If vDelegate.SequenceNumber.Length > 0 Then vList("SequenceNumber") = vDelegate.SequenceNumber
                            DataHelper.UpdateEventDelegate(vList)
                            vSequenceLists.Add(vRowIndex.ToString, vList)
                        End If
                        If vDelegate.NewDelegateRequired Then                                           'Title, Forenames, Surname or EmailAddress changed
                            Dim vList As ParameterList = GetContactDetails(vDelegate, vOrganisationNumber, vAddressNumber)
                            vList("EventDelegateNumber") = vDelegate.EventDelegateNumber
                            vList("Position") = vDelegate.Position
                            vList("OrganisationName") = vDelegate.OrganisationName
                            If vDelegate.SequenceNumber.Length > 0 Then vList("SequenceNumber") = vDelegate.SequenceNumber
                            vSequenceLists.Add(vRowIndex.ToString, vList)
                        End If
                    Else
                        If vDelegate.Surname.Length > 0 Then
                            Dim vList As ParameterList = GetContactDetails(vDelegate, vOrganisationNumber, vAddressNumber)
                            If vDelegate.Position.Length > 0 Then vList("Position") = vDelegate.Position
                            If vDelegate.OrganisationName.Length > 0 Then vList("OrganisationName") = vDelegate.OrganisationName
                            If vDelegate.SequenceNumber.Length > 0 Then vList("SequenceNumber") = vDelegate.SequenceNumber
                            vSequenceLists.Add(vRowIndex.ToString, vList)
                        End If
                    End If
                    vRowIndex += 1
                Next

                AssignDelegateRecordToDisplayRow(vSequenceLists)

                vRowIndex = 1
                'Call web services to persist data
                For Each vDelegate As EventDelegate In mvDelegates.Values
                    If vDelegate.Existing _
            And Not vDelegate.NewDelegateRequired _
            And vDelegate.EmailAdded Then 'not Title, Forenames, Surname or EmailAddress changes
                        Dim vList As New ParameterList(HttpContext.Current)
                        vList("ContactNumber") = vDelegate.ContactNumber
                        vList("Device") = DataHelper.ControlValue(DataHelper.ControlValues.email_device)
                        vList("Number") = vDelegate.Email
                        DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctNumber, vList)
                    Else
                        If vSequenceLists.ContainsKey(vRowIndex.ToString) Then
                            If vSequenceLists(vRowIndex.ToString).Contains("EventDelegateNumber") Then
                                DataHelper.UpdateEventDelegate(vSequenceLists(vRowIndex.ToString))
                            Else
                                DataHelper.AddEventDelegate(vSequenceLists(vRowIndex.ToString))
                            End If
                            vSequenceLists.Remove(vRowIndex.ToString)  'This prevents it from being submitted again (i.e. when there are more than one display row with same seat number)
                        End If
                    End If
                    vRowIndex += 1
                Next

                GoToSubmitPage("&BN=" & mvBookingNumber)
            End If
        Catch vEx As ThreadAbortException
            Throw vEx
        Catch vEx As CareException
            ProcessError(vEx)
        Catch vException As Exception
            ProcessError(vException)
        End Try
    End Sub
    ''' <summary>
    ''' Rebuilds delegates dictionary ordered by Sequence Number including newly initialised delegate for empty rows
    ''' </summary>
    Private Sub RebuildDelegates()
        Dim vDelegatesCopy As New Dictionary(Of String, EventDelegate)
        For vIndex As Integer = 1 To mvDelegateCount
            Dim vDelegateSequenceFound As Boolean = False
            Dim vDelegateItem As New EventDelegate
            For Each vDelegateItem In mvDelegates.Values
                If IntegerValue(vDelegateItem.SequenceNumber) = vIndex Then
                    vDelegateSequenceFound = True
                    Exit For
                End If
            Next
            If Not vDelegateSequenceFound Then
                vDelegateItem = New EventDelegate
            End If
            vDelegatesCopy.Add(vIndex.ToString, vDelegateItem)
        Next
        mvDelegates = vDelegatesCopy
    End Sub

    Private Function GetContactDetails(pDelegate As EventDelegate, pOrganisationNumber As Integer, pAddressNumber As Integer) As ParameterList
        Dim vList As New ParameterList(HttpContext.Current)
        If pDelegate.Title.Length > 0 Then vList("Title") = pDelegate.Title
        If pDelegate.Forenames.Length > 0 Then vList("Forenames") = pDelegate.Forenames
        If pDelegate.Surname.Length > 0 Then vList("Surname") = pDelegate.Surname
        If pDelegate.Email.Length > 0 Then vList("EmailAddress") = pDelegate.Email
        If InitialParameters("DeDuplicate").ToString = "Y" Then vList("DeDuplicate") = "Y"
        vList("Source") = DefaultParameters("Source")
        If pOrganisationNumber > 0 Then vList("OrganisationNumber") = pOrganisationNumber.ToString
        If pAddressNumber > 0 Then vList("AddressNumber") = pAddressNumber.ToString
        Dim vReturnList As ParameterList = DataHelper.AddItem(CareNetServices.XMLMaintenanceControlTypes.xmctContact, vList)
        vList = New ParameterList(HttpContext.Current)
        vList("BookingNumber") = mvBookingNumber
        vList("ContactNumber") = vReturnList("ContactNumber")
        vList("AddressNumber") = vReturnList("AddressNumber")
        Return vList
    End Function
    Private Sub AssignDelegateRecordToDisplayRow(pSequenceList As Dictionary(Of String, ParameterList))

        'Assign existing delegate record id appropriate for sequence/seat contact 
        For Each vSequence As String In pSequenceList.Keys
            If pSequenceList(vSequence).Contains("ContactNumber") Then 'i.e.Change in Contact Number
                Dim vNewContact As String = pSequenceList(vSequence)("ContactNumber").ToString
                'Contact has been changed therefore the delegates record ID for this row may not be correct
                '- Find if another row has been assigned the delegates record ID for this new contact
                Dim vFound As Boolean = False
                For Each vExistingDelegate As EventDelegate In mvDelegates.Values
                    If Not vExistingDelegate.ContactNumber Is Nothing Then
                        If vExistingDelegate.ContactNumber = vNewContact Then
                            'This other row (vExistingDelegate) has the delegates record ID for the new contact
                            '- assign this delegate record ID to this row
                            pSequenceList(vSequence)("EventDelegateNumber") = vExistingDelegate.EventDelegateNumber
                            vExistingDelegate.EventDelegateNumber = ""
                            vFound = True
                            Exit For
                        End If
                    End If
                Next
                If Not vFound Then
                    pSequenceList(vSequence).Remove("EventDelegateNumber")
                End If
            Else
                'This may be an update without contact changes, so remove delegate ID so that it cannot be re-assigned to a different row. 
                mvDelegates(vSequence).EventDelegateNumber = ""
            End If
        Next vSequence

        'Assign non-assigned delegate record id to contact on same sequence/seat.
        '(This may not be correct, but it follows the logic of previous code - i.e. it is not making it worse)
        For Each vSequence As String In pSequenceList.Keys
            If pSequenceList(vSequence).Contains("ContactNumber") Then 'i.e.Change in Contact Number
                If Not pSequenceList(vSequence).Contains("EventDelegateNumber") Then
                    If mvDelegates.ContainsKey(vSequence) Then
                        If Not mvDelegates(vSequence).ContactNumber Is Nothing Then
                            If mvDelegates(vSequence).EventDelegateNumber.Length > 0 Then
                                '- assign this delegate record ID to this row
                                pSequenceList(vSequence)("EventDelegateNumber") = mvDelegates(vSequence).EventDelegateNumber
                                mvDelegates(vSequence).EventDelegateNumber = ""
                            End If
                        End If
                    End If
                End If
            End If
        Next vSequence

        'Assign non-assigned delegate record id to contacts without assigned delegate record id.
        '(This may not be strictly correct but previous portal function does these type of assignment - we are not making it worse) 
        For Each vSequence As String In pSequenceList.Keys
            If pSequenceList(vSequence).Contains("ContactNumber") Then 'i.e.Change in Contact Number
                If Not pSequenceList(vSequence).Contains("EventDelegateNumber") Then
                    For Each vChangedSequence As String In pSequenceList.Keys
                        Dim vExistingDelegate As EventDelegate = mvDelegates(vChangedSequence)
                        If Not vExistingDelegate.ContactNumber Is Nothing Then
                            If vExistingDelegate.EventDelegateNumber.Length > 0 Then
                                '- assign this delegate record ID to this row
                                pSequenceList(vSequence)("EventDelegateNumber") = vExistingDelegate.EventDelegateNumber
                                vExistingDelegate.EventDelegateNumber = ""
                                Exit For
                            End If
                        End If
                    Next vChangedSequence
                End If
            End If
        Next vSequence

    End Sub
    Private Class EventDelegate
        Private mvEventDelegateNumber As String
        Private mvContactNumber As String
        Private mvAddressNumber As String
        Private mvSequenceNumber As String
        Private mvTitle As String
        Private mvForenames As String
        Private mvSurname As String
        Private mvEmail As String
        Private mvPosition As String
        Private mvOrganisationName As String

        Private mvExisting As Boolean
        Private mvUpdateRequired As Boolean
        Private mvNewDelegateRequired As Boolean
        Private mvEmailAdded As Boolean
        Private mvMandatoryFieldsNotSet As Boolean
        Private mvAnyFieldsSet As Boolean

        Public Property AnyFieldsSet As Boolean
            Get
                Return mvAnyFieldsSet
            End Get
            Set(value As Boolean)
                mvAnyFieldsSet = value
            End Set
        End Property

        Public Property MandatoryFieldsNotSet As Boolean
            Get
                Return mvMandatoryFieldsNotSet
            End Get
            Set(value As Boolean)
                mvMandatoryFieldsNotSet = value
            End Set
        End Property

        Public ReadOnly Property Existing() As Boolean
            Get
                Return mvExisting
            End Get
        End Property

        Public ReadOnly Property UpdateRequired() As Boolean
            Get
                Return mvUpdateRequired
            End Get
        End Property

        Public ReadOnly Property NewDelegateRequired() As Boolean
            Get
                Return mvNewDelegateRequired
            End Get
        End Property

        Public ReadOnly Property EmailAdded() As Boolean
            Get
                Return mvEmailAdded
            End Get
        End Property

        Public Property EventDelegateNumber() As String
            Get
                Return mvEventDelegateNumber
            End Get
            Set(value As String)
                mvEventDelegateNumber = value
            End Set
        End Property
        Public ReadOnly Property ContactNumber() As String
            Get
                Return mvContactNumber
            End Get
        End Property
        Public ReadOnly Property AddressNumber() As String
            Get
                Return mvAddressNumber
            End Get
        End Property
        Public Property SequenceNumber() As String
            Get
                Return mvSequenceNumber
            End Get
            Set(value As String)
                If mvExisting AndAlso value <> mvSequenceNumber Then
                    mvUpdateRequired = True
                End If
                mvSequenceNumber = value
            End Set
        End Property
        Public Property Title() As String
            Get
                Return mvTitle
            End Get
            Set(value As String)
                If mvExisting AndAlso value <> mvTitle Then
                    mvNewDelegateRequired = True
                End If
                mvTitle = value
            End Set
        End Property
        Public Property Forenames() As String
            Get
                Return mvForenames
            End Get
            Set(value As String)
                If mvExisting AndAlso value <> mvForenames Then
                    mvNewDelegateRequired = True
                End If
                mvForenames = value
            End Set
        End Property
        Public Property Surname() As String
            Get
                Return mvSurname
            End Get
            Set(value As String)
                If mvExisting AndAlso value <> mvSurname Then
                    mvNewDelegateRequired = True
                End If
                mvSurname = value
            End Set
        End Property
        Public Property Email() As String
            Get
                Return mvEmail
            End Get
            Set(value As String)
                If mvExisting Then
                    If value.Length > 0 AndAlso mvEmail.Length = 0 Then
                        mvEmailAdded = True             'Email address has been added to existing delegate
                    ElseIf value <> mvEmail Then
                        mvNewDelegateRequired = True    'Changed email address could be new delegate
                    End If
                End If
                Utilities.ValidateEmailAddress(value)
                mvEmail = value
            End Set
        End Property
        Public Property Position As String
            Get
                Return mvPosition
            End Get
            Set(value As String)
                If mvExisting AndAlso value <> mvPosition Then
                    mvUpdateRequired = True
                End If
                mvPosition = value
            End Set
        End Property
        Public Property OrganisationName As String
            Get
                Return mvOrganisationName
            End Get
            Set(value As String)
                If mvExisting AndAlso value <> mvOrganisationName Then
                    mvUpdateRequired = True
                End If
                mvOrganisationName = value
            End Set
        End Property

        Public Sub New()
            'Not an existing delegate
            mvTitle = ""
            mvForenames = ""
            mvSurname = ""
            mvEmail = ""
            mvPosition = ""
            mvOrganisationName = ""
            mvSequenceNumber = ""
        End Sub

        Public Sub New(ByVal pContactNumber As String, ByVal pEventDelegateNumber As String, ByVal pAddressNumber As String, pSequenceNumber As String, ByVal pTitle As String, ByVal pForenames As String, ByVal pSurname As String, ByVal pEmail As String, ByVal pPosition As String, ByVal pOrganisationName As String)
            mvEventDelegateNumber = pEventDelegateNumber
            mvContactNumber = pContactNumber
            mvAddressNumber = pAddressNumber
            mvSequenceNumber = pSequenceNumber
            mvTitle = pTitle
            mvForenames = pForenames
            mvSurname = pSurname
            mvEmail = pEmail
            mvPosition = pPosition
            mvOrganisationName = pOrganisationName
            mvExisting = True
        End Sub
    End Class
End Class