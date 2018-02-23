Imports System.Text.RegularExpressions
Public Class frmAddCriteria

#Region "Private Variables"
  Private mvMailingSelection As MailingInfo
  Private mvStandardExclusions As Boolean
  Private mvValueValid As Boolean
  Private mvAreaInfo As SearchAreaInfo
  Private mvValueNull As Boolean
  Private mvEndValueValid As Boolean
  Private mvNullMainDTPValue As Boolean
  Private mvSubValueValid As Boolean
  Private mvSubValueNull As Boolean
  Private mvEndSubValueValid As Boolean
  Private mvNullSubDTPValue As Boolean
  Private mvLastType As String
  Private mvSearchAreaValid As Boolean
  Private mvValueVariable As Boolean
  Private mvIsSourceLookupValueLoaded As Boolean
  Private mvIsSourceLookupEndValueLoaded As Boolean
  Private mvArrayListValue As New ArrayList()
  Private mvArrayListSubValue As New ArrayList()
  Private mvArrayListPeriod As New ArrayList()
  Private mvControlInfo As SelectionControlInfo
  Private mvSubValuesInit As Boolean

  Private mvMainValidationItem As New ValidationItem
  Private mvSubValidationItem As New ValidationItem

  Private Const NOT_TEXTBOX_LOOKUP As String = "NotTextBoxOrLookupControl"
#End Region

#Region "Structure SearchAreaInfo"
  Private Structure SearchAreaInfo
    Dim MainValue As Boolean
    Dim SubValue As Boolean
    Dim CorO As String
    Dim Period As Boolean
    Dim MainValidate As Boolean
    Dim SubValidate As Boolean
  End Structure
#End Region

#Region "Structure SelectionControlInfo"
  Private Structure SelectionControlInfo
    Dim MainAttr As String
    Dim MainValueHeading As String
    Dim SubAttr As String
    Dim SubValueHeading As String
    Dim ValidationTable As String
    Dim MainValidationAttr As String
    Dim SubValidationAttr As String
    Dim MainDataType As String
    Dim SubDataType As String
    Dim MainLength As Integer
    Dim SubLength As Integer
    Dim TableName As String
    Dim Pattern As String
  End Structure
#End Region

#Region "Enum DTPFields"
  Private Enum DTPFields
    dtpMainFrom
    dtpMainTo
    dtpPeriodFrom
    dtpPeriodTo
    dtpSubFrom
    dtpSubTo
  End Enum
#End Region

  Public Sub New(ByVal pMailingSelection As MailingInfo, Optional ByVal pExclusions As Boolean = False)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()
    ' Add any initialization after the InitializeComponent() call.

    InitialiseControls(pMailingSelection, pExclusions)
  End Sub

#Region "Button Events"

  Private Sub cmdValueVar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValueVar.Click
    Try
      If dtpValue.Visible = True Or mvControlInfo.MainDataType <> "D" Then
        dtpValue.Visible = False
        txtValue.Visible = True
        txtValue.Text = ControlText.TxtDollar
        SetFocusOnLookupValueControl(True)
        txtValue.Select(1, txtValue.Text.Length)
        txtValue.Focus()
        txtValue.MaxLength = 0
        cmdAddValue.Enabled = False
        cmdDeleteValue.Enabled = False
      Else
        dtpValue.Visible = True
        txtValue.Visible = False
        dtpValue.Focus()
        cmdAddValue.Enabled = Not dtpValue.Checked ' IsNull(dtpValue)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAddValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddValue.Click
    Dim vValue As String
    Dim vControlName As String = NOT_TEXTBOX_LOOKUP
    Try
      If chkValueRange.Checked = True Then
        'A range has been selected
        'Check if the range is valid
        If dtpValue.Visible Then
          If dtpEndValue.Checked Then
            If dtpEndValue.Value <= dtpValue.Value Then
              ShowInformationMessage(InformationMessages.ImFromDateGTToDate)
              dtpEndValue.Focus()
              Exit Sub
            End If
            vValue = dtpValue.Value.ToString(AppValues.DateFormat) & " to " & dtpEndValue.Value.ToString(AppValues.DateFormat) ' Format$(dtpValue.Value, DateFormat) & " to " & Format$(dtpEndValue.Value, DateFormat)
          Else
            vValue = dtpValue.Value.ToString(AppValues.DateFormat) 'Format$(dtpValue.Value, DateFormat)
          End If
        Else
          If CheckRange(GetTextBoxValue().Text, GetTextBoxEndValue().Text, True) = False Then
            ShowInformationMessage(String.Format(InformationMessages.ImFieldInvalidRange, GetTextBoxValue().Text, GetTextBoxEndValue().Text))
            GetTextBoxEndValue().Focus()
            Exit Sub
          End If
          'BR12095: changes to add quotes for Range values
          vValue = HandleSpacesAndSingleQuotes(GetTextBoxValue().Text, mvControlInfo.MainDataType) & " to " & HandleSpacesAndSingleQuotes(GetTextBoxEndValue().Text, mvControlInfo.MainDataType)
        End If
      Else
        If dtpValue.Visible Then
          If dtpValue.Checked = False Then
            vValue = "NULL"
            mvNullMainDTPValue = True
          Else
            vValue = dtpValue.Value.ToString(AppValues.DateFormat) ' Format$(dtpValue.Value, DateFormat)
            mvNullMainDTPValue = False
          End If
        ElseIf chkYN.Visible Then
          vValue = CStr(IIf(chkYN.Checked = True, "Y", "N"))
        Else
          vValue = GetTextBoxValue().Text 'txtValue.Text
        End If
      End If
      If ValidateVariable(vValue, GetTextBoxValue()) Then
        If Not ItemExists(lstValue, vValue) AndAlso Not ItemExists(lstValue, "'" & vValue & "'") Then  'BR16014 check vValue with and without quotes
          If chkValueRange.Checked = False And dtpValue.Visible = False And chkYN.Visible = False Then
            vValue = HandleSpacesAndSingleQuotes(vValue, mvControlInfo.MainDataType)      'BR12095: add quotes if there is a space in vValue
          End If
          lstValue.Items.Add(vValue)
          lstValue.SelectedIndex = lstValue.Items.Count - 1
          cmdDeleteValue.Enabled = True
          SetVisibleControls()
          CheckCriteria()
          If chkYN.Visible Then cmdAddValue.Enabled = False
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdDeleteValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteValue.Click
    Try
      Dim vIndex As Integer

      vIndex = lstValue.SelectedIndex
      lstValue.Items.RemoveAt(vIndex)
      If vIndex >= lstValue.Items.Count Then vIndex = vIndex - 1
      lstValue.SelectedIndex = vIndex
      If lstValue.Items.Count = 0 Then
        cmdDeleteValue.Enabled = False
        cmdDeleteSubValue.Enabled = False
        lstSubValue.Items.Clear()
      Else
        If lstValue.SelectedIndex >= 0 Then lstValue_Click(lstValue, Nothing)
      End If
      mvNullMainDTPValue = False
      SetVisibleControls()
      CheckCriteria()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdSubValueVar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSubValueVar.Click
    Try
      If dtpSubValue.Visible = True Or mvControlInfo.SubDataType <> "D" Then
        dtpSubValue.Visible = False
        txtSubValue.Visible = True
        txtSubValue.Enabled = True
        txtSubValue.Text = ControlText.TxtDollar
        SetFocusOnLookupSubValueControl(True)
        txtSubValue.Select(1, txtValue.Text.Length)
        SafeSetFocus(txtSubValue)
        txtSubValue.MaxLength = 0
        cmdAddSubValue.Enabled = False
        cmdDeleteValue.Enabled = False
      Else
        dtpSubValue.Visible = True
        txtSubValue.Visible = False
        dtpSubValue.Focus()
        cmdAddSubValue.Enabled = Not dtpSubValue.Checked ' IsNull(dtpSubValue.Value)
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAddSubValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddSubValue.Click
    Dim vValue As String
    Dim vControlName As String = NOT_TEXTBOX_LOOKUP
    Try
      If chkSubValueRange.Checked = True Then
        'A range has been selected
        'Check if the range is valid
        If dtpSubValue.Visible Then
          If dtpEndSubValue.Checked Then
            If dtpEndSubValue.Value <= dtpSubValue.Value Then
              ShowInformationMessage(InformationMessages.ImFromDateGTToDate)
              dtpEndSubValue.Focus()
              Exit Sub
            End If
            vValue = dtpSubValue.Value.ToString(AppValues.DateFormat) & " to " & dtpEndSubValue.Value.ToString(AppValues.DateFormat)
          Else
            vValue = dtpSubValue.Value.ToString(AppValues.DateFormat)
          End If
        Else
          If CheckRange(GetTextBoxSubValue().Text, GetTextBoxEndSubValue().Text, False) = False Then
            ShowInformationMessage(String.Format(InformationMessages.ImFieldInvalidRange, GetTextBoxSubValue().Text, GetTextBoxEndSubValue().Text))
            GetTextBoxEndSubValue().Focus()
            Exit Sub
          End If
          'BR12095: changes to add quotes for Range values
          vValue = HandleSpacesAndSingleQuotes(GetTextBoxSubValue().Text, mvControlInfo.SubDataType) & " to " & HandleSpacesAndSingleQuotes(GetTextBoxEndSubValue().Text, mvControlInfo.SubDataType)
        End If
      Else
        If dtpSubValue.Visible Then
          If dtpSubValue.Checked Then
            vValue = "NULL"
            mvNullSubDTPValue = True
          Else
            vValue = dtpSubValue.Value.ToString(AppValues.DateFormat) 'Format$(dtpSubValue.Value, DateFormat)
            mvNullSubDTPValue = False
          End If
        Else
          vValue = GetTextBoxSubValue().Text 'txtValue.Text
        End If
      End If
      If ValidateVariable(vValue, txtSubValue) Then
        If Not ItemExists(lstSubValue, vValue) Then
          If chkSubValueRange.Checked = False And dtpSubValue.Visible = False Then
            vValue = HandleSpacesAndSingleQuotes(vValue, mvControlInfo.SubDataType)      'BR12095: add quotes if there is a space in vValue
          End If
          lstSubValue.Items.Add(vValue)
          lstSubValue.SelectedIndex = lstSubValue.Items.Count - 1
          cmdDeleteSubValue.Enabled = True
          CheckCriteria()
          cmdAddValue_Click(Me, Nothing)
        End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub


  Private Sub cmdDeleteSubValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeleteSubValue.Click
    Dim vIndex As Integer
    Try
      vIndex = lstSubValue.SelectedIndex
      lstSubValue.Items.RemoveAt(vIndex)

      If vIndex >= lstSubValue.Items.Count Then vIndex = vIndex - 1
      lstSubValue.SelectedIndex = vIndex
      If lstSubValue.Items.Count = 0 Then cmdDeleteSubValue.Enabled = False
      CheckCriteria()
      mvNullSubDTPValue = False
      SetVisibleControls()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdPeriodVar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPeriodVar.Click
    Try
      If dtpFrom.Visible = True Then
        dtpFrom.Visible = False
        lblPeriodFrom.Visible = False
        dtpTo.Visible = False
        lblPeriodTo.Visible = False
        txtPeriodVar.Visible = True
        lblPeriodVar.Visible = True
        txtPeriodVar.Text = ControlText.TxtDollar
        txtPeriodVar.Select(1, txtValue.Text.Length)
        txtPeriodVar.Focus()
        cmdAddPeriod.Enabled = False
        cmdDeletePeriod.Enabled = False
      Else
        dtpFrom.Visible = True
        lblPeriodFrom.Visible = True
        dtpTo.Visible = True
        lblPeriodTo.Visible = True
        txtPeriodVar.Visible = False
        lblPeriodVar.Visible = False
        dtpFrom.Focus()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdAddPeriod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddPeriod.Click
    Dim vControlName As String = NOT_TEXTBOX_LOOKUP
    Try
      If dtpFrom.Visible = True Then
        If dtpTo.Checked Then
          If CDate(dtpTo.Value) <= CDate(dtpFrom.Value) Then
            ShowInformationMessage(InformationMessages.ImFromDateGTToDate)
            dtpTo.Focus()
            Exit Sub
          End If
          lstPeriod.Items.Add(dtpFrom.Value.ToString(AppValues.DateFormat) & " to " & dtpTo.Value.ToString(AppValues.DateFormat))
        Else
          lstPeriod.Items.Add(dtpFrom.Value.ToString(AppValues.DateFormat))
        End If
        lstPeriod.SelectedIndex = lstPeriod.Items.Count - 1
        cmdDeletePeriod.Enabled = True
      ElseIf txtPeriodVar.Visible = True And txtPeriodVar.Text.Length > 1 AndAlso Strings.Left(txtPeriodVar.Text, 1) = "$" Then
        If Not ValidateVariable(txtPeriodVar.Text, txtPeriodVar) Then Exit Sub
        vControlName = txtPeriodVar.Name
        lstPeriod.Items.Add(txtPeriodVar.Text)
        lstPeriod.SelectedIndex = lstPeriod.Items.Count - 1
        cmdDeletePeriod.Enabled = True
      End If
      CheckCriteria()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdDeletePeriod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeletePeriod.Click
    Dim vIndex As Integer
    Try
      vIndex = lstPeriod.SelectedIndex
      If vIndex <> -1 Then
        lstPeriod.Items.RemoveAt(vIndex)

        If vIndex >= lstPeriod.Items.Count - 1 Then vIndex = vIndex - 1
        lstPeriod.SelectedIndex = vIndex
        If lstPeriod.Items.Count = 0 Then cmdDeletePeriod.Enabled = False
        CheckCriteria()
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Try
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    Try
      Dim vSearchArea As String
      Dim vCO As String
      Dim vIE As String
      Dim vValue As String = String.Empty
      Dim vSubValue As String = String.Empty
      Dim vPeriod As String = String.Empty
      Dim vIndex As Integer

      vSearchArea = txtLookupArea.Text
      vCO = CStr(IIf(optPerson.Checked = True, "C", "O"))
      vIE = CStr(IIf(optInclude.Checked = True, "I", "E"))
      'Set values and sub values
      If lstValue.Visible = True Then
        For vIndex = 0 To lstValue.Items.Count - 1
          If vIndex <> 0 Then vValue = vValue + ","
          vValue = vValue + lstValue.Items.Item(vIndex).ToString
        Next
        If lstSubValue.Visible = True Then
          For vIndex = 0 To lstSubValue.Items.Count - 1
            If vIndex <> 0 Then vSubValue = vSubValue + ","
            vSubValue = vSubValue + lstSubValue.Items.Item(vIndex).ToString
          Next
        End If
      End If
      'Set the periods
      If lstPeriod.Visible = True Then
        For vIndex = 0 To lstPeriod.Items.Count - 1
          If vIndex <> 0 Then vSubValue = vSubValue + ","
          vPeriod = vPeriod + lstPeriod.Items.Item(vIndex).ToString
        Next
      End If
      With mvMailingSelection.CurrentCriteria
        .SearchArea = vSearchArea
        .CO = vCO
        .IE = vIE
        .MainValue = CStr(IIf(vValue Is Nothing, String.Empty, vValue))
        .SubsidiaryValue = CStr(IIf(vSubValue Is Nothing, String.Empty, vSubValue))
        .Period = CStr(IIf(vPeriod Is Nothing, String.Empty, vPeriod)) ' vPeriod
        .Valid = True
      End With
      Me.Close()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Function HandleSpacesAndSingleQuotes(ByVal pValue As String, ByVal pDataType As String) As String
    ' BR11993 - Need to deal with any spaces in the 'Main value' value going into

    Dim vTemp As String = String.Empty
    Dim vTempArray() As String
    Dim vCount As Integer

    vTempArray = Split(pValue, Environment.NewLine)  ' Chr$(10)
    For vCount = 0 To UBound(vTempArray)
      If pDataType = "C" Then
        If Strings.InStr(1, vTempArray(vCount), " ") > 0 Then
          If ((Strings.Left(vTempArray(vCount), 1) = """") And Strings.Right(vTempArray(vCount), 1) = """") Or ((Strings.Left(vTempArray(vCount), 1) = "'") And (Strings.Right(vTempArray(vCount), 1) = "'")) Then '  (Right(vTempArray(vCount), 1) = """")
            ' The user has already put the quotes in for us - do nothing
          ElseIf (Strings.Left(vTempArray(vCount), 1) = "'") Then
            vTempArray(vCount) = vTempArray(vCount) & "'"
          ElseIf (Strings.Right(vTempArray(vCount), 1) = "'") Then
            vTempArray(vCount) = "'" & vTempArray(vCount)
          Else
            vTempArray(vCount) = "'" & vTempArray(vCount) & "'"
          End If
        Else
          While (Strings.Left(vTempArray(vCount), 1) = "'")      'Remove all single quotes in the begining
            vTempArray(vCount) = Strings.Mid(vTempArray(vCount), 2, vTempArray(vCount).Length)
          End While
          While (Strings.Right(vTempArray(vCount), 1) = "'")     'Remove all single quotes in the end
            vTempArray(vCount) = Strings.Mid(vTempArray(vCount), 1, vTempArray(vCount).Length - 1) '
          End While

          If ((Strings.Left(vTempArray(vCount), 1) = """") And (Strings.Right(vTempArray(vCount), 1) <> """")) Or ((Strings.Right(vTempArray(vCount), 1) = """") And (Strings.Left(vTempArray(vCount), 1) <> """")) Then
            vTempArray(vCount) = "'" & vTempArray(vCount) & "'"
          End If
        End If
      Else
        ' Its a date type field, do nowt!
      End If
    Next vCount

    For vCount = 0 To UBound(vTempArray)
      vTemp = vTemp & vTempArray(vCount) & ","
    Next vCount
    If Strings.Right(vTemp, 1) = "," Then vTemp = Strings.Left(vTemp, vTemp.Length - 1)

    Return vTemp

  End Function

  Private Function ItemExists(ByVal pListBox As ListBox, ByVal pValue As String) As Boolean
    Dim vIndex As Integer
    Dim vFound As Boolean

    For vIndex = 0 To pListBox.Items.Count - 1
      If pListBox.Items.Item(vIndex).ToString = pValue Then
        vFound = True
        Exit For
      End If
    Next
    Return vFound
  End Function

  Private Function CheckRange(ByVal pStartVal As String, ByVal pEndVal As String, ByVal pMain As Boolean) As Boolean
    Dim vDataType As String
    Dim vInRange As Boolean

    If pMain = True Then
      vDataType = mvControlInfo.MainDataType
    Else
      vDataType = mvControlInfo.SubDataType
    End If
    vInRange = True

    If Strings.Left(pStartVal, 1) = "$" OrElse Strings.Left(pEndVal, 1) = "$" Then
      vInRange = False
    Else
      Select Case vDataType
        Case "C"  'Character
          If pEndVal <= pStartVal Then vInRange = False
        Case "I", "N" 'Integer, Number
          If Val(pEndVal) <= Val(pStartVal) Then vInRange = False
        Case "D"
          If IsDate(pStartVal) And IsDate(pEndVal) Then
            If CDate(pEndVal) < CDate(pStartVal) Then vInRange = False
          Else
            vInRange = False  'Should never happen
          End If
        Case Else
          vInRange = False  'Should never happen
      End Select
    End If
    Return vInRange
  End Function

  Public Function ValidateVariable(ByVal pValue As String, ByVal pControl As Control) As Boolean
    Dim vValid As Boolean
    Dim vToDateBeforeFromDate As Boolean = False

    'BR18706 - max characters in database column
    If pValue.Length > 1000 Then
      ShowInformationMessage(InformationMessages.ImInvalidVariableCharacterCount)
      vValid = False
      pControl.Focus()
    Else
      If pValue.ToUpper.StartsWith("$TODAY") Then
        If pValue.Length > 6 Then
          'Ensure that the $TODAY variable is in the format: $TODAY [+|-n [+|-n2]], where:
          'n = the # of days before (-) or after (+) the current date the lower end of the date range will be.  This value cannot be greater than n2.
          'n2 = the # of days before (-) or after (+) the current date the uppper end of the date range will be.  This value cannot be less than n.
          'Both n & n2 must be numeric with a value between zero and 9999.
          'vValid will be set to False if/when the format doesn't match the above description.
          Dim vParamString As String = pValue.Substring(6)
          Dim vOperator1 As String = ""
          Dim vValue1 As String = ""
          Dim vOperator2 As String = ""
          Dim vValue2 As String = ""
          Dim vPos As Integer
          If vParamString.StartsWith("+") OrElse vParamString.StartsWith("-") Then
            vOperator1 = vParamString.Substring(0, 1)       'We have the first operator
            vParamString = vParamString.Remove(0, 1)                       'Remove it from the parameter string
            If vParamString.Length > 0 Then
              vPos = vParamString.IndexOfAny("+-".ToCharArray)
              If vPos >= 0 Then
                vOperator2 = vParamString.Substring(vPos, 1)
                vValue1 = vParamString.Substring(0, vPos)
                vValue2 = vParamString.Substring(vPos + 1)
              Else
                vValue1 = vParamString
              End If
            End If
            If IsNumeric(vValue1) Then
              If CInt(vValue1) >= 0 AndAlso CInt(vValue1) <= 9999 Then
                If vOperator2.Length = 0 Then
                  vValid = True
                Else
                  If IsNumeric(vValue2) Then
                    If CInt(vValue2) >= 0 AndAlso CInt(vValue2) <= 9999 Then
                      If CInt(vOperator2 & vValue2) > CInt(vOperator1 & vValue1) Then
                        vValid = True
                        'BR19001 invalid date ranges now showing appropriate message.
                      Else
                        vValid = False
                        vToDateBeforeFromDate = True
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        ElseIf pValue.Length = 6 Then        'Variable just contains $TODAY, so that's OK
          vValid = True
        End If
        If Not vValid Then
          If vToDateBeforeFromDate Then
            ShowInformationMessage(InformationMessages.ImInvalidVariableDateRange)
          Else
            ShowInformationMessage(InformationMessages.ImInvalidVariableFormat)
          End If
          pControl.Focus()
        End If
      Else
        If pValue.StartsWith("$") Then
          Dim vRegExStartChar As Regex = New Regex("[^a-zA-Z_]") 'Must start with a letter or underscore
          Dim vRegEx As Regex = New Regex("[^a-zA-Z0-9_.-]") 'Only allow alphanumeric, _ - .
          'BR17896 - Client side validation added. Avoid Server side errors.
          'BR18706 - Variable names in Campaign Manager must be restricted as they are processed as xml element names
          '          Element names are case-sensitive and must start with a letter or underscore. An element name can contain letters, digits, hyphens, underscores, and periods.
          If pValue.ToUpper.StartsWith("$XML") Then
            ShowInformationMessage(InformationMessages.ImInvalidVariableNameStartXML) 'Cannot start with letters "xml" upper or lowercase or combination therof
            vValid = False
            pControl.Focus()
          ElseIf vRegExStartChar.Match(pValue.Substring(1, 1)).Success Then
            ShowInformationMessage(InformationMessages.ImInvalidVariableNameStartChar) 'Must start with a letter or underscore
            vValid = False
            pControl.Focus()
          ElseIf vRegEx.Match(pValue.Remove(0, 2)).Success Then
            Dim vRegExInvalidCharacters As Regex = New Regex("[a-zA-Z0-9_.-]")
            Dim vInvalidCharacters As String = vRegExInvalidCharacters.Replace(pValue.Remove(0, 2), "") 'Now check validity of rest of string
            ShowInformationMessage(InformationMessages.ImInvalidVariableName, vInvalidCharacters)
            vValid = False
            pControl.Focus()
          Else
            vValid = True
          End If
        Else
          vValid = True 'It is not a variable, so cannot be an invalid variable.
        End If
      End If
    End If
    Return vValid
  End Function

  Private Sub SetFocusOnLookupValueControl(ByVal pTxtValueFocus As Boolean)
    If pTxtValueFocus Then
      txtLookupValue.TextBox.Text = ""
      txtLookupValue.SetComboString("")
    Else
      If txtLookupValue.Text.Length > 0 Then txtValue.Text = ""
    End If
  End Sub


  Private Sub SetFocusOnLookupSubValueControl(ByVal pTxtSubValueFocus As Boolean)
    If pTxtSubValueFocus Then
      txtLookupSubValue.TextBox.Text = ""
      txtLookupSubValue.SetComboString("")
    Else
      txtSubValue.Text = ""
    End If
  End Sub
#End Region

#Region "ListBox Events"
  Private Sub lstValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstValue.Click
    Try

      Dim vPos As Integer
      Dim vSource As String

      If lstValue.SelectedIndex >= 0 Then
        vSource = lstValue.Text ' lstValue.List(lstValue.ListIndex)
        vPos = Strings.InStr(vSource, " to ")
        If vPos > 0 Then
          If mvControlInfo.MainDataType = "D" Then
            dtpValue.Checked = True
            dtpValue.Value = CDate(Strings.Mid(vSource, 1, vPos - 1)) 'Mid$(vSource, 1, vPos - 1)
            dtpEndValue.Checked = True
            dtpEndValue.Value = CDate(Strings.Mid(vSource, vPos + 4)) 'Mid$(vSource, vPos + 4)
          Else
            Dim vControlName As String = NOT_TEXTBOX_LOOKUP  'GetControlNameFromArrayList(mvArrayListValue, lstValue.Text)

            If Strings.Left(vSource, 1) = "'" Then               'Remove any single quotes to ONLY display in text boxes
              txtValue.Text = Strings.Mid(vSource, 2, vPos - 3) ' Mid$(vSource, 2, vPos - 3)
            Else
              Dim vTextBox As TextBox
              vTextBox = GetTextBoxFromControlValue(vControlName, txtLookupValue, Strings.Mid(vSource, 1, vPos - 1))
              If vTextBox Is Nothing Then vTextBox = txtValue
              vTextBox.Text = Strings.Mid(vSource, 1, vPos - 1) 'Mid$(vSource, 1, vPos - 1)
              SetBlankForAlternateTextbox(vTextBox, "value")
            End If
            If Strings.Right(vSource, 1) = "'" Then ' Right$(vSource, 1)
              txtEndValue.Text = Strings.Mid(vSource, vPos + 5, (vSource.Length - vPos - 5)) ' Mid$(vSource, vPos + 5, (Len(vSource) - vPos - 5))
            Else
              Dim vTextBox1 As TextBox
              vTextBox1 = GetTextBoxFromControlValue(vControlName, txtLookupEndValue, Strings.Mid(vSource, vPos + 4))
              If vTextBox1 Is Nothing Then vTextBox1 = txtEndValue
              vTextBox1.Text = Strings.Mid(vSource, vPos + 4) ' Mid$(vSource, vPos + 4)
              SetBlankForAlternateTextbox(vTextBox1, "endvalue")
            End If
          End If
          chkValueRange.Checked = True
        Else
          Dim vControlName As String = NOT_TEXTBOX_LOOKUP ' GetControlNameFromArrayList(mvArrayListValue, lstValue.Text)
          Dim vTextBox As TextBox = GetTextBoxFromControlValue(vControlName, txtLookupValue, lstValue.Text)
          If vTextBox Is Nothing Then vTextBox = txtValue
          If Strings.Left(vSource, 1) = "$" Then
            vTextBox.MaxLength = 0
          Else
            vTextBox.MaxLength = mvControlInfo.MainLength
          End If
          If vSource.Length > 0 Then
            If vSource = "NULL" And vTextBox.MaxLength < 4 Then
              vTextBox.MaxLength = 4
            ElseIf vSource = "NOTNULL" And vTextBox.MaxLength < 7 Then
              vTextBox.MaxLength = 7
            End If
          End If
          If Strings.Left(vSource, 1) = "'" Then                     'Remove any single quotes to ONLY display in text boxes
            vTextBox.Text = Strings.Mid(vSource, 2, vSource.Length - 2) ' Mid$(vSource, 2, Len(vSource) - 2)
          Else
            vTextBox.Text = vSource
          End If
          SetBlankForAlternateTextbox(vTextBox, "value")
          chkValueRange.Checked = False
        End If
        cmdDeleteValue.Enabled = True
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub lstSubValue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstSubValue.Click
    Try
      Dim vPos As Integer
      Dim vSource As String

      If lstSubValue.SelectedIndex >= 0 Then
        vSource = lstSubValue.Text ' lstSubValue.List(lstSubValue.ListIndex)
        vPos = InStr(vSource, " to ")
        If vPos > 0 Then
          If mvControlInfo.SubDataType = "D" Then
            dtpSubValue.Checked = True
            dtpSubValue.Value = CDate(Strings.Mid(vSource, 1, vPos - 1)) 'Mid$(vSource, 1, vPos - 1)
            dtpEndSubValue.Checked = True
            dtpEndSubValue.Value = CDate(Strings.Mid(vSource, vPos + 4)) 'Mid$(vSource, vPos + 4)
          Else
            Dim vControlName As String = NOT_TEXTBOX_LOOKUP ' GetControlNameFromArrayList(mvArrayListSubValue, lstSubValue.Text)
            If Strings.Left(vSource, 1) = "'" Then
              txtSubValue.Text = Strings.Mid(vSource, 2, vPos - 3) ' Mid$(vSource, 2, vPos - 3)
            Else
              Dim vTextBox As TextBox
              vTextBox = GetTextBoxFromControlValue(vControlName, txtLookupSubValue, Strings.Mid(vSource, 1, vPos - 1))
              If vTextBox Is Nothing Then vTextBox = txtSubValue
              vTextBox.Text = Strings.Mid(vSource, 1, vPos - 1)
              SetBlankForAlternateTextbox(vTextBox, "subvalue")
            End If
            If Strings.Right(vSource, 1) = "'" Then  '  Right$(vSource, 1)
              txtEndSubValue.Text = Strings.Mid(vSource, vPos + 5, (vSource.Length - vPos - 5)) ' Mid$(vSource, vPos + 5, (Len(vSource) - vPos - 5))
            Else
              Dim vTextBox1 As TextBox
              vTextBox1 = GetTextBoxFromControlValue(vControlName, txtLookupEndSubValue, Strings.Mid(vSource, vPos + 4))
              If vTextBox1 Is Nothing Then vTextBox1 = txtEndSubValue
              vTextBox1.Text = Strings.Mid(vSource, vPos + 4)
              SetBlankForAlternateTextbox(vTextBox1, "endsubvalue")
            End If
          End If
          chkSubValueRange.Checked = True
        Else
          Dim vControlName As String = NOT_TEXTBOX_LOOKUP ' GetControlNameFromArrayList(mvArrayListSubValue, lstSubValue.Text)
          Dim vTextBox As TextBox = GetTextBoxFromControlValue(vControlName, txtLookupSubValue, lstSubValue.Text)
          If vTextBox Is Nothing Then vTextBox = txtSubValue
          If Strings.Left(vSource, 1) = "$" Then
            vTextBox.MaxLength = 0
          Else
            vTextBox.MaxLength = mvControlInfo.SubLength
          End If
          If vSource.Length > 0 Then
            If vSource = "NULL" And vTextBox.MaxLength < 4 Then
              vTextBox.MaxLength = 4
            ElseIf vSource = "NOTNULL" And vTextBox.MaxLength < 7 Then
              vTextBox.MaxLength = 7
            End If
          End If
          If Strings.Left(vSource, 1) = "'" Then
            vTextBox.Text = Strings.Mid(vSource, 2, vSource.Length - 2) 'Mid$(vSource, 2, Len(vSource) - 2)
          Else
            vTextBox.Text = vSource
          End If
          SetBlankForAlternateTextbox(vTextBox, "subvalue")
          chkSubValueRange.Checked = False
        End If
        cmdDeleteSubValue.Enabled = True
      End If

    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub lstPeriod_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstPeriod.Click
    Try
      If lstPeriod.SelectedIndex >= 0 Then
        cmdDeletePeriod.Enabled = True
      Else
        cmdDeletePeriod.Enabled = False
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
#End Region

#Region "Radio Button Events"
  Private Sub optOrganisation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optOrganisation.Click
    Try
      'Re-read the selection control information
      If optOrganisation.Checked AndAlso mvLastType <> "O" Then GetSelectionControl()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub optContact_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optPerson.Click
    Try
      'Re-read the selection control information
      If optPerson.Checked AndAlso mvLastType <> "C" Then GetSelectionControl()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
#End Region

#Region "CheckBox Events"

  Private Sub chkYN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkYN.CheckedChanged
    Try
      mvValueValid = True
      mvValueVariable = False
      mvValueNull = False
      SetVisibleControls()
      CheckCriteria()
      If lstValue.Items.Count > 0 Then cmdAddValue.Enabled = False
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub chkSubValueRange_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSubValueRange.CheckedChanged
    Try
      SetVisibleControls()
      CheckCriteria()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub chkValueRange_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkValueRange.CheckedChanged
    Try
      SetVisibleControls()
      CheckCriteria()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
#End Region

#Region "Form Events"
  Private Sub frmAddCriteria_Activated(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Activated
    Try
      If mvMailingSelection.CurrentCriteria.Valid Then
        SetFromCurrentCriteria()
        mvMailingSelection.CurrentCriteria.Valid = False
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
#End Region

#Region "TextBox Events"

  Private Sub txtValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtValue.TextChanged
    Try
      OnChangeTextLookupValue()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtValue_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtValue.KeyPress
    Try
      If txtLookupValue.Visible Then txtLookupValue.SetTextBoxString("") ' Blank out Lookup Value control if txt value is entered
      If mvControlInfo.MainAttr <> "logname" And Strings.Left(GetTextBoxValue().Text, 1) <> "$" Then
        'There is a validation code and combo so make it upper case
        ' TODO: KeyAscii ?
        If e.KeyChar = "*" AndAlso mvMainValidationItem.CaseConversion.Length > 0 Then
          ' Leave
        Else
          If Strings.Mid(txtValue.Text, 1, 2) = "NU" Then
            'Assume that NULL is being entered
            If mvControlInfo.MainLength < 4 Then GetTextBoxValue().MaxLength = 4
          ElseIf Strings.Mid(txtValue.Text, 1, 2) = "NO" Then
            'Assume that NOTNULL is being entered
            If mvControlInfo.MainLength < 7 Then GetTextBoxValue().MaxLength = 7
          End If
        End If
        'If KeyAscii = Asc("*") And Len(mvMainValidationItem.CaseConversion) > 0 Then
        '  'Leave
        'mvMainValidationItem.HandleKeyPress(KeyAscii)
        'End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtEndValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEndValue.TextChanged
    Try
      OnChangeTextLookupEndValue()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtEndValue_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEndValue.KeyPress
    Try
      If txtLookupEndValue.Visible Then txtLookupEndValue.SetTextBoxString("") ' Blank out Lookup EndValue control if txtEndValue is entered
      ' TODO: Any key typed should be entered in UPPER CASE
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtSubValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSubValue.TextChanged
    Try
      If Not txtLookupSubValue Is Nothing Then mvSubValueValid = (txtLookupSubValue.ComboBox.SelectedIndex > 0) '  mvSubValueLookup.SetDesc(txtSubValue, cboSubValueDesc)
      'Set the ui (for the range checkbox)
      If Not mvSubValueValid Then
        txtSubValue.MaxLength = mvControlInfo.SubLength
        If txtSubValue.Text = "NULL" Or txtSubValue.Text = "NOTNULL" Or Strings.InStr(txtSubValue.Text, "*") > 0 Or InStr(txtSubValue.Text, "?") > 0 Then
          mvSubValueValid = True
        ElseIf txtSubValue.Text.Length > 1 And Strings.Left(txtSubValue.Text, 1) = "$" Then
          mvSubValueValid = True
        ElseIf txtLookupSubValue.Visible = False Then 'cboSubValueDesc.Visible = False Then
          mvSubValueValid = True
        End If
      End If
      SetVisibleControls()
      CheckCriteria()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtSubValue_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSubValue.KeyPress
    Try
      If txtLookupSubValue.Visible Then txtLookupSubValue.SetTextBoxString("") ' Blank out LookupSubValue control if txtSubValue is entered
      If Strings.Left(GetTextBoxSubValue().Text, 1) <> "$" Then '  txtSubValue
        ' TODO: KeyAscii ?
        'If KeyAscii = Asc("*") And Len(mvMainValidationItem.CaseConversion) > 0 Then
        If e.KeyChar = "*" AndAlso mvMainValidationItem.CaseConversion.Length > 0 Then
          '  'Leave
        Else
          If Strings.Mid(GetTextBoxSubValue().Text, 1, 2) = "NU" Then  ' txtSubValue
            'Assume that NULL is being entered
            If mvControlInfo.MainLength < 4 Then GetTextBoxSubValue().MaxLength = 4
          ElseIf Strings.Mid(GetTextBoxSubValue().Text, 1, 2) = "NO" Then
            'Assume that NOTNULL is being entered
            If mvControlInfo.MainLength < 7 Then txtSubValue.MaxLength = 7
          End If
        End If
        '
        'mvSubValidationItem.HandleKeyPress(KeyAscii)
        'End If
      End If
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtEndSubValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEndSubValue.TextChanged
    Try
      OnChangeTextLookupEndSubValue()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtEndSubValue_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEndSubValue.KeyPress
    Try
      If txtLookupEndSubValue.Visible Then txtLookupEndSubValue.SetTextBoxString("") ' Blank out Lookup EndSubValue control if txtEndSubValue is entered
      ' TODO: Any key typed should be entered in UPPER CASE
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub txtPeriodVar_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPeriodVar.TextChanged
    Try
      cmdAddPeriod.Enabled = (txtPeriodVar.Text.Length > 1 AndAlso Strings.Left(txtPeriodVar.Text, 1) = "$")
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub
#End Region

#Region "Private Methods"

  Private Sub InitialiseControls(ByVal pMailingSelection As MailingInfo, Optional ByVal pExclusions As Boolean = False)
    Try
      SetControlTheme()
      mvMailingSelection = pMailingSelection
      mvStandardExclusions = pExclusions
      mvAreaInfo = New SearchAreaInfo
      mvControlInfo = New SelectionControlInfo

      Me.Text = mvMailingSelection.Caption & " -  Edit Criteria"

      Dim vList As New ParameterList()
      vList("ValidationTable") = "search_areas"
      vList("ValidationAttribute") = "search_area"
      InitTextLookupBox(txtLookupArea, vList)

      ResetUI()
      cmdOK.Enabled = False

      dtpValue.CustomFormat = AppValues.DateFormat
      dtpEndValue.CustomFormat = AppValues.DateFormat
      dtpSubValue.CustomFormat = AppValues.DateFormat
      dtpEndSubValue.CustomFormat = AppValues.DateFormat
      dtpFrom.CustomFormat = AppValues.DateFormat
      dtpTo.CustomFormat = AppValues.DateFormat
      dtpValue.Checked = False
      dtpSubValue.Checked = False
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub InitTextLookupBox(ByVal pTxtLookup As TextLookupBox, ByVal pList As ParameterList) 'ByVal pTableName As String)
    Dim vParamList As New ParameterList(True)
    vParamList("TableName") = pList("ValidationTable")
    vParamList("FieldName") = pList("ValidationAttribute")
    vParamList("FieldType") = "C"  ' Character FieldType
    Dim vParams As ParameterList = DataHelper.GetMaintenanceData(vParamList)
    vParams("AttributeName") = pList("ValidationAttribute")
    vParams("ValidationAttribute") = pList("ValidationAttribute")
    vParams("ValidationTable") = pList("ValidationTable")
    If pList.Contains("RestrictionAttribute") Then vParams("RestrictionAttribute") = pList("RestrictionAttribute")

    'pTxtLookup = New TextLookupBox()
    pTxtLookup.BackColor = Me.BackColor
    Dim vPanelItem As PanelItem = New PanelItem(pTxtLookup, vParams("ValidationAttribute"))
    'vPanelItem.SetValidationData(pTableName, pAttributeName, True)

    vPanelItem.InitFromMaintenanceData(vParams)
    If vPanelItem.ValidationAttribute.ToLower = "relationship" Then vPanelItem.RemoveLookupRestriction = True 'Want to select all Relationship codes

    pTxtLookup.Tag = vPanelItem
    pTxtLookup.Name = vPanelItem.ParameterName
    pTxtLookup.NotEditPanel = True
    pTxtLookup.ComboBox.DataSource = Nothing
    pTxtLookup.Init(vPanelItem, False, False)
    pTxtLookup.TotalWidth = pTxtLookup.Width
    pTxtLookup.SetBounds(pTxtLookup.Location.X, pTxtLookup.Location.Y, 80, EditPanelInfo.DefaultHeight)
    pTxtLookup.PreventHistoricalSelection = False
  End Sub

  Private Sub GetInitialCodeRestrictionsHandler(ByVal sender As System.Object, ByVal pParameterName As System.String, ByRef pList As CDBNETCL.ParameterList) Handles txtLookupArea.GetInitialCodeRestrictions, txtLookupValue.GetInitialCodeRestrictions, txtLookupEndValue.GetInitialCodeRestrictions, txtLookupSubValue.GetInitialCodeRestrictions, txtLookupEndSubValue.GetInitialCodeRestrictions
    If pList Is Nothing Then pList = New ParameterList(True)
    If pParameterName = "SearchArea" Then
      pList("ApplicationName") = mvMailingSelection.MailingTypeCode
    ElseIf pParameterName = "GeographicalRegion" Then
      pList("GeographicalRegionType") = txtLookupValue.Text
    End If
  End Sub

  Private Sub LookupValidatingHandler(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtLookupArea.Validating, txtLookupValue.Validating, txtLookupEndValue.Validating, txtLookupSubValue.Validating, txtLookupEndSubValue.Validating
    Try
      LookupHandler(sender, e)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub LookupChangedHandler(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLookupArea.TextChanged, txtLookupValue.TextChanged, txtLookupEndValue.TextChanged, txtLookupSubValue.TextChanged, txtLookupEndSubValue.TextChanged
    Try
      LookupHandler(sender, e)
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub LookupHandler(ByVal sender As Object, ByVal e As Object)
    If TypeOf (sender) Is TextLookupBox Then
      Dim vTextLookupBox As TextLookupBox = DirectCast(sender, TextLookupBox)
      Dim vPanelItem As PanelItem = DirectCast(vTextLookupBox.Tag, PanelItem)
      If vTextLookupBox Is txtLookupArea Then
        OnChangeTextLookupArea()
      ElseIf vTextLookupBox Is txtLookupValue Then
        SetFocusOnLookupValueControl(False)
        OnChangeTextLookupValue()
      ElseIf vTextLookupBox Is txtLookupEndValue Then
        OnChangeTextLookupEndValue()
      ElseIf vTextLookupBox Is txtLookupSubValue Then
        SetFocusOnLookupSubValueControl(False)
        txtSubValue_TextChanged(sender, CType(e, EventArgs))
      ElseIf vTextLookupBox Is txtLookupEndSubValue Then
        OnChangeTextLookupEndSubValue()
      End If
    End If
  End Sub

  Private Sub OnChangeTextLookupValue()
    If txtLookupValue.Visible = True AndAlso txtLookupValue.ComboBox.Visible = True Then
      'mvSuppressComboClick = True
      mvValueNull = False
      mvValueValid = False
      mvValueVariable = False
      If txtValue.Text = "NULL" Or txtValue.Text = "NOTNULL" Or InStr(txtValue.Text, "*") > 0 Or InStr(txtValue.Text, "?") > 0 Then
        mvValueNull = True
        'cboValueDesc.ListIndex = -1
        'txtLookupValue.ComboBox.SelectedIndex = -1
        txtLookupValue.SetComboString("")
      ElseIf txtValue.Text.Length > 1 AndAlso Strings.Left(txtValue.Text, 1) = "$" Then
        mvValueVariable = True
        'cboValueDesc.ListIndex = -1
        'txtLookupValue.ComboBox.SelectedIndex = -1
        txtLookupValue.SetComboString("")
      Else
        txtValue.MaxLength = mvControlInfo.MainLength
        txtLookupValue.TextBox.MaxLength = mvControlInfo.MainLength
        If Not txtLookupValue Is Nothing Then mvValueValid = (txtLookupValue.ComboBox.SelectedIndex > 0) 'mvValueLookup.SetDesc(txtValue, cboValueDesc)
      End If
      'mvSuppressComboClick = False
    ElseIf txtLookupValue.Visible = True Then 'AndAlso txtLookupValue.Label.Visible = True Then
      mvValueNull = False
      mvValueValid = False
      mvValueVariable = False
      If txtValue.Text = "NULL" Or txtValue.Text = "NOTNULL" Or InStr(txtValue.Text, "*") > 0 Or InStr(txtValue.Text, "?") > 0 Then
        mvValueNull = True
        txtLookupValue.Label.Text = ""
        'lblValueDesc.Caption = gvNull
      ElseIf txtValue.Text.Length > 1 And Strings.Left(txtValue.Text, 1) = "$" Then
        mvValueVariable = True
        txtLookupValue.Label.Text = "" 'lblValueDesc.Caption = gvNull
      Else
        txtValue.MaxLength = mvControlInfo.MainLength
        txtLookupValue.TextBox.MaxLength = mvControlInfo.MainLength
        mvValueValid = (txtLookupValue.Description.Length > 0)
      End If
    Else
      'There is no combo and therefore no validation
      If Strings.Left(txtValue.Text, 1) = "$" Then
        mvValueValid = False
        mvValueVariable = (txtValue.Text.Length > 1)
      Else
        mvValueVariable = False
        mvValueValid = CBool(IIf(txtValue.Text.Trim.Length > 0, True, False))
      End If
    End If
    'Set the ui (for the range checkbox)
    SetVisibleControls()
    'If we have a valid value then set up the sub values (if any)
    If mvValueValid Then SetSubValues()
    CheckCriteria()
  End Sub
  Private Sub OnChangeTextLookupEndValue()
    If txtLookupEndValue.Visible Then
      mvEndValueValid = txtLookupEndValue.IsValid
    Else
      mvEndValueValid = txtEndValue.Text.Length > 0
    End If
    CheckCriteria()
  End Sub
  Private Sub OnChangeTextLookupEndSubValue()
    If txtLookupEndSubValue.Visible Then
      mvEndSubValueValid = (txtLookupEndSubValue.Text.Length > 0)
    ElseIf txtEndSubValue.Text.Length > 0 Then
      mvEndSubValueValid = True
    End If
    CheckCriteria()
  End Sub

  Private Sub OnChangeTextLookupArea()

    mvSearchAreaValid = False
    'If mvAreasLookup.SetDesc(txtArea, cboAreaDesc) Then
    If txtLookupArea.IsValid AndAlso txtLookupArea.Text.Length > 0 Then
      mvSearchAreaValid = True
      'We have a valid area so get the area information
      GetAreaInfo(txtLookupArea.Text)
      'Now set up the Person or Organisation option buttons
      Select Case mvAreaInfo.CorO
        Case "C"
          optPerson.Checked = True
          optOrganisation.Enabled = False
          optPerson.Enabled = True
        Case "O"
          optOrganisation.Checked = True
          optOrganisation.Enabled = True
          optPerson.Enabled = False
        Case "B"
          optOrganisation.Enabled = True
          optPerson.Enabled = True
      End Select
      'And read the selection control information
      GetSelectionControl()
      'Now set up the rest of the UI
      SetVisibleControls()
      'And read any main values
      SetMainValues()
      'Handle special cases
      SpecialAreas()
    Else
      'make sure that none of the controls are visible (in case search area has been removed after being selected)
      mvAreaInfo.MainValue = False
      mvControlInfo = Nothing
      SetVisibleControls()
    End If
    CheckCriteria()
    mvSubValuesInit = False
  End Sub

  Private Sub GetAreaInfo(ByVal pSearchArea As String)
    'Given a search area code retrieve the search area information

    If pSearchArea.Length > 0 Then
      'vRecordset = mvConn.GetRecordSet("SELECT * FROM search_areas WHERE search_area = '" & pSearchArea & "' AND application_name = '" & mvMailingSelection.MailingTypeCode & "'")
      Dim vList As New ParameterList(True)
      vList("SearchArea") = pSearchArea
      vList("ApplicationName") = mvMailingSelection.MailingTypeCode
      Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSearchAreas, vList)
      If vTable IsNot Nothing Then
        If vTable.Rows.Count = 1 Then
          mvAreaInfo.MainValue = CBool(IIf(vTable.Rows(0)("MainValue").ToString = "Y", True, False)) ' CBool(vTable.Rows(0)("MainValue")) ' vRecordset.Fields("main_value").Bool
          mvAreaInfo.SubValue = CBool(IIf(vTable.Rows(0)("SubsidiaryValue").ToString = "Y", True, False)) ' CBool(vTable.Rows(0)("SubsidiaryValue")) ' vRecordset.Fields("subsidiary_value").Bool
          mvAreaInfo.Period = CBool(IIf(vTable.Rows(0)("Period").ToString = "Y", True, False)) ' CBool(vTable.Rows(0)("Period")) ' vRecordset.Fields("period").Bool
          mvAreaInfo.CorO = vTable.Rows(0)("CO").ToString  ' vRecordset.Fields("c_o").Value
          Exit Sub
        End If
      End If
      ShowInformationMessage(Strings.Format(InformationMessages.ImSearchAreaMissing, pSearchArea))
    End If
  End Sub

  Private Sub SetProperValues(ByVal pList As ParameterList, ByVal pValidationTable As String, ByVal pValidationAttr As String)
    Select Case pValidationTable
      Case "geographical_regions"
        pList("ValidationTable") = "geographical_region_types"
      Case Else
        pList("ValidationTable") = pValidationTable
    End Select

    Select Case pValidationAttr
      Case Else
        pList("ValidationAttribute") = pValidationAttr
    End Select
  End Sub

  Private Sub SetMainValues()
    'Set the value and sub value combos according to type
    Dim vRestriction As String = String.Empty
    Dim vList As New ParameterList()

    chkValueRange.Checked = False
    'Force a change event to occur
    'txtLookupValue.Text = " "
    'txtValue = gvNull
    'txtEndValue = gvNull
    txtLookupValue.Text = ""
    txtLookupEndValue.Text = ""
    txtValue.Text = ""
    lstValue.Items.Clear()
    cmdDeleteValue.Enabled = False

    If mvAreaInfo.MainValue And mvAreaInfo.MainValidate = True Then
      If mvControlInfo.Pattern.Length > 0 Then
        vList("ValidationTable") = mvControlInfo.TableName
        vList("ValidationAttribute") = mvControlInfo.MainAttr
        vList("Pattern") = mvControlInfo.Pattern
      Else
        If mvMailingSelection.MailingType = CareNetServices.MailingTypes.mtyNonMemberFulfilment And mvControlInfo.MainAttr = "reason_for_despatch" Then
          ' TODO: require testing
          vRestriction = "reason_for_despatch NOT IN (SELECT membership_type FROM membership_types)"
        End If
        With mvControlInfo
          'vList("AttributeName") = .MainAttr
          'vList("AttributeNameDesc") = .MainValueHeading
          'vList.IntegerValue("EntryLength") = .MainLength
          'vList("Pattern") = .Pattern
          'vList("ValidationTable") = .ValidationTable
          'vList("ValidationAttribute") = .MainValidationAttr
          SetProperValues(vList, .ValidationTable, .MainValidationAttr)
          vList("RestrictionAttribute") = vRestriction
        End With
      End If
      If txtLookupValue.Visible Then
        InitTextLookupBox(txtLookupValue, vList) ' mvControlInfo.MainValidationAttr, mvControlInfo.ValidationTable)
        InitTextLookupBox(txtLookupEndValue, vList)
      End If
    End If

    chkSubValueRange.Checked = False
    txtSubValue.Text = ""
    txtLookupEndSubValue.Text = ""
    txtEndSubValue.Text = ""
    txtLookupEndSubValue.Text = ""
    lstSubValue.Items.Clear()
    cmdDeleteSubValue.Enabled = False
    'vse2.Refresh()
  End Sub

  Private Sub SpecialAreas()
    'Handle the special areas of owner and co-owner
    If txtLookupArea.Text = "OWNER" Or txtLookupArea.Text = "CO-OWNER" Then
      lstValue.Items.Add(DataHelper.UserInfo.Department)  '.AddItem(txtValue)
      lstValue.SelectedIndex = 0
      lstValue_Click(lstValue, Nothing)
      If txtLookupValue.Visible Then txtLookupValue.Text = DataHelper.UserInfo.Department
    End If
  End Sub


  Private Sub CheckCriteria()
    'Determine if there is a valid criteria set
    'This will enable the OK command button etc.
    Dim vValid As Boolean
    Dim vValueValid As Boolean
    Dim vSubValueValid As Boolean

    If mvSearchAreaValid Then
      vValid = True
      If mvAreaInfo.MainValue Then
        If lstValue.Items.Count = 0 Then vValid = False
        If mvValueValid Then
          If chkValueRange.Checked = True Then
            If mvEndValueValid Then vValueValid = True
          Else
            'No range
            vValueValid = True
            If mvAreaInfo.SubValue Then
              '            If lstSubValue.ListCount = 0 Then vValid = False
              If mvSubValueValid Then
                If chkSubValueRange.Checked = True Then
                  If mvEndSubValueValid Then vSubValueValid = True
                Else
                  'No range
                  'If txtSubValue.Text.Length > 0 Then vSubValueValid = True
                  If GetTextBoxSubValue().Text.Length > 0 Then vSubValueValid = True
                End If
              End If
            Else
              'No sub value
            End If
          End If
        Else                      'Main value is not valid
          If mvValueNull Or mvValueVariable Then vValueValid = True 'Is it NULL or a variable
        End If
      Else
        'No main value
        If mvAreaInfo.Period Then
          If lstPeriod.Items.Count = 0 Then vValid = False
        Else
          vValid = False
        End If
      End If
    End If
    cmdAddValue.Enabled = vValueValid
    If dtpSubValue.Visible Then
      cmdAddSubValue.Enabled = True
    Else
      cmdAddSubValue.Enabled = vSubValueValid
    End If
    cmdOK.Enabled = vValid
  End Sub

  Private Sub SetBlankForAlternateTextbox(ByVal pTextBox As TextBox, ByVal pControl As String)
    Select Case pControl.ToLower()
      Case "value"
        If pTextBox.Name = txtValue.Name Then txtLookupValue.Text = "" Else txtValue.Text = ""
      Case "endvalue"
        If pTextBox.Name = txtEndValue.Name Then txtLookupEndValue.Text = "" Else txtEndValue.Text = ""
      Case "subvalue"
        If pTextBox.Name = txtSubValue.Name Then txtLookupSubValue.Text = "" Else txtSubValue.Text = ""
      Case "endsubvalue"
        If pTextBox.Name = txtEndSubValue.Name Then txtLookupEndSubValue.Text = "" Else txtEndSubValue.Text = ""
    End Select
  End Sub

  Private Function GetTextBoxValue() As TextBox
    If txtLookupValue.Visible = True AndAlso (txtLookupValue.ComboBox.SelectedIndex > 0 OrElse txtLookupValue.Description().Trim.Length > 0) Then
      Return txtLookupValue.TextBox
    Else
      Return txtValue
    End If
  End Function

  Private Function GetTextBoxEndValue() As TextBox
    If txtLookupEndValue.Visible = True AndAlso (txtLookupEndValue.ComboBox.SelectedIndex > 0 OrElse txtLookupEndValue.Description().Trim.Length > 0) Then
      Return txtLookupEndValue.TextBox
    Else
      Return txtEndValue
    End If
  End Function

  Private Function GetTextBoxSubValue() As TextBox
    If txtLookupSubValue.Visible = True AndAlso (txtLookupSubValue.ComboBox.SelectedIndex > 0 OrElse txtLookupSubValue.Description().Trim.Length > 0) Then
      Return txtLookupSubValue.TextBox
    Else
      Return txtSubValue
    End If
  End Function

  Private Function GetTextBoxEndSubValue() As TextBox
    If txtLookupEndSubValue.Visible = True AndAlso (txtLookupEndSubValue.ComboBox.SelectedIndex > 0 OrElse txtLookupEndSubValue.Description().Trim.Length > 0) Then
      Return txtLookupEndSubValue.TextBox
    Else
      Return txtEndSubValue
    End If
  End Function

  Private Function GetTextBoxFromControlValue(ByVal pControlName As String, ByVal pTextLookupBox As TextLookupBox, ByVal pSearchText As String) As TextBox
    If pControlName.Trim() <> NOT_TEXTBOX_LOOKUP Then
      Dim vObject As Object = DirectCast(FindControl(Me, pControlName), Object)
      If CType(vObject, TextBox) IsNot Nothing Then
        If vObject Is txtValue Then Return txtValue
        If vObject Is txtEndValue Then Return txtEndValue
        If vObject Is txtSubValue Then Return txtSubValue
        If vObject Is txtEndSubValue Then Return txtEndSubValue
      ElseIf CType(vObject, TextLookupBox) IsNot Nothing Then
        If vObject Is txtLookupValue Then Return txtLookupValue.TextBox
        If vObject Is txtLookupEndValue Then Return txtLookupEndValue.TextBox
        If vObject Is txtLookupSubValue Then Return txtLookupSubValue.TextBox
        If vObject Is txtLookupEndSubValue Then Return txtLookupEndSubValue.TextBox
      End If
    Else
      ' Check If pControlName = "NotTextBoxOrLookupControl", 
      ' then try to find the selected listbox value is present in adjacent lookup control
      If pControlName.Trim.Length > 0 Then
        If pControlName.Trim() = NOT_TEXTBOX_LOOKUP Then '"NotTextBoxOrLookupControl"
          'If pTextLookupBox.Visible = True AndAlso pSearchText.Length > 0 Then
          'pTextLookupBox.Text = pSearchText
          SelectComboBoxItem(pTextLookupBox.ComboBox, pSearchText, True)
          ' After setting this, if value is auto populated into combobox or label that means the control you are finding is lookup control
          If pTextLookupBox.ComboBox.SelectedIndex > 0 OrElse pTextLookupBox.Label.Text.Trim.Length > 0 Then Return pTextLookupBox.TextBox
          'End If
        End If
      End If
    End If
    Return Nothing
  End Function

  Private Sub SetSubValues()
    'Set the subsidiary value combo according to type
    Dim vRestriction As String

    chkSubValueRange.Checked = False
    GetTextBoxSubValue().Text = ""
    GetTextBoxEndSubValue().Text = ""
    lstSubValue.Items.Clear()

    'It is valid to have a blank sub-value
    cmdAddSubValue.Enabled = dtpSubValue.Visible  'False
    If mvAreaInfo.SubValue AndAlso ((mvAreaInfo.MainValidate = True) OrElse (mvControlInfo.MainValidationAttr = "event_number")) AndAlso mvValueValid Then
      If mvControlInfo.MainDataType = "I" Or mvControlInfo.MainDataType = "L" Then
        'vRestriction = mvControlInfo.MainValidationAttr & " = " & GetTextBoxValue().Text
        vRestriction = txtLookupValue.Name
      Else
        'vRestriction = mvControlInfo.MainValidationAttr & " = '" & GetTextBoxValue().Text & "'"
        vRestriction = txtLookupValue.Name
      End If

      If Not mvSubValuesInit Then
        Dim vList As New ParameterList()
        With mvControlInfo
          vList("AttributeName") = .SubAttr
          'vList("AttributeNameDesc") = .MainValueHeading
          'vList.IntegerValue("EntryLength") = .MainLength
          vList("Pattern") = .Pattern
          'vList("ValidationTable") = .ValidationTable
          'vList("ValidationAttribute") = .SubValidationAttr
          SetProperValues(vList, .ValidationTable, .SubAttr)
          'vList("RestrictionAttribute") = vRestriction

          If .SubValidationAttr.Length > 0 Then
            InitTextLookupBox(txtLookupSubValue, vList)
            InitTextLookupBox(txtLookupEndSubValue, vList)
            mvSubValuesInit = True
            SelectComboBoxItem(txtLookupValue.ComboBox, txtLookupValue.Text, True)
          End If
        End With
      End If
      If mvControlInfo.SubAttr = "geographical_region" Then
        txtLookupSubValue.FillComboWithRestriction(txtLookupValue.Text)
        txtLookupEndSubValue.FillComboWithRestriction(txtLookupValue.Text)
      End If

    End If
  End Sub

  Private Sub ResetUI()
    'Reset the mvAreaInfo structure to reflect no specified area
    mvAreaInfo.MainValue = False
    mvAreaInfo.SubValue = False
    mvAreaInfo.CorO = "B"
    mvAreaInfo.Period = False
    'Assume that validation will be required
    mvAreaInfo.MainValidate = True
    mvAreaInfo.SubValidate = True

    'Clear the selected search area
    txtLookupArea.Text = ""
    txtLookupArea.SetComboString("")
    lblValue.Text = ""
    lblSubValue.Text = ""

    dtpFrom.Checked = False
    dtpTo.Checked = False
    cmdValueVar.Visible = True
    cmdSubValueVar.Visible = True
    cmdPeriodVar.Visible = True

    If mvStandardExclusions Then
      optInclude.Enabled = False
      optInclude.Checked = False
      optExclude.Checked = True
    End If
    SetVisibleControls()
  End Sub

  Private Sub SetVisibleControls()
    'Set up the UI of the form to reflect the selected area
    'and current state. This will make controls visible or not
    Dim vState As Boolean
    Dim vRangestate As Boolean
    Dim vValidate As Boolean

    vRangestate = False
    If mvAreaInfo.MainValue Then
      vState = True
      vValidate = mvAreaInfo.MainValidate
      If mvValueValid Then vRangestate = True
    Else
      vState = False
      vValidate = False
    End If
    txtValue.Visible = vState
    lblValue.Visible = vState
    If Strings.Left(txtValue.Text, 1) <> "$" Then txtValue.MaxLength = mvControlInfo.MainLength
    cmdValueVar.Visible = vState
    lstValue.Visible = vState
    cmdAddValue.Visible = vState
    cmdDeleteValue.Visible = vState
    'If no validation is required then don't show the combo or finder
    txtLookupValue.Visible = vValidate
    'Show the range checkbox if a value has been entered
    chkValueRange.Visible = vRangestate
    If Not chkValueRange.Visible Then chkValueRange.Checked = False
    'Show the end value dependant on the range checkbox
    If vRangestate = True Then vRangestate = (chkValueRange.Checked = True)
    txtEndValue.Visible = vRangestate
    txtEndValue.MaxLength = txtValue.MaxLength
    If vRangestate AndAlso txtValue.Visible AndAlso txtValue.Text.Length > 0 Then txtEndValue.Visible = True
    'If no validation is required then don't show the end combo or finder
    If vValidate = False Then vRangestate = False
    txtLookupEndValue.Visible = vRangestate
    If vRangestate AndAlso txtLookupValue.Visible AndAlso txtLookupValue.TextBox.Text.Length > 0 Then txtLookupEndValue.Visible = True
    chkYN.Visible = False

    If mvControlInfo.MainDataType = "D" AndAlso Strings.Left(txtValue.Text, 1) <> "$" Then
      txtValue.Visible = False
      txtLookupValue.Visible = False
      dtpValue.Visible = True
      cmdAddValue.Enabled = True
      mvValueValid = dtpValue.Checked ' IsNull(dtpValue.Value)
      mvValueNull = Not mvValueValid
      chkValueRange.Visible = mvValueValid
      txtEndValue.Visible = False
      txtLookupEndValue.Visible = False
      dtpEndValue.Visible = (chkValueRange.Checked = True)
      If dtpEndValue.Visible Then
        mvEndValueValid = dtpEndValue.Checked ' IsNull(dtpEndValue.Value)
      Else
        dtpEndValue.Checked = False ' dtpEndValue.Value = Null
      End If
      dtpValue.Enabled = Not mvNullMainDTPValue
      cmdValueVar.Enabled = Not mvNullMainDTPValue
    Else
      dtpValue.Visible = False
      dtpEndValue.Visible = False
      If mvControlInfo.Pattern = "[YN]" Then
        chkYN.Visible = True
        lblValue.Visible = False
        mvValueValid = True
        txtValue.Visible = False
        txtLookupValue.Visible = False
        chkValueRange.Visible = False
        txtEndValue.Visible = False
        txtLookupEndValue.Visible = False
        cmdValueVar.Visible = False
      End If
    End If

    If vState AndAlso (mvControlInfo.MainValidationAttr = "source") AndAlso mvIsSourceLookupValueLoaded = False Then
      Dim vList As New ParameterList()
      vList("ValidationTable") = "sources"
      vList("ValidationAttribute") = mvControlInfo.MainValidationAttr
      InitTextLookupBox(txtLookupValue, vList)
      mvIsSourceLookupValueLoaded = True
    ElseIf mvControlInfo.MainValidationAttr <> "source" Then
      mvIsSourceLookupValueLoaded = False
    End If
    If vRangestate AndAlso (mvControlInfo.MainValidationAttr = "source") AndAlso mvIsSourceLookupEndValueLoaded = False Then
      Dim vList As New ParameterList()
      vList("ValidationTable") = "sources"
      vList("ValidationAttribute") = mvControlInfo.MainValidationAttr
      InitTextLookupBox(txtLookupEndValue, vList)
      mvIsSourceLookupEndValueLoaded = True
    ElseIf mvControlInfo.MainValidationAttr <> "source" Then
      mvIsSourceLookupEndValueLoaded = False
    End If

    If mvControlInfo.MainValidationAttr = "event_number" Then
      txtLookupValue.Visible = True
      If txtEndValue.Visible Then txtLookupEndValue.Visible = True 'cmdFindEndValue.Visible = True
    End If

    vRangestate = False
    If mvAreaInfo.SubValue AndAlso chkValueRange.Checked = False AndAlso mvValueValid AndAlso lstValue.Items.Count < 2 Then
      vState = True
      vValidate = mvAreaInfo.SubValidate
      If mvSubValueValid Then vRangestate = True
    Else
      vState = False
      vValidate = False
    End If

    txtSubValue.Visible = vState
    If Strings.Left(txtSubValue.Text, 1) <> "$" Then txtSubValue.MaxLength = mvControlInfo.SubLength
    cmdSubValueVar.Visible = vState
    lstSubValue.Visible = vState
    cmdAddSubValue.Visible = vState
    cmdDeleteSubValue.Visible = vState
    'If no validation is required then don't show the combo or finder
    txtLookupSubValue.Visible = vValidate
    'Show the range checkbox if a value has been entered
    chkSubValueRange.Visible = vRangestate
    'Show the end value dependant on the range checkbox
    If vRangestate = True Then vRangestate = (chkSubValueRange.Checked = True)
    txtEndSubValue.Visible = vRangestate
    txtEndSubValue.MaxLength = txtSubValue.MaxLength
    'If no validation is required then don't show the end combo or finder
    If vValidate = False Then vRangestate = False
    'If Not txtEndSubValue.Visible Then txtLookupEndSubValue.Visible = vRangestate Else txtLookupEndSubValue.Visible = False
    txtLookupEndSubValue.Visible = vRangestate
    cmdAddSubValue.Enabled = vState

    If mvControlInfo.SubDataType = "D" AndAlso Strings.Left(txtValue.Text, 1) <> "$" Then
      txtSubValue.Visible = False
      txtLookupSubValue.Visible = False
      dtpSubValue.Visible = mvValueValid
      cmdAddSubValue.Enabled = True
      mvSubValueValid = dtpSubValue.Checked ' IsNull(dtpSubValue.Value)
      mvSubValueNull = Not mvSubValueNull
      chkSubValueRange.Visible = mvSubValueValid
      txtEndSubValue.Visible = False
      'txtEndSubValue.Visible = False
      txtLookupEndSubValue.Visible = False
      dtpEndSubValue.Visible = chkSubValueRange.Checked = True
      If dtpEndSubValue.Visible Then
        mvEndSubValueValid = dtpEndSubValue.Checked ' IsNull(dtpEndSubValue.Value)
      Else
        dtpEndSubValue.Checked = False
      End If
      dtpSubValue.Enabled = Not mvNullSubDTPValue
      cmdSubValueVar.Enabled = Not mvNullSubDTPValue
    Else
      dtpSubValue.Visible = False
      dtpEndSubValue.Visible = False
    End If

    If mvAreaInfo.Period Then
      vState = True
    Else
      vState = False
    End If

    txtPeriodVar.Visible = False
    lblPeriodVar.Visible = False
    dtpFrom.Visible = vState
    lblPeriodFrom.Visible = vState
    cmdPeriodVar.Visible = vState
    dtpTo.Visible = vState
    lblPeriodTo.Visible = vState
    lstPeriod.Visible = vState
    cmdAddPeriod.Visible = vState
    cmdDeletePeriod.Visible = vState
    If vState Then
      dtpFrom.Value = Date.Today ' TodaysDate()
      cmdDeletePeriod.Enabled = False
    End If

    lblSubValue.Visible = (txtSubValue.Visible OrElse txtLookupSubValue.Visible OrElse dtpSubValue.Visible)

  End Sub

  Private Sub SafeSetFocus(ByVal pControl As Control)
    If pControl.Enabled = True And pControl.Visible = True Then pControl.Focus()
  End Sub

  Private Sub GetSelectionControl()
    'Get the selection control information relevant to a specified search area and type
    Dim vCO As String
    Dim vCase As String
    Dim vAttrs As String = String.Empty
    Dim vList As New ParameterList(True)

    'Ignore null codes
    If txtLookupArea.Text.Trim() = String.Empty Then Exit Sub
    vCO = CStr(IIf((optPerson.Checked = True), "C", "O"))

    vList("SearchArea") = txtLookupArea.Text
    vList("ApplicationName") = mvMailingSelection.MailingTypeCode
    vList("CO") = vCO
    Dim vTable As DataTable = DataHelper.GetLookupData(CareNetServices.XMLLookupDataTypes.xldtSelectionControls, vList)

    If vTable IsNot Nothing Then
      If vTable.Rows.Count = 1 Then
        With mvControlInfo
          .MainAttr = vTable.Rows(0)("MainAttribute").ToString  'vRecordset.Fields("main_attribute").Value
          .MainValueHeading = vTable.Rows(0)("MainValueHeading").ToString  'vRecordset.Fields("main_value_heading").Value
          .SubAttr = vTable.Rows(0)("SubsidiaryAttribute").ToString  'vRecordset.Fields("subsidiary_attribute").Value
          .SubValueHeading = vTable.Rows(0)("SubsidiaryValueHeading").ToString  'vRecordset.Fields("subsidiary_value_heading").Value
          .ValidationTable = vTable.Rows(0)("ValidationTable").ToString  'vRecordset.Fields("validation_table").Value
          .MainValidationAttr = vTable.Rows(0)("MainValidationAttribute").ToString  'vRecordset.Fields("main_validation_attribute").Value
          Select Case .MainValidationAttr
            Case "contact_number", "organisation_number", "address_number", "batch_number", "mailing_number", "sort_code"
              .MainValidationAttr = ""
          End Select
          .SubValidationAttr = vTable.Rows(0)("SubsidiaryValidationAttribut").ToString  'vRecordset.Fields(mvConn.DBAttrName("subsidiary_validation_attribute")).Value
          Select Case .SubValidationAttr
            Case "contact_number", "organisation_number", "address_number", "batch_number", "mailing_number", "sort_code"
              .SubValidationAttr = ""
          End Select
          .MainDataType = vTable.Rows(0)("MainDataType").ToString  'vRecordset.Fields("main_data_type").Value
          .SubDataType = vTable.Rows(0)("SubsidiaryDataType").ToString  'vRecordset.Fields("subsidiary_data_type").Value

          .MainLength = CInt(IIf(vTable.Rows(0)("MainLength").ToString.Length > 0, vTable.Rows(0)("MainLength").ToString, 0))  'vRecordset.Fields("main_length").LongValue
          .SubLength = CInt(IIf(vTable.Rows(0)("SubsidiaryLength").ToString.Length > 0, vTable.Rows(0)("SubsidiaryLength").ToString, 0))  'vRecordset.Fields("subsidiary_length").LongValue
          .TableName = vTable.Rows(0)("TableName").ToString  'vRecordset.Fields("table_name").Value
          .Pattern = vTable.Rows(0)("Pattern").ToString  'vRecordset.Fields("pattern").Value

          txtValue.Tag = .MainValueHeading
          lblValue.Text = .MainValueHeading
          lblValue.Visible = True
          dtpValue.Tag = .MainValueHeading
          chkYN.Text = .MainValueHeading
          txtSubValue.Tag = .SubValueHeading
          dtpSubValue.Tag = .SubValueHeading
          lblSubValue.Text = .SubValueHeading
          lblSubValue.Visible = True
          mvAreaInfo.MainValidate = CBool(IIf((.ValidationTable <> "") And (.MainValidationAttr <> ""), True, False))
          If Not mvAreaInfo.MainValidate Then
            mvAreaInfo.MainValidate = .Pattern.Length > 0
          End If
          'If .MainValidationAttr = "event_number" Then mvAreaInfo.MainValidate = False

          mvAreaInfo.SubValidate = CBool(IIf((.ValidationTable <> "") AndAlso (.SubValidationAttr <> ""), True, False))
          mvLastType = vCO

          vCase = CStr(IIf(.MainDataType = "C" And .ValidationTable <> "", "U", ""))
          mvMainValidationItem.InitFromParameters(.MainLength, vCase)
          vCase = CStr(IIf(.SubDataType = "C" And .ValidationTable <> "", "U", ""))
          mvSubValidationItem.InitFromParameters(.SubLength, vCase)
        End With
      End If
    End If
  End Sub

  Private Sub SetFromCurrentCriteria()
    With mvMailingSelection.CurrentCriteria
      If .CO = "C" Then
        optPerson.Checked = True
      Else
        optOrganisation.Checked = True
      End If
      txtLookupArea.Text = .SearchArea
      lstValue.Focus()
      If .IE = "I" Then
        optInclude.Checked = True
      Else
        optExclude.Checked = True
      End If
      FillListFromString(lstValue, .MainValue)
      lstValue.SelectedIndex = CInt(IIf(lstValue.Items.Count > 0, 0, -1))
      If lstValue.SelectedIndex >= 0 Then lstValue_Click(lstValue, Nothing)
      'LookupChangedHandler(txtLookupValue, Nothing)
      FillListFromString(lstSubValue, .SubsidiaryValue)
      lstSubValue.SelectedIndex = CInt(IIf(lstSubValue.Items.Count > 0, 0, -1))
      If lstSubValue.SelectedIndex >= 0 Then lstSubValue_Click(lstSubValue, Nothing)
      dtpFrom.Checked = False
      dtpTo.Checked = False
      FillListFromString(lstPeriod, .Period)
      lstPeriod.SelectedIndex = CInt(IIf(lstPeriod.Items.Count > 0, 0, -1))
      If lstPeriod.SelectedIndex >= 0 Then lstPeriod_Click(lstPeriod, Nothing)
    End With
  End Sub

  Private Sub FillListFromString(ByVal pList As ListBox, ByVal pSource As String)
    'Fill a list box by parsing out items from the string
    'The items in the string should be separated by CHR$(10)
    Dim vPos As Integer = 0

    pList.Items.Clear()
    If pSource <> "" Then
      Dim vTempArray() As String
      Dim vCount As Integer
      vTempArray = Split(pSource, Environment.NewLine)  ' Chr$(10)
      For vCount = 0 To UBound(vTempArray)
        pList.Items.Add(vTempArray(vCount)) '  Mid$(pSource, vPos)
      Next
    End If
  End Sub

#End Region

#Region "DateTimePicker Events"

  Private Sub SetFieldsFromDTP(ByVal pField As DTPFields)
    Select Case pField
      Case DTPFields.dtpMainFrom
        If dtpValue.Checked = False Then
          chkValueRange.Checked = False
          mvValueNull = True
        Else
          'chkValueRange.Checked = True   'NFPNGEX70 The range value checkbox should only be checked by the user as this will disable the Add button
          mvValueNull = False
        End If
        SetVisibleControls()
        CheckCriteria()
      Case DTPFields.dtpMainTo
        cmdAddValue.Enabled = (dtpEndValue.Checked = True) 'IsNull(dtpEndValue.Value)
      Case DTPFields.dtpPeriodFrom
        cmdAddPeriod.Enabled = (dtpFrom.Checked = True) 'IsNull(dtpFrom.Value)
      Case DTPFields.dtpPeriodTo
        '
      Case DTPFields.dtpSubFrom
        If dtpSubValue.Checked = False Then   ' IsNull(dtpSubValue.Value)
          chkSubValueRange.Checked = False
          mvSubValueNull = True
        Else
          chkSubValueRange.Checked = True
          mvSubValueNull = False
        End If
        SetVisibleControls()
      Case DTPFields.dtpSubTo
        cmdAddSubValue.Enabled = (dtpEndSubValue.Checked = True) ' dtpEndSubValue.Checked
    End Select
  End Sub

  Private Sub dtpValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpValue.TextChanged
    SetFieldsFromDTP(DTPFields.dtpMainFrom)
  End Sub

  Private Sub dtpValue_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpValue.DropDown
    SetFieldsFromDTP(DTPFields.dtpMainFrom)
  End Sub

  Private Sub dtpEndValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpEndValue.TextChanged
    SetFieldsFromDTP(DTPFields.dtpMainTo)
  End Sub

  Private Sub dtpEndValue_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpEndValue.DropDown
    SetFieldsFromDTP(DTPFields.dtpMainTo)
  End Sub

  Private Sub dtpSubValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpSubValue.TextChanged
    SetFieldsFromDTP(DTPFields.dtpSubFrom)
  End Sub

  Private Sub dtpSubValue_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpSubValue.DropDown
    SetFieldsFromDTP(DTPFields.dtpSubFrom)
  End Sub

  Private Sub dtpEndSubValue_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpEndSubValue.TextChanged
    SetFieldsFromDTP(DTPFields.dtpSubTo)
  End Sub

  Private Sub dtpEndSubValue_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpEndSubValue.DropDown
    SetFieldsFromDTP(DTPFields.dtpSubTo)
  End Sub

  Private Sub dtpFrom_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.TextChanged
    SetFieldsFromDTP(DTPFields.dtpPeriodFrom)
  End Sub


  Private Sub dtpFrom_DropDown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtpFrom.DropDown
    SetFieldsFromDTP(DTPFields.dtpPeriodFrom)
  End Sub
#End Region

End Class


Public Class ValidationItem

  Private mvEntryLength As Integer
  Private mvCaseConversion As String

  Public ReadOnly Property EntryLength() As Integer
    Get
      Return mvEntryLength
    End Get
  End Property


  Public ReadOnly Property CaseConversion() As String
    Get
      Return mvCaseConversion
    End Get
  End Property

  Public Sub InitFromParameters(ByVal pEntryLength As Integer, ByVal pCaseConversion As String)
    mvEntryLength = pEntryLength
    mvCaseConversion = pCaseConversion
  End Sub
End Class
