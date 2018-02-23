Public Class frmDateFormat

#Region "Private Members"

  Private mvDayMonth As String = String.Empty
  Private mvYear As String = String.Empty
  Private mvSep As String = String.Empty
  Private mvYearFirst As Boolean
  Private mvChosenDateFormat As String = String.Empty
  Private mvExpiryDate As Boolean

#End Region

#Region "Constructor"

  Public Sub New(ByVal pDateFormat As String, ByVal pExpiryDate As Boolean)
    ' This call is required by the Windows Form Designer.
    InitializeComponent()

    mvChosenDateFormat = pDateFormat
    mvExpiryDate = pExpiryDate
    InitialiseControls()
  End Sub

#End Region

#Region "Public Properties"

  Public ReadOnly Property DateFormat() As String
    Get
      Return mvChosenDateFormat
    End Get
  End Property

#End Region

#Region "Control Events"

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    mvChosenDateFormat = lblDateFormat.Text.Trim
    Me.Close()
  End Sub

  Private Sub optDefFormat_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDefFormat.CheckedChanged
    SetDefineFormat(True)
    BuildFormatString()
  End Sub

  Private Sub optAS400Format_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optAS400Format.CheckedChanged
    SetDefineFormat(False)
    lblDateFormat.Text = ControlText.LblDateFormat
  End Sub

  Private Sub optYear_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optYearYY.CheckedChanged, optYearYYYY.CheckedChanged
    If optYearYY.Checked Then mvYear = "yy" Else mvYear = "yyyy"
    BuildFormatString()
  End Sub

  Private Sub optMY_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMY.CheckedChanged, optYM.CheckedChanged
    If optMY.Checked Then
      mvDayMonth = "MM" & mvSep
      mvYearFirst = False
    Else
      mvDayMonth = mvSep & "MM"
      mvYearFirst = True
    End If
    BuildFormatString()
  End Sub

  Private Sub optOrder_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optOrderDMY.CheckedChanged, optOrderMDY.CheckedChanged, optOrderYMD.CheckedChanged
    If optOrderDMY.Checked Then
      mvDayMonth = "dd" & mvSep & "MM" & mvSep
      mvYearFirst = False
    ElseIf optOrderMDY.Checked Then
      mvDayMonth = "MM" & mvSep & "dd" & mvSep
      mvYearFirst = False
    Else
      mvDayMonth = mvSep & "MM" & mvSep & "dd"
      mvYearFirst = True
    End If
    BuildFormatString()
  End Sub

  Private Sub txtSeperator_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSeperator.TextChanged
    If mvSep.Length = 0 Then
      If mvYearFirst Then
        If mvExpiryDate Then
          mvDayMonth = txtSeperator.Text & mvDayMonth
        Else
          mvDayMonth = txtSeperator.Text & mvDayMonth.Substring(0, 2) & txtSeperator.Text & Mid(mvDayMonth, 3, 2)
        End If
      Else
        If mvExpiryDate Then
          mvDayMonth = mvDayMonth & txtSeperator.Text
        Else
          mvDayMonth = mvDayMonth.Substring(0, 2) & txtSeperator.Text & Mid(mvDayMonth, 3, 2) & txtSeperator.Text
        End If
      End If
    Else
      mvDayMonth = mvDayMonth.Replace(mvSep, txtSeperator.Text)
    End If

    BuildFormatString()
    mvSep = txtSeperator.Text
  End Sub

#End Region

#Region "Private Methods"

  Private Sub BuildFormatString()
    If mvYearFirst Then
      lblDateFormat.Text = mvYear & mvDayMonth
    Else
      lblDateFormat.Text = mvDayMonth & mvYear
    End If
  End Sub

  Private Sub InitialiseControls()
    SetControlTheme()

    Dim vOrder As String
    vOrder = mvChosenDateFormat.Substring(0, 1)
    If mvExpiryDate Then
      grpMMYY.Visible = True
      grpOrder.Visible = False
      If mvChosenDateFormat.Contains("yyyy") Then
        optYearYYYY.Checked = True
        If mvChosenDateFormat.Length > 6 Then
          If vOrder = "M" Then
            mvSep = Mid(mvChosenDateFormat, 3, 1)
          Else
            mvSep = Mid(mvChosenDateFormat, 5, 1)
          End If
        Else
          mvSep = String.Empty
        End If
      Else
        optYearYY.Checked = True
        If mvChosenDateFormat.Length > 4 Then
          mvSep = Mid(mvChosenDateFormat, 3, 1)
        Else
          mvSep = String.Empty
        End If
      End If
      If vOrder = "M" Then
        optMY.Checked = True
      Else
        optYM.Checked = True
      End If
    Else
      grpMMYY.Visible = False
      grpOrder.Visible = True
      If mvChosenDateFormat = "cyymmdd" Then
        optAS400Format.Checked = True
      Else
        optDefFormat.Checked = True
        If mvChosenDateFormat.Contains("yyyy") Then
          optYearYYYY.Checked = True
          If mvChosenDateFormat.Length > 8 Then
            If vOrder = "d" OrElse vOrder = "M" Then
              mvSep = Mid(mvChosenDateFormat, 3, 1)
            Else
              mvSep = Mid(mvChosenDateFormat, 5, 1)
            End If
          Else
            mvSep = String.Empty
          End If
        Else
          optYearYY.Checked = True
          If mvChosenDateFormat.Length > 6 Then
            mvSep = Mid(mvChosenDateFormat, 3, 1)
          Else
            mvSep = String.Empty
          End If
        End If

        If vOrder = "d" Then
          optOrderDMY.Checked = True
        ElseIf vOrder = "M" Then
          optOrderMDY.Checked = True
        Else
          optOrderYMD.Checked = True
        End If
      End If
    End If
    'Replace the default seperator with the custom one
    If txtSeperator.Text <> mvSep Then mvDayMonth = mvDayMonth.Replace(txtSeperator.Text, mvSep)
    txtSeperator.Text = mvSep
    lblDateFormat.Text = mvChosenDateFormat
  End Sub

  Private Sub SetDefineFormat(ByVal pValue As Boolean)
    grpOrder.Enabled = pValue
    grpMMYY.Enabled = pValue
    grpYear.Enabled = pValue
  End Sub

#End Region

End Class