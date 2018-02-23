Public Class DialogService
  Implements IDialogService

  Public Function FindDataDialog(pType As CareNetServices.XMLDataFinderTypes, pParams As ParameterList, Optional pOwner As Form = Nothing, Optional pAllContactGroups As Boolean = False, Optional pNewContactAtOrg As Boolean = False) As Integer Implements IDialogService.FindDataDialog
    Return AppHelper.ShowModalFinder(pType, pParams, pOwner, pAllContactGroups, pNewContactAtOrg)
  End Function

  Public Function QueryDataDialog(pDataContext As Object, ownerHandle As IntPtr) As Boolean Implements IDialogService.QueryDataDialog

    Dim vRtn As Boolean = False

    Dim vDlg As New Advanced.Client.Forms.QueryForm()
    vDlg.DataContext = pDataContext
    'vDlg.WindowStartupLocation = Windows.WindowStartupLocation.Manual
    'Dim vScreenBounds As Rectangle = Screen.FromControl(MainHelper.MainForm).Bounds
    'If pBounds.IsEmpty Then pBounds = vScreenBounds
    'If pBounds.Width < vDlg.Width Then
    '  pBounds.Location = New Point(vScreenBounds.Left, pBounds.Top)
    'End If
    'If pBounds.Height < vDlg.Height Then
    '  pBounds.Location = New Point(pBounds.Left, vScreenBounds.Top)
    'End If
    'vDlg.Left = pBounds.Left
    'vDlg.Top = pBounds.Top
    'vDlg.Width = pBounds.Width
    'vDlg.Height = pBounds.Height
    Dim iTrop As New System.Windows.Interop.WindowInteropHelper(vDlg)
    iTrop.Owner = ownerHandle
    vDlg.ShowDialog()
    If vDlg.DialogResult.HasValue Then vRtn = vDlg.DialogResult.Value

    Return vRtn

  End Function
End Class