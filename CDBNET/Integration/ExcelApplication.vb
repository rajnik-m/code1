Option Strict Off
Imports System.IO

Public Class ExcelApplication
  Inherits ExternalApplication

  'Private WithEvents mvExcel As Excel.Application
  Private mvExcel As Object
  Private mvMergeFileName As String
  Private mvExcelActive As Boolean

  Public Overrides Sub ProcessAppActive()
    If mvExcelActive Then
      If ExcelActive() Then
        Dim vFileInfo As New FileInfo(mvFileName)
        Select Case mvAction
          Case DocumentActions.daCreating, DocumentActions.daEditing, DocumentActions.daPrinting, DocumentActions.daViewing
            For Each vBook As Object In mvExcel.Workbooks       'Was Excel.Workbook
              If vBook.Name = vFileInfo.Name Then
                Select Case mvAction
                  Case DocumentActions.daCreating, DocumentActions.daEditing
                    vBook.Close(1)                              'Save the Changes
                  Case DocumentActions.daViewing, DocumentActions.daPrinting
                    vBook.Close(0)                              'Do Not Save the Changes
                End Select
                Exit For
              End If
            Next
        End Select
        mvExcelActive = False
        mvExcel.Quit()
      End If
      mvExcel = Nothing
      MyBase.ProcessActionComplete()
    End If
  End Sub

  Protected Overrides Sub DoEditDocument()
    InitExcel(DocumentActions.daEditing)
    Dim vBook As Object = mvExcel.Workbooks.Open(mvFileName)     'Was Excel.Workbook
    mvExcel.Visible = True
    vBook.Activate()
  End Sub

  Protected Overrides Sub DoViewDocument()
    InitExcel(DocumentActions.daViewing)
    Dim vBook As Object = mvExcel.Workbooks.Open(mvFileName, ReadOnly:=True)    'Was Excel.Workbook
    mvExcel.Visible = True
    vBook.Activate()
  End Sub

  Protected Overrides Sub DoPrintDocument()
    InitExcel(DocumentActions.daPrinting)
    Dim vBook As Object = mvExcel.Workbooks.Open(mvFileName, ReadOnly:=True)    'Was Excel.Workbook
    vBook.PrintOut()
    mvExcelActive = False
    mvExcel.Quit()
    mvExcel = Nothing
    MyBase.ProcessActionComplete()
  End Sub

  Protected Overrides Sub DoEditNewDocument(ByVal pList As ParameterList)
    InitExcel(DocumentActions.daCreating)
    Dim vBook As Object = mvExcel.Workbooks.Add                                 'Was Excel.Workbook
    vBook.SaveAs(CType(mvFileName, Object))
    mvExcel.Visible = True
    vBook.Activate()
  End Sub

  Public Overrides Sub EditNewStandardDocument(ByVal pList As ParameterList, ByVal pExtension As String, ByVal pMailMerge As Boolean)
    Try
      Dim vStandardDocument As String = pList("StandardDocument")
      mvFileName = DataHelper.GetStandardDocumentFile(vStandardDocument, pExtension)
      InitExcel(DocumentActions.daCreating)
      Dim vBook As Object = mvExcel.Workbooks.Open(mvFileName)                  'Was Excel.Workbook
      mvExcel.Visible = True
      vBook.Activate()
    Catch vException As Exception
      DataHelper.HandleException(vException)
    End Try
  End Sub

  Private Sub InitExcel(ByVal pAction As DocumentActions)
    If mvExcel Is Nothing Then
      mvExcel = CreateObject("Excel.Application")
      'mvExcel = New Excel.Application     
      If mvExcel Is Nothing Then Throw New CareException(CareException.ErrorNumbers.enCannotRunExcel)
    End If
    mvAction = pAction
    mvExcelActive = True
  End Sub

  Private Function ExcelActive() As Boolean
    Try
      If Not mvExcel Is Nothing AndAlso mvExcel.Visible Then ExcelActive = True
    Catch vException As System.Runtime.InteropServices.COMException
      mvExcelActive = False
      Debug.Print(vException.ToString)
    End Try
  End Function
  Public Overrides Sub MergeStandardDocument(ByVal pStandardDocument As String, ByVal pExtension As String, ByVal pMergeFileName As String, Optional ByVal pInstantPrint As Boolean = False)

  End Sub
End Class
