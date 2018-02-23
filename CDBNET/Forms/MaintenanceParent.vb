Public Interface MaintenanceParent

  Sub RefreshData()
  Sub RefreshData(ByVal pType As CareServices.XMLMaintenanceControlTypes)
  ReadOnly Property SizeMaintenanceForm() As Boolean
  ReadOnly Property ContactInfo() As ContactInfo
  ReadOnly Property ContactDataType() As CareServices.XMLContactDataSelectionTypes

End Interface

Public Class MaintenanceParentForm
  Inherits PersistentForm
  Implements MaintenanceParent

  Overridable ReadOnly Property SizeMaintenanceForm() As Boolean Implements MaintenanceParent.SizeMaintenanceForm
    Get
      Return False
    End Get
  End Property

  Overridable ReadOnly Property ContactInfo() As ContactInfo Implements MaintenanceParent.ContactInfo
    Get
      Return New ContactInfo(ContactInfo.ContactTypes.ctContact, "")
    End Get
  End Property

  Overridable ReadOnly Property ContactDataType() As CareServices.XMLContactDataSelectionTypes Implements MaintenanceParent.ContactDataType
    Get
      Return CareServices.XMLContactDataSelectionTypes.xcdtNone
    End Get
  End Property

  Overridable Sub RefreshData(ByVal pType As CareServices.XMLMaintenanceControlTypes) Implements MaintenanceParent.RefreshData

  End Sub

  Overridable Sub RefreshData() Implements MaintenanceParent.RefreshData

  End Sub

  Private Sub InitializeComponent()
    Me.SuspendLayout()
    '
    'MaintenanceParentForm
    '
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
    Me.ClientSize = New System.Drawing.Size(292, 260)
    Me.Name = "MaintenanceParentForm"
    Me.ResumeLayout(False)

  End Sub
End Class