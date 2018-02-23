﻿Namespace Access

  Public Class PopProductionDetail
    Inherits CARERecord

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum PopProductionDetailFields
      AllFields = 0
      PopProductionNumber
      Amount
      CreatedBy
      CreatedOn
      AmendedBy
      AmendedOn
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("pop_production_number", CDBField.FieldTypes.cftLong)
        .Add("amount", CDBField.FieldTypes.cftNumeric)
        .Add("created_by")
        .Add("created_on", CDBField.FieldTypes.cftDate)

        .Item(PopProductionDetailFields.Amount).PrefixRequired = True

        .SetControlNumberField(PopProductionDetailFields.PopProductionNumber, "PPN")
        .Item(PopProductionDetailFields.PopProductionNumber).PrimaryKey = True

      End With

    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return True
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "ppn"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "pop_production_details"
      End Get
    End Property

    '--------------------------------------------------
    'Default constructor
    '--------------------------------------------------
    Public Sub New(ByVal pEnv As CDBEnvironment)
      MyBase.New(pEnv)
    End Sub

    '--------------------------------------------------
    'Public property procedures
    '--------------------------------------------------
    Public ReadOnly Property PopProductionNumber() As Integer
      Get
        Return mvClassFields(PopProductionDetailFields.PopProductionNumber).IntegerValue
      End Get
    End Property
    Public Property Amount() As Double
      Get
        Return mvClassFields(PopProductionDetailFields.Amount).DoubleValue
      End Get
      Set(value As Double)
        mvClassFields(PopProductionDetailFields.Amount).DoubleValue = value
      End Set
    End Property
    Public ReadOnly Property CreatedBy() As String
      Get
        Return mvClassFields(PopProductionDetailFields.CreatedBy).Value
      End Get
    End Property
    Public ReadOnly Property CreatedOn() As String
      Get
        Return mvClassFields(PopProductionDetailFields.CreatedOn).Value
      End Get
    End Property
    
#End Region

#Region "Non-AutoGenerated Code"

    Protected Overrides Sub SetDefaults()
      MyBase.SetDefaults()
      mvClassFields.Item(PopProductionDetailFields.Amount).DoubleValue = 0
    End Sub

#End Region

  End Class
End Namespace
