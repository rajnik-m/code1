﻿Imports System.Linq
Namespace Access

  Public Class Reversal
    Inherits CARERecord

    Private mvBatchTransactionAnalysis As BatchTransactionAnalysis

#Region "AutoGenerated Code"

    '--------------------------------------------------
    'Enum defining all the fields in the table
    '--------------------------------------------------
    Private Enum ClassFieldItems
      AllFields = 0
      BatchNumber
      TransactionNumber
      LineNumber
      WasBatchNumber
      WasTransactionNumber
      WasLineNumber
      WasOphNumber
    End Enum

    '--------------------------------------------------
    'Required overrides for the class
    '--------------------------------------------------
    Protected Overrides Sub AddFields()
      With mvClassFields
        .Add("batch_number", CDBField.FieldTypes.cftInteger)
        .Add("transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("line_number", CDBField.FieldTypes.cftInteger)
        .Add("was_batch_number", CDBField.FieldTypes.cftInteger)
        .Add("was_transaction_number", CDBField.FieldTypes.cftInteger)
        .Add("was_line_number", CDBField.FieldTypes.cftInteger)
        .Add("was_oph_status", CDBField.FieldTypes.cftCharacter)

        .Item(ClassFieldItems.BatchNumber).PrefixRequired = True
        .Item(ClassFieldItems.TransactionNumber).PrefixRequired = True
        .Item(ClassFieldItems.LineNumber).PrefixRequired = True
        .Item(ClassFieldItems.WasBatchNumber).PrefixRequired = True
        .Item(ClassFieldItems.WasTransactionNumber).PrefixRequired = True
        .Item(ClassFieldItems.WasLineNumber).PrefixRequired = True

        .Item(ClassFieldItems.BatchNumber).SetPrimaryKeyOnly()
        .Item(ClassFieldItems.TransactionNumber).SetPrimaryKeyOnly()
        .Item(ClassFieldItems.LineNumber).SetPrimaryKeyOnly()

      End With

    End Sub

    Protected Overrides ReadOnly Property SupportsAmendedOnAndBy() As Boolean
      Get
        Return False
      End Get
    End Property
    Protected Overrides ReadOnly Property TableAlias() As String
      Get
        Return "rev"
      End Get
    End Property
    Protected Overrides ReadOnly Property DatabaseTableName() As String
      Get
        Return "reversals"
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

#End Region

    Public Property BatchNumber() As Integer
      Get
        Return mvClassFields(ClassFieldItems.BatchNumber).IntegerValue
      End Get
      Set(ByVal value As Integer)
        mvClassFields(ClassFieldItems.BatchNumber).IntegerValue = value
      End Set
    End Property
    Public Property TransactionNumber() As Integer
      Get
        Return mvClassFields(ClassFieldItems.TransactionNumber).IntegerValue
      End Get
      Set(ByVal value As Integer)
        mvClassFields(ClassFieldItems.TransactionNumber).IntegerValue = value
      End Set
    End Property
    Public Property LineNumber() As Integer
      Get
        Return mvClassFields(ClassFieldItems.LineNumber).IntegerValue
      End Get
      Set(ByVal value As Integer)
        mvClassFields(ClassFieldItems.LineNumber).IntegerValue = value
      End Set
    End Property
    Private Property WasBatchNumber() As Integer
      Get
        Return mvClassFields(ClassFieldItems.WasBatchNumber).IntegerValue
      End Get
      Set(ByVal value As Integer)
        mvClassFields(ClassFieldItems.WasBatchNumber).IntegerValue = value
      End Set
    End Property
    Private Property WasTransactionNumber() As Integer
      Get
        Return mvClassFields(ClassFieldItems.WasTransactionNumber).IntegerValue
      End Get
      Set(ByVal value As Integer)
        mvClassFields(ClassFieldItems.WasTransactionNumber).IntegerValue = value
      End Set
    End Property
    Private Property WasLineNumber() As Integer
      Get
        Return mvClassFields(ClassFieldItems.WasLineNumber).IntegerValue
      End Get
      Set(ByVal value As Integer)
        mvClassFields(ClassFieldItems.WasLineNumber).IntegerValue = value
      End Set
    End Property

    Public Property OriginalBatchNumber As Integer
      Get
        Return Me.WasBatchNumber
      End Get
      Set(value As Integer)
        Me.WasBatchNumber = value
      End Set
    End Property
    Public Property OriginalTransactionNumber As Integer
      Get
        Return Me.WasTransactionNumber
      End Get
      Set(value As Integer)
        Me.WasTransactionNumber = value
      End Set
    End Property
    Public Property OriginalLineNumber As Integer
      Get
        Return Me.WasLineNumber
      End Get
      Set(value As Integer)
        Me.WasLineNumber = value
      End Set
    End Property

    Public Property BatchTransactionAnalysis As BatchTransactionAnalysis
      Get
        If mvBatchTransactionAnalysis Is Nothing Then
          InitBTA()
        End If
        Return mvBatchTransactionAnalysis
      End Get
      Private Set(value As BatchTransactionAnalysis)
        mvBatchTransactionAnalysis = value
      End Set
    End Property

    Private Sub InitBTA()
      Dim vWhere As CDBFields = Me.CreateWhere(New List(Of Integer)({ClassFieldItems.BatchNumber,
                                                                     ClassFieldItems.TransactionNumber,
                                                                     ClassFieldItems.LineNumber}))
      Dim vBTA As New BatchTransactionAnalysis(Me.Environment)
      vBTA.InitWithPrimaryKey(vWhere)
      If vBTA.Existing Then Me.BatchTransactionAnalysis = vBTA
    End Sub

    Public Property OriginalBatchTransactionAnalysis As BatchTransactionAnalysis
      Get
        If mvBatchTransactionAnalysis Is Nothing Then
          Me.OriginalBatchTransactionAnalysis = GetOriginalBTAInstance()
        End If
        Return mvBatchTransactionAnalysis
      End Get
      Private Set(value As BatchTransactionAnalysis)
        mvBatchTransactionAnalysis = value
      End Set
    End Property

    Private Function GetOriginalBTAInstance() As BatchTransactionAnalysis
      Dim vResult As BatchTransactionAnalysis = Nothing
      Dim vTest As New BatchTransactionAnalysis(Me.Environment)
      vTest.Init(Me.WasBatchNumber, Me.WasTransactionNumber, Me.WasLineNumber)
      If vTest.Existing Then vResult = vTest
      Return vResult
    End Function

    Protected Overloads Sub InitClassFields()
      MyBase.InitClassFields()
      Me.BatchTransactionAnalysis = Nothing
    End Sub

  End Class
End Namespace