Namespace Access

  Public Class ExamBookingLine

    Private mvLineNumber As Integer
    Private mvProduct As String
    Private mvRateCode As String
    Private mvRateDesc As String
    Private mvQuantity As Integer
    Private mvAmount As Double
    Private mvVATAmount As Double
    Private mvVATRateCode As String
    Private mvVATPercentage As Double
    Private mvNotes As String
    Private mvExamUnitId As Integer
    Private mvExamUnitProductId As Integer
    Private mvBatchNumber As Integer
    Private mvTransactionNumber As Integer
    Private mvTransactionLineNumber As Integer

    Private mvAlternateRateCode As String
    Private mvAlternateRateDesc As String
    Private mvAlternateAmount As Double
    Private mvAlternateVATAmount As Double

    Public Sub New(ByVal pLineNumber As Integer, ByVal pExamUnitId As Integer, ByVal pExamUnitProductId As Integer, ByVal pProduct As String, ByVal pRateCode As String, ByVal pRateDesc As String, ByVal pQuantity As Integer, ByVal pAmount As Double, ByVal pVATAmount As Double, ByVal pVATRateCode As String, ByVal pVATPercentage As Double, ByVal pNotes As String)
      mvLineNumber = pLineNumber
      mvExamUnitId = pExamUnitId
      mvExamUnitProductId = pExamUnitProductId
      mvProduct = pProduct
      mvRateCode = pRateCode
      mvRateDesc = pRateDesc
      mvQuantity = pQuantity
      mvAmount = pAmount
      mvVATAmount = pVATAmount
      mvVATRateCode = pVATRateCode
      mvVATPercentage = pVATPercentage
      mvNotes = pNotes
    End Sub

    Public Sub SetAlternateRate(ByVal pRateCode As String, ByVal pRateDesc As String, ByVal pAmount As Double, ByVal pVATAmount As Double)
      mvAlternateRateCode = pRateCode
      mvAlternateRateDesc = pRateDesc
      mvAlternateAmount = pAmount
      mvAlternateVATAmount = pVATAmount
    End Sub

    Public ReadOnly Property Amount As Double
      Get
        Return mvAmount
      End Get
    End Property

    Public ReadOnly Property ProductCode As String
      Get
        Return mvProduct
      End Get
    End Property

    Public ReadOnly Property RateCode As String
      Get
        Return mvRateCode
      End Get
    End Property

    Public ReadOnly Property ExamUnitId As Integer
      Get
        Return mvExamUnitId
      End Get
    End Property

    Public ReadOnly Property ExamUnitProductId As Integer
      Get
        Return mvExamUnitProductId
      End Get
    End Property

    Public ReadOnly Property LineNumber As Integer
      Get
        Return mvLineNumber
      End Get
    End Property

    Public ReadOnly Property Notes As String
      Get
        Return mvNotes
      End Get
    End Property

    Public Property BatchNumber As Integer
      Get
        Return mvBatchNumber
      End Get
      Set(ByVal value As Integer)
        mvBatchNumber = value
      End Set
    End Property

    Public Property TransactionNumber As Integer
      Get
        Return mvTransactionNumber
      End Get
      Set(ByVal value As Integer)
        mvTransactionNumber = value
      End Set
    End Property

    Public Property TransactionLineNumber As Integer
      Get
        Return mvTransactionLineNumber
      End Get
      Set(ByVal value As Integer)
        mvTransactionLineNumber = value
      End Set
    End Property


    Public Function GetDataAsString() As String
      Dim vSB As New StringBuilder
      vSB.Append(mvLineNumber).Append(",")
      vSB.Append(mvExamUnitId).Append(",")
      vSB.Append(mvExamUnitProductId).Append(",")
      vSB.Append(mvProduct).Append(",")
      vSB.Append(mvRateCode).Append(",")
      vSB.Append(mvRateDesc).Append(",")
      vSB.Append(mvQuantity).Append(",")
      vSB.Append(mvAmount).Append(",")
      vSB.Append(mvVATAmount).Append(",")
      vSB.Append(mvVATRateCode).Append(",")
      vSB.Append(mvVATPercentage).Append(",")
      vSB.Append(mvNotes).Append(",")
      vSB.Append(mvAlternateRateCode).Append(",")
      vSB.Append(mvAlternateRateDesc).Append(",")
      vSB.Append(mvAlternateAmount).Append(",")
      vSB.Append(mvAlternateVATAmount)
      Return vSB.ToString
    End Function

    Public ReadOnly Property AlternateRateCode As String
      Get
        Return mvAlternateRateCode
      End Get
    End Property
    Public ReadOnly Property AlternateAmount As Double
      Get
        Return mvAlternateAmount
      End Get
    End Property

  End Class
End Namespace