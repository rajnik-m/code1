Imports System.Data

Public Class ExpressionEvaluator
  Inherits ArrayList

  Private Class BinaryNode
    Public Value As String
    Public NodeOperator As String
    Public Name As String
    Public LeftNode As BinaryNode
    Public RightNode As BinaryNode
    Public DataTable As System.Data.DataTable

    Public Sub New(ByVal pValue As String)
      Value = pValue
    End Sub
  End Class

  Private mvInvalid As Boolean
  Private mvOperatorStack As New Stack(Of String)
  Private mvOperandStack As New Stack(Of BinaryNode)

  Public Sub New(ByVal pExpression As String)
    Dim vInQuotes As Boolean
    Dim vLastChar As Char = Nothing
    Dim vItem As New StringBuilder
    Dim vLastItem As String = ""
    Dim vBracketCount As Integer
    Dim vItemCount As Integer
    For Each vChar As Char In pExpression.ToCharArray
      Select Case vChar
        Case """"c
          vItem.Append(vChar)
          If Not vInQuotes AndAlso vLastChar = """" Then ' Two quotes make a quote
            vLastChar = vChar
          End If
          vInQuotes = Not vInQuotes
        Case ","c, "'"c, "~"c, "!"c
          'Ignore
        Case " "c, "("c, ")"c, "&"c, "|"c
          If Not vInQuotes Then
            If vItem.Length > 0 Then
              MyBase.Add(vItem.ToString)
              vItemCount += 1
              vItem = New StringBuilder
            End If
            If vChar <> " "c Then MyBase.Add(vChar)
            If vChar = "(" Then
              vBracketCount += 1
            ElseIf vChar = ")" Then
              vBracketCount -= 1
            End If
          Else
            vItem.Append(vChar) 'in quotes
          End If
        Case Else
          vItem.Append(vChar)   'not an operator character
      End Select
      vLastChar = vChar
    Next
    If vItem.Length > 0 Then
      MyBase.Add(vItem.ToString)
      vItemCount += 1
    End If
    If vInQuotes OrElse vBracketCount > 0 OrElse vItemCount = 0 Then
      mvInvalid = True
    Else
      'Add any missing ands
      Dim vItemInserted As Boolean
      Do
        vItemInserted = False
        vLastItem = ""
        For vIndex As Integer = 0 To Me.Count - 1
          Dim vString As String = Me(vIndex).ToString
          Select Case vString.ToLower
            Case "and", "or", "&", "|", ")"
              'Do Nothing
            Case Else
              Select Case vLastItem.ToLower
                Case "and", "or", "&", "|"
                  'Do Nothing
                Case ""
                  'Do nothing
                Case Else
                  vString = "and"
                  MyBase.Insert(vIndex, vString)
                  vItemInserted = True
                  Exit For
              End Select
          End Select
          If vString <> "(" And vString <> ")" Then vLastItem = vString
        Next
      Loop While vItemInserted
      If Me.Count > 0 Then
        Select Case Me(Me.Count - 1).ToString
          Case "and", "or", "&", "|"
            mvInvalid = True
        End Select
      End If
    End If
  End Sub

  Public ReadOnly Property Invalid() As Boolean
    Get
      Return mvInvalid
    End Get
  End Property

  Public ReadOnly Property Expression() As String
    Get
      Dim vList As New System.Text.StringBuilder
      For Each vString As String In Me
        If vList.Length > 0 Then vList.Append(" ")
        vList.Append(vString)
      Next
      Return vList.ToString
    End Get
  End Property

  Public Function ParseExpression() As Integer
    mvOperatorStack.Clear()
    mvOperandStack.Clear()
    Dim vOperandCount As Integer
    For Each vString As String In Me
      Select Case vString.ToLower
        Case "("
          mvOperatorStack.Push(vString)
        Case "&", "|", "and", "or"
          If (mvOperatorStack.Count = 0) OrElse (mvOperatorStack.Peek = "(") Then
            mvOperatorStack.Push(vString)
          Else 'clear operator stack and push new one onto it 
            Do
              PopConnectPush()
            Loop While (mvOperatorStack.Count > 0) AndAlso (mvOperatorStack.Peek <> "(")
            mvOperatorStack.Push(vString)
          End If
        Case ")"  'clear operator stack back to the '(' 
          While (mvOperatorStack.Count > 0) AndAlso (mvOperatorStack.Peek <> "(")
            PopConnectPush()
          End While
          'Get rid of the open paren that matches this closing one??
          If mvOperatorStack.Count > 0 AndAlso mvOperatorStack.Peek = "(" Then mvOperatorStack.Pop()
        Case ""
          'Just in case
          'Case "(", "&", "|", "and", "or"
          '  mvOperatorStack.Push(vString)
          'Case ")"
          '  Dim vNode As New BinaryNode("")
          '  Do
          '    vNode.NodeOperator = mvOperatorStack.Pop
          '  Loop While vNode.NodeOperator = "("
          '  vNode.RightNode = mvOperandStack.Pop
          '  vNode.LeftNode = mvOperandStack.Pop
          '  mvOperandStack.Push(vNode)
        Case Else
          mvOperandStack.Push(New BinaryNode(vString))
          vOperandCount += 1
      End Select
    Next
    While mvOperatorStack.Count > 0
      PopConnectPush()
    End While
    Return vOperandCount
  End Function

  Private Sub PopConnectPush()
    Dim vNode As New BinaryNode("")
    vNode.NodeOperator = mvOperatorStack.Pop
    vNode.RightNode = mvOperandStack.Pop
    vNode.LeftNode = mvOperandStack.Pop
    mvOperandStack.Push(vNode)
  End Sub

  Public Delegate Function GetExpressionDataTable(ByVal pSearchTerm As String, ByVal pName As String) As System.Data.DataTable

  Public Function EvaluateExpression(ByVal pGetDataTable As GetExpressionDataTable) As System.Data.DataTable
    Dim vNode As BinaryNode = mvOperandStack.Peek
    If vNode IsNot Nothing Then
      TraverseTree(vNode, 1)
      EvaluateNode(vNode, pGetDataTable)
      Return vNode.DataTable
    Else
      Return Nothing
    End If
  End Function

  Private Sub TraverseTree(ByVal pNode As BinaryNode, ByRef pCount As Integer)
    If pNode Is Nothing Then Return
    If pNode.LeftNode IsNot Nothing Then TraverseTree(pNode.LeftNode, pCount)
    If pNode.RightNode IsNot Nothing Then TraverseTree(pNode.RightNode, pCount)
    pNode.Name = "Node" & pCount
    pCount += 1
    Debug.Print(String.Format("{0} Value {1} Operator '{2}'", pNode.Name, pNode.Value, pNode.NodeOperator))
    If pNode.LeftNode IsNot Nothing Then Debug.Print(String.Format("  Left {0}", pNode.LeftNode.Name))
    If pNode.RightNode IsNot Nothing Then Debug.Print(String.Format("  Right {0}", pNode.RightNode.Name))
  End Sub

  Private Sub EvaluateNode(ByVal pNode As BinaryNode, ByVal pGetDataTable As GetExpressionDataTable)
    If Not String.IsNullOrEmpty(pNode.LeftNode.Value) Then
      Debug.Print(String.Format("Evaluating Node {0} LeftNode Name {1} Value {2} ", pNode.Name, pNode.LeftNode.Name, pNode.LeftNode.Value))
      pNode.LeftNode.DataTable = pGetDataTable(pNode.LeftNode.Value, pNode.LeftNode.Name)
    Else
      EvaluateNode(pNode.LeftNode, pGetDataTable)
    End If
    'Check for early out optimisation
    If (pNode.LeftNode.DataTable Is Nothing OrElse pNode.LeftNode.DataTable.Rows.Count = 0) AndAlso _
       (pNode.NodeOperator IsNot Nothing AndAlso pNode.NodeOperator = "and") Then
      pNode.DataTable = Nothing
      pNode.LeftNode.DataTable = Nothing
      pNode.RightNode.DataTable = Nothing
      Exit Sub
    End If

    If Not String.IsNullOrEmpty(pNode.RightNode.Value) Then
      Debug.Print(String.Format("Evaluating Node {0} RightNode Name {1} Value {2} ", pNode.Name, pNode.RightNode.Name, pNode.RightNode.Value))
      pNode.RightNode.DataTable = pGetDataTable(pNode.RightNode.Value, pNode.RightNode.Name)
    Else
      EvaluateNode(pNode.RightNode, pGetDataTable)
    End If
    If Not String.IsNullOrEmpty(pNode.NodeOperator) Then
      Select Case pNode.NodeOperator
        Case "and"
          If pNode.LeftNode.DataTable IsNot Nothing AndAlso pNode.LeftNode.DataTable.Rows.Count > 0 AndAlso _
             pNode.RightNode.DataTable IsNot Nothing AndAlso pNode.RightNode.DataTable.Rows.Count > 0 Then
            pNode.DataTable = JoinTables(pNode.LeftNode.DataTable, pNode.RightNode.DataTable, BooleanQueryJoinType.AndJoin)
            pNode.DataTable.TableName = pNode.Name
            pNode.LeftNode.DataTable.DataSet.Tables.Add(pNode.DataTable)
          Else
            pNode.DataTable = Nothing
          End If
          pNode.LeftNode.DataTable = Nothing
          pNode.RightNode.DataTable = Nothing
        Case "or"
          If pNode.LeftNode.DataTable IsNot Nothing AndAlso pNode.LeftNode.DataTable.Rows.Count > 0 AndAlso _
             pNode.RightNode.DataTable IsNot Nothing AndAlso pNode.RightNode.DataTable.Rows.Count > 0 Then
            pNode.DataTable = JoinTables(pNode.LeftNode.DataTable, pNode.RightNode.DataTable, BooleanQueryJoinType.OrJoin)
            pNode.DataTable.TableName = pNode.Name
            pNode.LeftNode.DataTable.DataSet.Tables.Add(pNode.DataTable)
          ElseIf pNode.LeftNode.DataTable IsNot Nothing AndAlso pNode.LeftNode.DataTable.Rows.Count > 0 Then
            pNode.DataTable = pNode.LeftNode.DataTable
          ElseIf pNode.RightNode.DataTable IsNot Nothing AndAlso pNode.RightNode.DataTable.Rows.Count > 0 Then
            pNode.DataTable = pNode.RightNode.DataTable
          End If
          pNode.LeftNode.DataTable = Nothing
          pNode.RightNode.DataTable = Nothing
      End Select
    End If
  End Sub

  Public Enum BooleanQueryJoinType
    AndJoin         'Result only if in both tables
    OrJoin          'Result if in either table
    ExcludeLeft     'All Results from left table and Results from right table if not already in left table
  End Enum

  Public Function MergeTableRows(ByVal pTable As DataTable) As DataTable
    Dim vRow As Integer = 0
    Dim vCheckRow As Integer = 0
    Dim vAllRowsProcessed As Boolean = False
    While vAllRowsProcessed = False
      If vRow >= pTable.Rows.Count Then
        vAllRowsProcessed = True
      Else
        Dim vAllRowsChecked As Boolean = False
        vCheckRow = vRow + 1
        While vAllRowsChecked = False
          If vCheckRow >= pTable.Rows.Count Then
            vAllRowsChecked = True
          Else
            'Check this row and merge / delete
            If pTable.Rows(vRow)("item_number").ToString = pTable.Rows(vCheckRow)("item_number").ToString AndAlso _
               pTable.Rows(vRow)("item_type").ToString = pTable.Rows(vCheckRow)("item_type").ToString Then
              MergeRowText(pTable.Rows(vRow), pTable.Rows(vCheckRow))
              pTable.Rows.RemoveAt(vCheckRow)
            Else
              vCheckRow += 1
            End If
          End If
        End While
        vRow += 1
      End If
    End While
    Return pTable
  End Function

  Private Sub MergeRowText(ByVal pDestRow As DataRow, ByVal pMergeRow As DataRow)
    Dim vDestSource() As String = pDestRow("item_source").ToString.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
    Dim vMergeSource() As String = pMergeRow("item_source").ToString.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
    Dim vDestText() As String = pDestRow("item_text").ToString.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
    Dim vMergeText() As String = pMergeRow("item_text").ToString.Split(vbCrLf.ToCharArray, StringSplitOptions.RemoveEmptyEntries)
    Dim vFound As Boolean = False
    For vDestIndex As Integer = 0 To vDestSource.Length - 1
      For vMergeIndex As Integer = 0 To vMergeSource.Length - 1
        If vDestSource(vDestIndex) = vMergeSource(vMergeIndex) Then
          If vDestText(vDestIndex) = vMergeText(vMergeIndex) Then
            vFound = True
            Exit For
          End If
        End If
      Next
      If vFound Then Exit For
    Next
    If vFound = False Then
      pDestRow("item_source") = pDestRow("item_source").ToString & vbCrLf & pMergeRow("item_source").ToString
      pDestRow("item_text") = pDestRow("item_text").ToString & vbCrLf & pMergeRow("item_text").ToString
    End If
  End Sub

  Public Function JoinTables(ByVal pLeftTable As DataTable, ByVal pRightTable As DataTable, ByVal pJoinType As BooleanQueryJoinType) As DataTable
    Dim vDestTable As DataTable = pLeftTable.Clone
    Dim vDataSet As DataSet = pLeftTable.DataSet
    Dim vParentCols(1) As DataColumn
    Dim vChildCols(1) As DataColumn

    vParentCols(0) = pLeftTable.Columns("item_number")
    vParentCols(1) = pLeftTable.Columns("item_type")
    vChildCols(0) = pRightTable.Columns("item_number")
    vChildCols(1) = pRightTable.Columns("item_type")

    ' create a relationship between the two tables
    vDataSet.Relations.Add(New DataRelation("__RELATIONSHIP__", vParentCols, vChildCols, False))
    Dim vRows() As DataRow
    Dim vDestRow As DataRow
    For Each vLeftRow As DataRow In pLeftTable.Rows
      ' Get the related rows from the "right" table
      vRows = vLeftRow.GetChildRows("__RELATIONSHIP__")
      ' For inner joins, we don't record anything unless there is a matching row
      If UBound(vRows) >= 0 OrElse pJoinType = BooleanQueryJoinType.OrJoin OrElse pJoinType = BooleanQueryJoinType.ExcludeLeft Then
        vDestRow = vDestTable.NewRow
        vDestRow.ItemArray = vLeftRow.ItemArray
        ' There are three possibilities... there are no matching rows, there is
        ' only one related row, there are many related rows.
        Select Case UBound(vRows)
          Case -1
            ' Just record the row as it is now (with just the columns from the left table).
            vDestTable.Rows.Add(vDestRow)
          Case 0
            MergeRowText(vDestRow, vRows(0))
            vDestTable.Rows.Add(vDestRow)
            If pJoinType = BooleanQueryJoinType.OrJoin Then pRightTable.Rows.Remove(vRows(0))
          Case Else
            For Each vFoundRow As DataRow In vRows
              MergeRowText(vDestRow, vFoundRow)
              If pJoinType = BooleanQueryJoinType.OrJoin Then pRightTable.Rows.Remove(vFoundRow)
            Next
            vDestTable.Rows.Add(vDestRow)
        End Select
      End If
    Next
    If pJoinType = BooleanQueryJoinType.OrJoin Then
      For Each vRightRow As DataRow In pRightTable.Rows
        vDestRow = vDestTable.NewRow
        vDestRow.ItemArray = vRightRow.ItemArray
        vDestTable.Rows.Add(vDestRow)
      Next
    End If
    ' delete the temporary relationship we created above
    vDataSet.Relations.Remove("__RELATIONSHIP__")
    Return vDestTable
  End Function


End Class