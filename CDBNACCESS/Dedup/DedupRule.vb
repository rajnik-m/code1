Imports System.Xml.Serialization
Imports System.Linq

Namespace Access.Deduplication
  <System.Xml.Serialization.XmlType(Namespace:="http://tempuri.org/ANFP/Access/Dedup")>
  <System.Xml.Serialization.XmlRoot(ElementName:="DedupRule", Namespace:="http://tempuri.org/ANFP/Access/Dedup", IsNullable:=True)>
  Public Class DedupRule
    Implements IComparable(Of DedupRule)

    Public Sub New()
      Me.Clauses = New List(Of DedupClause)
    End Sub
    <XmlAttribute>
    Public Property ID As String

    <XmlAttribute>
    Public Property Description As String

    <XmlAttribute>
    Public Property RuleRank As DedupRank

    <XmlArray(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified),
     XmlArrayItem(Form:=System.Xml.Schema.XmlSchemaForm.Unqualified)> _
    Public Property Clauses As List(Of DedupClause)

    Public Function CompareTo(other As DedupRule) As Integer Implements IComparable(Of DedupRule).CompareTo
      Dim vCompareResult As Integer = 0
      If other IsNot Nothing Then
        vCompareResult = Me.RuleRank.CompareTo(other.RuleRank)
        If vCompareResult = 0 Then
          vCompareResult = Me.ID.CompareTo(other.ID)
        End If
      End If
      Return vCompareResult
    End Function

    Public Overrides Function ToString() As String
      Return String.Format("{0} Description: {1} Rank:{2}", MyBase.ToString(), Me.Description, Me.RuleRank)
    End Function
  End Class

End Namespace