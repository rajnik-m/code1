Imports System.Xml.Serialization

Namespace Access.Deduplication
  <XmlType(Namespace:="http://tempuri.org/ANFP/Access/Dedup")>
Public Class DedupClause
    'Implements IXmlSerializable

    Dim mvTable As String
    Dim mvAlias As String

    Public Sub New()
      Me.Joins = New List(Of Join) 'If this gets too complicated it should be replaced by a class
    End Sub

    <XmlAttribute>
    Public Property Parameter As String

    <XmlAttribute>
    Public Property Attribute As String

    <XmlAttribute>
    Public Property Match As DedupMatch

    <XmlAttribute>
    Public Property Table As String
      Get
        Return mvTable
      End Get
      Set(value As String)
        mvTable = Utilities.Common.ExtractWord(value, 0)
        mvAlias = Utilities.Common.ExtractWord(value, 1)
      End Set
    End Property

    Public ReadOnly Property TableAlias As String
      Get
        Return If(Not String.IsNullOrWhiteSpace(mvAlias), mvAlias, mvTable)
      End Get
    End Property

    Public Property Joins As List(Of Join)
    'Public Function GetSchema() As Xml.Schema.XmlSchema Implements IXmlSerializable.GetSchema
    '  Return Nothing
    'End Function

    'Public Sub ReadXml(reader As Xml.XmlReader) Implements IXmlSerializable.ReadXml
    '  If reader.MoveToContent() = System.Xml.XmlNodeType.Element AndAlso reader.LocalName = Me.GetType().Name Then
    '    Me.Parameter = reader.GetAttribute("Parameter")
    '    Me.Table = reader.GetAttribute("Table")
    '    Me.Attribute = reader.GetAttribute("Attribute")
    '    [Enum].TryParse(Of DedupMatch)(reader.GetAttribute("Match"), Me.Match)
    '    reader.Read()
    '    If reader.MoveToContent() = System.Xml.XmlNodeType.Element AndAlso reader.LocalName = "Joins" Then
    '      reader.ReadStartElement("Joins")
    '      While reader.LocalName = "Clause"
    '        Dim vClause As New Join() With {.Table = reader.GetAttribute("Table"), .AnchorPart = reader.GetAttribute("AnchorPart"), .AnchorValue = reader.GetAttribute("JoinPart")}
    '        Me.Joins.Add(vClause)
    '        reader.Read()
    '      End While
    '    End If
    '  End If
    'End Sub

    'Public Sub WriteXml(writer As Xml.XmlWriter) Implements IXmlSerializable.WriteXml
    '  writer.WriteAttributeString("Parameter", Me.Parameter)
    '  writer.WriteAttributeString("Table", Me.Table)
    '  writer.WriteAttributeString("Attribute", Me.Attribute)
    '  writer.WriteAttributeString("Match", Me.Match.ToString())
    '  If Me.Joins IsNot Nothing AndAlso Me.Joins.Count > 0 Then
    '    writer.WriteStartElement("Joins")
    '    For Each vJoin As Join In Me.Joins
    '      writer.WriteStartElement("Clause")
    '      writer.WriteAttributeString("Table", vJoin.Table)
    '      writer.WriteAttributeString("AnchorPart", vJoin.AnchorPart)
    '      writer.WriteAttributeString("JoinPart", vJoin.AnchorValue)
    '      writer.WriteEndElement()
    '    Next
    '    writer.WriteEndElement()
    '  End If
    'End Sub

    Public Class Join
      Private mvTable As String = String.Empty
      Private mvAlias As String = String.Empty

      ''' <summary>
      ''' The SQL table that the join clause uses.  The Get of this property is purely for Xml serialization, or if you want the table name as it was originall passed.  Use the TableAlias and Table properties instead.
      ''' </summary>
      ''' <value></value>
      ''' <remarks></remarks>
      <XmlAttribute("Table")>
      Public Property RawTableName As String
        Get
          Return Me.TableAndAlias
        End Get
        Set(value As String)
          mvTable = Utilities.Common.ExtractWord(value, 0)
          mvAlias = Utilities.Common.ExtractWord(value, 1)
        End Set
      End Property

      Public ReadOnly Property [Alias] As String
        Get
          Return If(Not String.IsNullOrWhiteSpace(mvAlias), mvAlias, mvTable)
        End Get
      End Property
      Public ReadOnly Property Table As String
        Get
          Return mvTable
        End Get
      End Property

      Public ReadOnly Property TableAndAlias As String
        Get
          Return Me.Table + If(Me.Alias.Equals(Me.Table), "", String.Format(" {0}", Me.Alias))
        End Get
      End Property

      <XmlAttribute>
      Public Property AnchorPart As String
      <XmlAttribute>
      Public Property JoinPart As String

    End Class

  End Class
End Namespace