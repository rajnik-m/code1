Public Interface IDedupDataGenerator

  ReadOnly Property Rules As List(Of Access.Deduplication.DedupRule)
  Property Connection As CDBConnection
  Function GenerateSQLStatement(pDedupRule As Access.Deduplication.DedupRule) As SQLStatement
  Property Parent As Deduplication.DedupDataSelection
  Property ResultColumns As String
  Property SelectColumns As String
  Property Headings As String
  Property RequiredItems As String
  Property Code As String
  Property Widths As String


End Interface
