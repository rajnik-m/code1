Imports System.Xml.Serialization

Namespace Access.Deduplication

  <XmlType(Namespace:="http://tempuri.org/ANFP/Access/Dedup")>
  Public Enum DedupMatch
    Equals
    IsLike
    Contains
  End Enum
End Namespace
