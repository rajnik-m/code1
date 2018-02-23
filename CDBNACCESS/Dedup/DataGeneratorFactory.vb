Imports System.Linq

Namespace Access.Deduplication
  Public Class DedupDataGeneratorFactory
    Public Shared Function GetDedupDataGenerator(pEnv As CDBEnvironment, pGeneratorType As DedupDataSelection.DataSelectionTypes) As IDedupDataGenerator
      Dim vGenerator As IDedupDataGenerator = Nothing
      Dim vType As Type = GetType(IDedupDataGenerator)
      Dim vFactoryFlags As List(Of DedupDataSelection.DataSelectionTypes) = GetSystemFactoryAttributes(pGeneratorType, pEnv) 'Get the right factory type for the current system's installation
      Dim vClasses As IEnumerable(Of Type) = GetType(DedupDataSelection).Assembly.GetTypes().Where(Function(p) vType.IsAssignableFrom(p))
      Dim vMatchedClasses As New Dictionary(Of Type, List(Of DedupDataSelection.DataSelectionTypes))
      For Each vClass As Type In vClasses
        Dim vClassAttributes() As Object = vClass.GetCustomAttributes(GetType(EnumEquivalentAttribute), False)
        If vClassAttributes IsNot Nothing AndAlso vClassAttributes.Length > 0 Then
          Dim vClassFlags As New List(Of DedupDataSelection.DataSelectionTypes)
          For Each vAttribute As EnumEquivalentAttribute In vClassAttributes
            Dim vEnumEquiv As EnumEquivalentAttribute = DirectCast(vAttribute, EnumEquivalentAttribute)
            vClassFlags.Add(vEnumEquiv.GetValue(Of DedupDataSelection.DataSelectionTypes)())
          Next
          If New HashSet(Of DedupDataSelection.DataSelectionTypes)(vClassFlags).SetEquals(vFactoryFlags) Then ' Check if all required factory flags are defined on the class
            'If vClassFlags.Intersect(vFactoryFlags).Count = vFactoryFlags.Count Then
            vGenerator = CType(Activator.CreateInstance(vClass), IDedupDataGenerator)
            If vGenerator IsNot Nothing Then
              vGenerator.Connection = pEnv.Connection
            End If
            Exit For
          End If
        End If
      Next
      Return vGenerator
    End Function
    ''' <summary>
    ''' Returns the Factory Type that needs to be loaded based on how the system is configured.
    ''' </summary>
    ''' <param name="pFactoryType"></param>
    ''' <returns></returns>
    ''' <remarks>The Factory Type that is returned will be the one that is needed for the system's current implementation.
    ''' For example, if the call is for a DedupDataSelectionType.Contacts but Uniserv is installed, then the factory type will be changed to DedupDataSelectionType.Contacts | DedupDataSelectionType.Uniserv
    ''' </remarks>
    Private Shared Function GetSystemFactoryAttributes(pFactoryType As DedupDataSelection.DataSelectionTypes, pEnv As CDBEnvironment) As List(Of DedupDataSelection.DataSelectionTypes)
      Dim vResult As New List(Of DedupDataSelection.DataSelectionTypes)
      vResult.Add(pFactoryType)
      If pEnv.UniservInterface.MailActive() Then
        vResult.Add(DedupDataSelection.DataSelectionTypes.Uniserv)
      End If
      Return vResult
    End Function
  End Class
End Namespace
