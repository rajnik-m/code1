

Namespace Access
  Public Class PostcodeProximityOrgs

    Private mvEnv As CDBEnvironment
    Private mvCol As Collection
    Private mvNearest As CDBDataTable
    Public Sub Init(ByRef pEnv As CDBEnvironment)
      Dim vPostcodeGroup As String
      Dim vRecordSet As CDBRecordSet
      Dim vPPO As PostcodeProximityOrg
      Dim vExclStatus As String
      Dim vFields As New CDBFields

      mvEnv = pEnv
      vPostcodeGroup = pEnv.GetConfig("cd_postcode_group")

      vFields.Add("o.organisation_group", CDBField.FieldTypes.cftCharacter, vPostcodeGroup)
      vFields.Add("a.address_number", CDBField.FieldTypes.cftLong, "o.address_number")
      vFields.Add("pgr.postcode", CDBField.FieldTypes.cftLong, "a.postcode")
      vExclStatus = pEnv.GetConfig("proximity_exclusion_statuses")
      If Len(vExclStatus) > 0 Then
        If Right(vExclStatus, 1) = "|" Then vExclStatus = Left(vExclStatus, Len(vExclStatus) - 1)
        If Left(vExclStatus, 1) = "|" Then vExclStatus = Right(vExclStatus, Len(vExclStatus) - 1)
        If Len(vExclStatus) > 0 Then
          vExclStatus = Replace(vExclStatus, "|", "','")
          vExclStatus = "'" & vExclStatus & "'"
          vFields.Add("status", CDBField.FieldTypes.cftCharacter, "", CDBField.FieldWhereOperators.fwoOpenBracket)
          vFields.Add("o.status", CDBField.FieldTypes.cftCharacter, vExclStatus, CDBField.FieldWhereOperators.fwoNotIn Or CDBField.FieldWhereOperators.fwoOR Or CDBField.FieldWhereOperators.fwoCloseBracket)
        End If
      End If
      mvCol = New Collection
      vRecordSet = pEnv.Connection.GetRecordSet("SELECT organisation_number, easting, northing, a.postcode FROM organisations o, addresses a, postcode_grid_references pgr WHERE " & pEnv.Connection.WhereClause(vFields))
      With vRecordSet
        While .Fetch() = True
          vPPO = New PostcodeProximityOrg
          vPPO.Init(mvEnv, .Fields(1).IntegerValue, .Fields(2).IntegerValue, .Fields(3).IntegerValue, .Fields(4).Value)
          mvCol.Add(vPPO)
        End While
        .CloseRecordSet()
      End With
    End Sub
    Public Sub InitNearest(ByVal pNumberRequired As Integer, ByVal pEasting As Integer, ByVal pNorthing As Integer, ByVal pPostCode As String)
      'User specifies they want nearest (pNumberRequired) number of
      'Postcode Proximity Organisations to a Reference point specified by location
      'pEasting, pNorthing
      Dim vPPO As PostcodeProximityOrg
      Dim vIndex As Integer
      Dim vDistance As Double
      Dim vIndex2 As Integer
      Dim vDR As CDBDataRow

      'Because of the problems in converting to .NET we must dimension this array before we start
      'We will assume that we will never be asked for more than 20 items
      Dim vNearest(2, 20) As Double
      If pNumberRequired > 20 Then pNumberRequired = 20

      'Initialise all lines in array to Organisation Number 0, Distance 999999999.99
      For vIndex = 1 To pNumberRequired
        vNearest(1, vIndex) = 0
        vNearest(2, vIndex) = 999999999.99
      Next
      'Process each Postcode Proximity Organisation
      For Each vPPO In mvCol
        'For Northern Ireland addresses (postcodes beginning with BT) only include those organisations that are located in Northern Ireland.
        'For all other postcodes exclude any organisations that are located in Northern Ireland.
        If ((Left(pPostCode, 2) = "BT" And Left(vPPO.PostCode, 2) = "BT") Or (Left(pPostCode, 2) <> "BT" And Left(vPPO.PostCode, 2) <> "BT")) Then
          vDistance = System.Math.Sqrt((pEasting - vPPO.Easting) ^ 2 + (pNorthing - vPPO.Northing) ^ 2)
          For vIndex = 1 To pNumberRequired
            If vDistance < vNearest(2, vIndex) Then
              'Shuffle any below this row down before adding this value
              For vIndex2 = pNumberRequired To vIndex Step -1
                vNearest(1, vIndex2) = vNearest(1, vIndex2 - 1)
                vNearest(2, vIndex2) = vNearest(2, vIndex2 - 1)
              Next
              vNearest(1, vIndex) = vPPO.OrganisationNumber
              vNearest(2, vIndex) = vDistance
              Exit For
            End If
          Next
        End If
      Next vPPO
      'Load the value of the array into a module-level data table
      mvNearest = New CDBDataTable
      With mvNearest
        .AddColumn("OrganisationNumber", CDBField.FieldTypes.cftLong)
        .AddColumn("Distance", CDBField.FieldTypes.cftNumeric)
      End With
      For vIndex = 1 To pNumberRequired
        If vNearest(1, vIndex) > 0 Then
          vDR = mvNearest.AddRow
          vDR.Item("OrganisationNumber") = CStr(vNearest(1, vIndex))
          vDR.Item("Distance") = CStr(vNearest(2, vIndex))
        End If
      Next
    End Sub
    Public Sub GetNearest(ByVal pItemNumber As Integer, ByRef pOrganisationNumber As Integer, ByRef pDistance As Double)
      If mvNearest.Rows.Count() >= pItemNumber Then
        pOrganisationNumber = CInt(mvNearest.Rows.Item(pItemNumber - 1).Item(1))
        pDistance = CDbl(mvNearest.Rows.Item(pItemNumber - 1).Item(2))
      Else
        pOrganisationNumber = 0
        pDistance = 999999999.99
      End If
    End Sub
    Public ReadOnly Property NearestCount() As Integer
      Get
        If mvNearest Is Nothing Then mvNearest = New CDBDataTable
        NearestCount = mvNearest.Rows.Count()
      End Get
    End Property

    'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub Class_Initialize_Renamed()
      mvCol = New Collection
    End Sub
    Public Sub New()
      MyBase.New()
      Class_Initialize_Renamed()
    End Sub

  End Class
End Namespace
