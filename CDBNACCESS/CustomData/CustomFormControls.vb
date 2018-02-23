Public Class CustomFormControls
  Inherits CollectionList(Of CustomFormControl)

  Protected mvMaxHeight As Long

  Public Overloads Function AddFromRecordSet(ByVal pEnv As CDBEnvironment, ByVal pRecordSet As CDBRecordSet, ByVal pStartNumberingAt As Integer) As CustomFormControl
    Dim vCustomFormControl As New CustomFormControl(pEnv)
    vCustomFormControl.InitFromRecordSet(pRecordSet)
    If vCustomFormControl.ParameterName.Length = 0 Then
      Dim vBaseName As String = ProperName(vCustomFormControl.AttributeName)
      If pStartNumberingAt > 0 Then vBaseName = vBaseName & pStartNumberingAt
      Dim vName As String
      If MyBase.ContainsKey(vBaseName) Then
        Dim vCount As Integer
        If pStartNumberingAt > 0 Then
          vBaseName = Substring(vBaseName, 0, vBaseName.Length - pStartNumberingAt.ToString.Length)
          vCount = pStartNumberingAt
        Else
          vCount = 1                              'Start with 2
        End If
        Do
          vCount = vCount + 1
          vName = vBaseName & vCount
        Loop While MyBase.ContainsKey(vName)
      Else
        vName = vBaseName
      End If
      vCustomFormControl.ParameterName = vName
    End If
    If vCustomFormControl.ControlTop + vCustomFormControl.ControlHeight > mvMaxHeight Then mvMaxHeight = vCustomFormControl.ControlTop + vCustomFormControl.ControlHeight
    MyBase.Add(vCustomFormControl.ParameterName, vCustomFormControl)
    Return vCustomFormControl
  End Function

  Public Sub Init(ByVal pEnv As CDBEnvironment, ByVal pCustomForm As Integer)
    Dim vCustomForm As New CustomForm(pEnv)
    vCustomForm.Init(pCustomForm)
    If Not (pEnv.GetDataStructureInfo(CDBEnvironment.cdbDataStructureConstants.cdbCustomFormWebPage) AndAlso vCustomForm.CustomFormUrl.Length > 0) Then
      'Jira 1414: Only get custom form controls when the custom_form_url value is not set (i.e. when we are not displaying a web browser instead of the controls)
      Dim vCaptionWidth As Integer = vCustomForm.TagWidth
      Dim vWhereFields As New CDBFields
      vWhereFields.Add("custom_form", pCustomForm)
      Dim vAnsiJoins As New AnsiJoins
      Dim vCustomFormControl As New CustomFormControl(pEnv)
      Dim vMaintenanceAttribute As New MaintenanceAttribute(pEnv)
      vMaintenanceAttribute.InitForAdditionalInfo(True)
      vAnsiJoins.AddLeftOuterJoin("maintenance_attributes ma", "cfc.table_name", "ma.table_name", "cfc.attribute_name", "ma.attribute_name")
      Dim vSQL As New SQLStatement(pEnv.Connection, vCustomFormControl.GetRecordSetFields & "," & vMaintenanceAttribute.GetRecordSetFields, "custom_form_controls cfc", vWhereFields, "tab_number, cfc.sequence_number", vAnsiJoins)
      Dim vRecordSet As CDBRecordSet = vSQL.GetRecordSet
      While vRecordSet.Fetch()
        vCustomFormControl = AddFromRecordSet(pEnv, vRecordSet, 0)
        vCustomFormControl.AdjustWidth(vCaptionWidth)
      End While
      vRecordSet.CloseRecordSet()
    End If
  End Sub

  Public Function TableNames() As CDBParameters
    Dim vTables As New CDBParameters
    For Each vControl As CustomFormControl In Me
      If vControl.TableName.Length > 0 Then
        If Not vTables.Exists(vControl.TableName) Then vTables.Add(vControl.TableName)
      End If
    Next vControl
    Return vTables
  End Function

  Public Function Fields(ByVal pEnv As CDBEnvironment, ByVal pTableName As String, ByVal pParams As CDBParameters) As CDBFields
    Dim vFields As New CDBFields

    For Each vControl As CustomFormControl In Me
      If vControl.TableName = pTableName Then
        If pParams.Exists((vControl.ParameterName)) Then
          If vControl.ParameterName = "AmendedOn" And vControl.DefaultValue = "#DATE" Then
            vFields.Add(vControl.AttributeName, vControl.MaintenanceAttribute.FieldType, TodaysDate())
          ElseIf vControl.ParameterName = "AmendedBy" And vControl.DefaultValue = "#USER" Then
            vFields.Add(vControl.AttributeName, vControl.MaintenanceAttribute.FieldType, pEnv.User.UserID)
          Else
            vFields.Add(vControl.AttributeName, vControl.MaintenanceAttribute.FieldType, pParams((vControl.ParameterName)).Value)
          End If
        End If
      End If
    Next vControl
    Fields = vFields
  End Function

  Public Function WhereFields(ByRef pTableName As String, ByVal pParams As CDBParameters) As CDBFields
    Dim vWhereFields As New CDBFields

    For Each vControl As CustomFormControl In Me
      If vControl.TableName = pTableName And vControl.MaintenanceAttribute.PrimaryKey Then
        If pParams.Exists(vControl.ParameterName) Then
          vWhereFields.Add(vControl.AttributeName, vControl.MaintenanceAttribute.FieldType, pParams(vControl.ParameterName).Value)
        End If
      End If
    Next vControl
    Return vWhereFields
  End Function
End Class
