Attribute VB_Name = "FunctionData"
Option Explicit

Function GetConnectionString() As String
    GetConnectionString = "Provider=SQLOLEDB.1;Password=" & CONST_PWD & "  ;Persist Security Info=True;User ID=" & CONST_UID & ";Initial Catalog=" & CONST_DATABASE & ";Data Source=" & CONST_SERVER & ""
    
 

End Function

Public Function SQLStaticRecordset(ByVal strSQL As String, ParamArray Param() As Variant) As ADODB.Recordset

    Dim RS  As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim pmt As ADODB.Parameter
    Dim i   As Long
    
    On Error GoTo ErrHandler

    Set cmd = New ADODB.Command
    Set RS = New ADODB.Recordset
    Set pmt = New ADODB.Parameter
    
    With cmd
        .ActiveConnection = GetConnectionString()
        .CommandType = adCmdText
        .CommandText = strSQL
         For Each pmt In .Parameters
            cmd.Parameters(i).Value = Param(i) ''pmt.Value = Param(i)
            i = i + 1
         Next
    End With

    With RS
        .CursorLocation = adUseClient
        .Open cmd, , adOpenStatic
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
        Set .ActiveConnection = Nothing
    End With

    Set SQLStaticRecordset = RS
    Exit Function

ErrHandler:
    Set SQLStaticRecordset = Nothing
    Set RS = Nothing
    Set cmd = Nothing
    Set pmt = Nothing
    MsgBox Err.Description
    'RaiseError str_Module, "GetEditRecordset(" & strSQL & ", ...)"
End Function

Public Function SQLEditRecordsetP(ByVal strSQL As String, ParamArray Param() As Variant) As ADODB.Recordset

    Dim RS  As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim pmt As ADODB.Parameter
    Dim i   As Long
    
    On Error GoTo ErrHandler

    Set cmd = New ADODB.Command
    Set RS = New ADODB.Recordset
    Set pmt = New ADODB.Parameter
    
    With cmd
        .ActiveConnection = GetConnectionString()
        .CommandType = adCmdText
        .CommandText = strSQL
         For Each pmt In .Parameters
            cmd.Parameters(i).Value = Param(i) ''pmt.Value = Param(i)
            i = i + 1
         Next
    End With

    With RS
        .CursorLocation = adUseClient
        .Open cmd, , adOpenKeyset, adLockOptimistic
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
        Set .ActiveConnection = Nothing
    End With

    Set SQLEditRecordsetP = RS
    Exit Function

ErrHandler:
    Set SQLEditRecordsetP = Nothing
    Set RS = Nothing
    Set cmd = Nothing
    Set pmt = Nothing
    'RaiseError str_Module, "GetEditRecordset(" & strSQL & ", ...)"
End Function
Public Function GetEditRecordset(ByVal strSQL As String) As ADODB.Recordset
    On Error GoTo ErrHandler
    
    ' Create the ADO objects
    Dim RS As ADODB.Recordset, cmd As ADODB.Command
    Set RS = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    ' Init the ADO objects  & the stored proc parameters
    cmd.ActiveConnection = GetConnectionString
    cmd.CommandText = strSQL
    cmd.CommandType = adCmdText
    
    RS.Open cmd, , adOpenKeyset, adLockOptimistic
    Set GetEditRecordset = RS
    Exit Function

ErrHandler:
    Set GetEditRecordset = Nothing
    Set RS = Nothing
    Set cmd = Nothing
  '  RaiseError str_Module, "GetEditRecordset(" & strSQL & ", ...)"
End Function

Public Function RSRequery(ByRef RSRequeryNew As ADODB.Recordset, ParamArray Param() As Variant)

    'Dim RS  As ADODB.Recordset
    Dim cmd As ADODB.Command
    Dim pmt As ADODB.Parameter
    Dim i   As Long
    
    On Error GoTo ErrHandler

    Set cmd = New ADODB.Command
    'Set RS = New ADODB.Recordset
    Set pmt = New ADODB.Parameter
    
    With cmd
        .ActiveConnection = GetConnectionString()
        .CommandType = adCmdText
       ' .CommandText = strSQL
         For Each pmt In .Parameters
            cmd.Parameters(i).Value = Param(i) ''pmt.Value = Param(i)
            i = i + 1
         Next
    End With

    'With RS
     '   .CursorLocation = adUseClient
     '   .Open cmd, , adOpenKeyset, adLockOptimistic
        RSRequeryNew.Requery
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
        Set RSRequeryNew.ActiveConnection = Nothing
    'End With

    'Set SQLEditRecordsetP = RS
    Exit Function

ErrHandler:
    'Set SQLEditRecordsetP = Nothing
    'Set RS = Nothing
    Set cmd = Nothing
    Set pmt = Nothing
    'RaiseError str_Module, "GetEditRecordset(" & strSQL & ", ...)"
End Function
