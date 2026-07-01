Attribute VB_Name = "ModLog"
Option Explicit

Private m_logFile As String
Private m_logEnabled As Boolean

Public Sub IniciarLogDesarrollo(Optional ByVal habilitar As Boolean = True)
    On Error Resume Next
    Dim logDir As String
    Dim fecha As String
    
    m_logEnabled = habilitar
    If Not m_logEnabled Then Exit Sub
    
    logDir = App.Path & "\Logs"
    If Dir(logDir, vbDirectory) = "" Then
        MkDir logDir
    End If
    
    fecha = Format(Date, "yyyymmdd")
    m_logFile = logDir & "\OSBusiness_" & fecha & ".log"
    
    Dim f As Integer
    f = FreeFile
    Open m_logFile For Append As #f
    Print #f, "========================================"
    Print #f, "OSBusiness - Inicio de Sesion: " & Format(Now, "dd/mm/yyyy HH:MM:ss")
    Print #f, "Usuario: " & Environ("USERNAME")
    Print #f, "PC: " & Environ("COMPUTERNAME")
    Print #f, "========================================"
    Close #f
End Sub

Public Sub LogInfo(ByVal mensaje As String, Optional ByVal modulo As String = "GENERAL")
    If Not m_logEnabled Then Exit Sub
    If m_logFile = "" Then Exit Sub
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open m_logFile For Append As #f
    Print #f, Format(Now, "HH:MM:ss") & " [INFO] [" & modulo & "] " & mensaje
    Close #f
End Sub

Public Sub LogError(ByVal mensaje As String, Optional ByVal modulo As String = "GENERAL")
    If Not m_logEnabled Then Exit Sub
    If m_logFile = "" Then Exit Sub
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open m_logFile For Append As #f
    Print #f, Format(Now, "HH:MM:ss") & " [ERROR] [" & modulo & "] " & mensaje
    Close #f
End Sub

Public Sub LogTrace(ByVal mensaje As String, Optional ByVal modulo As String = "GENERAL")
    If Not m_logEnabled Then Exit Sub
    If m_logFile = "" Then Exit Sub
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open m_logFile For Append As #f
    Print #f, Format(Now, "HH:MM:ss") & " [TRACE] [" & modulo & "] " & mensaje
    Close #f
End Sub

Public Function ObtenerRutaLog() As String
    ObtenerRutaLog = m_logFile
End Function
