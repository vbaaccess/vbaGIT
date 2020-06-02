Option Compare Database
Option Explicit

Private Const CurrentModName = "clsGIT"

Private m_ObjectToExport As String

Private Sub Class_Initialize()
    '--- Konstruktor ---
End Sub

Private Sub Class_Terminate()
    '--- Destruktor ---
End Sub

Property Let ObjectToExport(ObjectName As String)
  m_ObjectToExport = ObjectName
End Property

Property Get ObjectToExport() As String
  ObjectToExport = m_ObjectToExport
End Property

Public Sub CheckForUpdates()
    'TO DO
End Sub

Public Sub Update()
    Debug.Print "Checking if there is a new version..."   'TO DO - zaczynamy
    
    'TO DO - sparwdzamy czy jest nowa wersja
    
    'TO DO - jesli jest uzytkownik musi potwierdzić
    
    Debug.Print "No New Version."
End Sub

Public Function Export()
    If Len(m_ObjectToExport) > 0 Then
        Debug.Print "Object export " & m_ObjectToExport
        Call ExportProcedure
    Else
        Debug.Print "Object not set"
    End If
End Function

Private Sub ExportProcedure()
    Dim sObjectName As String
    Dim ObjectType As AcObjectType
    Dim strCreateFile As String
    Dim strFolderToUse As String
    
    ObjectType = acForm
    sObjectName = m_ObjectToExport
    strFolderToUse = Application.CurrentProject.Path
    strCreateFile = "" & sObjectName & ".form.bas"
    
    Application.SaveAsText ObjectType, sObjectName, strFolderToUse & "\" & strCreateFile
End Sub
