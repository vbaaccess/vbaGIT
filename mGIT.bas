Option Compare Database
Option Explicit

Private Const CurrentModeName = "mGIT"
Public Const vbaGIVersionNumber = "200602.134825"

Public GIT As New clsGIT
Public GIVersionNumber As New clsUpdateVersionNumber

Public Function UpdateNumber()
    GIVersionNumber.Update
End Function
