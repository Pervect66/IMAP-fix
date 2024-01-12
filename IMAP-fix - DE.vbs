' Geïmporteerde IMAP structuur vervangen door IPF.Note
' © 2023 Jan Boezeman voor Microsoft M365 Support 

Dim i

MsgBox "Machen Sie Outlook-Elemente nach dem IMAP-Import sichtbar", vbInfo+vbSystemModal, "IMAP-fix"

Call FolderSelect()

Public Sub FolderSelect()
  Dim objOutlook
  Set objOutlook = CreateObject("Outlook.Application")

  Dim F, Folders
  Set F = objOutlook.Session.PickFolder

  If Not F Is Nothing Then
    Dim Result
    Result = MsgBox("Möchten Sie auch alle Unterordner reparieren??", vbYesNo+vbDefaultButton2+vbSystemModal, "Alle Ordner?")

    i = 0
    FixIMAPFolder(F)

    If Result = 6 Then
      Set Folders = F.Folders
      LoopFolders Folders
    End If

    Result = MsgBox("Fertig!" & vbNewLine & i & " Alle Ordner wurden repariert!", vbInfo+vbSystemModal, "IMAP-Ordner wiederhergestellt")
  
    Set F = Nothing
    Set Folders = Nothing
    Set objOutlook = Nothing
  End If
End Sub

Private Sub LoopFolders(Folders)
  Dim F
  
  For Each F In Folders
    FixIMAPFolder(F)
    LoopFolders F.Folders
  Next
End Sub

Private Sub FixIMAPFolder(F)
  Dim oPA, PropName, Value, FolderType

  PropName = "http://schemas.microsoft.com/mapi/proptag/0x3613001E"
  Value = "IPF.Note"

  On Error Resume Next
  Set oPA = F.PropertyAccessor
  FolderType = oPA.GetProperty(PropName)

  'MsgBox (F.Name & " - " & FolderType)

  If FolderType = "IPF.Imap" Then
    oPA.SetProperty PropName, Value
    i = i + 1
  End If

  Set oPA = Nothing
End Sub