Attribute VB_Name = "modAnalyse"
Dim ID As Integer

Public Sub StartAnalyse(lvw As ListView)
   lvw.ListItems.Clear
   BootStrap lvw
   ID = 0
End Sub

Sub BootStrap(lvw As ListView)
Dim sLine As String
Dim sTmp As String
Dim iFile As Integer

   iFile = FreeFile
   
   Open sFile For Input As iFile
      Do Until sLine = "[Bootstrap]"
      
         Line Input #iFile, sLine
      
      Loop
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      If sTmp = "" Then
         lvw.ListItems.Add , , "SetupTitle", , 1
         ID = ID + 1
         lvw.ListItems(ID).ListSubItems.Add , , "There are no setup title to this project, this is not needed."
      End If
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      If sTmp = "" Then
         lvw.ListItems.Add , , "SetupText", , 1
         ID = ID + 1
         lvw.ListItems(ID).ListSubItems.Add , , "There are no setup text to this project, this is not needed, but it's a good idea to include it."
      End If
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      If sTmp = "" Then
         lvw.ListItems.Add , , "CabFile", , 2
         ID = ID + 1
         lvw.ListItems(ID).ListSubItems.Add , , "No cab file. If you have no cab file, it will increase the amount of space this setup takes, include a cab file."
      End If
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      If sTmp <> "Setup1.exe" Then
         lvw.ListItems.Add , , "Spawn", , 2
         ID = ID + 1
         lvw.ListItems(ID).ListSubItems.Add , , "You do not start the Setup1.exe after pre install, if not configured propably this might cause problems."
      End If
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      If sTmp = "" Then
         lvw.ListItems.Add , , "Uninstall", , 1
         ID = ID + 1
         lvw.ListItems(ID).ListSubItems.Add , , "You have no uninstall file."
      End If
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      If sTmp = "" Then
         lvw.ListItems.Add , , "TmpDir", , 2
         ID = ID + 1
         lvw.ListItems(ID).ListSubItems.Add , , "Without a temp dir, the cab file will not work."
      End If
      
      Line Input #iFile, sLine
      sTmp = CutRight(sLine)
      If sTmp = "" Then
         lvw.ListItems.Add , , "Cabs", , 2
         ID = ID + 1
         lvw.ListItems(ID).ListSubItems.Add , , "You have no cabs."
      End If
      
   Close iFile
End Sub
