Attribute VB_Name = "ModLoader"
Sub Main()
Dim iFile As Long, iPos As Long, ExeHeadFile As String, sBuff1 As String, sHeader As String _
, StrBuff2 As String, DatFile As String, sFormFileA As String

    iFile = FreeFile
    
    ExeHeadFile = FixPath(App.Path) & App.EXEName & ".exe" ' the path and file name of the exe were looking at now
    
    Open ExeHeadFile For Binary As #iFile
        sBuff1 = Space(LOF(iFile))
        Get #iFile, , sBuff1
    Close #iFile
    
    sHeader = "<DATA>" & Chr(5)
    iPos = InStr(1, sBuff1, sHeader, vbBinaryCompare)
    
    If iPos <= 0 Then
        MsgBox "Unable to locate the main data stream.", vbCritical, "Invaild Data Stream"
        End
    End If
    
    StrBuff2 = Mid(sBuff1, iPos + Len(sHeader), Len(sBuff1))
    ExeHeadFile = ""
    sHeader = ""
    sBuff1 = ""
    iPos = 0
    
    DatFile = DMGetTempPath & "gTmp.o"
    sFormFileA = DMGetTempPath & "gTmp.bfm"
    
    Open DatFile For Binary As #1
        Put #1, , StrBuff2
    Close #1
    
    Open DatFile For Binary As #2
        Get #2, , MakeExe
    Close #2
    
    Kill DatFile
    
    Open sFormFileA For Binary As #3
        Put #3, , MakeExe.Win32FormData
    Close #3
    
    TProject.UnitFile = Encrypt(MakeExe.Win32CodeData)
    TProject.ProgLan = MakeExe.Win32Lan
    
    If Not LoadForm(frmWorkArea, sFormFileA) Then
        MsgBox "There was an error while loading the appliaction.", vbCritical, "Appliaction Error"
        End
    End If
    
    Kill sFormFileA
    MakeExe.Win32CodeData = ""
    MakeExe.Win32FormData = ""
    MakeExe.Win32Lan = ""
    DatFile = ""
    sFormFileA = ""
    StrBuff2 = ""
    
    ' setup the script
    Set dScript.DialogObject = frmWorkArea
    dScript.DialogStrName = "Dialog"
    dScript.mLanguage = TProject.ProgLan
    dScript.SetupControl
    dScript.RunCode TProject.UnitFile
    frmWorkArea.Show
End Sub

