Attribute VB_Name = "ModTools"
Enum EditOp
    nCut = 1
    nCopy
    nPaste
    nSelectAll
    nDelete
End Enum

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Function RemoveControlButton(Frm As Form, dMnuPosition As Integer, En As Boolean)
Dim hMenu As Long, iRet As Long
    hMenu = GetSystemMenu(Frm.hwnd, En)
    iRet = DeleteMenu(hMenu, dMnuPosition, MF_BYPOSITION)
End Function

Public Function IsFileHere(lzFilename As String) As Boolean
    If Dir(lzFilename) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function FixPath(lzPath As String) As String
    If Right(FixPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Function FindDir(lzPath As String) As Boolean
    If Not Dir(FixPath(lzPath), vbDirectory) = "." Then
        FindDir = False
        Exit Function
    Else
        FindDir = True
    End If
End Function

Public Function EnablePaste() As Boolean
   If Len(Clipboard.GetText(vbCFText)) > 0 Then
        EnablePaste = True
    Else
        EnablePaste = False
    End If
End Function

Public Function EditMenu(mOption As EditOp, txtBox As TextBox)
    Select Case mOption
        Case nCut
            Clipboard.SetText txtBox.SelText
            txtBox.SelText = ""
        Case nCopy
            Clipboard.SetText txtBox.SelText
        Case nPaste
            txtBox.SelText = Clipboard.GetText(vbCFText)
        Case nSelectAll
            txtBox.SelStart = 0
            txtBox.SelLength = Len(txtBox.Text)
            txtBox.SetFocus
        Case nDelete
            txtBox.SelText = ""
    End Select
End Function

Public Sub DrawGrid(Frm As Form, Optional GridColor As Long = vbBlack)
    Frm.AutoRedraw = True
    For X = 0 To Frm.ScaleWidth Step 115
        For Y = 0 To Frm.ScaleHeight Step 115
            Frm.PSet (X, Y), GridColor
        Next
    Next
    Frm.Refresh
    Frm.AutoRedraw = False
    X = 0: Y = 0
End Sub

Public Function OpenFile(lzFilename As String) As String
Dim StrA As String, iFile As Long
    iFile = FreeFile
        
    Open lzFilename For Binary As #iFile
        StrA = Space(LOF(iFile))
        Get #iFile, , StrA
    Close #iFile
    
    OpenFile = StrA
    StrA = ""
End Function

Public Sub lstVBFunctions(lzCode As String, cboLst As ComboBox)
Dim icnt As Long, Ipart, lPart As Long, X As Long, Y As Long, ch As Long
Dim LnStr, Strbuff As String, FuncName As String, SubName As String, strln As String
On Error Resume Next
' VB Function names Added by Ben jones
' Ok this does work to a level but may need touching up in parts but as am example it will ok


    cboLst.Clear
    cboLst.AddItem "(General)"
    Strbuff = lzCode & vbCrLf
   
    For icnt = 1 To Len(Strbuff)
        ch = Asc(Mid$(Strbuff, icnt, 1))
        If ch <> 13 Then
            strln = strln & Chr(ch)
        Else
    
        
            Ipart = InStr(1, strln, "Function ", vbTextCompare) ' Start of function name
            lPart = InStr(1, strln, "(")    ' End of function name
            If Ipart > 0 And lPart > 0 Then
                FuncName = Trim$(Mid$(strln, Ipart + Len("Function"), lPart - Ipart - Len("Function")))
                cboLst.AddItem " " & FuncName
            End If
            
            X = InStr(1, strln, "Sub ", vbTextCompare)
           
            
            Y = InStr(1, strln, "(")
            If X > 0 And Y > 0 Then
                SubName = Trim(Mid(strln, X + Len("Sub"), Y - X - Len("Sub")))
                cboLst.AddItem " " & SubName
            End If
            strln = ""
            icnt = icnt + 1
        End If
    Next icnt
    cboLst.ListIndex = 0
    icnt = 0: Ipart = 0: lPart = 0: X = 0: Y = 0
    FuncName = ""
    Strbuff = ""
    LnStr = ""
    ch = ""
    
End Sub

Function GetAbsPath(lzPath As String) As String
Dim ipos As Long, I As Long
    For I = 1 To Len(lzPath)
        If InStr(I, lzPath, "\", vbBinaryCompare) Then
            ipos = I
        End If
    Next
    
    If ipos = 0 Then
        GetAbsPath = lzPath
    Else
        GetAbsPath = Mid(lzPath, 1, ipos)
    End If
    
    ipos = 0
End Function

Public Function WriteToFile(lzFile As String, lzData As String)
Dim iFile As Long
    iFile = FreeFile
    Open lzFile For Binary As #iFile
        Put #iFile, , lzData
    Close #iFile
    
End Function

Sub AppendErrorLog(StrError As String)
    Open FixPath(App.Path) & "error.log" For Append As #1
        Print #1, StrError
    Close #1
    
End Sub
