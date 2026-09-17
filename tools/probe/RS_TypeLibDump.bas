Attribute VB_Name = "RS_TypeLibDump"
Option Explicit
'=============================================================================
' Reads the MS Project TYPE LIBRARY and writes what's in it as text.
'
' WHY THIS INSTEAD OF THE PROBE
'
' RS_ModelProbe asks "can I read this member?", which is not the question. A
' 438 comes back for an absent member, for an object with no default property,
' and for a write-only or hidden one - and Project.NewTasksAreManual is exactly
' that last case: undocumented, unreadable, and definitely real, because our own
' code assigns to it from an early-bound variable and compiles.
'
' The type library is the actual answer. It is what VBA itself compiles against,
' so what it lists IS what exists - including hidden members the documentation
' omits. And nothing gets invoked to find out.
'
' TWO ROUTES. It tries both, in order:
'
'   A. TypeLib Information (TLI / tlbinf32.dll). Enumerates every interface and
'      member with its invoke kind (get / put / method) and argument counts -
'      everything the model needs, properly structured. Only present on some
'      machines, and 32-bit only; this Project reports a 32-bit Office, so it
'      has a chance. Reported as unavailable rather than failing if it isn't
'      registered.
'
'   B. Identifier scan of MSPRJ.OLB. Every type and member name is stored as
'      text in the library's string heap, so reading the file and pulling out
'      identifier-shaped runs gives the complete set of REAL names, with no
'      dependency on anything being installed. It loses the structure - you
'      can't tell a property from a method - but a name either appears or it
'      doesn't, which settles existence.
'
' Route B runs only when route A doesn't. TLI already reports every name
' WITH its invoke kind, and the identifier scan is a few hundred KB of text -
' which matters when the only way back from work is the body of an email.
' ScanOlbOnly runs it deliberately.
'
' TO RUN:
'   1. Alt+F11, Insert > Module, paste this in
'   2. click inside DumpTypeLib and press F5
'   3. it writes a text file and shows you the path
'
' READ-ONLY, and this time by construction rather than by assertion: it opens
' one file for binary input and reads a COM type library. It never touches your
' project, and it does not call a single Project method - the whole reason the
' last one closed Project.
'=============================================================================

Private mOut As String
Private mCount As Long

Public Sub DumpTypeLib()
    Dim app As Object, olb As String, outPath As String

    Set app = Application
    olb = FindOlb(app)

    mOut = ""
    mCount = 0
    Say "RS TYPELIB DUMP v1"
    Say "olb=" & olb
    Say "office=" & HostBit(app)
    Say ""

    If Len(olb) = 0 Then
        MsgBox "Couldn't locate MSPRJ.OLB. Nothing to read.", vbExclamation, _
               "RS TypeLib Dump"
        Exit Sub
    End If

    ' Route B is a fallback, not a supplement. When TLI works it has already
    ' given us every name WITH its invoke kind and argument counts, and the
    ' identifier scan would add a few hundred KB of text to an email for
    ' nothing. You can't attach files at work, so size is the constraint.
    If TryTli(olb) Then
        Say "[NAMES]"
        Say "skipped - TLI above already covers every name, with structure."
        Say "Run ScanOlbOnly if you want the raw identifier list as well."
    Else
        ScanOlb olb
    End If

    outPath = OutFolder(app) & "\RS_TypeLib.txt"
    If WriteOut(outPath) Then
        MsgBox "Written to:" & vbCrLf & vbCrLf & outPath & vbCrLf & vbCrLf & _
               mCount & " names found. Send that file back.", _
               vbInformation, "RS TypeLib Dump"
    Else
        MsgBox "Couldn't write the file - the report is in the Immediate " & _
               "window (Ctrl+G) instead.", vbExclamation, "RS TypeLib Dump"
        Debug.Print mOut
    End If
End Sub

' Route B on demand, when route A worked but the raw names are wanted anyway.
Public Sub ScanOlbOnly()
    Dim app As Object, olb As String, outPath As String
    Set app = Application
    olb = FindOlb(app)
    mOut = ""
    mCount = 0
    Say "RS TYPELIB DUMP v1 - identifier scan only"
    Say "olb=" & olb
    Say ""
    If Len(olb) = 0 Then Exit Sub
    ScanOlb olb
    outPath = OutFolder(app) & "\RS_TypeLibNames.txt"
    If WriteOut(outPath) Then
        MsgBox "Written to:" & vbCrLf & vbCrLf & outPath & vbCrLf & vbCrLf & _
               mCount & " identifiers. This one is large - a few hundred KB.", _
               vbInformation, "RS TypeLib Dump"
    End If
End Sub

'--- route A: structured, if TLI happens to be registered --------------------
' InvokeKinds: 1 = method, 2 = property get, 4 = property put, 8 = put by ref.
' That distinction is the thing the probe could never see, and it's what makes
' a write-only member identifiable instead of indistinguishable from an absent
' one.
Private Function TryTli(ByVal olb As String) As Boolean
    Dim tli As Object, tlb As Object, ti As Object, mi As Object
    Dim kinds As String, n As Long

    Say "[TLI]"
    On Error Resume Next
    Set tli = CreateObject("TLI.TLIApplication")
    If tli Is Nothing Or Err.Number <> 0 Then
        Say "unavailable - tlbinf32.dll isn't registered here (err " & _
            Err.Number & "). Route B below still covers existence."
        Err.Clear
        On Error GoTo 0
        Say ""
        Exit Function
    End If

    Set tlb = tli.TypeLibInfoFromFile(olb)
    If Err.Number <> 0 Then
        Say "couldn't open the library: err " & Err.Number & " " & Err.Description
        Err.Clear
        On Error GoTo 0
        Say ""
        Exit Function
    End If

    Say "name=" & tlb.Name & "|ver=" & tlb.MajorVersion & "." & tlb.MinorVersion
    Say "# class|member|kind|args|optional"

    For Each ti In tlb.TypeInfos
        For Each mi In ti.Members
            Select Case mi.InvokeKind
                Case 1: kinds = "METHOD"
                Case 2: kinds = "GET"
                Case 4: kinds = "PUT"
                Case 8: kinds = "PUTREF"
                Case Else: kinds = "KIND" & mi.InvokeKind
            End Select
            Say ti.Name & "|" & mi.Name & "|" & kinds & "|" & _
                mi.Parameters.Count & "|" & mi.Parameters.OptionalCount
            n = n + 1
        Next mi
    Next ti
    Err.Clear
    On Error GoTo 0

    Say "# " & n & " members enumerated"
    mCount = mCount + n
    TryTli = (n > 0)
    Say ""
End Function

'--- route B: the names, straight out of the file ----------------------------
' Type libraries store names as plain single-byte strings, so every type and
' member name is sitting in the file as text. Pulling out identifier-shaped
' runs gives a superset of the real names - some unrelated strings come along -
' but nothing real is missed, which is what makes it useful for settling
' whether a member exists.
Private Sub ScanOlb(ByVal olb As String)
    Dim f As Integer, buf As String, i As Long, ch As Long
    Dim cur As String, seen As Object, n As Long

    Say "[NAMES]"
    On Error GoTo Nope

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 0                     ' case-sensitive: casing matters

    f = FreeFile
    Open olb For Binary Access Read As #f
    buf = Space$(LOF(f))
    Get #f, , buf
    Close #f

    For i = 1 To Len(buf)
        ch = Asc(Mid$(buf, i, 1))
        If (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Or _
           (ch >= 48 And ch <= 57) Or ch = 95 Then
            cur = cur & Chr$(ch)
        Else
            If Len(cur) >= 3 And Len(cur) <= 64 Then
                ' must start with a letter to be an identifier
                ch = Asc(Left$(cur, 1))
                If (ch >= 65 And ch <= 90) Or (ch >= 97 And ch <= 122) Then
                    If Not seen.Exists(cur) Then
                        seen.Add cur, 1
                        n = n + 1
                    End If
                End If
            End If
            cur = ""
        End If
    Next i

    Dim k As Variant
    For Each k In seen.Keys
        Say CStr(k)
    Next k

    Say "# " & n & " distinct identifiers"
    mCount = mCount + n
    Exit Sub

Nope:
    On Error Resume Next
    Close #f
    Say "scan failed: err " & Err.Number & " " & Err.Description
    Err.Clear
End Sub

'--- where things are -------------------------------------------------------
Private Function FindOlb(ByVal app As Object) As String
    Dim r As Object, p As String

    ' The references list knows the exact path, and needs the same Trust Center
    ' setting the rest of this project already needs.
    On Error Resume Next
    For Each r In app.VBE.ActiveVBProject.References
        If StrComp(r.Name, "MSProject", vbTextCompare) = 0 Then
            FindOlb = r.FullPath
            Err.Clear
            Exit Function
        End If
    Next r
    Err.Clear

    ' Failing that, next to the running Project.
    p = app.Path & "\MSPRJ.OLB"
    If Len(Dir$(p)) > 0 Then FindOlb = p
    Err.Clear
    On Error GoTo 0
End Function

Private Function HostBit(ByVal app As Object) As String
    On Error Resume Next
    HostBit = app.Version & " build " & app.Build & " / " & app.OperatingSystem
    Err.Clear
End Function

Private Function OutFolder(ByVal app As Object) As String
    On Error Resume Next
    OutFolder = app.ActiveProject.Path
    Err.Clear
    If Len(OutFolder) = 0 Then OutFolder = Environ$("TEMP")
    If Len(OutFolder) = 0 Then OutFolder = Environ$("USERPROFILE")
    On Error GoTo 0
End Function

'--- output -----------------------------------------------------------------
Private Sub Say(ByVal s As String)
    mOut = mOut & s & vbCrLf
End Sub

Private Function WriteOut(ByVal outPath As String) As Boolean
    Dim f As Integer
    On Error GoTo Nope
    f = FreeFile
    Open outPath For Output As #f
    Print #f, mOut
    Close #f
    WriteOut = True
    Exit Function
Nope:
    On Error Resume Next
    Close #f
    WriteOut = False
End Function
