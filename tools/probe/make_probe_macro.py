#!/usr/bin/env python3
"""Generate tools/probe/RS_ModelProbe.bas from the Project model.

The probe is how the curated model gets checked against a real MS Project.
Nothing here can be verified from a Mac: the model is assembled from Microsoft's
published reference, and published references drift from shipped type libraries.
So the macro asks Project itself, member by member, and writes a text file to
bring back.

Generated rather than hand-written so the probe and the model can't disagree
about what's being tested - the failure mode where a check silently stops
covering what it claims to cover is the one that cost us a week of form builds.

VBA constraints that shape the output:
  - no line continuation past 25 lines, and lines die well before 1024 chars,
    so member names go out as pipe-delimited chunks fed through Split()
  - every Project member is reached late-bound through Object, so a member this
    version lacks is a runtime 438 we record, not a compile error that stops
    the whole module
  - PROPERTIES ONLY. CallByName with VbGet on a zero-argument method CALLS it,
    which is how version 1 of this generator ran LevelNow and Quit and closed
    the user's Project. Methods are reported as NOT-PROBED-METHOD; there is no
    way to test a method's existence without invoking it.
"""
import json
import os
import sys

ROOT = os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__))))
# The model now lives in the fork, where --host project reads it.
MODEL = os.path.join(ROOT, "vendor", "VBAlidator", "src", "models",
                     "project.json")
OUT = os.path.join(ROOT, "tools", "probe", "RS_ModelProbe.bas")

# Each probed class needs a live object to ask. Project can't hand us one by
# name, so the macro builds them by navigation - and every one of these is a
# read-only path.
SAMPLES = {
    "Project": 'Set o = app.ActiveProject',
    "Tasks": 'Set o = app.ActiveProject.Tasks',
    "Task": 'Set o = FirstTask(app)',
    "Resources": 'Set o = app.ActiveProject.Resources',
    "Resource": 'Set o = FirstRes(app)',
    "Assignments": 'Set o = FirstTask(app).Assignments',
    "Assignment": 'Set o = FirstAsg(app)',
}

# Never probed, whatever the model says about them. Anything that acts rather
# than reports, plus Task.Stop/Task.Resume - documented as dates, but the cost
# of skipping two properties is nothing against the cost of being wrong once.
DENY = set("""
    Quit Delete Save SaveAs Add Remove Clear Activate Level LevelNow
    LevelClearDates CalculateAll CalculateProject UndoClear
    OpenUndoTransaction CloseUndoTransaction Replan Split SetField AppendNotes
    CheckIn CheckOut CheckoutProject ExportAsFixedFormat SetCustomUI
    MakeServerURLTrusted HideCheckoutMsgBar DeliverablesClearAll
    DeliverableRefreshServerCache OutlineIndent OutlineOutdent
    OutlineHideSubTasks OutlineShowSubTasks OutlineShowAllTasks
    LinkPredecessors LinkSuccessors UnlinkPredecessors UnlinkSuccessors
    Pause Stop Resume Reset Refresh Apply Copy Cut Paste Print Close Open New
""".split())

CHUNK = 800


def chunks(names):
    out, cur = [], ""
    for n in names:
        if cur and len(cur) + len(n) + 1 > CHUNK:
            out.append(cur)
            cur = ""
        cur = n if not cur else cur + "|" + n
    if cur:
        out.append(cur)
    return out


def main():
    if not os.path.exists(MODEL):
        sys.exit(f"no model at {MODEL} - build it first")
    model = json.load(open(MODEL))
    classes = model.get("classes", {})

    lines = []
    w = lines.append
    w('Attribute VB_Name = "RS_ModelProbe"')
    w("Option Explicit")
    w("'" + "=" * 77)
    w("' Asks THIS copy of MS Project which object-model members it actually has,")
    w("' and writes the answers to a text file to bring back.")
    w("'")
    w("' WHY: the type model we check our VBA against was assembled from Microsoft's")
    w("' published reference. Published references drift from shipped type libraries,")
    w("' and a model that claims a member exists when it doesn't is worse than no")
    w("' model - it blesses a typo. This closes that gap with evidence.")
    w("'")
    w("' TO RUN:")
    w("'   1. Alt+F11 in MS Project")
    w("'   2. Insert > Module, paste all of this in")
    w("'   3. click inside ExportProjectModel and press F5")
    w("'   4. a MsgBox gives you the file path - send that file back")
    w("'")
    w("' Open a real project file first (one with tasks, resources and assignments),")
    w("' or the members that need a sample object can't be asked. Use a COPY.")
    w("'")
    w("' WHAT THIS READS, AND WHY THAT DISTINCTION MATTERS")
    w("'")
    w("' Version 1 of this macro claimed to be read-only because it only ever used")
    w("' CallByName with VbGet. That was wrong, and it closed MS Project.")
    w("'")
    w("'   CallByName(obj, name, VbGet) on a zero-argument METHOD does not read")
    w("'   anything. It CALLS it.")
    w("'")
    w("' Probing Application in alphabetical order therefore ran CalculateAll,")
    w("' CalculateProject, CloseUndoTransaction, LevelNow - which re-levels the whole")
    w("' project and can move dates - and then Quit, which is why Project vanished.")
    w("' It stopped there only because Application sorts before Task; another few")
    w("' names and it would have called Task.Delete the same way.")
    w("'")
    w("' So this version:")
    w("'   - probes PROPERTIES only. Methods are listed as NOT-PROBED-METHOD and")
    w("'     never touched, because there is no way to test whether a method exists")
    w("'     without running it.")
    w("'   - skips Application entirely. Its members ARE the command surface, and")
    w("'     nothing about it is type-checked anyway.")
    w("'   - refuses a hard-coded list of names whatever the model says about them,")
    w("'     in case a method is ever mislabelled as a property.")
    w("'")
    w("' Reads are still late-bound, so a member your version lacks is a runtime")
    w("' error this records rather than a compile error that stops the module.")
    w("'")
    w("' GENERATED by tools/make_probe_macro.py - edit the model, not this.")
    w("'" + "=" * 77)
    w("")
    w("Private mOut As String")
    w("")
    w("Public Sub ExportProjectModel()")
    w("    Dim app As Object, path As String, n As Long")
    w("    Set app = Application")
    w("")
    w("    ' Asking first, because the previous version of this macro closed")
    w("    ' Project without warning.")
    w("    If MsgBox(\"Read the object model of this project and write a report?\" & _")
    w("              vbCrLf & vbCrLf & \"Project: \" & ProjName(app) & vbCrLf & vbCrLf & _")
    w("              \"Properties are read; methods are NOT called. Nothing is \" & _")
    w("              \"written to your project.\" & vbCrLf & vbCrLf & _")
    w("              \"Even so, use a copy of the file if it matters.\", _")
    w("              vbOKCancel + vbQuestion, \"RS Model Probe\") <> vbOK Then Exit Sub")
    w("")
    w("    mOut = \"\"")
    w("    Say \"RS PROJECT MODEL PROBE v3 - properties only, Set retry\"")
    w("    Say \"generated-from: models/msproject.json\"")
    w("    Say \"\"")
    w("")
    w("    Header app")
    w("    Refs app")
    w("    Fields app")
    w("    n = Members(app)")
    w("")
    w("    Say \"\"")
    w("    Say \"END - \" & n & \" members probed\"")
    w("")
    w("    path = OutPath(app)")
    w("    If WriteOut(path) Then")
    w("        MsgBox \"Written to:\" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & _")
    w("               \"Send that file back. \" & n & \" members were probed.\", _")
    w("               vbInformation, \"RS Model Probe\"")
    w("    Else")
    w("        MsgBox \"Couldn't write the file. The report is in the Immediate \" & _")
    w("               \"window instead (Ctrl+G) - copy from there.\", _")
    w("               vbExclamation, \"RS Model Probe\"")
    w("        Debug.Print mOut")
    w("    End If")
    w("End Sub")
    w("")
    w("Private Function ProjName(ByVal app As Object) As String")
    w("    On Error Resume Next")
    w("    ProjName = app.ActiveProject.Name")
    w("    If Len(ProjName) = 0 Then ProjName = \"(none open)\"")
    w("    Err.Clear")
    w("End Function")
    w("")
    w("'--- what version are we even talking to ------------------------------------")
    w("Private Sub Header(ByVal app As Object)")
    w("    Say \"[HOST]\"")
    w("    SayTry app, \"Name\"")
    w("    SayTry app, \"Version\"")
    w("    SayTry app, \"Build\"")
    w("    SayTry app, \"OperatingSystem\"")
    w("    SayTry app, \"Path\"")
    w("    On Error Resume Next")
    w("    Say \"ActiveProject.Name=\" & app.ActiveProject.Name")
    w("    Say \"Tasks.Count=\" & app.ActiveProject.Tasks.Count")
    w("    Say \"Resources.Count=\" & app.ActiveProject.Resources.Count")
    w("    Err.Clear")
    w("    On Error GoTo 0")
    w("    Say \"\"")
    w("End Sub")
    w("")
    w("' One member, reported whatever happens. Late-bound, so a member this version")
    w("' hasn't got is an error string rather than a module that won't compile.")
    w("Private Sub SayTry(ByVal o As Object, ByVal nm As String)")
    w("    On Error Resume Next")
    w("    Dim v As Variant")
    w("    v = CallByName(o, nm, VbGet)")
    w("    If Err.Number = 0 Then")
    w("        Say nm & \"=\" & CStr(v)")
    w("    Else")
    w("        Say nm & \"=<err \" & Err.Number & \">\"")
    w("    End If")
    w("    Err.Clear")
    w("    On Error GoTo 0")
    w("End Sub")
    w("")
    w("'--- which type libraries this project references ---------------------------")
    w("' The exact GUID and version is what tells us which MSPRJ typelib the model")
    w("' should be matched against. Needs \"Trust access to the VBA project object")
    w("' model\" in Trust Center > Macro Settings; says so plainly when it's blocked.")
    w("Private Sub Refs(ByVal app As Object)")
    w("    Dim vbp As Object, r As Object, n As Long")
    w("    Say \"[REFERENCES]\"")
    w("    On Error Resume Next")
    w("    Set vbp = app.VBE.ActiveVBProject")
    w("    If vbp Is Nothing Or Err.Number <> 0 Then")
    w("        Say \"BLOCKED - no VBE access (Trust Center > Macro Settings >\"")
    w("        Say \"          Trust access to the VBA project object model)\"")
    w("        Err.Clear")
    w("        Say \"\"")
    w("        Exit Sub")
    w("    End If")
    w("    For Each r In vbp.References")
    w("        n = n + 1")
    w("        Say \"ref\" & n & \"|name=\" & r.Name & \"|guid=\" & r.GUID & _")
    w("            \"|ver=\" & r.Major & \".\" & r.Minor & \"|builtin=\" & r.BuiltIn")
    w("        Say \"      desc=\" & r.Description")
    w("        Say \"      path=\" & r.FullPath")
    w("    Next r")
    w("    Err.Clear")
    w("    On Error GoTo 0")
    w("    Say \"\"")
    w("End Sub")
    w("")
    w("'--- field names and the constants behind them ------------------------------")
    w("' Our code finds columns by their LABEL, so the mapping between a label, its")
    w("' Text number and its pj constant is the part we most need to be right about.")
    w("' pjStandard didn't exist and cost a module; this is where that gets caught.")
    w("Private Sub Fields(ByVal app As Object)")
    w("    Dim i As Long, k As Long, nm As String, c As Variant, loc As Variant")
    w("    Say \"[FIELDS]\"")
    w("    Say \"# name | FieldNameToFieldConstant(name, kind) per kind 0,1,2 |\"")
    w("    Say \"#        CustomFieldGetName(constant) = the label in THIS file\"")
    w("    On Error Resume Next")
    w("    For k = 0 To 2")
    w("        For i = 1 To 30")
    w("            nm = \"Text\" & i")
    w("            Err.Clear")
    w("            c = app.FieldNameToFieldConstant(nm, k)")
    w("            If Err.Number = 0 Then")
    w("                Err.Clear")
    w("                loc = app.CustomFieldGetName(c)")
    w("                If Err.Number <> 0 Then loc = \"\"")
    w("                Say \"kind\" & k & \"|\" & nm & \"|const=\" & c & \"|label=\" & loc")
    w("            Else")
    w("                Say \"kind\" & k & \"|\" & nm & \"|const=<err \" & Err.Number & \">\"")
    w("            End If")
    w("        Next i")
    w("    Next k")
    w("    Err.Clear")
    w("    On Error GoTo 0")
    w("    Say \"\"")
    w("End Sub")
    w("")
    w("'--- the actual point: does each member in our model exist here? ------------")
    w("Private Function Members(ByVal app As Object) As Long")
    w("    Dim o As Object, names As Variant, i As Long, n As Long")
    w("    Say \"[MEMBERS]\"")
    w("    Say \"# class|member|OK        member exists and reads\"")
    w("    Say \"#              ARGS      exists but needs arguments (still exists)\"")
    w("    Say \"#              MISSING   438 - not a member of this version\"")
    w("    Say \"#              ERR<n>    exists, read failed for another reason\"")

    for cls in sorted(classes):
        if cls not in SAMPLES:
            continue
        allm = classes[cls].get("members", {})
        # Properties only, and DENY on top of that. See the header comment: a
        # method here doesn't get read, it gets RUN.
        members = sorted(n for n, spec in allm.items()
                         if spec.get("type") != "Function" and n not in DENY)
        skipped = sorted(set(allm) - set(members))
        if not members:
            continue
        w("")
        w(f"    ' ---- {cls} ({len(members)} members)")
        w("    Set o = Nothing")
        w("    On Error Resume Next")
        w(f"    {SAMPLES[cls]}")
        w("    Err.Clear")
        w("    On Error GoTo 0")
        w(f'    If o Is Nothing Then')
        w(f'        Say "{cls}|<no sample object - open a project with tasks, '
          'resources and assignments>"')
        w("    Else")
        for ch in chunks(skipped):
            w(f'        names = Split("{ch}", "|")')
            w("        For i = LBound(names) To UBound(names)")
            w(f'            Say "{cls}|" & names(i) & "|NOT-PROBED-METHOD"')
            w("        Next i")
        for ch in chunks(members):
            w(f'        names = Split("{ch}", "|")')
            w("        For i = LBound(names) To UBound(names)")
            w(f'            Probe "{cls}", o, CStr(names(i))')
            w("            n = n + 1")
            w("        Next i")
        w("    End If")

    w("")
    w("    Members = n")
    w("End Function")
    w("")
    w("' WHAT 438 DOES AND DOESN'T TELL US")
    w("'")
    w("' \"Object doesn't support this property or method\" comes back for THREE")
    w("' different situations, and only one of them means the member is absent:")
    w("'")
    w("'   1. it really isn't a member of this version")
    w("'   2. it returns an object with no default property, so assigning it to a")
    w("'      Variant fails even though the member is fine - retried with Set below")
    w("'   3. it is write-only, or hidden in the type library and undocumented.")
    w("'      Nothing can distinguish this from case 1 by reading.")
    w("'")
    w("' Case 3 is not hypothetical. Project.NewTasksAreManual reports 438 here,")
    w("' is absent from Microsoft's published property list, and is nonetheless")
    w("' REAL - our own code assigns to it from an early-bound Project variable,")
    w("' which would not compile at all if the member didn't exist.")
    w("'")
    w("' So the label is NOT-READABLE, never MISSING. Removing something from the")
    w("' model needs this plus the documentation agreeing, not this alone.")
    w("Private Sub Probe(ByVal cls As String, ByVal o As Object, ByVal nm As String)")
    w("    Dim v As Variant, o2 As Object, e As Long")
    w("    On Error Resume Next")
    w("    v = CallByName(o, nm, VbGet)")
    w("    e = Err.Number")
    w("    Err.Clear")
    w("")
    w("    If e = 438 Then")
    w("        ' Could just be an object with no default property.")
    w("        Set o2 = CallByName(o, nm, VbGet)")
    w("        If Err.Number = 0 Then")
    w("            Err.Clear")
    w("            On Error GoTo 0")
    w("            Say cls & \"|\" & nm & \"|OK-OBJECT\"")
    w("            Exit Sub")
    w("        End If")
    w("        Err.Clear")
    w("    End If")
    w("    On Error GoTo 0")
    w("")
    w("    Select Case e")
    w("        Case 0")
    w("            Say cls & \"|\" & nm & \"|OK\"")
    w("        Case 449, 450")
    w("            Say cls & \"|\" & nm & \"|ARGS\"")
    w("        Case 438")
    w("            Say cls & \"|\" & nm & \"|NOT-READABLE\"")
    w("        Case Else")
    w("            Say cls & \"|\" & nm & \"|ERR\" & e")
    w("    End Select")
    w("End Sub")
    w("")
    w("'--- sample objects, by navigation, read-only -------------------------------")
    w("' Not Tasks(1): a blank row or a row deleted earlier leaves a Nothing in the")
    w("' collection, and Project raises on it rather than returning Nothing.")
    w("Private Function FirstTask(ByVal app As Object) As Object")
    w("    Dim t As Object")
    w("    On Error Resume Next")
    w("    For Each t In app.ActiveProject.Tasks")
    w("        If Not t Is Nothing Then")
    w("            Set FirstTask = t")
    w("            Exit Function")
    w("        End If")
    w("    Next t")
    w("    Err.Clear")
    w("End Function")
    w("")
    w("Private Function FirstRes(ByVal app As Object) As Object")
    w("    Dim r As Object")
    w("    On Error Resume Next")
    w("    For Each r In app.ActiveProject.Resources")
    w("        If Not r Is Nothing Then")
    w("            Set FirstRes = r")
    w("            Exit Function")
    w("        End If")
    w("    Next r")
    w("    Err.Clear")
    w("End Function")
    w("")
    w("Private Function FirstAsg(ByVal app As Object) As Object")
    w("    Dim t As Object, a As Object")
    w("    On Error Resume Next")
    w("    For Each t In app.ActiveProject.Tasks")
    w("        If Not t Is Nothing Then")
    w("            For Each a In t.Assignments")
    w("                Set FirstAsg = a")
    w("                Exit Function")
    w("            Next a")
    w("        End If")
    w("    Next t")
    w("    Err.Clear")
    w("End Function")
    w("")
    w("'--- output ----------------------------------------------------------------")
    w("Private Sub Say(ByVal s As String)")
    w("    mOut = mOut & s & vbCrLf")
    w("End Sub")
    w("")
    w("' Next to the project file when there is one, TEMP otherwise. Never the Office")
    w("' install folder - that's read-only on a managed machine.")
    w("Private Function OutPath(ByVal app As Object) As String")
    w("    Dim folder As String")
    w("    On Error Resume Next")
    w("    folder = app.ActiveProject.Path")
    w("    Err.Clear")
    w("    If Len(folder) = 0 Then folder = Environ$(\"TEMP\")")
    w("    If Len(folder) = 0 Then folder = Environ$(\"USERPROFILE\")")
    w("    On Error GoTo 0")
    w("    OutPath = folder & \"\\RS_ProjectModel.txt\"")
    w("End Function")
    w("")
    w("Private Function WriteOut(ByVal path As String) As Boolean")
    w("    Dim f As Integer")
    w("    On Error GoTo Nope")
    w("    f = FreeFile")
    w("    Open path For Output As #f")
    w("    Print #f, mOut")
    w("    Close #f")
    w("    WriteOut = True")
    w("    Exit Function")
    w("Nope:")
    w("    On Error Resume Next")
    w("    Close #f")
    w("    WriteOut = False")
    w("End Function")

    with open(OUT, "w", newline="") as fh:
        fh.write("\r\n".join(lines) + "\r\n")

    probed = sum(len(classes[c].get("members", {})) for c in classes if c in SAMPLES)
    print(f"wrote {OUT}")
    print(f"  {len(lines)} lines, probing {probed} members across "
          f"{sum(1 for c in classes if c in SAMPLES)} classes")
    missing = [c for c in classes if c not in SAMPLES]
    if missing:
        print(f"  no sample object, so not probed: {', '.join(sorted(missing))}")


if __name__ == "__main__":
    main()
