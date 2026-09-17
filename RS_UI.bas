Attribute VB_Name = "RS_UI"
Option Explicit
'==============================================================================
' ROLE STAFFING - RS_UI
'
' Everything you click or read: the menu, the setup check and its form, the
' control panel, the ribbon, and the reports.
'
' All four modules share one namespace, so nothing needs importing between them -
' but RoleStaffing holds the state they all use, so import that one first.
'==============================================================================

'==============================================================================
' MENU
'==============================================================================

Public Sub RS_Menu()
    Dim c As String, m As String

    ' Two levels, because an InputBox can't show thirty lines - the bottom of a
    ' long list simply isn't on screen, which is how the form options went
    ' missing. Everything is still one keystroke from here.
    Do
        m = "WHAT DO YOU WANT TO DO?" & vbCrLf & vbCrLf
        m = m & "0   Staff and schedule - the whole thing, one go" & vbCrLf
        m = m & "1   Check and prep this file" & vbCrLf
        m = m & "2   Validate - what's wrong, changes nothing" & vbCrLf
        m = m & "3   What-if - project the finish date, writes nothing" & vbCrLf & vbCrLf
        m = m & "S   Staffing, step by step ..." & vbCrLf
        m = m & "R   Reports ..." & vbCrLf
        m = m & "T   Task data - roles, SameWorkAs, spans ..." & vbCrLf
        m = m & "P   People - edit RoleAvail ..." & vbCrLf
        m = m & "F   Forms and ribbon ..." & vbCrLf
        m = m & "X   Setup and maintenance ..."

        c = UCase$(Trim$(InputBox(m, "Role Staffing")))
        If c = "" Then Exit Sub

        Select Case c
            Case "0": RS_AutoStaff: Exit Sub
            Case "1": RS_PrepFile: Exit Sub
            Case "2": RS_Validate: Exit Sub
            Case "3": RS_Schedule: Exit Sub
            Case "S": If SubMenu(MenuStaffing()) Then Exit Sub
            Case "R": If SubMenu(MenuReports()) Then Exit Sub
            Case "T": If SubMenu(MenuTasks()) Then Exit Sub
            Case "P": If SubMenu(MenuPeople()) Then Exit Sub
            Case "F": If SubMenu(MenuForms()) Then Exit Sub
            Case "X": If SubMenu(MenuSetup()) Then Exit Sub
            Case Else: Exit Sub
        End Select
    Loop
End Sub

' True if something ran. Empty input goes back to the main menu.
Private Function SubMenu(ByVal m As String) As Boolean
    Dim c As String
    c = UCase$(Trim$(InputBox(m & vbCrLf & vbCrLf & "Blank to go back.", "Role Staffing")))
    If c = "" Then Exit Function
    SubMenu = True

    Select Case c
        Case "1": RS_Suggest
        Case "2": RS_Apply
        Case "3": RS_Unstaff
        Case "4": RS_Level
        Case "5": RS_ClearSuggestions
        Case "6": RS_AddLevelOfEffort

        Case "10": RS_RoleLoad
        Case "11": RS_RuleLinks
        Case "12": RS_Schedule
        Case "13": RS_Validate

        Case "20": RS_MarkRoles
        Case "21": RS_BackfillRoles
        Case "22": RS_SyncCovers
        Case "23": RS_SpanTask

        Case "30": RS_EditRoleAvail
        Case "31": RS_RoleForm

        Case "40": RS_RebuildForms
        Case "41": RS_BuildPanel
        Case "42": RS_BuildCheckForm
        Case "43": RS_BuildRoleForm
        Case "44": RS_Panel
        Case "45": RS_CheckForm
        Case "46": RS_BuildRibbon
        Case "47": RS_ShowRibbonFile
        Case "48": RS_TidyForms

        Case "50": RS_Setup
        Case "51": RS_Check
        Case "52": RS_ExportAll
        Case "53": RS_EventTest
        Case "58": RS_AddContextXml
        Case "59": RS_RemoveContextXml
        Case "54"
            If gRS_GuardOn Then
                RS_StopGuard
                MsgBox "Manual-edit guard is OFF.", vbInformation, "Role Staffing"
            Else
                RS_StartGuard
                MsgBox "Manual-edit guard is ON.", vbInformation, "Role Staffing"
            End If
        Case Else: SubMenu = False
    End Select
End Function

Private Function MenuStaffing() As String
    MenuStaffing = "STAFFING, STEP BY STEP" & vbCrLf & vbCrLf & _
        "1   Suggest - who should do what" & vbCrLf & _
        "2   Apply - put them in Resource Names" & vbCrLf & _
        "3   Unstaff - everyone back to role rows" & vbCrLf & _
        "4   Level - move dates to fit availability" & vbCrLf & _
        "5   Clear suggestions" & vbCrLf & _
        "6   Add level-of-effort to selected tasks"
End Function

Private Function MenuReports() As String
    MenuReports = "REPORTS - none of these change anything" & vbCrLf & vbCrLf & _
        "10  Role load and drift" & vbCrLf & _
        "11  Rule links - what covers what" & vbCrLf & _
        "12  Forward schedule - project the finish date" & vbCrLf & _
        "13  Validate"
End Function

Private Function MenuTasks() As String
    MenuTasks = "TASK DATA" & vbCrLf & vbCrLf & _
        "20  Mark role rows (migrating an existing file)" & vbCrLf & _
        "21  Re-record roles from assignments" & vbCrLf & _
        "22  Re-point " & RS_COVERS_LABEL & " after moving rows" & vbCrLf & _
        "23  Span another task (support across a deliverable)"
End Function

Private Function MenuPeople() As String
    MenuPeople = "PEOPLE - select someone on the resource sheet first" & vbCrLf & vbCrLf & _
        "30  Edit their RoleAvail" & vbCrLf & _
        "31  ...as a form"
End Function

Private Function MenuForms() As String
    MenuForms = "FORMS AND RIBBON" & vbCrLf & vbCrLf & _
        "44  Open the control panel" & vbCrLf & _
        "45  Open the setup check form" & vbCrLf & vbCrLf & _
        "Rebuild after updating the code, or they stay on the old layout:" & vbCrLf & _
        "40  Rebuild all three forms" & vbCrLf & _
        "41  ...just the control panel" & vbCrLf & _
        "42  ...just the setup check form" & vbCrLf & _
        "43  ...just the RoleAvail form" & vbCrLf & _
        "48  Remove every form (tidies up numbered strays)" & vbCrLf & vbCrLf & _
        "46  Build a Role Staffing ribbon tab" & vbCrLf & _
        "47  Show the ribbon customisation file"
End Function

Private Function MenuSetup() As String
    MenuSetup = "SETUP AND MAINTENANCE" & vbCrLf & vbCrLf & _
        "50  Set up this project file" & vbCrLf & _
        "51  Setup check - what is and isn't set up" & vbCrLf & _
        "52  Export all files to a folder (for committing)" & vbCrLf & _
        "53  Are the events firing? (diagnostic)" & vbCrLf & _
        "58  Right-click menu: add it via customUI" & vbCrLf & _
        "59  ...take it back out" & vbCrLf & _
        "54  Turn manual-edit guard " & IIf(gRS_GuardOn, "OFF", "ON")
End Function


'==============================================================================
' ROLEAVAIL EDITOR
'
' Select a person on the resource sheet and edit their roles without typing the
' string by hand.
'
' The engine here is the same whichever way you drive it, so RS_EditRoleAvail
' works with no form at all, and the form is only a nicer front end for it.
'==============================================================================

' "Check:40*,Draft:20" -> a collection of Array(role, percent, isPrimary)
Public Function RoleAvailRows(ByVal avail As String) As Collection
    Dim c As New Collection, d As Object, k As Variant
    Set RoleAvailRows = c
    Set d = ParseRoleAvail(avail)
    For Each k In d.Keys
        c.Add Array(CStr(d(k)(2)), CDbl(d(k)(0)) * 100#, CBool(d(k)(1)))
    Next k
End Function

' ...and back again, in the order given.
Public Function RowsToRoleAvail(rows As Collection) As String
    Dim x As Variant, s As String
    For Each x In rows
        If s <> "" Then s = s & ","
        s = s & CStr(x(0)) & ":" & Format$(CDbl(x(1)), "0.##")
        If CBool(x(2)) Then s = s & "*"
    Next x
    RowsToRoleAvail = s
End Function

Public Function RowsTotal(rows As Collection) As Double
    Dim x As Variant
    For Each x In rows
        RowsTotal = RowsTotal + CDbl(x(1))
    Next x
End Function

' Scales every percentage so they add to 100, keeping their relative sizes.
' The shares are a division of one person's time, so 40/40/10/10 and
' 80/80/20/20 mean the same thing - this just states it in the usual terms.
Public Sub NormaliseRows(rows As Collection)
    Dim t As Double, i As Long, x As Variant, out As Collection
    t = RowsTotal(rows)
    If t <= 0 Then Exit Sub

    Set out = New Collection
    For Each x In rows
        out.Add Array(CStr(x(0)), CDbl(x(1)) * 100# / t, CBool(x(2)))
    Next x
    Do While rows.Count > 0
        rows.Remove 1
    Loop
    For Each x In out
        rows.Add x
    Next x
End Sub

' The selected resource, or Nothing with a word about why.
Public Function SelectedPerson(p As Project) As Resource
    Dim rs As Resources, r As Resource
    On Error Resume Next
    Set rs = ActiveSelection.Resources
    On Error GoTo 0
    If rs Is Nothing Then Exit Function
    If rs.Count = 0 Then Exit Function
    For Each r In rs
        If Not r Is Nothing Then
            Set SelectedPerson = r
            Exit Function
        End If
    Next r
End Function

' No form needed. Shows what they have, takes a new list, checks it adds up.
Public Sub RS_EditRoleAvail()
    Dim p As Project, r As Resource, rows As Collection, x As Variant
    Dim cur As String, shown As String, answer As String, t As Double
    Dim avail As String, k As Variant

    If Not HaveProject(p) Then Exit Sub
    BuildCatalog p

    Set r = SelectedPerson(p)
    If r Is Nothing Then
        MsgBox "Select a person on the resource sheet first." & vbCrLf & vbCrLf & _
               "A resource view - the RS Roles view, or Resource Sheet.", _
               vbInformation, "Role Staffing"
        Exit Sub
    End If
    If IsBucket(r) Then
        MsgBox r.Name & " is a role row, not a person - it has an IsRole, so RoleAvail " & _
               "doesn't apply to it.", vbInformation, "Role Staffing"
        Exit Sub
    End If

    cur = Trim$(r.Text1)
    Set rows = RoleAvailRows(cur)
    For Each x In rows
        shown = shown & "   " & PadR(CStr(x(0)), 16) & PadL(Format$(CDbl(x(1)), "0.##") & "%", 7) & _
                IIf(CBool(x(2)), "   * their job", "") & vbCrLf
    Next x
    If shown = "" Then shown = "   (nothing yet - they can't be suggested for anything)" & vbCrLf

    For Each k In RoleNames()
        If avail <> "" Then avail = avail & ", "
        avail = avail & CStr(k)
    Next k

    answer = InputBox( _
        r.Name & "  -  Max Units " & Pct(ToFraction(r.MaxUnits)) & vbCrLf & vbCrLf & _
        "Now:" & vbCrLf & shown & vbCrLf & _
        "Roles on this file: " & IIf(avail = "", "none yet", avail) & vbCrLf & vbCrLf & _
        "Type the list. A * marks a role as their job." & vbCrLf & _
        "   Check:40*, Draft:20" & vbCrLf & vbCrLf & _
        "Listed at all means allowed. Leave a role out and they'll never be " & _
        "suggested for it.", _
        "RoleAvail - " & r.Name, cur)
    If StrPtr(answer) = 0 Then Exit Sub

    Set rows = RoleAvailRows(answer)
    If rows.Count = 0 And Trim$(answer) <> "" Then
        MsgBox "Couldn't read that. Use Role:percent, separated by commas.", _
               vbExclamation, "Role Staffing"
        Exit Sub
    End If

    t = RowsTotal(rows)
    If rows.Count > 0 And Abs(t - 100#) > 0.5 Then
        Select Case MsgBox("Those add up to " & Format$(t, "0.##") & "%, not 100%." & vbCrLf & vbCrLf & _
                           "YES  - scale them to 100% keeping the same proportions" & vbCrLf & _
                           "NO   - leave them exactly as typed" & vbCrLf & _
                           "CANCEL - go back" & vbCrLf & vbCrLf & _
                           "They're shares of this person's time, so the proportions are what " & _
                           "matter - " & Format$(t, "0") & "% and 100% rank people the same way.", _
                           vbYesNoCancel + vbQuestion, "Role Staffing")
            Case vbYes: NormaliseRows rows
            Case vbCancel: Exit Sub
        End Select
    End If

    Application.OpenUndoTransaction "Role Staffing: RoleAvail"
    gRS_Suppress = gRS_Suppress + 1
    On Error Resume Next
    r.Text1 = RowsToRoleAvail(rows)
    On Error GoTo 0
    gRS_Suppress = gRS_Suppress - 1
    Application.CloseUndoTransaction

    BuildCatalog p
    MsgBox r.Name & " is now:" & vbCrLf & vbCrLf & "   " & _
           IIf(Trim$(r.Text1) = "", "(nothing)", Trim$(r.Text1)), _
           vbInformation, "Role Staffing"
End Sub


'==============================================================================
' THE ROLEAVAIL FORM
'
' Master-detail rather than a grid: a list of the person's roles, and one set of
' fields that edits whichever is selected. A row-per-role grid needs about
' thirty generated controls, and generating controls is the part of this that
' keeps failing - this needs nine.
'
' The form holds no logic. Every button calls into this module, so if the form
' won't build, RS_EditRoleAvail does the same job.
'==============================================================================

Public Sub RS_RoleForm()
    Dim nm As String
    On Error GoTo NoForm
    UserForms.Add(RS_FORM_ROLES).Show
    Exit Sub
NoForm:
    Err.Clear
    nm = FindFormCalling("RS_RoleFormFill")
    If nm <> "" Then
        On Error GoTo NoForm2
        UserForms.Add(nm).Show
        Exit Sub
    End If
NoForm2:
    Err.Clear
    If MsgBox("The RoleAvail form isn't built yet. Build it now?" & vbCrLf & vbCrLf & _
              "If it won't build, RS_EditRoleAvail does the same job without one.", _
              vbYesNo + vbQuestion, "Role Staffing") = vbYes Then RS_BuildRoleForm
End Sub

' Rebuilds all three forms. The RS_*Form wrappers only build one that ISN'T
' there, so after a layout change they'd go on showing the old one - this is how
' you take a new version.
' Removes every form of ours, numbered strays included. A failed rename used to
' leave frmRSCheck1, 2, 3... behind, and a stale marker meant the opener never
' recognised them - so each run made another one.
Public Sub RS_TidyForms()
    Dim n As Long, names As String

    n = KillForms(RS_FORM_PANEL, "RS_PanelStatus", names)
    n = n + KillForms(RS_FORM_CHECK, "RS_CheckRowsBuild", names)
    n = n + KillForms(RS_FORM_ROLES, "RS_RoleFormFill", names)
    ' Anything parked on the way out and not collected - shouldn't happen, but a
    ' leftover would otherwise sit there for good, invisible to the prefixes above.
    n = n + KillForms(RS_DEAD, "", names)

    If n = 0 Then
        MsgBox "No forms of ours to remove." & vbCrLf & vbCrLf & _
               "Build them with Menu > F > 40.", vbInformation, "Role Staffing"
        Exit Sub
    End If

    MsgBox n & " form(s) removed:" & vbCrLf & vbCrLf & names & vbCrLf & _
           "Rebuild them with Menu > F > 40." & vbCrLf & vbCrLf & _
           "Do that now, in the same sitting - names only come free properly " & _
           "once Project has been round the houses, and the builders count on it.", _
           vbInformation, "Role Staffing"
End Sub

Public Sub RS_RebuildForms()
    If MsgBox("Rebuild all three forms from the current code?" & vbCrLf & vbCrLf & _
              "   the control panel" & vbCrLf & _
              "   the setup check form" & vbCrLf & _
              "   the RoleAvail form" & vbCrLf & vbCrLf & _
              "Any that won't build say so and the rest carry on. Each has a " & _
              "no-form equivalent either way." & vbCrLf & vbCrLf & _
              "Close Project and save Global.MPT afterwards, or they're gone.", _
              vbOKCancel + vbQuestion, "Role Staffing") <> vbOK Then Exit Sub

    RS_TidyForms
    RS_BuildPanel
    RS_BuildCheckForm
    RS_BuildRoleForm
End Sub

Public Sub RS_BuildRoleForm()
    Dim vbp As Object, vbc As Object, c As Object, ctl As Object, code As String
    Dim stage As String, want As String

    On Error GoTo Failed
    stage = "finding the VBA project"
    Set vbp = OwnVBProject()
    If vbp Is Nothing Then
        MsgBox "Couldn't find the VBA project holding this code.", vbExclamation, "Role Staffing"
        Exit Sub
    End If

    stage = "removing any previous copy"
    Dim gone As String
    KillForms RS_FORM_ROLES, "RS_RoleFormFill", gone

    stage = "adding the form"
    Set vbc = vbp.VBComponents.Add(3)

    stage = "naming it"
    want = NameForm(vbc, RS_FORM_ROLES)

    stage = "sizing it"
    On Error Resume Next
    vbc.Properties("Caption") = "Role availability"
    vbc.Properties("Width") = 340
    vbc.Properties("Height") = 330
    Err.Clear
    On Error GoTo Failed

    stage = "adding the controls"
    AddCtl vbc, "Forms.Label.1", "lblWho", 12, 8, 46, 14
    AddCtl vbc, "Forms.ComboBox.1", "cboPerson", 60, 6, 180, 18
    AddCtl vbc, "Forms.Label.1", "lblUnits", 246, 8, 76, 14
    AddCtl vbc, "Forms.ListBox.1", "lstRoles", 12, 30, 310, 106
    AddCtl vbc, "Forms.Label.1", "lblRole", 12, 144, 40, 14
    AddCtl vbc, "Forms.ComboBox.1", "cboRole", 54, 142, 130, 18
    AddCtl vbc, "Forms.Label.1", "lblPct", 194, 144, 20, 14
    AddCtl vbc, "Forms.TextBox.1", "txtPct", 216, 142, 44, 18
    AddCtl vbc, "Forms.CheckBox.1", "chkPrimary", 266, 142, 60, 18
    AddCtl vbc, "Forms.CommandButton.1", "cmdSet", 12, 168, 100, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdDel", 120, 168, 100, 22
    AddCtl vbc, "Forms.Label.1", "lblTotal", 12, 198, 200, 14
    AddCtl vbc, "Forms.CommandButton.1", "cmdNorm", 216, 194, 106, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdApply", 12, 228, 100, 24
    AddCtl vbc, "Forms.CommandButton.1", "cmdClose", 222, 228, 100, 24

    On Error Resume Next
    vbc.Designer.Controls("lblWho").Caption = "Person"
    vbc.Designer.Controls("lblRole").Caption = "Role"
    vbc.Designer.Controls("lblPct").Caption = "%"
    vbc.Designer.Controls("chkPrimary").Caption = "their job"
    vbc.Designer.Controls("cmdSet").Caption = "Add / update"
    vbc.Designer.Controls("cmdDel").Caption = "Remove"
    vbc.Designer.Controls("cmdNorm").Caption = "Scale to 100%"
    vbc.Designer.Controls("cmdApply").Caption = "Apply"
    vbc.Designer.Controls("cmdClose").Caption = "Close"
    Err.Clear
    On Error GoTo Failed

    stage = "writing the code behind it"
    code = "Private Sub UserForm_Initialize()" & vbCrLf & _
           "    RS_RoleFormFill Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cboPerson_Change()" & vbCrLf & _
           "    RS_RoleFormLoad Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub lstRoles_Click()" & vbCrLf & _
           "    RS_RoleFormPick Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cmdSet_Click()" & vbCrLf & _
           "    RS_RoleFormSet Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cmdDel_Click()" & vbCrLf & _
           "    RS_RoleFormDel Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cmdNorm_Click()" & vbCrLf & _
           "    RS_RoleFormNorm Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cmdApply_Click()" & vbCrLf & _
           "    RS_RoleFormApply Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cmdClose_Click()" & vbCrLf & _
           "    Unload Me" & vbCrLf & _
           "End Sub" & vbCrLf
    vbc.CodeModule.AddFromString code

    NameWarn want, RS_FORM_ROLES
    MsgBox "Built as """ & vbc.Name & """." & vbCrLf & vbCrLf & _
           "Select a person on the resource sheet, then run RS_RoleForm." & vbCrLf & vbCrLf & _
           "Close Project and say Yes to saving Global.MPT, or it won't be there next time.", _
           vbInformation, "Role Staffing"
    Exit Sub

Failed:
    MsgBox "Couldn't build the RoleAvail form." & vbCrLf & vbCrLf & _
           "Failed while: " & stage & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "RS_EditRoleAvail does the same job without a form, and the README has " & _
           "instructions for drawing this one by hand.", vbExclamation, "Role Staffing"
End Sub

Private Sub AddCtl(vbc As Object, ByVal kind As String, ByVal nm As String, _
                   ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    Dim ctl As Object
    Set ctl = vbc.Designer.Controls.Add(kind)
    ctl.Name = nm
    ctl.Left = x
    ctl.Top = y
    ctl.Width = w
    ctl.Height = h
End Sub

' ---- what the form calls ----------------------------------------------------

Public Sub RS_RoleFormFill(f As Object)
    Dim p As Project, r As Resource, k As Variant, sel As Resource

    On Error Resume Next
    If Application.Projects.Count = 0 Then Exit Sub
    Set p = ActiveProject
    BuildCatalog p

    ' Everyone who isn't a role row. Picking from here means the form stands on
    ' its own - no need to have selected the right row on the sheet first.
    f.cboPerson.Clear
    For Each r In p.Resources
        If IsWorkPerson(r) Then f.cboPerson.AddItem r.Name
    Next r

    f.cboRole.Clear
    For Each k In RoleNames()
        f.cboRole.AddItem CStr(k)
    Next k

    Set sel = SelectedPerson(p)
    If Not sel Is Nothing Then
        If IsWorkPerson(sel) Then f.cboPerson.Text = sel.Name
    End If
    If Trim$(f.cboPerson.Text) = "" And f.cboPerson.ListCount > 0 Then
        f.cboPerson.ListIndex = 0
    End If

    RS_RoleFormLoad f
End Sub

' Load whoever the dropdown is on.
Public Sub RS_RoleFormLoad(f As Object)
    Dim p As Project, r As Resource

    On Error Resume Next
    If Application.Projects.Count = 0 Then Exit Sub
    Set p = ActiveProject

    Set r = ResourceByName(p, Trim$(f.cboPerson.Text))
    If r Is Nothing Then
        f.lblUnits.Caption = ""
        Set gRS_FormRows = New Collection
        gRS_FormResId = 0
        RS_RoleFormPaint f
        Exit Sub
    End If

    gRS_FormResId = r.ID
    f.lblUnits.Caption = "Max Units " & Pct(ToFraction(r.MaxUnits))
    Set gRS_FormRows = RoleAvailRows(Trim$(r.Text1))
    RS_RoleFormPaint f
End Sub

Public Sub RS_RoleFormPaint(f As Object)
    Dim x As Variant, t As Double
    On Error Resume Next
    f.lstRoles.Clear
    If gRS_FormRows Is Nothing Then Exit Sub
    For Each x In gRS_FormRows
        f.lstRoles.AddItem IIf(CBool(x(2)), "* ", "  ") & PadR(CStr(x(0)), 18) & _
                           Format$(CDbl(x(1)), "0.##") & "%"
    Next x
    t = RowsTotal(gRS_FormRows)
    f.lblTotal.Caption = "Total " & Format$(t, "0.##") & "%" & _
                         IIf(Abs(t - 100) > 0.5 And gRS_FormRows.Count > 0, "  - should be 100", "")
End Sub

Public Sub RS_RoleFormPick(f As Object)
    Dim i As Long, x As Variant
    On Error Resume Next
    i = f.lstRoles.ListIndex
    If i < 0 Or gRS_FormRows Is Nothing Then Exit Sub
    x = gRS_FormRows(i + 1)
    f.cboRole.Text = CStr(x(0))
    f.txtPct.Text = Format$(CDbl(x(1)), "0.##")
    f.chkPrimary.Value = CBool(x(2))
End Sub

Public Sub RS_RoleFormSet(f As Object)
    Dim nm As String, pct As Double, i As Long, found As Long
    On Error Resume Next
    nm = Trim$(f.cboRole.Text)
    If nm = "" Then
        MsgBox "Pick a role first.", vbInformation, "Role Staffing"
        Exit Sub
    End If
    If Not IsNumeric(Replace(f.txtPct.Text, "%", "")) Then
        MsgBox "The percentage needs to be a number.", vbInformation, "Role Staffing"
        Exit Sub
    End If
    pct = CDbl(Replace(f.txtPct.Text, "%", ""))
    If gRS_FormRows Is Nothing Then Set gRS_FormRows = New Collection

    For i = 1 To gRS_FormRows.Count
        If StrComp(CStr(gRS_FormRows(i)(0)), nm, vbTextCompare) = 0 Then found = i
    Next i

    If found > 0 Then
        gRS_FormRows.Remove found
        gRS_FormRows.Add Array(nm, pct, CBool(f.chkPrimary.Value)), , found
    Else
        gRS_FormRows.Add Array(nm, pct, CBool(f.chkPrimary.Value))
    End If
    RS_RoleFormPaint f
End Sub

Public Sub RS_RoleFormDel(f As Object)
    Dim i As Long
    On Error Resume Next
    i = f.lstRoles.ListIndex
    If i < 0 Or gRS_FormRows Is Nothing Then Exit Sub
    gRS_FormRows.Remove i + 1
    RS_RoleFormPaint f
End Sub

Public Sub RS_RoleFormNorm(f As Object)
    On Error Resume Next
    If gRS_FormRows Is Nothing Then Exit Sub
    NormaliseRows gRS_FormRows
    RS_RoleFormPaint f
End Sub

Public Function RS_RoleFormApply(f As Object) As Boolean
    Dim p As Project, r As Resource, t As Double
    On Error Resume Next
    If Application.Projects.Count = 0 Then Exit Function
    Set p = ActiveProject
    Set r = p.Resources(gRS_FormResId)
    If r Is Nothing Then Exit Function
    If gRS_FormRows Is Nothing Then Exit Function

    t = RowsTotal(gRS_FormRows)
    If gRS_FormRows.Count > 0 And Abs(t - 100#) > 0.5 Then
        Select Case MsgBox("These add up to " & Format$(t, "0.##") & "%, not 100%." & vbCrLf & vbCrLf & _
                           "YES  - scale to 100%, same proportions" & vbCrLf & _
                           "NO   - save them as they are" & vbCrLf & _
                           "CANCEL - go back", vbYesNoCancel + vbQuestion, "Role Staffing")
            Case vbYes: NormaliseRows gRS_FormRows
            Case vbCancel: Exit Function
        End Select
    End If

    Application.OpenUndoTransaction "Role Staffing: RoleAvail"
    gRS_Suppress = gRS_Suppress + 1
    r.Text1 = RowsToRoleAvail(gRS_FormRows)
    gRS_Suppress = gRS_Suppress - 1
    Application.CloseUndoTransaction

    BuildCatalog p
    RS_RoleFormApply = True

    ' Stays open - with a person dropdown you'd usually do several in a row.
    f.lblUnits.Caption = "Max Units " & Pct(ToFraction(r.MaxUnits)) & "   saved"
    RS_RoleFormPaint f
End Function


'==============================================================================
' EXPORTING
'==============================================================================

' Writes every Role Staffing component out to a folder, so you can commit it.
' This is how the form's .frx gets into the repo: build the form once here,
' export, commit, and every machine after that is a plain file import.
Public Sub RS_ExportAll()
    Dim vbp As Object, c As Object, folder As String, ext As String
    Dim n As Long, done As Collection, dflt As String

    Set done = New Collection
    dflt = Environ$("USERPROFILE") & "\Documents\RoleStaffingVBA"

    folder = InputBox("Export the Role Staffing files to which folder?" & vbCrLf & vbCrLf & _
                      "Point this at your local clone and the files are ready to commit.", _
                      "Role Staffing - Export", dflt)
    If Trim$(folder) = "" Then Exit Sub
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)

    On Error Resume Next
    If Dir$(folder, vbDirectory) = "" Then MkDir folder
    If Err.Number <> 0 Then
        MsgBox "Couldn't create " & folder & ": " & Err.Description, vbExclamation, "Role Staffing"
        Exit Sub
    End If
    Err.Clear
    On Error GoTo NoTrust

    Set vbp = OwnVBProject()
    If vbp Is Nothing Then
        MsgBox "Couldn't find the VBA project holding this code.", vbExclamation, "Role Staffing"
        Exit Sub
    End If

    For Each c In vbp.VBComponents
        ext = ""
        Select Case c.Type
            Case 1: ext = ".bas"
            Case 2: ext = ".cls"
            Case 3: ext = ".frm"      ' .frx is written alongside automatically
        End Select
        If ext <> "" And IsOurs(c.Name) Then
            c.Export folder & "\" & c.Name & ext
            done.Add c.Name & ext
            n = n + 1
        End If
    Next c

    If n = 0 Then
        MsgBox "Nothing to export - no Role Staffing components found.", vbExclamation, "Role Staffing"
    Else
        MsgBox n & " file(s) written to" & vbCrLf & folder & vbCrLf & vbCrLf & _
               JoinCollection(done, vbCrLf) & vbCrLf & vbCrLf & _
               "The form also writes a .frx next to its .frm - commit both, or the " & _
               "form won't import anywhere else.", vbInformation, "Role Staffing"
    End If
    Exit Sub

NoTrust:
    MsgBox "Couldn't export: " & Err.Description & vbCrLf & vbCrLf & _
           "Most likely ""Trust access to the VBA project object model"" is off." & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings > Macro Settings." & vbCrLf & vbCrLf & _
           "Without it, export each component by hand: right-click it in the Project " & _
           "Explorer > Export File.", vbExclamation, "Role Staffing"
End Sub


Private Function IsOurs(ByVal nm As String) As Boolean
    IsOurs = (StrComp(nm, "RoleStaffing", vbTextCompare) = 0) Or _
             (StrComp(Left$(nm, 5), "clsRS", vbTextCompare) = 0) Or _
             (StrComp(Left$(nm, 7), "frmRole", vbTextCompare) = 0)
End Function


' The VBA project this code is sitting in - not whatever is selected in the editor.
Private Function OwnVBProject() As Object
    Dim vbp As Object, c As Object
    On Error Resume Next
    For Each vbp In Application.VBE.VBProjects
        For Each c In vbp.VBComponents
            If StrComp(c.Name, "RoleStaffing", vbTextCompare) = 0 Then
                Set OwnVBProject = vbp
                Exit Function
            End If
        Next c
    Next vbp
    Set OwnVBProject = Application.VBE.ActiveVBProject
End Function










'==============================================================================
' SETUP
'==============================================================================

Public Sub RS_Setup()
    Dim p As Project
    If Not HaveProject(p) Then Exit Sub
    If MsgBox("Set this file up for Role Staffing - ALL OF IT, without asking again." & vbCrLf & vbCrLf & _
              "If you'd rather see what's missing and pick, close this and use " & _
              "Set up and check on the ribbon (or Menu > X > 51). It lists every item " & _
              "with a tick or a cross and fixes only what you choose." & vbCrLf & vbCrLf & _
              "ADDS WHAT'S MISSING, leaves the rest alone:" & vbCrLf & _
              "  - names a free Text field after each column it needs" & vbCrLf & _
              "  - adds a bucket resource for any role that hasn't got one" & vbCrLf & _
              "  - records the role on assignments that already use a role row" & vbCrLf & vbCrLf & _
              "REBUILDS EVERY TIME:" & vbCrLf & _
              "  - the RS Roles and RS Staffing views and their tables" & vbCrLf & _
              "  - the RS by Role group and the RS Role: filters" & vbCrLf & _
              "  so any column widths or extras you added to THOSE are lost." & vbCrLf & vbCrLf & _
              "Nothing else in the file is touched. Fill in RoleAvail first if you can - " & _
              "buckets come from the roles it finds there." & vbCrLf & vbCrLf & "Continue?", _
              vbOKCancel + vbQuestion, "Role Staffing") <> vbOK Then Exit Sub

    Dim msgs As New Collection, x As Variant, r As Resource, nm As String

    On Error Resume Next

    PA.CustomFieldRename FieldID:=pjCustomResourceText1, NewName:="RoleAvail"
    PA.CustomFieldRename FieldID:=pjCustomResourceText2, NewName:="IsRole"

    ' Name free fields after any label nothing carries. Never renames a field
    ' that already has a name of its own.
    ResolveFields
    Dim namedN As Long, ranOut As Boolean
    namedN = NameMissingColumns(ranOut)
    If namedN > 0 Then msgs.Add "Named " & namedN & " column(s): " & FieldsFound()
    If ranOut Then
        msgs.Add "Ran out of unused Text fields. Free one up, or rename an existing field"
        msgs.Add "   to the label yourself - any Text field works, the name is what matters."
    End If
    If namedN = 0 And Not ranOut Then msgs.Add "Columns already named: " & FieldsFound()
    Err.Clear

    ' Buckets, one per discovered role
    BuildCatalog p
    For Each x In RoleNames()
        If BucketFor(p, CStr(x)) Is Nothing Then
            nm = CStr(x)
            If Not ResourceByName(p, nm) Is Nothing Then nm = nm & " (role)"
            Set r = p.Resources.Add(Name:=nm)
            r.Text2 = CStr(x)
            r.MaxUnits = ToDbl(r.MaxUnits) * RS_BUCKET_UNITS
            msgs.Add "Created bucket resource """ & nm & """ for role " & x & "."
        End If
    Next x
    If RoleNames().Count = 0 Then
        msgs.Add "No roles found yet. Fill in RoleAvail on the resource sheet, then re-run."
    End If
    If Err.Number <> 0 Then msgs.Add "Bucket problem: " & Err.Description
    Err.Clear

    BuildViews p
    msgs.Add "Rebuilt the views, group and filters."
    Err.Clear
    On Error GoTo 0

    Dim msg As String
    For Each x In msgs
        msg = msg & CStr(x) & vbCrLf
    Next x
    ' If tasks already carry role rows, record the role on those assignments now.
    BackfillRoles p, False

    msg = msg & vbCrLf & "Next: open ""RS Roles"", set each person's Max Units and RoleAvail, " & _
          "then assign bucket resources to the tasks that need staffing." & vbCrLf & vbCrLf & _
          "Existing schedule with Engineer / Checker / Draft already in Resource Names? " & _
          "Run RS_MarkRoles (menu 13) to mark those rows as roles and copy the role onto " & _
          "every assignment that uses them."
    MsgBox msg, vbInformation, "Role Staffing"
End Sub


'==============================================================================
' SETUP CHECK
'
' What's set up, what isn't, and fix only the ones you choose. Nothing is
' created behind your back - the previous RS_Setup renamed fields and added
' resources on sight, which is no way to treat somebody's file.
'
' Findings are Array(ok, text, fixCode). RS_Check shows them numbered and asks
' which to fix; a form can show the same list with tick boxes and call
' ApplyFix for each one selected. Both drive the same two functions.
'==============================================================================

Public Sub BuildViews(p As Project)
    Dim x As Variant                 ' the role-filter loop below; Option Explicit
    On Error Resume Next
    ' Resource view
    PA.TableEditEx Name:="RS Roles", TaskTable:=False, Create:=True, OverwriteExisting:=True, _
                FieldName:="ID", Width:=6, ShowInMenu:=True
    PA.TableEditEx Name:="RS Roles", TaskTable:=False, NewFieldName:="Name", Width:=24
    PA.TableEditEx Name:="RS Roles", TaskTable:=False, NewFieldName:="Max Units", Width:=10
    PA.TableEditEx Name:="RS Roles", TaskTable:=False, NewFieldName:="Text1", Width:=40
    PA.TableEditEx Name:="RS Roles", TaskTable:=False, NewFieldName:="Text2", Width:=12
    ' (resource RoleAvail/IsRole are fixed at Text1/Text2 - only the task
    '  columns move around, because those are the ones users rename.)
    PA.ViewEditSingle Name:="RS Roles", Create:=True, Screen:=pjResourceSheet, _
                   Table:="RS Roles", Filter:="All Resources", ShowInMenu:=True
    If Err.Number <> 0 Then Debug.Print "Roles view problem: " & Err.Description Else Debug.Print "Created view ""RS Roles""."
    Err.Clear

    ' Task Usage view for reviewing suggestions
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, Create:=True, OverwriteExisting:=True, _
                FieldName:="ID", Width:=6, ShowInMenu:=True
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Name", Width:=34
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Text" & mF_Role, Width:=12
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Text" & mF_Suggested, Width:=20
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Work", Width:=9
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Start", Width:=11
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Finish", Width:=11
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Text" & mF_SameWork, Width:=12
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Text" & mF_Spans, Width:=10
    PA.TableEditEx Name:="RS Staffing", TaskTable:=True, NewFieldName:="Text" & mF_Calc, Width:=18
    PA.ViewEditSingle Name:="RS Staffing", Create:=True, Screen:=pjTaskUsage, _
                   Table:="RS Staffing", Filter:="All Tasks", ShowInMenu:=True
    If Err.Number <> 0 Then Debug.Print "Staffing view problem: " & Err.Description Else Debug.Print "Created view ""RS Staffing""."
    Err.Clear

    ' Group assignments by role, for Resource Usage
    PA.GroupEditEx Name:="RS by Role", TaskGroup:=False, Create:=True, OverwriteExisting:=True, _
                FieldName:="Text" & mF_Role, GroupAssignments:=True, ShowInMenu:=True
    If Err.Number <> 0 Then Debug.Print "Group problem (create ""RS by Role"" by hand if you want it): " & Err.Description
    Err.Clear

    ' One filter per role
    For Each x In RoleNames()
        PA.FilterEdit Name:="RS Role: " & x, TaskFilter:=False, Create:=True, OverwriteExisting:=True, _
                   FieldName:="Text1", Test:="contains", Value:=CStr(x), ShowInMenu:=True
    Next x
    If Err.Number <> 0 Then Debug.Print "Filter problem: " & Err.Description Else Debug.Print "Created ""RS Role: <role>"" filters."
    On Error GoTo 0
End Sub

Public Sub RS_Check()
    Dim p As Project
    If Not HaveProject(p) Then Exit Sub

    Dim f As Collection, x As Variant, i As Long, n As Long
    Dim body As String, pick As String, done As String, bad As Long

    Set f = BuildFindings(p)

    i = 0
    For Each x In f
        i = i + 1
        If CBool(x(0)) Then
            body = body & "  [ok]      " & CStr(x(1)) & vbCrLf
        Else
            bad = bad + 1
            body = body & "  [  ]  " & Format$(i, "00") & "  " & CStr(x(1)) & vbCrLf
        End If
    Next x

    If bad = 0 Then
        MsgBox "Everything's set up." & vbCrLf & vbCrLf & body, vbInformation, "Role Staffing"
        Exit Sub
    End If

    pick = InputBox( _
        body & vbCrLf & _
        "Type the numbers to fix, e.g. 1,3,4 - or ALL." & vbCrLf & _
        "Leave blank to change nothing.", _
        "Role Staffing - setup check")
    If Trim$(pick) = "" Then Exit Sub

    i = 0
    For Each x In f
        i = i + 1
        If Not CBool(x(0)) Then
            If StrComp(Trim$(pick), "ALL", vbTextCompare) = 0 Or _
               InStr(1, "," & Replace(Trim$(pick), " ", "") & ",", "," & i & ",") > 0 Then
                done = done & "  " & ApplyFix(p, CStr(x(2))) & vbCrLf
                n = n + 1
            End If
        End If
    Next x

    MsgBox n & " fixed:" & vbCrLf & vbCrLf & done & vbCrLf & _
           "Run this again to see where things stand.", vbInformation, "Role Staffing"
End Sub

' Array(ok, what it says, code for ApplyFix)
' Is a view of this name in the file? Checks both collections: a view made by
' ViewEditSingle is a single view, and depending on the Project version it may
' be listed under ViewsSingle, Views, or both. Late-bound and error-swallowing
' because a missing collection should read as "no view", not blow up the check.
Private Function HaveView(p As Project, ByVal nm As String) As Boolean
    Dim v As Object

    On Error Resume Next
    For Each v In p.Views
        If StrComp(v.Name, nm, vbTextCompare) = 0 Then
            HaveView = True
            Err.Clear
            Exit Function
        End If
    Next v
    Err.Clear

    For Each v In p.ViewsSingle
        If StrComp(v.Name, nm, vbTextCompare) = 0 Then
            HaveView = True
            Err.Clear
            Exit Function
        End If
    Next v
    Err.Clear
End Function

Public Function BuildFindings(p As Project) As Collection
    Dim f As New Collection, x As Variant, k As Variant, r As Resource
    Dim needed As Object, nm As String, people As Long
    Dim absent As String

    Set BuildFindings = f
    ResolveFields
    BuildCatalog p
    BuildRules

    ' --- columns ---------------------------------------------------------
    f.Add Array(ResourceFieldNamed("RoleAvail") > 0, _
                "RoleAvail column on the resource sheet", "RESCOL:RoleAvail")
    f.Add Array(ResourceFieldNamed("IsRole") > 0, _
                "IsRole column on the resource sheet", "RESCOL:IsRole")

    For Each x In Array(RS_L_ROLE, RS_L_SAMEWORK, RS_L_SUGGESTED, RS_L_SPANS, RS_L_CALC, _
                        RS_L_UNITSTASH, RS_L_SWSTASH, RS_L_SPANSTASH)
        f.Add Array(TaskFieldNamed(CStr(x)) > 0, CStr(x) & " column on tasks", "COL:" & CStr(x))
    Next x

    ' --- roles the rules mention must exist at all -----------------------
    Set needed = NewDict()
    If Not mRules Is Nothing Then
        For Each x In mRules
            needed(UCase$(CStr(x(0)))) = CStr(x(0))
            For Each k In x(1).Keys
                needed(UCase$(CStr(x(1)(k)))) = CStr(x(1)(k))
            Next k
        Next x
    End If
    For Each k In needed.Keys
        nm = CStr(needed(k))
        f.Add Array(NormalizeRole(nm) <> "", _
                    "Role """ & nm & """ exists - RS_RULES needs it", "ROLE:" & nm)
    Next k

    ' --- a bucket for every role -----------------------------------------
    For Each x In RoleNames()
        f.Add Array(Not BucketFor(p, CStr(x)) Is Nothing, _
                    "Bucket resource for " & x & IIf(needed.Exists(UCase$(CStr(x))), _
                    " (needed by RS_RULES)", ""), "BUCKET:" & CStr(x))
    Next x

    ' --- people ----------------------------------------------------------
    For Each r In p.Resources
        If IsWorkPerson(r) Then
            If Trim$(r.Text1) <> "" Then people = people + 1
        End If
    Next r
    f.Add Array(people > 0, people & " person(s) with RoleAvail filled in", "NONE:people")
    f.Add Array(RoleNames().Count > 0, RoleNames().Count & " role(s) found", "NONE:roles")

    ' --- views -----------------------------------------------------------
    ' This row used to be hardcoded False, so it showed a red cross for ever:
    ' clicking Fix rebuilt the views perfectly well and the row still came back
    ' a cross, because nothing ever asked whether they were there. An action
    ' wearing a checkbox. Now it's an actual check.
    If Not HaveView(p, "RS Roles") Then absent = "RS Roles"
    If Not HaveView(p, "RS Staffing") Then
        If Len(absent) > 0 Then absent = absent & " and "
        absent = absent & "RS Staffing"
    End If
    If Len(absent) = 0 Then
        f.Add Array(True, "Views RS Roles and RS Staffing are built", "VIEWS:")
    Else
        f.Add Array(False, "Build the views, group and filters - " & absent & _
                    " missing", "VIEWS:")
    End If
End Function

Public Function ApplyFix(p As Project, ByVal code As String) As String
    Dim what As String, arg As String, i As Long, free As Long
    Dim fid As Variant, r As Resource, nm As String

    i = InStr(code, ":")
    what = Left$(code, i - 1)
    arg = Mid$(code, i + 1)

    On Error Resume Next
    Select Case what
        Case "COL"
            free = FirstFreeTextField()
            If free = 0 Then
                ApplyFix = arg & ": no unused Text field left - free one up first"
            Else
                fid = PA.FieldNameToFieldConstant("Text" & free)
                PA.CustomFieldRename FieldID:=fid, NewName:=arg
                ResolveFields
                ApplyFix = arg & " -> Text" & free
            End If

        Case "RESCOL"
            free = FirstFreeResourceField()
            If free = 0 Then
                ApplyFix = arg & ": no unused resource Text field left"
            Else
                fid = PA.FieldNameToFieldConstant("Text" & free, 1)
                PA.CustomFieldRename FieldID:=fid, NewName:=arg
                ApplyFix = arg & " -> resource Text" & free
            End If

        Case "BUCKET", "ROLE"
            nm = arg
            If Not ResourceByName(p, nm) Is Nothing Then nm = nm & " (role)"
            Set r = p.Resources.Add(Name:=nm)
            r.Text2 = arg
            r.MaxUnits = ToDbl(r.MaxUnits) * RS_BUCKET_UNITS
            BuildCatalog p
            ApplyFix = "added resource """ & nm & """ for role " & arg

        Case "VIEWS"
            BuildViews p
            ApplyFix = "rebuilt the views, group and filters"

        Case Else
            ApplyFix = "nothing to do"
    End Select
    If Err.Number <> 0 Then ApplyFix = ApplyFix & " (" & Err.Description & ")"
    On Error GoTo 0
End Function

' Which resource Text field carries this name, or 0.
Public Function ResourceFieldNamed(ByVal label As String) As Long
    Dim n As Long, fid As Variant, nm As String
    On Error Resume Next
    For n = 1 To 30
        fid = PA.FieldNameToFieldConstant("Text" & n, 1)
        If Err.Number = 0 Then
            nm = ""
            nm = PA.CustomFieldGetName(fid)
            If StrComp(nm, label, vbTextCompare) = 0 Then
                ResourceFieldNamed = n
                Exit Function
            End If
        End If
        Err.Clear
    Next n
End Function

Public Function FirstFreeResourceField() As Long
    Dim n As Long, fid As Variant, nm As String
    On Error Resume Next
    For n = 1 To 30
        fid = PA.FieldNameToFieldConstant("Text" & n, 1)
        If Err.Number = 0 Then
            nm = ""
            nm = PA.CustomFieldGetName(fid)
            Err.Clear
            If nm = "" Or StrComp(nm, "Text" & n, vbTextCompare) = 0 Then
                FirstFreeResourceField = n
                Exit Function
            End If
        End If
        Err.Clear
    Next n
End Function

Public Function TaskFieldNamed(ByVal label As String) As Long
    Dim n As Long, fid As Variant, nm As String
    On Error Resume Next
    For n = 1 To 30
        fid = PA.FieldNameToFieldConstant("Text" & n)
        If Err.Number = 0 Then
            nm = ""
            nm = PA.CustomFieldGetName(fid)
            If StrComp(nm, label, vbTextCompare) = 0 Then
                TaskFieldNamed = n
                Exit Function
            End If
        End If
        Err.Clear
    Next n
End Function


'==============================================================================
' THE CHECK FORM
'
' Three controls: a multi-select list of findings, Fix selected, Close. Far less
' to go wrong than the twelve-button panel, which this build's VBE refused.
'
' If it still won't build, the instructions to draw it by hand are in the README
' and it's a two-minute job - the code behind it is four lines, all of which
' call into this module.
'==============================================================================

Public Sub RS_CheckForm()
    On Error GoTo NoForm
    UserForms.Add(RS_FORM_CHECK).Show
    Exit Sub
NoForm:
    Err.Clear
    Dim nm As String
    nm = FindFormCalling("RS_CheckRowsBuild")
    If nm <> "" Then
        On Error GoTo NoForm2
        UserForms.Add(nm).Show
        Exit Sub
    End If
NoForm2:
    Err.Clear
    If MsgBox("The check form isn't built yet. Build it now?" & vbCrLf & vbCrLf & _
              "Needs ""Trust access to the VBA project object model"" in Trust Center > " & _
              "Macro Settings." & vbCrLf & vbCrLf & _
              "If it won't build, RS_Check does the same job as a numbered list.", _
              vbYesNo + vbQuestion, "Role Staffing") = vbYes Then RS_BuildCheckForm
End Sub

Public Sub RS_BuildCheckForm()
    Dim vbp As Object, vbc As Object, c As Object, code As String
    Dim stage As String, want As String

    On Error GoTo Failed
    stage = "finding the VBA project"
    Set vbp = OwnVBProject()
    If vbp Is Nothing Then
        MsgBox "Couldn't find the VBA project holding this code.", vbExclamation, "Role Staffing"
        Exit Sub
    End If

    stage = "removing any previous copy"
    Dim gone As String
    KillForms RS_FORM_CHECK, "RS_CheckRowsBuild", gone

    stage = "adding the form"
    Set vbc = vbp.VBComponents.Add(3)

    stage = "naming it"
    want = NameForm(vbc, RS_FORM_CHECK)

    stage = "sizing it"
    On Error Resume Next
    vbc.Properties("Caption") = "Role Staffing - setup check"
    vbc.Properties("Width") = 480
    vbc.Properties("Height") = 400
    Err.Clear
    On Error GoTo Failed

    ' Three controls only. Every row is created when the form opens, so the
    ' designer never holds more than this however many findings there are - and
    ' whatever stopped the control panel building at thirty controls can't
    ' apply here.
    stage = "adding the frame and buttons"
    AddCtl vbc, "Forms.Label.1", "lblHead", 12, 8, 450, 14
    AddCtl vbc, "Forms.Frame.1", "fraRows", 12, 26, 452, 300
    AddCtl vbc, "Forms.CommandButton.1", "cmdRefresh", 12, 336, 120, 24
    AddCtl vbc, "Forms.CommandButton.1", "cmdClose", 344, 336, 120, 24

    On Error Resume Next
    vbc.Designer.Controls("cmdRefresh").Caption = "Refresh"
    vbc.Designer.Controls("cmdClose").Caption = "Close"
    vbc.Designer.Controls("fraRows").Caption = ""
    Err.Clear
    On Error GoTo Failed

    stage = "writing the code behind it"
    code = "Private Sub UserForm_Initialize()" & vbCrLf & _
           "    RS_CheckRowsBuild Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub UserForm_Terminate()" & vbCrLf & _
           "    RS_CheckRowsDone" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cmdRefresh_Click()" & vbCrLf & _
           "    RS_CheckRowsBuild Me" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf & _
           "Private Sub cmdClose_Click()" & vbCrLf & _
           "    Unload Me" & vbCrLf & _
           "End Sub" & vbCrLf
    vbc.CodeModule.AddFromString code

    NameWarn want, RS_FORM_CHECK
    MsgBox "Built as """ & vbc.Name & """." & vbCrLf & vbCrLf & _
           "Run RS_CheckForm to open it." & vbCrLf & vbCrLf & _
           "Close Project and say Yes to saving Global.MPT, or it won't be there " & _
           "next time.", vbInformation, "Role Staffing"
    Exit Sub

Failed:
    MsgBox "Couldn't build the check form." & vbCrLf & vbCrLf & _
           "Failed while: " & stage & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "RS_Check does the same job as a numbered list.", vbExclamation, "Role Staffing"
End Sub

' ---- the rows, built when the form opens -----------------------------------

' One row per finding: a tick or cross in green or red, and a button whose
' caption is the finding - click it to fix that one. Rows are made here rather
' than in the designer so there are exactly as many as there are findings.
Public Sub RS_CheckRowsBuild(f As Object)
    Dim p As Project, fnd As Collection, x As Variant
    Dim fra As Object, lbl As Object, btn As Object, h As Object
    Dim i As Long, y As Long, bad As Long

    On Error Resume Next
    If Application.Projects.Count = 0 Then Exit Sub
    Set p = ActiveProject

    Set fra = f.Controls("fraRows")
    RS_CheckRowsDone                      ' drop the previous lot
    Do While fra.Controls.Count > 0
        fra.Controls.Remove 0
    Loop

    Set gRS_RowBtns = New Collection
    Set fnd = BuildFindings(p)

    y = 4
    For Each x In fnd
        i = i + 1

        Set lbl = fra.Controls.Add("Forms.Label.1", "lblRow" & i, True)
        lbl.Left = 4
        lbl.Top = y + 3
        lbl.Width = 18
        lbl.Height = 14
        If CBool(x(0)) Then
            lbl.Caption = ChrW(10004)                 ' tick
            lbl.ForeColor = RGB(0, 128, 0)
        Else
            lbl.Caption = ChrW(10006)                 ' cross
            lbl.ForeColor = RGB(192, 0, 0)
            bad = bad + 1
        End If

        Set btn = fra.Controls.Add("Forms.CommandButton.1", "btnRow" & i, True)
        btn.Left = 24
        btn.Top = y
        btn.Width = 400
        btn.Height = 20
        btn.Caption = CStr(x(1))
        btn.Enabled = Not CBool(x(0))                 ' nothing to do on a tick

        If Not CBool(x(0)) Then
            Set h = New clsRS_RowBtn
            h.Attach btn, f, i
            gRS_RowBtns.Add h                         ' kept alive here, or the
        End If                                        ' buttons stop responding

        y = y + 24
    Next x

    fra.ScrollBars = 2                                ' vertical
    fra.ScrollHeight = y + 8

    f.Controls("lblHead").Caption = fnd.Count & " item(s), " & bad & " needing attention." & _
        IIf(bad > 0, "  Click one to fix it.", "  Nothing to do.")
End Sub

Public Sub RS_CheckRowsDone()
    Set gRS_RowBtns = Nothing
End Sub

' Called by a row's button. Fixes that finding and redraws.
Public Sub RS_FixRow(f As Object, ByVal rowIndex As Long)
    Dim p As Project, fnd As Collection, msg As String

    On Error Resume Next
    If Application.Projects.Count = 0 Then Exit Sub
    Set p = ActiveProject

    Set fnd = BuildFindings(p)
    If rowIndex < 1 Or rowIndex > fnd.Count Then Exit Sub
    If CBool(fnd(rowIndex)(0)) Then Exit Sub          ' already fine

    msg = ApplyFix(p, CStr(fnd(rowIndex)(2)))
    RS_CheckRowsBuild f
    f.Controls("lblHead").Caption = msg
End Sub


'==============================================================================
' CONTROL PANEL
'
' A UserForm keeps all its controls in a binary .frx file, so a form can't be
' shipped as readable text. Instead RS_BuildPanel writes one into your
' Global.MPT for you - run it once, then use RS_Panel from then on.
'
' It needs File > Options > Trust Center > Trust Center Settings >
' Macro Settings > "Trust access to the VBA project object model" ticked.
' If your IT department won't allow that, the README has instructions for
' drawing the same form by hand in five minutes, and RS_Menu works regardless.
'==============================================================================




Public Sub RS_Panel()
    Dim nm As String
    On Error GoTo NoForm

    UserForms.Add(RS_FORM_PANEL).Show
    Exit Sub

NoForm:
    ' Not under the name we wanted. Some builds refuse to rename a new
    ' component, so look for the form by what's in it instead.
    Err.Clear
    nm = FindPanel()
    If nm <> "" Then
        On Error GoTo NoForm2
        UserForms.Add(nm).Show
        Exit Sub
    End If
NoForm2:
    If MsgBox("The control panel hasn't been built yet." & vbCrLf & vbCrLf & _
              "Build it now? (One-off. Needs ""Trust access to the VBA project " & _
              "object model"" in Trust Center > Macro Settings.)", _
              vbYesNo + vbQuestion, "Role Staffing") = vbYes Then
        RS_BuildPanel
    End If
End Sub

' Removes every form whose name starts with base, numbered strays included, and
' appends what went to log. Used by all three builders and by RS_TidyForms.
'
' The rename is the whole point of this, not tidiness:
'
'   VBComponents.Remove is DEFERRED. The VBE marks the component and doesn't
'   actually process it until our code stops running - so the name a removed
'   form held is STILL TAKEN for the rest of this procedure, and a form added
'   straight afterwards cannot have it.
'
'   Renaming is immediate. Parking the doomed form under RS_DEAD frees the real
'   name on that line, so the Add that follows can take it.
'
' Leaving this out is what left the check form called UserForm1: removing
' frmRSCheck plus frmRSCheck1..4 put all five names in the pending state, the
' fallback ladder tried those same five, every one failed, and the form kept the
' name VBA gave it. One removal used to work only because the ladder's next
' number happened to be a name nothing had ever held.
Private Function KillForms(ByVal stem As String, ByVal marker As String, _
                           ByRef report As String) As Long
    Dim vbp As Object, c As Object, victim As Object
    Dim doomed As Collection, cm As Object
    Dim nm As String, n As Long, ours As Boolean

    On Error Resume Next
    Set vbp = OwnVBProject()
    If vbp Is Nothing Then Exit Function

    ' Collect first, act second, and never look at the collection again. Because
    ' the removal is deferred, a form we've already dealt with is STILL LISTED
    ' for the rest of this run - a loop that rescans finds it over and over and
    ' spins until its guard runs out. Scanning once is safe here: nothing adds a
    ' form until this returns.
    Set doomed = New Collection
    For Each c In vbp.VBComponents
        If c.Type = 3 Then
            ours = (StrComp(Left$(c.Name, Len(stem)), stem, vbTextCompare) = 0)
            If Not ours And marker <> "" Then
                Set cm = c.CodeModule
                If cm.CountOfLines > 0 Then
                    ours = (InStr(1, cm.Lines(1, cm.CountOfLines), marker, vbTextCompare) > 0)
                End If
            End If
            If ours Then doomed.Add c
        End If
    Next c

    For Each victim In doomed
        nm = victim.Name
        Err.Clear
        victim.Name = RS_DEAD & n & Format$(Timer, "000000")
        If Err.Number <> 0 Then
            ' Can't even rename it - it's loaded, or the project is locked.
            ' Removing it now would put its name in the pending state and cost us
            ' the name we came for, so leave it whole and say so.
            Err.Clear
            report = report & "   " & nm & " - in use, left alone" & vbCrLf
        Else
            vbp.VBComponents.Remove victim
            Err.Clear
            report = report & "   " & nm & vbCrLf
            n = n + 1
        End If
    Next victim

    Err.Clear
    KillForms = n
End Function

' Gives the form the name we want and returns the name it actually got.
' Retries the SAME name before falling to a numbered one - a numbered form is a
' form every opener has to go hunting for, so it's a last resort, not a first
' response to one failed attempt.
Private Function NameForm(ByVal vbc As Object, ByVal want As String) As String
    Dim i As Long, alt As String

    alt = want
    On Error Resume Next
    For i = 1 To 12
        Err.Clear
        vbc.Name = alt
        If Err.Number = 0 And StrComp(vbc.Name, alt, vbTextCompare) = 0 Then Exit For
        DoEvents
        If i >= 6 Then alt = want & (i - 5)
    Next i
    Err.Clear
    NameForm = vbc.Name
End Function

' Says so when a form couldn't have the name it was built for. It still works -
' the openers find a form by the code inside it - but silence here is what let
' UserForm1 look like a success.
Private Sub NameWarn(ByVal got As String, ByVal want As String)
    If StrComp(got, want, vbTextCompare) = 0 Then Exit Sub
    MsgBox "The form was built, but Project wouldn't let it be called " & want & _
           " - it's """ & got & """ instead." & vbCrLf & vbCrLf & _
           "It still works: everything finds it by the code inside it, not by " & _
           "name." & vbCrLf & vbCrLf & _
           "To get the proper name back: close Project saving Global.MPT, " & _
           "reopen, then Menu > F > 48 and F > 40 without doing anything else " & _
           "in between.", vbExclamation, "Role Staffing"
End Sub

' The panel, whatever it ended up being called: the one whose code calls
' RS_PanelStatus. Needs trust access, which building it needed anyway.
' A form by what its code calls, so the name doesn't matter.
Public Function FindFormCalling(ByVal marker As String) As String
    Dim vbp As Object, c As Object, cm As Object
    On Error Resume Next
    Set vbp = OwnVBProject()
    If vbp Is Nothing Then Exit Function
    For Each c In vbp.VBComponents
        If c.Type = 3 Then
            Set cm = c.CodeModule
            If cm.CountOfLines > 0 Then
                If InStr(1, cm.Lines(1, cm.CountOfLines), marker, vbTextCompare) > 0 Then
                    FindFormCalling = c.Name
                    Exit Function
                End If
            End If
        End If
    Next c
End Function

' Was its own copy of the search above, with the marker written out a second
' time. Two copies of a marker is how one of them goes stale unnoticed - and a
' stale marker is exactly what had the check form building a new copy of itself
' every run. One search, and the checker can see every marker that feeds it.
Private Function FindPanel() As String
    FindPanel = FindFormCalling("RS_PanelStatus")
End Function

Public Sub RS_BuildPanel()
    Dim vbp As Object, vbc As Object, c As Object, code As String
    Dim stage As String, want As String

    On Error GoTo Failed
    stage = "finding the VBA project"
    Set vbp = OwnVBProject()
    If vbp Is Nothing Then
        MsgBox "Couldn't find the VBA project holding this code.", vbExclamation, "Role Staffing"
        Exit Sub
    End If

    stage = "removing any previous copy"
    Dim gone As String
    KillForms RS_FORM_PANEL, "RS_PanelStatus", gone

    stage = "adding the form"
    Set vbc = vbp.VBComponents.Add(3)

    stage = "naming it"
    want = NameForm(vbc, RS_FORM_PANEL)

    stage = "sizing it"
    On Error Resume Next
    vbc.Properties("Caption") = "Role Staffing"
    vbc.Properties("Width") = 330
    vbc.Properties("Height") = 400
    Err.Clear
    On Error GoTo Failed

    ' One control per row, captions carrying their own explanation. The previous
    ' version used a button AND a label per row, plus Font.Bold on the designer
    ' controls - about thirty controls and a property this VBE may not allow.
    ' The two forms that build cleanly use exactly this shape, so it does too.
    stage = "adding the buttons"
    AddCtl vbc, "Forms.Label.1", "lblStatus", 12, 6, 300, 42

    AddCtl vbc, "Forms.CommandButton.1", "cmdAuto", 12, 52, 300, 28
    AddCtl vbc, "Forms.CommandButton.1", "cmdPrep", 12, 84, 300, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdValidate", 12, 110, 300, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdSched", 12, 136, 300, 22

    AddCtl vbc, "Forms.CommandButton.1", "cmdLoad", 12, 168, 300, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdLinks", 12, 194, 300, 22

    AddCtl vbc, "Forms.CommandButton.1", "cmdUnstaff", 12, 226, 300, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdRoles", 12, 252, 300, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdCheck", 12, 278, 300, 22
    AddCtl vbc, "Forms.CommandButton.1", "cmdMore", 12, 304, 300, 22

    AddCtl vbc, "Forms.CommandButton.1", "cmdClose", 12, 336, 300, 22

    stage = "captioning them"
    On Error Resume Next
    SetCap vbc, "cmdAuto", "STAFF AND SCHEDULE   -   pick people, assign, level"
    SetCap vbc, "cmdPrep", "Check && prep the file   -   fix task types and scheduling"
    SetCap vbc, "cmdValidate", "Validate   -   what's wrong. Changes nothing."
    SetCap vbc, "cmdSched", "What-if   -   project the finish date. Writes nothing."
    SetCap vbc, "cmdLoad", "Role load   -   hours per person, per role, per week"
    SetCap vbc, "cmdLinks", "Rule links   -   what covers what, and why"
    SetCap vbc, "cmdUnstaff", "Unstaff everyone   -   back to role rows, ready to re-run"
    SetCap vbc, "cmdRoles", "Edit RoleAvail   -   for whoever is selected"
    SetCap vbc, "cmdCheck", "Setup check   -   what is and isn't set up"
    SetCap vbc, "cmdMore", "More..."
    SetCap vbc, "cmdClose", "Close"
    Err.Clear
    On Error GoTo Failed

    stage = "writing the code behind it"
    code = "Private Sub UserForm_Initialize()" & vbCrLf & _
           "    lblStatus.Caption = RS_PanelStatus()" & vbCrLf & _
           "End Sub" & vbCrLf & vbCrLf
    code = code & Handler("cmdAuto", "RS_AutoStaff")
    code = code & Handler("cmdPrep", "RS_PrepFile")
    code = code & Handler("cmdValidate", "RS_Validate")
    code = code & Handler("cmdSched", "RS_Schedule")
    code = code & Handler("cmdLoad", "RS_RoleLoad")
    code = code & Handler("cmdLinks", "RS_RuleLinks")
    code = code & Handler("cmdUnstaff", "RS_Unstaff")
    code = code & Handler("cmdRoles", "RS_EditRoleAvail")
    code = code & Handler("cmdCheck", "RS_Check")
    code = code & Handler("cmdMore", "RS_Menu")
    code = code & "Private Sub cmdClose_Click()" & vbCrLf & _
                  "    Unload Me" & vbCrLf & _
                  "End Sub" & vbCrLf
    vbc.CodeModule.AddFromString code

    NameWarn want, RS_FORM_PANEL
    MsgBox "Built as """ & vbc.Name & """." & vbCrLf & vbCrLf & _
           "Run RS_Panel to open it." & vbCrLf & vbCrLf & _
           "Close Project and say Yes to saving Global.MPT, or it won't be there next time.", _
           vbInformation, "Role Staffing"
    Exit Sub

Failed:
    MsgBox "Couldn't build the control panel." & vbCrLf & vbCrLf & _
           "Failed while: " & stage & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "RS_Menu does everything the panel does.", vbExclamation, "Role Staffing"
End Sub

Private Sub SetCap(vbc As Object, ByVal nm As String, ByVal cap As String)
    On Error Resume Next
    vbc.Designer.Controls(nm).Caption = cap
End Sub

Private Function Handler(ByVal btn As String, ByVal macro As String) As String
    Handler = "Private Sub " & btn & "_Click()" & vbCrLf & _
              "    Unload Me" & vbCrLf & _
              "    " & macro & vbCrLf & _
              "End Sub" & vbCrLf & vbCrLf
End Function

' What the panel shows at the top: is anything actually set up yet?
Public Function RS_PanelStatus() As String
    Dim p As Project, t As Task, a As Assignment, r As Resource
    Dim openSlots As Long, people As Long, s As String

    On Error GoTo Bare
    If Application.Projects.Count = 0 Then
        RS_PanelStatus = "No project open."
        Exit Function
    End If
    Set p = ActiveProject
    BuildCatalog p
    BuildRules

    For Each r In p.Resources
        If IsWorkPerson(r) Then
            If Trim$(r.Text1) <> "" Then people = people + 1
        End If
    Next r
    For Each t In p.Tasks
        If Not t Is Nothing Then
            For Each a In t.Assignments
                If IsBucket(ResourceOf(p, a)) Then openSlots = openSlots + 1
            Next a
        End If
    Next t

    s = p.Name & vbCrLf
    s = s & "Roles: " & IIf(RoleNames().Count = 0, "none yet - fill in RoleAvail", _
            JoinCollection(RoleNames(), ", ")) & vbCrLf
    s = s & people & " people with RoleAvail" & "     " & openSlots & " slot(s) open" & vbCrLf
    s = s & "Rules: " & IIf(RulesText() = "", "none", RulesText()) & _
            "     Guard " & IIf(gRS_GuardOn, "ON", "OFF")
    RS_PanelStatus = s
    Exit Function

Bare:
    RS_PanelStatus = "Role Staffing"
End Function


'==============================================================================
' RIBBON
'
' Office keeps its ribbon customisation in a plain XML file:
'   %LOCALAPPDATA%\Microsoft\Office\MSProject.officeUI
'
' So we can just write it. No VBE automation, no trust setting, no UserForm -
' ordinary file I/O, which always works. Project reads it at startup.
'
' RS_ShowRibbonFile  prints the path and whatever is there now. Harmless, and
'                    the way to find out what your Project calls a macro.
' RS_BuildRibbon     backs the file up and writes a Role Staffing tab.
'==============================================================================

Private Function RibbonPath() As String
    RibbonPath = Environ$("LOCALAPPDATA") & "\Microsoft\Office\MSProject.officeUI"
End Function

Public Sub RS_ShowRibbonFile()
    Dim path As String, rep As New Collection, f As Integer, ln As String
    path = RibbonPath()

    rep.Add "Ribbon customisation file"
    rep.Add "-------------------------"
    rep.Add "   " & path
    rep.Add ""

    If Dir$(path) = "" Then
        rep.Add "Nothing there yet - you have no ribbon customisations."
        rep.Add "RS_BuildRibbon will create it."
    Else
        rep.Add "Current contents:"
        rep.Add ""
        On Error Resume Next
        f = FreeFile
        Open path For Input As #f
        Do Until EOF(f)
            Line Input #f, ln
            rep.Add ln
        Loop
        Close #f
        On Error GoTo 0
        rep.Add ""
        rep.Add "If a macro button is already in there, its onAction="""" attribute shows"
        rep.Add "exactly how this Project build names a macro - which is the one thing"
        rep.Add "that differs between hosts."
    End If
    ShowReport "Ribbon File", rep
End Sub

' Called from Auto_Open. Adds the tab if it isn't there, leaves everything else
' alone, and never blocks startup - a ribbon is not worth failing to load over.
'
' Merges rather than replaces: writing the whole file would wipe any ribbon
' customisation you'd made by hand, every single launch.
Public Sub EnsureRibbon()
    Dim path As String, cur As String, f As Integer, ln As String

    On Error GoTo Quiet
    If gRS_RibbonChecked Then Exit Sub
    gRS_RibbonChecked = True

    path = RibbonPath()
    If Dir$(Left$(path, InStrRev(path, "\") - 1), vbDirectory) = "" Then Exit Sub

    If Dir$(path) <> "" Then
        f = FreeFile
        Open path For Input As #f
        Do Until EOF(f)
            Line Input #f, ln
            cur = cur & ln & vbCrLf
        Loop
        Close #f
    End If

    ' already ours? nothing to do, and no message
    If InStr(1, cur, "id=""rsTab""", vbTextCompare) > 0 Then Exit Sub

    Dim merged As String
    If Trim$(cur) = "" Then
        WriteRibbon path, RibbonXml()
    Else
        merged = MergeTab(cur)
        If merged = "" Then Exit Sub           ' shape we don't understand - leave it
        On Error Resume Next
        Kill path & ".RoleStaffing.bak"
        Err.Clear
        FileCopy path, path & ".RoleStaffing.bak"
        On Error GoTo Quiet
        WriteRibbon path, merged
    End If

    MsgBox "Added a ""Role Staffing"" tab to your ribbon." & vbCrLf & vbCrLf & _
           "It appears next time you start Project." & vbCrLf & vbCrLf & _
           "Anything already on your ribbon was kept, and the old file is saved as" & vbCrLf & _
           RibbonPath() & ".RoleStaffing.bak", vbInformation, "Role Staffing"
    Exit Sub

Quiet:
    ' Never let this stop Project opening.
End Sub

Private Sub WriteRibbon(ByVal path As String, ByVal xml As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, xml
    Close #f
End Sub

Public Sub RS_BuildRibbon()
    Dim path As String, cur As String, f As Integer, ln As String, had As Boolean

    path = RibbonPath()

    On Error GoTo Fail
    If Dir$(path) <> "" Then
        f = FreeFile
        Open path For Input As #f
        Do Until EOF(f)
            Line Input #f, ln
            cur = cur & ln & vbCrLf
        Loop
        Close #f
    End If
    had = (InStr(1, cur, "id=""rsTab""", vbTextCompare) > 0)

    If MsgBox(IIf(had, "Rebuild", "Add") & " the Role Staffing ribbon tab?" & vbCrLf & vbCrLf & _
              path & vbCrLf & vbCrLf & _
              "Anything else on your ribbon is kept - only the Role Staffing tab is " & _
              "touched. The current file is copied to a .RoleStaffing.bak first." & vbCrLf & vbCrLf & _
              "Project has to be restarted to see it.", _
              vbOKCancel + vbQuestion, "Role Staffing") <> vbOK Then Exit Sub

    If Dir$(path) <> "" Then
        If Dir$(path & ".RoleStaffing.bak") <> "" Then Kill path & ".RoleStaffing.bak"
        FileCopy path, path & ".RoleStaffing.bak"
    End If

    Dim merged As String
    If Trim$(cur) = "" Then
        WriteRibbon path, RibbonXml()
    Else
        merged = MergeTab(cur)
        If merged <> "" Then
            WriteRibbon path, merged
        Else
            If MsgBox("That file isn't in a shape I recognise, so I can't merge into it." & vbCrLf & vbCrLf & _
                      "Replace it entirely? Your other ribbon customisations would be lost " & _
                      "(the .bak has them).", vbYesNo + vbExclamation, "Role Staffing") <> vbYes Then Exit Sub
            WriteRibbon path, RibbonXml()
        End If
    End If

    MsgBox IIf(had, "Rebuilt.", "Added.") & vbCrLf & vbCrLf & _
           "CLOSE AND REOPEN PROJECT to see the Role Staffing tab." & vbCrLf & vbCrLf & _
           "If the buttons appear but do nothing, this build names macros differently - " & _
           "run RS_ShowRibbonFile and send me what onAction says." & vbCrLf & vbCrLf & _
           "To undo: rename" & vbCrLf & path & ".RoleStaffing.bak" & vbCrLf & "back over it.", _
           vbInformation, "Role Staffing"
    Exit Sub

Fail:
    On Error Resume Next
    Close #f
    MsgBox "Couldn't write the ribbon file:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & _
           "If that folder doesn't exist, add any ribbon button by hand once - that " & _
           "creates it - then run this again.", vbExclamation, "Role Staffing"
End Sub

' Slots our tab into an existing customisation, whatever shape it's in, without
' disturbing anything else. Returns "" if the file isn't a shape we understand,
' in which case the caller leaves it alone rather than guessing.
'
'   <mso:tabs> ... </mso:tabs>   insert before the close
'   <mso:tabs/>                  self-closed - a QAT-only customisation. Has to
'                                be expanded, or the insert finds nothing and
'                                silently does nothing.
'   no <mso:tabs> at all         add the whole block inside <mso:ribbon>
Private Function MergeTab(ByVal xml As String) As String
    Dim body As String
    body = DropOurTab(xml)

    If InStr(1, body, "</mso:tabs>", vbTextCompare) > 0 Then
        MergeTab = Replace(body, "</mso:tabs>", RibbonTab() & "</mso:tabs>", 1, 1, vbTextCompare)
        Exit Function
    End If

    If InStr(1, body, "<mso:tabs/>", vbTextCompare) > 0 Then
        MergeTab = Replace(body, "<mso:tabs/>", "<mso:tabs>" & vbCrLf & RibbonTab() & "</mso:tabs>", _
                           1, 1, vbTextCompare)
        Exit Function
    End If

    If InStr(1, body, "<mso:tabs />", vbTextCompare) > 0 Then
        MergeTab = Replace(body, "<mso:tabs />", "<mso:tabs>" & vbCrLf & RibbonTab() & "</mso:tabs>", _
                           1, 1, vbTextCompare)
        Exit Function
    End If

    If InStr(1, body, "</mso:ribbon>", vbTextCompare) > 0 Then
        MergeTab = Replace(body, "</mso:ribbon>", _
                           "<mso:tabs>" & vbCrLf & RibbonTab() & "</mso:tabs>" & vbCrLf & "</mso:ribbon>", _
                           1, 1, vbTextCompare)
        Exit Function
    End If
End Function

' Drops any previous contextMenus block of ours, then adds the current one after
' the ribbon. Separate from MergeTab so the context menu can be taken back out
' on its own when an idMso turns out to be wrong.
Private Function MergeContextMenus(ByVal xml As String) As String
    Dim a As Long, b As Long, body As String

    body = xml
    a = InStr(1, body, "<mso:contextMenus>", vbTextCompare)
    If a > 0 Then
        b = InStr(a, body, "</mso:contextMenus>", vbTextCompare)
        If b > 0 Then body = Left$(body, a - 1) & Mid$(body, b + Len("</mso:contextMenus>"))
    End If

    If InStr(1, body, "</mso:ribbon>", vbTextCompare) = 0 Then Exit Function
    MergeContextMenus = Replace(body, "</mso:ribbon>", "</mso:ribbon>" & vbCrLf & ContextMenusXml(), _
                                1, 1, vbTextCompare)
End Function

Public Sub RS_AddContextXml()
    Dim path As String, cur As String, f As Integer, ln As String, merged As String

    path = RibbonPath()
    If MsgBox("Add ""Edit RoleAvail..."" to Project's right-click menus?" & vbCrLf & vbCrLf & _
              "Project keeps context menus in the same customUI file as the ribbon, so " & _
              "this writes them there." & vbCrLf & vbCrLf & _
              "THE RISK: nobody publishes what Project calls its context menus, so the " & _
              "names in RS_CTX_MENUS are educated guesses. A wrong one can make Office " & _
              "ignore the whole file - taking the Role Staffing ribbon tab with it." & vbCrLf & vbCrLf & _
              "A .bak is kept, and menu 58 takes just the context menus back out." & vbCrLf & vbCrLf & _
              "Restart Project afterwards. Continue?", _
              vbOKCancel + vbExclamation, "Role Staffing") <> vbOK Then Exit Sub

    On Error GoTo Fail
    If Dir$(path) <> "" Then
        f = FreeFile
        Open path For Input As #f
        Do Until EOF(f)
            Line Input #f, ln
            cur = cur & ln & vbCrLf
        Loop
        Close #f
        If Dir$(path & ".RoleStaffing.bak") <> "" Then Kill path & ".RoleStaffing.bak"
        FileCopy path, path & ".RoleStaffing.bak"
    End If

    If Trim$(cur) = "" Then
        WriteRibbon path, RibbonXml()
    Else
        merged = MergeContextMenus(cur)
        If merged = "" Then
            MsgBox "That file isn't a shape I can add to. Run RS_BuildRibbon first.", _
                   vbExclamation, "Role Staffing"
            Exit Sub
        End If
        WriteRibbon path, merged
    End If

    MsgBox "Written, for these menus:" & vbCrLf & vbCrLf & "   " & _
           Replace(RS_CTX_MENUS, ",", vbCrLf & "   ") & vbCrLf & vbCrLf & _
           "RESTART PROJECT, then right-click a resource." & vbCrLf & vbCrLf & _
           "If the Role Staffing ribbon tab has also vanished, one of those names is " & _
           "wrong - run menu 58 to take the context menus back out.", _
           vbInformation, "Role Staffing"
    Exit Sub

Fail:
    On Error Resume Next
    Close #f
    MsgBox "Couldn't write it: " & Err.Description, vbExclamation, "Role Staffing"
End Sub

Public Sub RS_RemoveContextXml()
    Dim path As String, cur As String, f As Integer, ln As String, a As Long, b As Long

    path = RibbonPath()
    If Dir$(path) = "" Then
        MsgBox "No ribbon file to change.", vbInformation, "Role Staffing"
        Exit Sub
    End If

    On Error GoTo Fail
    f = FreeFile
    Open path For Input As #f
    Do Until EOF(f)
        Line Input #f, ln
        cur = cur & ln & vbCrLf
    Loop
    Close #f

    a = InStr(1, cur, "<mso:contextMenus>", vbTextCompare)
    If a = 0 Then
        MsgBox "There are no context menus in it.", vbInformation, "Role Staffing"
        Exit Sub
    End If
    b = InStr(a, cur, "</mso:contextMenus>", vbTextCompare)
    If b = 0 Then Exit Sub

    WriteRibbon path, Left$(cur, a - 1) & Mid$(cur, b + Len("</mso:contextMenus>"))
    MsgBox "Taken out. Restart Project - the ribbon tab should be back.", _
           vbInformation, "Role Staffing"
    Exit Sub

Fail:
    On Error Resume Next
    Close #f
    MsgBox "Couldn't change it: " & Err.Description, vbExclamation, "Role Staffing"
End Sub

' Cut a previous Role Staffing tab out, so rebuilding replaces it instead of
' adding a second one. Everything else in the file is left exactly as it was.
Private Function DropOurTab(ByVal xml As String) As String
    Dim a As Long, b As Long
    DropOurTab = xml
    a = InStr(1, xml, "<mso:tab id=""rsTab""", vbTextCompare)
    If a = 0 Then Exit Function
    b = InStr(a, xml, "</mso:tab>", vbTextCompare)
    If b = 0 Then Exit Function
    DropOurTab = Left$(xml, a - 1) & Mid$(xml, b + Len("</mso:tab>"))
End Function

Private Function RibbonXml() As String
    Dim x As String
    x = "<mso:customUI xmlns:mso=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    x = x & "<mso:ribbon>" & vbCrLf
    x = x & "<mso:qat/>" & vbCrLf
    x = x & "<mso:tabs>" & vbCrLf
    x = x & RibbonTab()
    x = x & "</mso:tabs>" & vbCrLf
    x = x & "</mso:ribbon>" & vbCrLf
    x = x & ContextMenusXml()
    x = x & "</mso:customUI>"
    RibbonXml = x
End Function

' Context menus go in the SAME customUI file as the ribbon, after the ribbon
' element - that's the 2009/07 schema, which is what .officeUI already uses.
' CommandBars was simply the wrong door: Project 2013+ moved shortcut menus into
' customUI, which is why the collection has no Task or Resource menu in it.
'
' What isn't published anywhere is what Project CALLS its context menus. These
' are the plausible ones; add or change them in RS_CTX_MENUS if you find the
' real name. A wrong idMso is the risk - Office may ignore the whole file, which
' would take the ribbon tab with it. RS_BuildRibbon keeps a .bak for exactly
' that, and RS_RemoveContextXml puts things back without one.
Private Function ContextMenusXml() As String
    Dim x As String, nm As Variant

    If Trim$(RS_CTX_MENUS) = "" Then Exit Function

    x = "<mso:contextMenus>" & vbCrLf
    For Each nm In Split(RS_CTX_MENUS, ",")
        If Trim$(CStr(nm)) <> "" Then
            x = x & "<mso:contextMenu idMso=""" & Trim$(CStr(nm)) & """>" & vbCrLf
            x = x & "<mso:button id=""rsCtxRoleAvail"" label=""Edit RoleAvail..."" " & _
                    "onAction=""RS_EditRoleAvail"" insertBeforeMso=""FileClose"" visible=""true""/>" & vbCrLf
            x = x & "</mso:contextMenu>" & vbCrLf
        End If
    Next nm
    x = x & "</mso:contextMenus>" & vbCrLf
    ContextMenusXml = x
End Function

Private Function RibbonTab() As String
    Dim x As String
    x = "<mso:tab id=""rsTab"" label=""Role Staffing"" insertBeforeQ=""mso:TabFormat"">" & vbCrLf

    x = x & "<mso:group id=""rsRun"" label=""Run"" autoScale=""true"">" & vbCrLf
    x = x & Btn("rsAuto", "RS_AutoStaff", "Staff and" & vbLf & "Schedule", "Refresh", True)
    x = x & Btn("rsPrep", "RS_PrepFile", "Check and prep", "Spelling", False)
    x = x & Btn("rsVal", "RS_Validate", "Validate", "Info", False)
    x = x & Btn("rsSched", "RS_Schedule", "What-if", "FindDialog", False)
    x = x & "</mso:group>" & vbCrLf

    x = x & "<mso:group id=""rsPeople"" label=""People and setup"" autoScale=""true"">" & vbCrLf
    x = x & Btn("rsRoles", "RS_RoleForm", "Role" & vbLf & "availability", "PropertySheet", True)
    x = x & Btn("rsCheck", "RS_CheckForm", "Set up" & vbLf & "and check", "Spelling", True)
    x = x & "</mso:group>" & vbCrLf

    x = x & "<mso:group id=""rsRep"" label=""Reports"" autoScale=""true"">" & vbCrLf
    x = x & Btn("rsLoad", "RS_RoleLoad", "Role load", "FilePrintQuick", False)
    x = x & Btn("rsLinks", "RS_RuleLinks", "Rule links", "Hyperlink", False)
    x = x & "</mso:group>" & vbCrLf

    x = x & "<mso:group id=""rsSet"" label=""Reset"" autoScale=""true"">" & vbCrLf
    x = x & Btn("rsUnstaff", "RS_Unstaff", "Unstaff", "Undo", False)
    x = x & Btn("rsMark", "RS_MarkRoles", "Mark roles", "TagMarkComplete", False)
    x = x & Btn("rsMenu", "RS_Menu", "More...", "Help", False)
    x = x & "</mso:group>" & vbCrLf

    x = x & "</mso:tab>" & vbCrLf
    RibbonTab = x
End Function

Private Function Btn(ByVal id As String, ByVal macro As String, ByVal label As String, _
                     ByVal img As String, ByVal big As Boolean) As String
    Btn = "<mso:button id=""" & id & """ label=""" & label & """ imageMso=""" & img & _
          """ onAction=""" & macro & """ visible=""true""" & _
          IIf(big, " size=""large""", "") & "/>" & vbCrLf
End Function


'==============================================================================
' REPORTS
'==============================================================================

Public Sub RS_RuleLinks()
    Dim p As Project
    If Not HaveProject(p) Then Exit Sub
    BuildCatalog p
    BuildRules
    BuildRelations p

    Dim rep As New Collection, k As Variant, u As Variant, pr As Variant, n As Long

    rep.Add "Rules in force: " & IIf(RulesText() = "", "(none)", RulesText())
    rep.Add "Roles on the resource sheet: " & IIf(RoleNames().Count = 0, "NONE", _
                                                  JoinCollection(RoleNames(), ", "))
    rep.Add ""
    rep.Add "A rule only does anything if it names roles from that list. ""Draft&!Check"" does"
    rep.Add "nothing at all if your role rows say ""Avionics Drafter"" and ""Avionics Checker""."
    rep.Add "Roles not named in any rule are transparent - the walk passes through them."
    rep.Add ""
    rep.Add "Tasks below show ID (uid UniqueID)."
    If RS_COVERS_BY_ROW_ID Then
        rep.Add RS_COVERS_LABEL & " reads a plain number as the ID - the row number you can see."
        rep.Add "Row IDs SHIFT when tasks are inserted, deleted or moved, so check this report"
        rep.Add "again after reordering. Write u<number> to pin an entry to a Unique ID instead."
    Else
        rep.Add RS_COVERS_LABEL & " reads a plain number as the Unique ID (in brackets below)."
        rep.Add "Write #<number> to mean a row ID instead."
    End If
    rep.Add ""

    For Each k In TaskKeys()
        If TouchesRules(CStr(k)) Then
            n = n + 1
            rep.Add "Task " & TaskIdOf(CStr(k)) & " (uid " & k & ")  " & TaskNameOf(CStr(k)) & _
                    "   [" & RolesOn(CStr(k)) & "]   via " & CoverKind(CStr(k))
            If UpstreamOf(CStr(k)).Count = 0 Then
                rep.Add "   covers: (nothing found)"
            Else
                For Each u In UpstreamOf(CStr(k)).Keys
                    rep.Add "   covers: task " & TaskIdOf(CStr(u)) & " (uid " & u & ")  " & _
                            TaskNameOf(CStr(u)) & "   [" & RolesOn(CStr(u)) & "]   " & _
                            PeopleSummary(CStr(u))
                Next u
            End If
            For Each pr In TaskProblems(CStr(k))
                rep.Add "   ! " & pr
            Next pr
            rep.Add ""
        End If
    Next k

    If n = 0 Then
        rep.Add "No tasks carry a role that any rule mentions."
    End If
    ShowReport "Rule Links", rep
End Sub

Private Function RolesOn(ByVal key As String) As String
    Dim k As Variant, s As String
    If mTasks Is Nothing Then Exit Function
    If Not mTasks.Exists(key) Then Exit Function
    For Each k In mTasks(key)("roles").Keys
        If s <> "" Then s = s & ", "
        s = s & CStr(mTasks(key)("roles")(k))
    Next k
    RolesOn = s
End Function

Private Function PeopleSummary(ByVal key As String) As String
    Dim e As Variant, s As String, r As Resource
    If mTasks Is Nothing Then Exit Function
    If Not mTasks.Exists(key) Then Exit Function
    For Each e In mTasks(key)("people")
        Set r = Nothing
        On Error Resume Next
        Set r = mProject.Resources(CLng(e(0)))
        On Error GoTo 0
        If Not r Is Nothing Then
            If s <> "" Then s = s & ", "
            s = s & r.Name & " (" & CStr(e(1)) & ")"
        End If
    Next e
    PeopleSummary = s
End Function

Public Sub RS_Validate()
    Dim p As Project
    If Not HaveProject(p) Then Exit Sub
    FastOn "validating"
    ApplyPendingTags p
    BuildCatalog p
    BuildRules
    BuildLoad p
    BuildRelations p

    Dim errs As New Collection, warns As New Collection, infos As New Collection
    Dim r As Resource, t As Task, a As Assignment, res As Resource
    Dim x As Variant, k As Variant, v As Variant, role As String, raw As String
    Dim counts As Object, rep As New Collection, cap As Double, wkLoad As Double
    Dim over As Long, peak As Double
    Dim miss As Collection, missCount As Long, missList As String, taken As String

    ' 0. Are the columns named? Report only - naming fields is RS_Setup's job.
    Set miss = MissingLabels()
    missCount = miss.Count
    For Each x In miss
        If missList <> "" Then missList = missList & ", "
        missList = missList & CStr(x)
    Next x
    If missCount > 0 Then
        warns.Add missCount & " column(s) don't exist yet: " & missList
        warns.Add "   Whatever needs them is switched off until they do."
        warns.Add "   Add them either way:"
        warns.Add "      - run RS_Setup, which names free fields after them"
        warns.Add "      - or Project > Custom Fields > Task, pick ANY unused Text field,"
        warns.Add "        Rename it to the label. Text3 or Text27, it makes no difference -"
        warns.Add "        the code looks for the name, not the number."
    End If

    ' 1. Roles discovered
    Set counts = NewDict()
    For Each r In p.Resources
        If IsWorkPerson(r) Then
            For Each x In RoleNames()
                If HasRole(r, CStr(x)) Then counts(CStr(x)) = ToDbl(counts(CStr(x))) + 1
            Next x
        End If
    Next r
    For Each x In RoleNames()
        If ToDbl(counts(CStr(x))) = 0 Then
            warns.Add "Role """ & x & """: nobody has it in RoleAvail, so it can never be staffed."
        ElseIf ToDbl(counts(CStr(x))) = 1 Then
            infos.Add "Role """ & x & """: only one person can do it. If that's a typo it'll look like this."
        End If
        If BucketFor(p, CStr(x)) Is Nothing Then
            infos.Add "Role """ & x & """: no bucket resource, so you can't mark tasks as needing it. Re-run RS_Setup."
        End If
    Next x

    ' 2. Rules reference real roles
    If mRules Is Nothing Then BuildRules
    For Each x In mRules
        If NormalizeRole(CStr(x(0))) = "" Then
            errs.Add "RS_RULES names """ & x(0) & """, but no resource row has that as its " & _
                     "IsRole, so this rule can never fire. Roles found: " & _
                     IIf(RoleNames().Count = 0, "none", JoinCollection(RoleNames(), ", ")) & _
                     ". Either rename the rule or the role."
        End If
        For Each k In x(1).Keys
            If NormalizeRole(CStr(x(1)(k))) = "" Then
                errs.Add "RS_RULES names """ & x(1)(k) & """, but no resource row has that as its " & _
                         "IsRole, so this rule can never fire. Roles found: " & _
                         IIf(RoleNames().Count = 0, "none", JoinCollection(RoleNames(), ", ")) & _
                         ". Either rename the rule or the role."
            End If
        Next k
    Next x

    ' 3. Resource sheet sanity
    For Each r In p.Resources
        If Not r Is Nothing Then
            If IsBucket(r) Then
                If NormalizeRole(Trim$(r.Text2)) = "" Then
                    warns.Add r.Name & ": IsRole is """ & Trim$(r.Text2) & """, which isn't a known role."
                End If
                If Trim$(r.Text1) <> "" Then
                    warns.Add r.Name & ": it's a bucket (IsRole set) but also has RoleAvail. Clear one."
                End If
            ElseIf r.Type = pjResourceTypeWork Then
                If Trim$(r.Text1) = "" Then
                    infos.Add r.Name & ": no RoleAvail, so they'll never be suggested for anything."
                End If
                If ToFraction(r.MaxUnits) <= 0 Then
                    warns.Add r.Name & ": Max Units is 0, so they have no time for this project."
                End If
            End If
        End If
    Next r

    ' 4. SameWorkAs pointing at itself
    For Each t In p.Tasks
        If Not t Is Nothing Then
            If Trim$(TaskText(t, mF_SameWork)) <> "" Then
                Dim refs As String
                refs = "," & Replace(Replace(Replace(Trim$(TaskText(t, mF_SameWork)), " ", ""), ";", ","), "#", "") & ","
                If InStr(1, refs, "," & t.ID & ",", vbTextCompare) > 0 Or _
                   InStr(1, refs, ",u" & t.UniqueID & ",", vbTextCompare) > 0 Or _
                   (Not RS_COVERS_BY_ROW_ID And InStr(1, refs, "," & t.UniqueID & ",", vbTextCompare) > 0) Then
                    errs.Add TaskLabel(t) & ": " & RS_COVERS_LABEL & " names this task itself. " & _
                             "It has to name the OTHER task - the one this is the same work as."
                End If
            End If
        End If
    Next t

    ' 4b. SpanCalc formulas
    For Each t In p.Tasks
        If Not t Is Nothing Then
            If Trim$(TaskText(t, mF_Calc)) <> "" Then
                If EvalCalc(t, Trim$(TaskText(t, mF_Calc))) < 0 Then
                    errs.Add TaskLabel(t) & ": SpanCalc """ & Trim$(TaskText(t, mF_Calc)) & """ - " & CalcError()
                End If
            End If
        End If
    Next t

    ' 5. Tasks and assignments
    For Each t In p.Tasks
        If Not t Is Nothing Then
            raw = GetTaskRoleRaw(t)
            If raw <> "" And NormalizeRole(raw) = "" Then
                warns.Add TaskLabel(t) & ": Role """ & raw & """ isn't a known role."
            End If
            For Each a In t.Assignments
                Set res = ResourceOf(p, a)
                If Not res Is Nothing Then
                    If IsBucket(res) Then
                        infos.Add TaskLabel(t) & ": unstaffed " & res.Name & _
                                  IIf(CBool(t.Summary), " (level-of-effort)", "")
                    ElseIf IsWorkPerson(res) Then
                        role = InferRole(a, res, t)
                        If role = "" Then
                            If RoleCount(res) > 1 Then
                                infos.Add TaskLabel(t) & ": " & res.Name & "'s role is unknown. " & _
                                          "Set Role on the task, or on the assignment row."
                            End If
                        ElseIf NormalizeRole(role) <> "" Then
                            If Not HasRole(res, role) Then
                                warns.Add TaskLabel(t) & ": " & res.Name & " is filling " & role & _
                                          " but has no " & role & " in RoleAvail."
                            End If
                        End If
                    End If
                End If
            Next a
        End If
    Next t

    ' 5. Rule breaches
    For Each x In Violations()
        Set res = Nothing
        On Error Resume Next
        Set res = p.Resources(CLng(x(2)))
        On Error GoTo 0
        If Not res Is Nothing Then
            If CStr(x(0)) = CStr(x(1)) Then
                errs.Add "Task " & TaskIdOf(CStr(x(0))) & " " & TaskNameOf(CStr(x(0))) & ": " & _
                         res.Name & " is both " & x(3) & " and " & x(4) & " on the same task."
            Else
                errs.Add "Task " & TaskIdOf(CStr(x(0))) & " " & TaskNameOf(CStr(x(0))) & ": " & _
                         res.Name & " is " & x(4) & " here but was " & x(3) & " on task " & _
                         TaskIdOf(CStr(x(1))) & " (" & TaskNameOf(CStr(x(1))) & ")."
            End If
        End If
    Next x

    ' 6. Real overallocation (the hard wall)
    For Each r In p.Resources
        If IsWorkPerson(r) Then
            cap = TotalCapacity(r)
            over = 0
            peak = 0
            If cap > 0 Then
                For Each v In LoadWeeks()
                    wkLoad = TotalLoad(r.ID, CStr(v))
                    If wkLoad > cap * 1.001 Then
                        over = over + 1
                        If wkLoad / cap > peak Then peak = wkLoad / cap
                    End If
                Next v
                If over > 0 Then
                    warns.Add r.Name & ": past real availability in " & over & " week(s), peak " & _
                              Pct(peak) & " of Max Units. Run RS_Level."
                End If
            End If
        End If
    Next r

    If gRS_CoversCleared > 0 Then
        warns.Add gRS_CoversCleared & " " & RS_COVERS_LABEL & " entr(ies) were cleared because " & _
                  "nothing in them was usable - usually a task naming itself, which can't be " & _
                  "its own upstream work."
    End If

    rep.Add "Rules in force: " & IIf(RulesText() = "", "(none)", RulesText())
    rep.Add "Roles found: " & JoinCollection(RoleNames(), ", ")
    rep.Add "Columns found by name: " & IIf(FieldsFound() = "", _
             "none - using the default field numbers", FieldsFound())
    If missCount > 0 Then
        rep.Add "Columns NOT named: " & missList
        rep.Add "   Until they're named, those fall back to a fixed field number, so a file"
        rep.Add "   already using that field would collide."
    End If
    rep.Add ""
    AddSection rep, "MUST FIX (" & errs.Count & ")", errs
    AddSection rep, "WORTH A LOOK (" & warns.Count & ")", warns
    AddSection rep, "FOR INFORMATION (" & infos.Count & ")", infos
    FastOff
    ShowReport "Validate", rep

    MsgBox errs.Count & " must fix, " & warns.Count & " worth a look, " & _
           infos.Count & " for information.", vbInformation, "Role Staffing"
End Sub


' Weekly load per person, and how far it drifts from their RoleAvail shares.
' Drift is information, not a fault - lending time between roles is the point.
Public Sub RS_RoleLoad()
    Dim p As Project
    If Not HaveProject(p) Then Exit Sub
    FastOn "adding up role load"
    ApplyPendingTags p
    BuildCatalog p
    BuildLoad p

    Dim rep As New Collection, r As Resource, x As Variant, v As Variant
    Dim cap As Double, wkLoad As Double, roleMin As Double, tot As Double
    Dim weeks As Variant, wkList As Collection, k As Variant
    Dim actual As Object, grand As Double

    Set wkList = SortedKeys(LoadWeeks())

    rep.Add "Hard constraint is Max Units. RoleAvail shares are preferences - unused share"
    rep.Add "lends to other roles, so ""actual"" drifting from ""target"" is normal."
    rep.Add ""

    For Each r In p.Resources
        If IsWorkPerson(r) Then
            cap = TotalCapacity(r)
            If cap > 0 Then
                Set actual = NewDict()
                grand = 0
                For Each x In RoleNames()
                    roleMin = 0
                    For Each k In wkList
                        roleMin = roleMin + RoleLoad(r.ID, CStr(x), CStr(k))
                    Next k
                    actual(CStr(x)) = roleMin
                    grand = grand + roleMin
                Next x

                rep.Add r.Name & "   Max Units " & Pct(ToFraction(r.MaxUnits)) & _
                        "   =  " & Format$(cap / 60, "0.#") & " h/week"
                rep.Add "   " & PadR("role", 20) & PadL("target", 9) & PadL("actual", 10) & PadL("of total", 10)
                Dim shareCol As String
                For Each x In RoleNames()
                    If RoleShare(r, CStr(x)) > 0 Or ToDbl(actual(CStr(x))) > 0 Then
                        ' Worked out separately: IIf evaluates BOTH arms, so an
                        ' inline division would blow up when grand is 0.
                        If grand > 0 Then
                            shareCol = Pct(ToDbl(actual(CStr(x))) / grand)
                        Else
                            shareCol = "-"
                        End If
                        rep.Add "   " & PadR(CStr(x), 20) & _
                                PadL(Pct(RoleShare(r, CStr(x))), 9) & _
                                PadL(Format$(ToDbl(actual(CStr(x))) / 60, "0.#") & " h", 10) & _
                                PadL(shareCol, 10)
                    End If
                Next x

                ' weeks past the hard wall
                Dim overList As String
                overList = ""
                For Each k In wkList
                    wkLoad = TotalLoad(r.ID, CStr(k))
                    If wkLoad > cap * 1.001 Then
                        If overList <> "" Then overList = overList & ", "
                        overList = overList & CStr(k) & " (" & Pct(wkLoad / cap) & ")"
                    End If
                Next k
                If overList <> "" Then
                    rep.Add "   OVER real availability: " & overList
                End If
                rep.Add ""
            End If
        End If
    Next r

    ' unstaffed demand
    Dim anyOpen As Boolean
    rep.Add "Unstaffed demand still sitting on buckets"
    rep.Add "-----------------------------------------"
    For Each x In RoleNames()
        tot = 0
        For Each k In wkList
            tot = tot + OpenLoad(CStr(x), CStr(k))
        Next k
        If tot > 0 Then
            rep.Add "   " & PadR(CStr(x), 20) & PadL(Format$(tot / 60, "0.#") & " h", 10)
            anyOpen = True
        End If
    Next x
    If Not anyOpen Then rep.Add "   (none - everything is staffed)"

    FastOff
    ShowReport "Role Load", rep
End Sub
