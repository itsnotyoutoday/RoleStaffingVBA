VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRS_AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================
' Role Staffing: the manual-edit guard.
'
' Fires when someone changes the resource on an assignment by hand, e.g. typing
' a name over a bucket in Task Usage.
'   - Blocks the edit if it breaks a rule in RS_RULES (a drafter checking their
'     own work, and so on).
'   - Asks for confirmation if the person has no authority for that role, i.e.
'     the role isn't in their RoleAvail.
'
' Bulk paste may not raise events at all, so RS_Validate is the real backstop.
'
' IF VBA SAYS "Procedure declaration does not match description of event":
'   1. delete just the Sub ... line for that event
'   2. LEFT dropdown at the top of this window  -> App
'   3. RIGHT dropdown                           -> the event
'      the editor writes the exact signature for your Project version
'   4. paste the body back inside it
'
' Signatures differ between Project versions, and between the old events and the
' ones ending in 2. The dropdowns always get it right.
'==============================================================================

Public WithEvents App As Application

Private Sub App_ProjectBeforeAssignmentChange2(ByVal asg As Assignment, ByVal Field As PjAssignmentField, _
                                               ByVal NewVal As Variant, ByVal Info As EventInfo)
    On Error GoTo Quiet
    If Not gRS_GuardOn Or gRS_Suppress > 0 Then Exit Sub
    If Field <> pjAssignmentResourceName And Field <> pjAssignmentResourceID Then Exit Sub

    Dim p As Project, newRes As Resource
    Dim role As String, why As String

    Set p = ActiveProject

    If Field = pjAssignmentResourceID Then
        Set newRes = p.Resources(CLng(NewVal))
    Else
        Set newRes = ResourceByName(p, CStr(NewVal))
    End If
    If newRes Is Nothing Then Exit Sub
    If Not IsWorkPerson(newRes) Then Exit Sub

    ' Works out the role for this row and checks it against the rules.
    why = GuardCheck(asg, newRes, role)
    If role = "" Then Exit Sub

    If why <> "" Then
        MsgBox newRes.Name & " can't be " & role & " here:" & vbCrLf & vbCrLf & _
               "They " & why & "." & vbCrLf & vbCrLf & _
               "Rule in force: " & RulesText(), vbCritical, "Role Staffing"
        Info.Cancel = True
        Exit Sub
    End If

    If Not HasRole(newRes, role) Then
        If MsgBox(newRes.Name & " has no " & role & " in their RoleAvail, so they aren't " & _
                  "authorised for it." & vbCrLf & vbCrLf & "Assign anyway?", _
                  vbYesNo + vbExclamation, "Role Staffing") = vbNo Then
            Info.Cancel = True
            Exit Sub
        End If
    End If

    ' Can't edit an assignment from inside a "before" event, so stash the role
    ' and let ApplyPendingTags write it on the next save or macro run.
    If gRS_PendingTags Is Nothing Then Set gRS_PendingTags = NewDict()
    gRS_PendingTags(CStr(asg.UniqueID)) = role
    Exit Sub

Quiet:
    ' Never block the user's edit because the guard itself hit an error.
End Sub

Private Sub App_ProjectBeforeSave2(ByVal pj As Project, ByVal SaveAsUi As Boolean, ByVal Info As EventInfo)
    On Error Resume Next
    ApplyPendingTags pj
    CoversMaybeStale pj
End Sub


'==============================================================================
' TO ADD: refuse a SameWorkAs that names its own task, as you type it
'
' Project refuses a task as its own predecessor. This does the same for
' SameWorkAs - but it needs an event, and event signatures differ between
' Project versions. One wrong argument stops this whole class compiling and
' takes the working guard above down with it, so the Sub line isn't shipped.
' Thirty seconds to add, once:
'
'   1. click anywhere in this code window
'   2. LEFT dropdown at the top   -> App
'   3. RIGHT dropdown             -> ProjectBeforeTaskChange
'      the editor writes an empty Sub with the right signature for your version
'   4. put ONE line inside it:
'
'          Cancel = RS_BlockSelfRef(tsk, Field, NewVal)
'
'      Use whatever names the editor gave the arguments. If it called the last
'      one Info rather than Cancel:
'
'          Info.Cancel = RS_BlockSelfRef(tsk, Field, NewVal)
'
' Everything else lives in RoleStaffing.bas, so that line is the whole job.
'
' Until it's added, a self-reference is stripped and the field cleared the next
' time anything runs, and RS_Validate reports it - corrected rather than
' prevented.
'
'
' SAME AGAIN, OPTIONAL: re-point SameWorkAs the moment a row moves
'
'   RIGHT dropdown -> ProjectCalculate, then inside it:
'
'          On Error Resume Next
'          If gRS_GuardOn Then CoversMaybeStale pj
'
' Without it, SameWorkAs re-points whenever a macro runs or the file is saved,
' or on demand with RS_SyncCovers.
'==============================================================================
