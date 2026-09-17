#!/usr/bin/env python3
"""Static checks for the VBA sources, for things the editor only finds at
Debug > Compile - by which point the code is already on someone's machine.

    python3 tools/check_vba.py

Checks, in the order they've actually bitten:

  1. CRLF line endings. Windows MS Project's importer can't parse a class
     header without them and reports "Expected: end of statement" on
     VERSION 1.0 CLASS.
  2. Reserved words used as identifiers. "Dim any As Boolean" is a syntax
     error because Any is a keyword (Declare ... As Any).
  3. Module-level declarations after the first procedure. VBA requires them
     all at the top: "only comments may appear after End Sub".
  4. Block balance: Sub/Function, If/End If, For/Next, Do/Loop, With/End With,
     Select/End Select.
  5. Constants that don't exist, e.g. pjStandard. A bare constant is
     compile-checked even inside a late-bound call.
"""
import re, sys, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILES = ["RoleStaffing.bas", "RS_TaskData.bas", "RS_Staffing.bas", "RS_UI.bas",
         "clsRS_AppEvents.cls", "clsRS_RowBtn.cls"]
# Generated, but it has to compile inside MS Project on a machine we can't see,
# so it is checked exactly like the rest.
GENERATED = ["tools/probe/RS_ModelProbe.bas", "tools/probe/RS_TypeLibDump.bas"]
PASTE = ["paste/RoleStaffing.txt", "paste/RS_TaskData.txt", "paste/RS_Staffing.txt",
         "paste/RS_UI.txt", "paste/clsRS_AppEvents.txt", "paste/clsRS_RowBtn.txt"]

RESERVED = set(w.lower() for w in """
Alias And Any As Base Binary Boolean ByRef Byte ByVal Call Case Circle Class Close Compare Const
Currency Date Debug Decimal Declare Dim Do Double Each Else ElseIf Empty End Enum Eqv Erase Error
Event Exit Explicit False For Friend Function Get Global GoSub GoTo If Imp Implements In Input Int
Integer Is Len Let Lib Like Line Load Lock Long Loop LSet Me Mid Mod Module New Next Not Nothing
Null Object Of On Open Option Optional Or Output ParamArray Preserve Print Private Property PSet
Public Put RaiseEvent Random Read ReDim Rem Resume Return RSet Scale Seek Select Set Shared Single
Spc Static Step Stop String Sub Tab Text Then Time To True Type TypeOf Unload Unlock Until Variant
Wend While Width With WithEvents Write Xor Name Left Right Val Str Format Year Month Day Hour
Log Exp Sqr Sgn Abs Fix Rnd Timer Now Space Split Join Replace Trim Asc Chr Hex Oct Sin Cos Tan Atn
Array Choose Switch IIf Dir Shell Environ Command StrConv Round
""".split())
DECL_KW = set("byval byref optional paramarray as dim static private public const sub function".split())

# Project constants this code is allowed to name. Anything else beginning "pj"
# is probably invented - pjStandard was, and it stopped the module compiling.
# Every pj* constant MS Project actually has: 4,907 of them across 172
# enumerations, from the extraction that feeds the type model
# (vendor/VBAlidator/tools/data/project_vba_enums.json).
#
# This was a hand-maintained list of seventeen names. Seventeen is enough to
# catch pjStandard, which is what it was written for, and not enough for
# anything else - every real constant we hadn't used yet read as invented. The
# extracted data has no such gap, and it cost nothing: the same file feeds the
# type model.
def _known_pj():
    # Lives in the VBAlidator submodule with the rest of the model tooling.
    path = os.path.join(ROOT, "vendor", "VBAlidator", "tools", "data",
                        "project_vba_enums.json")
    if not os.path.exists(path):
        return None
    import json
    data = json.load(open(path))
    return {m.lower() for e in data.values() for m in e.get("members", {})}


KNOWN_PJ = _known_pj()

def strip_comment(line):
    """Cut a trailing ' comment, but not an apostrophe inside a string literal.
    Getting this wrong made the checker itself report false imbalances on any
    statement containing a word like "everyone's"."""
    out, in_str = [], False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            # "" inside a string is an escaped quote
            if in_str and i + 1 < len(line) and line[i + 1] == '"':
                out.append('""'); i += 2; continue
            in_str = not in_str
        elif ch == "'" and not in_str:
            break
        out.append(ch); i += 1
    return "".join(out)


def statements(src):
    """(line number, folded statement) with continuations joined, keeping the
    line number the statement started on."""
    out, buf, start = [], "", 0
    for n, line in enumerate(src.split("\n"), 1):
        s = line.rstrip()
        if not buf:
            start = n
        if s.endswith("_"):
            buf += s[:-1] + " "
            continue
        out.append((start, (buf + s).strip()))
        buf = ""
    return out


problems = []


def note(f, msg):
    problems.append(f"{f}: {msg}")


def joined(src):
    """Source with line continuations folded, for block counting."""
    return re.sub(r"\s_\n\s*", " ", src)


def check_endings(path, raw):
    if b"\r\n" not in raw:
        note(path, "LF line endings - MS Project's importer needs CRLF")
    elif re.search(rb"(?<!\r)\n", raw):
        note(path, "mixed line endings - should be CRLF throughout")


def check_reserved(path, src):
    for n, line in statements(src):
        s = line.strip()
        if s.startswith("'"):
            continue
        s = strip_comment(s)
        names = []
        m = re.match(r"\s*(?:Dim|Static)\s+(.+)$", s, re.I)
        if m:
            names += [(re.match(r"\s*([A-Za-z_]\w*)", p), "Dim") for p in m.group(1).split(",")]
        # Module-level state and constants. Missed until a test went looking:
        # a reserved word in "Public Text As String" is the same compile error
        # as in "Dim Text As String", and this module is mostly Publics.
        m = re.match(r"\s*(?:Public|Private|Global)\s+(?:Const\s+)?"
                     r"([A-Za-z_]\w*)\s*(?:\(\))?\s+As\s", s, re.I)
        if m:
            names.append((m, "module-level"))
        m = re.match(r"\s*(?:Public |Private )?Const\s+([A-Za-z_]\w*)", s, re.I)
        if m:
            names.append((m, "Const"))

        m = re.match(r"\s*(?:Public |Private )?(?:Sub|Function)\s+\w+\s*\((.*?)\)", s, re.I)
        if m and m.group(1).strip():
            names += [(re.search(r"([A-Za-z_]\w*)\s*(?:\(\))?\s+As\s", p), "parameter")
                      for p in m.group(1).split(",")]
        for v, kind in names:
            if v and v.group(1).lower() in RESERVED and v.group(1).lower() not in DECL_KW:
                note(path, f"line {n}: '{v.group(1)}' is a VBA reserved word ({kind})")
        for v in (re.findall(r"\bFor\s+Each\s+([A-Za-z_]\w*)", s, re.I) +
                  re.findall(r"\bFor\s+([A-Za-z_]\w*)\s*=", s, re.I)):
            if v.lower() in RESERVED and v.lower() not in DECL_KW:
                note(path, f"line {n}: '{v}' is a VBA reserved word (loop variable)")


def check_decl_position(path, src):
    inproc = seen = False
    for n, line in enumerate(src.splitlines(), 1):
        s = line.strip()
        if not s or s.startswith("'"):
            continue
        if re.match(r"(public |private )?(sub|function|property) ", s, re.I):
            inproc, seen = True, True
            continue
        if re.match(r"end (sub|function|property)", s, re.I):
            inproc = False
            continue
        if not inproc and seen and re.match(
                r"(public|private|dim|const|type|enum|declare|option)\b", s, re.I):
            note(path, f"line {n}: module-level declaration after a procedure - "
                       f"VBA needs these at the top: {s[:60]}")


def check_blocks(path, src):
    """Walk the nesting per procedure, so a mismatch names the line."""
    depth, opened, proc = 0, [], "(module level)"
    for n, raw in statements(src):
        s = strip_comment(raw).strip()
        if not s:
            continue
        if re.match(r"(public |private )?(sub|function|property) ", s, re.I):
            if depth:
                note(path, f"line {n}: entering {s[:40]} with {depth} unclosed If "
                           f"(opened at {opened})")
            depth, opened, proc = 0, [], s[:50]
            continue
        if re.match(r"end (sub|function|property)\b", s, re.I):
            if depth:
                note(path, f"line {n}: {proc} ends with {depth} unclosed If "
                           f"(opened at {opened})")
            depth, opened = 0, []
            continue
        if re.match(r"if .*\bthen\s*$", s, re.I):
            depth += 1
            opened.append(n)
        elif re.match(r"end if\b", s, re.I):
            depth -= 1
            if opened:
                opened.pop()
            if depth < 0:
                note(path, f"line {n}: 'End If' with no matching If, in {proc}")
                depth = 0

    # the simpler pairings, counted
    code = [strip_comment(s).strip() for _, s in statements(src)]
    code = [l for l in code if l]

    def c(pat):
        return sum(1 for l in code if re.match(pat, l, re.I))

    for name, a, b in [
        ("Sub/Function", c(r"(public |private )?(sub|function) "), c(r"end (sub|function)")),
        ("For/Next", c(r"for (each )?"), c(r"next\b")),
        ("Do/Loop", c(r"do\b"), c(r"^loop")),
        ("With/End With", c(r"^with "), c(r"end with")),
        ("Select/End Select", c(r"select case"), c(r"end select")),
    ]:
        if a != b:
            note(path, f"unbalanced {name}: {a} open, {b} close")


def check_assign_to_call(path, src):
    """A call to one of our own Functions on the left of = is a compile error:
    "Function call on left-hand side of assignment must return Variant or
    Object". Easy to introduce when converting direct field reads into
    accessors, and invisible until the editor refuses it.

    Only names declared as Function in this file count - d(key) = v on a
    Dictionary and arr(i) = v on an array are both perfectly legal."""
    funcs = set()
    for _, raw in statements(src):
        m = re.match(r"\s*(?:Public |Private )?Function\s+(\w+)", strip_comment(raw), re.I)
        if m:
            funcs.add(m.group(1).lower())

    for n, raw in statements(src):
        line = strip_comment(raw)
        for m in re.finditer(r"\b([A-Za-z_]\w*)\(([^()]*)\)\s*=\s*(?![=<>])", line):
            if m.group(1).lower() not in funcs:
                continue
            pre = line[:m.start()].rstrip()
            if pre == "" or pre.lower().endswith(("then", "else", ":")):
                note(path, f"line {n}: assignment to the function "
                           f"'{m.group(1)}(...)' - needs a Set... procedure")


def check_continuations(path, src):
    """VBA allows 25 line continuations in one statement. Past that it reports
    "Too many line continuations" and the module won't compile. Easy to pass
    when a menu or a message grows an entry at a time."""
    n, start = 0, 0
    for i, line in enumerate(src.split("\n"), 1):
        if line.rstrip().endswith("_"):
            if n == 0:
                start = i
            n += 1
        else:
            if n > 24:
                note(path, f"line {start}: statement uses {n} line continuations "
                           f"- VBA allows 25. Build the string in pieces.")
            n = 0


def check_pairs(path, src):
    """OpenUndoTransaction/Close, the gRS_Suppress counter and FastOn/FastOff
    must balance. Closing one that was never opened is the dangerous direction:
    a stray FastOff inside an outer fast block turns screen updating back on
    mid-operation, and a stray suppress decrement re-arms the guard while a
    macro is still editing."""
    pairs = [("Application.OpenUndoTransaction", "Application.CloseUndoTransaction", "undo transaction"),
             ("gRS_Suppress = gRS_Suppress + 1", "gRS_Suppress = gRS_Suppress - 1", "suppress counter"),
             ("FastOn ", "FastOff", "fast mode")]
    cur, body = None, []
    for n, raw in statements(src):
        t = strip_comment(raw).strip()
        m = re.match(r"(?:Public |Private )?(?:Sub|Function)\s+(\w+)", t)
        if m:
            cur, body = m.group(1), []
            continue
        if re.match(r"End (Sub|Function)", t) and cur:
            for opn, cls, label in pairs:
                o = sum(1 for x in body if x.startswith(opn))
                c_ = sum(1 for x in body if cls in x)
                if c_ and not o:
                    note(path, f"{cur}: closes the {label} {c_}x but never opens it")
                if o and not c_:
                    note(path, f"{cur}: opens the {label} {o}x but never closes it")
            cur = None
            continue
        if cur is not None:
            body.append(t)


def check_constants(path, src):
    if KNOWN_PJ is None:
        note(path, "can't check pj* constants: the VBAlidator submodule isn't "
                   "checked out - git submodule update --init")
        return
    for n, line in enumerate(src.splitlines(), 1):
        if line.strip().startswith("'"):
            continue
        for m in re.findall(r"\bpj[A-Za-z]\w*", line):
            if m.lower() not in KNOWN_PJ:
                note(path, f"line {n}: '{m}' is not a Project constant - it is "
                           f"not one of the 4,907 in any documented enumeration")


def check_cross_module():
    """VBA modules share one namespace, but Private doesn't cross them. A
    Private procedure or variable called from another module compiles as
    "Variable not defined" - which points at the call, not the cause, and is
    exactly what splitting one module into four will produce."""
    import glob
    mods = [f for f in FILES if f.endswith(".bas")]
    owner = {}
    for rel in mods:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            continue
        src = open(path, newline="").read().replace("\r\n", "\n")
        for m in re.finditer(r"^Private (?:Sub|Function)\s+(\w+)", src, re.M):
            owner[m.group(1)] = rel
        for m in re.finditer(r"^Private (m\w+)\s", src, re.M):
            owner[m.group(1)] = rel

    for rel in FILES:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            continue
        src = open(path, newline="").read().replace("\r\n", "\n")
        src = re.sub(r"^\s*'.*$", "", src, flags=re.M)
        for name, home in owner.items():
            if home == rel:
                continue
            if re.search(r"\b" + re.escape(name) + r"\b", src):
                note(rel, f"uses '{name}', which is Private in {home} - "
                          f"make it Public, or move it")


def check_module_names():
    """A module can't share its name with a procedure - VBA reports "expected
    variable or procedure, not module" at the CALL, which says nothing about
    the cause. Easy to create when splitting one module into several and naming
    one after what it does."""
    procs = set()
    for rel in FILES:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            continue
        src = open(path, newline="").read().replace("\r\n", "\n")
        for m in re.finditer(r"^(?:Public |Private )?(?:Sub|Function)\s+(\w+)", src, re.M):
            procs.add(m.group(1))
    for rel in FILES:
        mod = os.path.basename(rel).rsplit(".", 1)[0]
        if mod in procs:
            note(rel, f"module is called '{mod}', which is also a procedure name - "
                      f"rename the module")


def check_name_collisions():
    """VBA identifiers are case-insensitive, so a constant RS_ROLEFORM and a Sub
    RS_RoleForm are one name - the procedure wins and the constant reads as
    "expected function or variable" wherever it's used. Distinguishing two
    things only by capitalisation looks fine and isn't."""
    import collections
    names = collections.defaultdict(list)
    for rel in FILES:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            continue
        src = open(path, newline="").read().replace("\r\n", "\n")
        for pat, kind in [(r"^(?:Public |Private )?(?:Sub|Function)\s+(\w+)", "procedure"),
                          (r"^(?:Public|Private) Const (\w+)", "constant"),
                          (r"^(?:Public|Private) ([mg]\w+)\s+As", "variable")]:
            for m in re.finditer(pat, src, re.M):
                names[m.group(1).lower()].append((rel, kind, m.group(1)))
    for _, entries in sorted(names.items()):
        if len(entries) > 1:
            where = "; ".join(f"{o} ({k} in {f})" for f, k, o in entries)
            note(entries[0][0], f"same name, different capitalisation: {where}")


def check_calls_resolve():
    """Every procedure called must be defined somewhere. Catches a deletion that
    took more than intended - removing three helpers by pattern also removed
    RS_ExportAll and OwnVBProject, and nothing complained until a menu item was
    pressed.

    Constants, module-level variables and module names are collected too, or
    every RS_RULES and RS_UI would read as a missing procedure."""
    known, called = set(), {}
    for rel in FILES:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            continue
        known.add(os.path.basename(rel).rsplit(".", 1)[0].lower())
        src = open(path, newline="").read().replace("\r\n", "\n")
        for pat in (r"^(?:Public |Private )?(?:Sub|Function)\s+(\w+)",
                    r"^(?:Public|Private) Const (\w+)",
                    r"^(?:Public|Private) (\w+)\s+As"):
            for m in re.finditer(pat, src, re.M):
                known.add(m.group(1).lower())

    for rel in FILES:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            continue
        src = open(path, newline="").read().replace("\r\n", "\n")
        src = re.sub(r"^\s*'.*$", "", src, flags=re.M)
        src = re.sub(r'"[^"]*"', '""', src)          # and not names inside strings
        for m in re.finditer(r"\b((?:RS_|Ensure|Refit|Sync|Resolve|Build)\w+)\b", src):
            if m.group(1).lower() not in known:
                called.setdefault(m.group(1), rel)

    for n, rel in sorted(called.items()):
        note(rel, f"calls '{n}', which isn't defined anywhere")


def check_form_markers():
    """RS_*Form finds an already-built form by a procedure its code-behind
    calls. If that procedure is renamed or removed the form is never
    recognised, so every run builds another numbered one - which is exactly
    what happened when the check form's ListBox helpers were replaced."""
    path = os.path.join(ROOT, "RS_UI.bas")
    if not os.path.exists(path):
        return
    src = open(path, newline="").read().replace("\r\n", "\n")

    defined = set()
    for rel in FILES:
        p = os.path.join(ROOT, rel)
        if os.path.exists(p):
            s2 = open(p, newline="").read().replace("\r\n", "\n")
            defined |= set(re.findall(r"^Public (?:Sub|Function)\s+(\w+)", s2, re.M))

    for m in re.finditer(r'(?:FindFormCalling\(|KillForms[ (][^,]+, )"(\w+)"', src):
        nm = m.group(1)
        if nm not in defined:
            note("RS_UI.bas", f"form marker '{nm}' names a procedure that no longer exists")
        else:
            # It has to appear in generated code, which is built as string
            # literals here. Not "<nm> Me": the panel's marker is called as a
            # function in an assignment, and testing for one calling convention
            # would have reported working code as broken.
            quoted = re.findall(r'"([^"]*)"', src)
            if not any(nm in q and q != nm for q in quoted):
                note("RS_UI.bas", f"form marker '{nm}' appears in no generated form code")


def check_probe_reads_only():
    """The probe macro must never name a METHOD.

    CallByName(obj, name, VbGet) on a zero-argument method CALLS it. The first
    version of the probe took that for a read, and probing Application in
    alphabetical order ran CalculateAll, CalculateProject,
    CloseUndoTransaction, LevelNow - which re-levels the project and moves
    dates - then Quit, closing MS Project on the user's machine. It stopped
    there only because Application sorts before Task; a few names later it
    would have called Task.Delete.

    So this reads the generated macro back, pulls out every name it will
    actually probe, and checks each one against the model. Reviewing the
    generator isn't enough - the generator is what got this wrong."""
    import json
    probe = os.path.join(ROOT, "tools", "probe", "RS_ModelProbe.bas")
    model_path = os.path.join(ROOT, "vendor", "VBAlidator", "src",
                              "models", "project.json")
    if not (os.path.exists(probe) and os.path.exists(model_path)):
        return

    src = open(probe, newline="").read().replace("\r\n", "\n")
    model = json.load(open(model_path))

    # Per class, not globally. UniqueID is a plain property on Task and a
    # parameterised lookup returning an object on Tasks - checking names across
    # all classes at once conflates the two and reports a safe read as a risk.
    kind = {}
    for cls_name, cls in model.get("classes", {}).items():
        for nm, spec in cls.get("members", {}).items():
            kind[(cls_name, nm)] = spec.get("type")

    # The names go out as pipe-delimited Split() literals, and the Probe line
    # that follows names the class. Only those matter - NOT-PROBED-METHOD chunks
    # are reported, never called.
    probed, block = set(), None
    for line in src.splitlines():
        m = re.search(r'names = Split\("([^"]+)", "\|"\)', line)
        if m:
            block = m.group(1).split("|")
            continue
        if block is None:
            continue
        m = re.search(r'Probe "(\w+)"', line)
        if m:
            probed.update((m.group(1), nm) for nm in block)
            block = None
        elif "NOT-PROBED-METHOD" in line:
            block = None

    if not probed:
        note("tools/probe/RS_ModelProbe.bas",
             "couldn't find any probed names - has the generator changed shape?")
        return

    bad = sorted(f"{c}.{n}" for c, n in probed if kind.get((c, n)) == "Function")
    if bad:
        note("tools/probe/RS_ModelProbe.bas",
             f"probes {len(bad)} METHOD(s), which CallByName would CALL: "
             f"{', '.join(bad[:8])}")

    names = {n for _, n in probed}
    for danger in ("Quit", "Delete", "LevelNow", "Save", "SaveAs", "UndoClear"):
        if danger in names:
            note("tools/probe/RS_ModelProbe.bas",
                 f"probes '{danger}' - CallByName would run it")


def check_generated():
    """The probe macro is generated, but it runs inside MS Project on a machine
    we never see. Same per-file checks as shipped code: a reserved word or an
    unbalanced block there costs a round trip to the office and back."""
    for rel in GENERATED:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            continue
        raw = open(path, newline="").read()
        if "\r\n" not in raw:
            note(rel, "LF endings - VBA's importer needs CRLF")
        src = raw.replace("\r\n", "\n")
        check_reserved(rel, src)
        check_decl_position(rel, src)
        check_blocks(rel, src)
        # VBA gives up past 1023 characters on one line, and the member names
        # go out as chunked string literals - so this is a real limit, not a
        # style rule.
        for n, line in enumerate(src.splitlines(), 1):
            if len(line) > 1000:
                note(rel, f"line {n}: {len(line)} chars - VBA's limit is 1023")


def main():
    for rel in FILES + PASTE:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            note(rel, "missing")
            continue
        raw = open(path, "rb").read()
        src = raw.decode("utf-8", "replace").replace("\r\n", "\n")
        check_endings(rel, raw)
        check_reserved(rel, src)
        if rel in FILES:
            check_decl_position(rel, src)
            check_assign_to_call(rel, src)
            check_continuations(rel, src)
            check_pairs(rel, src)
            check_blocks(rel, src)
            check_constants(rel, src)

    check_cross_module()
    check_module_names()
    check_name_collisions()
    check_calls_resolve()
    check_form_markers()
    check_generated()
    check_probe_reads_only()

    # paste/ must match its source, minus the header
    for src_rel, paste_rel in zip(FILES, PASTE):
        sp, pp = os.path.join(ROOT, src_rel), os.path.join(ROOT, paste_rel)
        if not (os.path.exists(sp) and os.path.exists(pp)):
            continue
        s = open(sp, newline="").read().replace("\r\n", "\n")
        p = open(pp, newline="").read().replace("\r\n", "\n")
        body = s[s.index("Option Explicit"):] if "Option Explicit" in s else s
        if p.strip() != body.strip():
            note(paste_rel, "out of step with its source - regenerate it")

    if problems:
        print(f"{len(problems)} problem(s):\n")
        for p in problems:
            print("  " + p)
        return 1
    print("VBA checks passed: endings, reserved words, declaration order, "
          "block balance, constants, paste/ in step.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
