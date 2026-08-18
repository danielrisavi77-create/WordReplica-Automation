"""Builds a Word ribbon add-in (.dotm) that runs word-replica from inside Word.

Word 2010 predates Office.js web add-ins (which need Word 2016+), so the only
route to a native ribbon tab here is a macro-enabled template: a VBA module
plus a Custom UI ribbon definition injected into the .dotm's OOXML package,
the same package-editing approach the Pure DOCX renderer already uses.

Requires "Trust access to the VBA project object model" enabled in Word's
Trust Center (Macro Settings) for this Windows account, purely so this
script can write the VBA module through Word's VBE object model. End users
who receive the finished .dotm do not need that setting themselves — they
are only loading an already-built macro project, not generating one.

Usage:
    python build_ribbon_addin.py [--install] [--visible] [--exe PATH]
"""
from __future__ import annotations

import argparse
import os
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUSTOM_UI_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"

CUSTOM_UI_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_OnLoad">
  <ribbon>
    <tabs>
      <tab id="tabWordReplica" label="Word Replica">
        <group id="grpOptions" label="Opcije">
          <dropDown id="ddRenderer" label="Renderer" onAction="OnRendererChange" getSelectedItemIndex="GetRendererSelectedIndex">
            <item id="auto" label="Automatski"/>
            <item id="word" label="Microsoft Word"/>
            <item id="docx" label="Pure DOCX"/>
          </dropDown>
          <dropDown id="ddVisibility" label="Vidljivost" onAction="OnVisibilityChange" getSelectedItemIndex="GetVisibilitySelectedIndex">
            <item id="visible" label="Vidljivo"/>
            <item id="background" label="U pozadini"/>
          </dropDown>
          <dropDown id="ddFidelity" label="Vjernost" onAction="OnFidelityChange" getSelectedItemIndex="GetFidelitySelectedIndex">
            <item id="clean" label="Čista replika"/>
            <item id="full" label="Puna vjernost"/>
          </dropDown>
          <dropDown id="ddMetadata" label="Metapodaci" onAction="OnMetadataChange" getSelectedItemIndex="GetMetadataSelectedIndex">
            <item id="fresh" label="Novi metapodaci"/>
            <item id="preserve" label="Sačuvaj legitimne"/>
          </dropDown>
        </group>
        <group id="grpRun" label="Rekonstrukcija">
          <button id="btnRebuild" label="Pokreni rekonstrukciju" size="large" onAction="RunRebuild" imageMso="FileSave"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"""

VBA_SOURCE = r'''Option Explicit

Private ribbonUI As IRibbonUI
Private gRendererIndex As Integer
Private gVisibilityIndex As Integer
Private gFidelityIndex As Integer
Private gMetadataIndex As Integer

Private Const RENDERER_IDS As String = "auto|word|docx"
Private Const VISIBILITY_IDS As String = "visible|background"
Private Const FIDELITY_IDS As String = "clean|full"
Private Const METADATA_IDS As String = "fresh|preserve"

Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set ribbonUI = ribbon
    gRendererIndex = 0
    gVisibilityIndex = 0
    gFidelityIndex = 0
    gMetadataIndex = 0
End Sub

Private Function IdAt(list As String, index As Integer) As String
    Dim parts() As String
    parts = Split(list, "|")
    IdAt = parts(index)
End Function

Sub OnRendererChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    gRendererIndex = selectedIndex
End Sub

Sub GetRendererSelectedIndex(control As IRibbonControl, ByRef index As Variant)
    index = gRendererIndex
End Sub

Sub OnVisibilityChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    gVisibilityIndex = selectedIndex
End Sub

Sub GetVisibilitySelectedIndex(control As IRibbonControl, ByRef index As Variant)
    index = gVisibilityIndex
End Sub

Sub OnFidelityChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    gFidelityIndex = selectedIndex
End Sub

Sub GetFidelitySelectedIndex(control As IRibbonControl, ByRef index As Variant)
    index = gFidelityIndex
End Sub

Sub OnMetadataChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    gMetadataIndex = selectedIndex
End Sub

Sub GetMetadataSelectedIndex(control As IRibbonControl, ByRef index As Variant)
    index = gMetadataIndex
End Sub

Private Function FindWordReplicaExe() As String
    Dim candidates(3) As String
    candidates(0) = Environ$("LOCALAPPDATA") & "\WordReplica\word-replica.exe"
    candidates(1) = "C:\WordReplica-Automation\.venv\Scripts\word-replica.exe"
    candidates(2) = Environ$("LOCALAPPDATA") & "\WordReplica\WordReplica.exe"
    candidates(3) = "C:\WordReplica-Automation\repo\dist\WordReplica.exe"
    Dim i As Integer
    For i = 0 To 3
        If Len(Dir$(candidates(i))) > 0 Then
            FindWordReplicaExe = candidates(i)
            Exit Function
        End If
    Next i
    FindWordReplicaExe = ""
End Function

Sub RunRebuild(control As IRibbonControl)
    DoRebuild
End Sub

' Parameterless so it is also callable directly via Application.Run for
' testing, since Application.Run cannot synthesize an IRibbonControl argument.
Sub DoRebuild()
    On Error GoTo Fail

    If Documents.Count = 0 Then
        MsgBox "Nema otvorenog dokumenta.", vbExclamation, "Word Replica"
        Exit Sub
    End If

    Dim docPath As String
    docPath = ActiveDocument.FullName
    If InStr(1, docPath, "\") = 0 Or LCase$(Right$(docPath, 5)) <> ".docx" Then
        MsgBox "Sačuvaj dokument kao .docx prije rekonstrukcije.", vbExclamation, "Word Replica"
        Exit Sub
    End If

    Dim exePath As String
    exePath = FindWordReplicaExe()
    If exePath = "" Then
        MsgBox "word-replica.exe nije pronađen. Instaliraj Word Replica ili prilagodi FindWordReplicaExe u makrou.", vbCritical, "Word Replica"
        Exit Sub
    End If

    Dim rendererArg As String, visibilityArg As String, fidelityArg As String, metadataArg As String
    rendererArg = IdAt(RENDERER_IDS, gRendererIndex)
    visibilityArg = IdAt(VISIBILITY_IDS, gVisibilityIndex)
    fidelityArg = IdAt(FIDELITY_IDS, gFidelityIndex)
    metadataArg = IdAt(METADATA_IDS, gMetadataIndex)

    ' /k (not /c) leaves the console window open with a fresh prompt after the
    ' command finishes, so the CLI's own output stays on screen until the user
    ' closes it. Shell (not WScript.Shell.Run with wait:=True) returns
    ' immediately: Word never blocks its own UI thread waiting for this.
    Dim cmd As String
    cmd = "cmd /k """"" & exePath & """ rebuild """ & docPath & """" & _
          " --renderer " & rendererArg & " --visibility " & visibilityArg & _
          " --fidelity " & fidelityArg & " --metadata " & metadataArg & """"

    Shell cmd, vbNormalFocus
    Exit Sub

Fail:
    MsgBox "Word Replica greška: " & Err.Description, vbCritical, "Word Replica"
End Sub
'''


def build_dotm(output_path: Path, *, visible: bool = False) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    try:
        word.DisplayAlerts = 0
        word.Visible = bool(visible)
        doc = word.Documents.Add()
        try:
            project = word.VBE.VBProjects(1)
        except Exception as exc:
            raise RuntimeError(
                "Cannot access the VBA project model. Enable 'Trust access to the "
                "VBA project object model' in Word's Trust Center > Macro Settings."
            ) from exc
        component = project.VBComponents.Add(1)  # vbext_ct_StdModule
        component.CodeModule.AddFromString(VBA_SOURCE)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        # wdFormatXMLTemplateMacroEnabled
        doc.SaveAs2(str(output_path), FileFormat=15)
        doc.Close(False)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()


def _xml(root_tag: str, encoding: str = "UTF-8") -> bytes:
    return f"<?xml version=\"1.0\" encoding=\"{encoding}\" standalone=\"yes\"?>\n<{root_tag}/>".encode(encoding)


def inject_custom_ui(dotm_path: Path) -> None:
    from lxml import etree

    with ZipFile(dotm_path, "r") as src:
        members = {name: src.read(name) for name in src.namelist()}

    content_types = etree.fromstring(members["[Content_Types].xml"])
    override_exists = content_types.xpath(
        "//*[local-name()='Override'][@PartName='/customUI/customUI14.xml']"
    )
    if not override_exists:
        override = etree.SubElement(content_types, f"{{{CT_NS}}}Override")
        override.set("PartName", "/customUI/customUI14.xml")
        override.set("ContentType", "application/xml")
    members["[Content_Types].xml"] = etree.tostring(
        content_types, xml_declaration=True, encoding="UTF-8", standalone=True
    )

    rels_path = "_rels/.rels"
    rels = etree.fromstring(members[rels_path])
    existing_ids = {rel.get("Id") for rel in rels}
    already_linked = rels.xpath(
        f"//*[local-name()='Relationship'][@Type='{CUSTOM_UI_RELATIONSHIP_TYPE}']"
    )
    if not already_linked:
        rel_id = "rIdCustomUI"
        while rel_id in existing_ids:
            rel_id += "x"
        relationship = etree.SubElement(rels, f"{{{REL_NS}}}Relationship")
        relationship.set("Id", rel_id)
        relationship.set("Type", CUSTOM_UI_RELATIONSHIP_TYPE)
        relationship.set("Target", "customUI/customUI14.xml")
    members[rels_path] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    members["customUI/customUI14.xml"] = CUSTOM_UI_XML.encode("utf-8")

    tmp_path = dotm_path.with_suffix(".tmp.dotm")
    with ZipFile(tmp_path, "w", ZIP_DEFLATED) as dst:
        for name, data in members.items():
            dst.writestr(name, data)
    tmp_path.replace(dotm_path)


def install_to_startup(dotm_path: Path) -> Path:
    startup_dir = Path(os.environ["APPDATA"]) / "Microsoft" / "Word" / "STARTUP"
    startup_dir.mkdir(parents=True, exist_ok=True)
    destination = startup_dir / dotm_path.name
    destination.write_bytes(dotm_path.read_bytes())
    return destination


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=Path(r"C:\WordReplica-Automation\repo\dist\WordReplicaRibbon.dotm"))
    parser.add_argument("--install", action="store_true", help="Copy the built .dotm into Word's STARTUP folder")
    parser.add_argument("--visible", action="store_true", help="Show Word while building (debugging)")
    args = parser.parse_args()

    build_dotm(args.output, visible=args.visible)
    inject_custom_ui(args.output)
    print(f"Built: {args.output}")

    if args.install:
        installed = install_to_startup(args.output)
        print(f"Installed to STARTUP: {installed}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
