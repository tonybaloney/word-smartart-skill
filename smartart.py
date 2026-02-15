"""
smartart.py - Create native SmartArt (DrawingML Diagrams) in Word documents.

This module creates real, editable SmartArt objects in .docx files. It works by
extracting diagram XML parts from Word-generated templates and injecting modified
versions into target documents using python-docx and lxml.

No images are generated - these are native SmartArt objects that render and are
fully editable in Microsoft Word.

Supported diagram types:
- Basic List: Simple block list of items
- Basic Process: Linear flow from left to right
- Hierarchy: Org chart / tree structure
- Cycle: Circular process
- Pyramid: Layered pyramid
- Radial: Central item with radiating connections

Requirements:
    pip install python-docx lxml

Usage:
    from docx import Document
    from smartart import SmartArt

    doc = Document()
    doc.add_heading("My Document", 0)

    SmartArt.add_basic_process(doc, "My Process", [
        "Step 1: Research",
        "Step 2: Design",
        "Step 3: Implement",
        "Step 4: Test",
    ])

    doc.save("output.docx")
"""

import os
import copy
import uuid
import zipfile
from typing import Optional
from lxml import etree
from docx import Document
from docx.opc.part import Part
from docx.opc.packuri import PackURI

# ─── Paths ──────────────────────────────────────────────────────────────────

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")

# ─── OpenXML Namespaces ────────────────────────────────────────────────────

NS = {
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "dsp": "http://schemas.microsoft.com/office/drawing/2008/diagram",
}

# Content types for SmartArt parts
CT_DGM_DATA = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
CT_DGM_LAYOUT = "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"
CT_DGM_STYLE = "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"
CT_DGM_COLORS = "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"
CT_DGM_DRAWING = "application/vnd.ms-office.drawingml.diagramDrawing+xml"

# Relationship types for diagram parts (must match what Word expects)
RT_DGM_DATA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
RT_DGM_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
RT_DGM_STYLE = "http://schemas.microsoft.com/office/2007/relationships/diagramStyle"
RT_DGM_COLORS = "http://schemas.microsoft.com/office/2007/relationships/diagramColors"
RT_DGM_DRAWING = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"

# Map template names to layout URIs
LAYOUT_URIS = {
    "basic_list": "urn:microsoft.com/office/officeart/2005/8/layout/default",
    "basic_process": "urn:microsoft.com/office/officeart/2005/8/layout/process1",
    "hierarchy": "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1",
    "cycle": "urn:microsoft.com/office/officeart/2005/8/layout/cycle1",
    "pyramid": "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1",
    "radial": "urn:microsoft.com/office/officeart/2005/8/layout/radial1",
}


def _qn(ns: str, tag: str) -> str:
    """Create a qualified name {namespace}tag."""
    return "{{{}}}{}".format(NS[ns], tag)


def _new_guid() -> str:
    """Generate a GUID in SmartArt format."""
    return "{{{}}}".format(str(uuid.uuid4()).upper())


def _next_diagram_id(doc: Document) -> int:
    """Find the next available diagram part index in the document."""
    existing = set()
    for rel in doc.part.rels.values():
        target = rel.target_ref if hasattr(rel, 'target_ref') else ""
        if "diagrams/data" in target:
            try:
                num = int(target.split("data")[1].split(".")[0])
                existing.add(num)
            except (ValueError, IndexError):
                pass
    return max(existing, default=0) + 1


# ─── Template Extraction ───────────────────────────────────────────────────

def _extract_template_parts(template_name: str) -> dict:
    """
    Extract diagram XML parts from a pre-generated Word template.

    Returns dict with keys: 'data', 'layout', 'style', 'colors', 'drawing'
    Each value is the raw XML bytes for that part.
    """
    template_path = os.path.join(TEMPLATE_DIR, "{}.docx".format(template_name))
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            "Template '{}' not found at {}. Run generate_templates.py first.".format(
                template_name, template_path
            )
        )

    parts = {}
    # Map of file name patterns to part keys
    part_map = {
        "data": "data",
        "layout": "layout",
        "quickStyle": "style",
        "colors": "colors",
        "drawing": "drawing",
    }

    with zipfile.ZipFile(template_path, 'r') as z:
        for name in z.namelist():
            if "diagrams/" not in name:
                continue
            for pattern, key in part_map.items():
                if pattern in name:
                    parts[key] = z.read(name)
                    break

    required = {'data', 'layout', 'style', 'colors', 'drawing'}
    missing = required - set(parts.keys())
    if missing:
        raise ValueError("Template missing parts: {}".format(missing))

    return parts


def _extract_rels_from_template(template_name: str) -> dict:
    """Extract the relationship types from the template's document.xml.rels."""
    template_path = os.path.join(TEMPLATE_DIR, "{}.docx".format(template_name))
    rel_types = {}

    with zipfile.ZipFile(template_path, 'r') as z:
        rels_xml = z.read("word/_rels/document.xml.rels")
        root = etree.fromstring(rels_xml)
        for rel in root:
            rtype = rel.get("Type", "")
            target = rel.get("Target", "")
            if "diagramData" in rtype:
                rel_types["data"] = rtype
            elif "diagramLayout" in rtype:
                rel_types["layout"] = rtype
            elif "diagramStyle" in rtype or "diagramQuickStyle" in rtype:
                rel_types["style"] = rtype
            elif "diagramColors" in rtype:
                rel_types["colors"] = rtype
            elif "diagramDrawing" in rtype:
                rel_types["drawing"] = rtype

    return rel_types


# ─── Data Model Modification ───────────────────────────────────────────────

def _build_data_model_flat(template_data_xml: bytes, items: list[str],
                           layout_uri: str) -> bytes:
    """
    Build a new data model XML for flat diagrams (list, process, cycle, etc.)
    based on the template's structure but with custom items.

    The template's data model has a known structure with parTrans/sibTrans points.
    We rebuild the data model with the correct number of items.
    """
    root = etree.Element(_qn("dgm", "dataModel"), nsmap={
        "dgm": NS["dgm"],
        "a": NS["a"],
        "r": NS["r"],
    })

    pt_lst = etree.SubElement(root, _qn("dgm", "ptLst"))
    cxn_lst = etree.SubElement(root, _qn("dgm", "cxnLst"))

    # Parse template to extract the doc point's prSet attributes
    tpl_root = etree.fromstring(template_data_xml)
    tpl_doc_pt = tpl_root.find(".//dgm:pt[@type='doc']", NS)
    tpl_pr_set = tpl_doc_pt.find("dgm:prSet", NS) if tpl_doc_pt is not None else None

    # Create doc point (root)
    doc_guid = _new_guid()
    doc_pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={
        "modelId": doc_guid,
        "type": "doc",
    })

    # Copy prSet from template or create default
    if tpl_pr_set is not None:
        pr_set = copy.deepcopy(tpl_pr_set)
        doc_pt.append(pr_set)
    else:
        etree.SubElement(doc_pt, _qn("dgm", "prSet"), attrib={
            "loTypeId": layout_uri,
            "loCatId": "list",
            "qsTypeId": "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
            "qsCatId": "simple",
            "csTypeId": "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
            "csCatId": "accent1",
            "phldr": "1",
        })

    etree.SubElement(doc_pt, _qn("dgm", "spPr"))
    t_elem = etree.SubElement(doc_pt, _qn("dgm", "t"))
    etree.SubElement(t_elem, _qn("a", "bodyPr"))
    etree.SubElement(t_elem, _qn("a", "lstStyle"))
    p = etree.SubElement(t_elem, _qn("a", "p"))
    etree.SubElement(p, _qn("a", "endParaRPr"), attrib={"lang": "en-US"})

    # Create data points with parTrans and sibTrans
    for i, item_text in enumerate(items):
        item_guid = _new_guid()
        par_trans_guid = _new_guid()
        sib_trans_guid = _new_guid()
        cxn_guid = _new_guid()

        # Data point
        pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={"modelId": item_guid})
        etree.SubElement(pt, _qn("dgm", "prSet"), attrib={"phldrT": "[Text]"})
        etree.SubElement(pt, _qn("dgm", "spPr"))
        t_elem = etree.SubElement(pt, _qn("dgm", "t"))
        etree.SubElement(t_elem, _qn("a", "bodyPr"))
        etree.SubElement(t_elem, _qn("a", "lstStyle"))
        p = etree.SubElement(t_elem, _qn("a", "p"))
        r = etree.SubElement(p, _qn("a", "r"))
        etree.SubElement(r, _qn("a", "rPr"), attrib={"lang": "en-US"})
        t = etree.SubElement(r, _qn("a", "t"))
        t.text = item_text

        # Parent transition point
        par_pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={
            "modelId": par_trans_guid,
            "type": "parTrans",
            "cxnId": cxn_guid,
        })
        etree.SubElement(par_pt, _qn("dgm", "prSet"))
        etree.SubElement(par_pt, _qn("dgm", "spPr"))
        par_t = etree.SubElement(par_pt, _qn("dgm", "t"))
        etree.SubElement(par_t, _qn("a", "bodyPr"))
        etree.SubElement(par_t, _qn("a", "lstStyle"))
        par_p = etree.SubElement(par_t, _qn("a", "p"))
        etree.SubElement(par_p, _qn("a", "endParaRPr"), attrib={"lang": "en-US"})

        # Sibling transition point
        sib_pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={
            "modelId": sib_trans_guid,
            "type": "sibTrans",
            "cxnId": cxn_guid,
        })
        etree.SubElement(sib_pt, _qn("dgm", "prSet"))
        etree.SubElement(sib_pt, _qn("dgm", "spPr"))
        sib_t = etree.SubElement(sib_pt, _qn("dgm", "t"))
        etree.SubElement(sib_t, _qn("a", "bodyPr"))
        etree.SubElement(sib_t, _qn("a", "lstStyle"))
        sib_p = etree.SubElement(sib_t, _qn("a", "p"))
        etree.SubElement(sib_p, _qn("a", "endParaRPr"), attrib={"lang": "en-US"})

        # Connection: doc -> item (parOf)
        etree.SubElement(cxn_lst, _qn("dgm", "cxn"), attrib={
            "modelId": cxn_guid,
            "type": "parOf",
            "srcId": doc_guid,
            "destId": item_guid,
            "srcOrd": str(i),
            "destOrd": "0",
            "parTransId": par_trans_guid,
            "sibTransId": sib_trans_guid,
        })

    # Background element (required by Word)
    bg = etree.SubElement(root, _qn("dgm", "bg"))
    whole = etree.SubElement(root, _qn("dgm", "whole"))

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _build_data_model_hierarchy(template_data_xml: bytes, hierarchy: dict,
                                layout_uri: str) -> bytes:
    """
    Build a data model for hierarchy diagrams from a nested dict.

    hierarchy format: {"Root": {"Child1": {"GC1": {}, "GC2": {}}, "Child2": {}}}
    """
    root = etree.Element(_qn("dgm", "dataModel"), nsmap={
        "dgm": NS["dgm"],
        "a": NS["a"],
        "r": NS["r"],
    })

    pt_lst = etree.SubElement(root, _qn("dgm", "ptLst"))
    cxn_lst = etree.SubElement(root, _qn("dgm", "cxnLst"))

    # Parse template doc point
    tpl_root = etree.fromstring(template_data_xml)
    tpl_doc_pt = tpl_root.find(".//dgm:pt[@type='doc']", NS)
    tpl_pr_set = tpl_doc_pt.find("dgm:prSet", NS) if tpl_doc_pt is not None else None

    doc_guid = _new_guid()
    doc_pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={
        "modelId": doc_guid,
        "type": "doc",
    })
    if tpl_pr_set is not None:
        doc_pt.append(copy.deepcopy(tpl_pr_set))
    else:
        etree.SubElement(doc_pt, _qn("dgm", "prSet"), attrib={
            "loTypeId": layout_uri, "loCatId": "hierarchy",
            "qsTypeId": "urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1",
            "qsCatId": "simple",
            "csTypeId": "urn:microsoft.com/office/officeart/2005/8/colors/accent1_2",
            "csCatId": "accent1", "phldr": "1",
        })
    etree.SubElement(doc_pt, _qn("dgm", "spPr"))
    t_elem = etree.SubElement(doc_pt, _qn("dgm", "t"))
    etree.SubElement(t_elem, _qn("a", "bodyPr"))
    etree.SubElement(t_elem, _qn("a", "lstStyle"))
    p = etree.SubElement(t_elem, _qn("a", "p"))
    etree.SubElement(p, _qn("a", "endParaRPr"), attrib={"lang": "en-US"})

    ord_counter = [0]

    def add_hierarchy_nodes(parent_guid, children_dict, depth=0):
        for i, (label, grandchildren) in enumerate(children_dict.items()):
            item_guid = _new_guid()
            par_trans_guid = _new_guid()
            sib_trans_guid = _new_guid()
            cxn_guid = _new_guid()

            # Data point
            pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={"modelId": item_guid})
            etree.SubElement(pt, _qn("dgm", "prSet"), attrib={"phldrT": "[Text]"})
            etree.SubElement(pt, _qn("dgm", "spPr"))
            t_elem = etree.SubElement(pt, _qn("dgm", "t"))
            etree.SubElement(t_elem, _qn("a", "bodyPr"))
            etree.SubElement(t_elem, _qn("a", "lstStyle"))
            p = etree.SubElement(t_elem, _qn("a", "p"))
            r = etree.SubElement(p, _qn("a", "r"))
            etree.SubElement(r, _qn("a", "rPr"), attrib={"lang": "en-US"})
            t = etree.SubElement(r, _qn("a", "t"))
            t.text = label

            # Transition points
            par_pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={
                "modelId": par_trans_guid, "type": "parTrans", "cxnId": cxn_guid,
            })
            etree.SubElement(par_pt, _qn("dgm", "prSet"))
            etree.SubElement(par_pt, _qn("dgm", "spPr"))
            par_t = etree.SubElement(par_pt, _qn("dgm", "t"))
            etree.SubElement(par_t, _qn("a", "bodyPr"))
            etree.SubElement(par_t, _qn("a", "lstStyle"))
            par_p = etree.SubElement(par_t, _qn("a", "p"))
            etree.SubElement(par_p, _qn("a", "endParaRPr"), attrib={"lang": "en-US"})

            sib_pt = etree.SubElement(pt_lst, _qn("dgm", "pt"), attrib={
                "modelId": sib_trans_guid, "type": "sibTrans", "cxnId": cxn_guid,
            })
            etree.SubElement(sib_pt, _qn("dgm", "prSet"))
            etree.SubElement(sib_pt, _qn("dgm", "spPr"))
            sib_t = etree.SubElement(sib_pt, _qn("dgm", "t"))
            etree.SubElement(sib_t, _qn("a", "bodyPr"))
            etree.SubElement(sib_t, _qn("a", "lstStyle"))
            sib_p = etree.SubElement(sib_t, _qn("a", "p"))
            etree.SubElement(sib_p, _qn("a", "endParaRPr"), attrib={"lang": "en-US"})

            # Connection
            etree.SubElement(cxn_lst, _qn("dgm", "cxn"), attrib={
                "modelId": cxn_guid, "type": "parOf",
                "srcId": parent_guid, "destId": item_guid,
                "srcOrd": str(i), "destOrd": "0",
                "parTransId": par_trans_guid, "sibTransId": sib_trans_guid,
            })

            if grandchildren:
                add_hierarchy_nodes(item_guid, grandchildren, depth + 1)

    add_hierarchy_nodes(doc_guid, hierarchy)

    etree.SubElement(root, _qn("dgm", "bg"))
    etree.SubElement(root, _qn("dgm", "whole"))

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ─── Minimal Drawing Part ──────────────────────────────────────────────────

def _build_empty_drawing() -> bytes:
    """Build a minimal drawing part. Word regenerates this on open."""
    root = etree.Element(_qn("dsp", "drawing"), nsmap={
        "dsp": NS["dsp"],
    })
    sp_tree = etree.SubElement(root, _qn("dsp", "spTree"))
    nvGrpSpPr = etree.SubElement(sp_tree, _qn("dsp", "nvGrpSpPr"))
    etree.SubElement(nvGrpSpPr, _qn("dsp", "cNvPr"), attrib={"id": "0", "name": ""})
    etree.SubElement(nvGrpSpPr, _qn("dsp", "cNvGrpSpPr"))
    etree.SubElement(sp_tree, _qn("dsp", "grpSpPr"))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


# ─── Part Injection ─────────────────────────────────────────────────────────

def _inject_smartart(doc: Document, template_name: str, data_xml: bytes,
                     width_emu: int = 5486400, height_emu: int = 3200400):
    """
    Inject SmartArt into a document by copying template parts and using custom data.

    Args:
        doc: python-docx Document
        template_name: Name of template (e.g. 'basic_list')
        data_xml: Custom diagram data XML
        width_emu: Width in EMUs
        height_emu: Height in EMUs
    """
    idx = _next_diagram_id(doc)
    parts = _extract_template_parts(template_name)
    rel_types = _extract_rels_from_template(template_name)

    # Part names for this diagram
    data_pn = PackURI("/word/diagrams/data{}.xml".format(idx))
    layout_pn = PackURI("/word/diagrams/layout{}.xml".format(idx))
    style_pn = PackURI("/word/diagrams/quickStyle{}.xml".format(idx))
    colors_pn = PackURI("/word/diagrams/colors{}.xml".format(idx))
    drawing_pn = PackURI("/word/diagrams/drawing{}.xml".format(idx))

    # Use custom data, template layout/style/colors, and empty drawing
    # (Word regenerates the drawing from data + layout on open)
    data_part = Part(data_pn, CT_DGM_DATA, data_xml, doc.part.package)
    layout_part = Part(layout_pn, CT_DGM_LAYOUT, parts['layout'], doc.part.package)
    style_part = Part(style_pn, CT_DGM_STYLE, parts['style'], doc.part.package)
    colors_part = Part(colors_pn, CT_DGM_COLORS, parts['colors'], doc.part.package)
    drawing_part = Part(drawing_pn, CT_DGM_DRAWING, _build_empty_drawing(), doc.part.package)

    # Add relationships using the exact types from the template
    r_data = doc.part.relate_to(data_part, rel_types.get('data', RT_DGM_DATA))
    r_layout = doc.part.relate_to(layout_part, rel_types.get('layout', RT_DGM_LAYOUT))
    r_style = doc.part.relate_to(style_part, rel_types.get('style', RT_DGM_STYLE))
    r_colors = doc.part.relate_to(colors_part, rel_types.get('colors', RT_DGM_COLORS))
    r_drawing = doc.part.relate_to(drawing_part, rel_types.get('drawing', RT_DGM_DRAWING))

    # Build inline drawing paragraph
    para_xml = (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        ' xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        '<w:r><w:drawing>'
        '<wp:inline distT="0" distB="0" distL="0" distR="0">'
        '<wp:extent cx="{w}" cy="{h}"/>'
        '<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        '<wp:docPr id="{did}" name="Diagram {idx}"/>'
        '<wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">'
        '<dgm:relIds r:dm="{dm}" r:lo="{lo}" r:qs="{qs}" r:cs="{cs}"/>'
        '</a:graphicData></a:graphic>'
        '</wp:inline>'
        '</w:drawing></w:r></w:p>'
    ).format(
        w=width_emu, h=height_emu, did=idx + 100, idx=idx,
        dm=r_data, lo=r_layout, qs=r_style, cs=r_colors,
    )

    para_element = etree.fromstring(para_xml.encode("utf-8"))
    doc.element.body.append(para_element)


def _add_smartart(doc: Document, template_name: str, items: list[str],
                  title: Optional[str] = None, hierarchy: Optional[dict] = None,
                  width_emu: int = 5486400, height_emu: int = 3200400):
    """
    High-level function to add SmartArt to a document.
    """
    if title:
        doc.add_heading(title, level=2)

    parts = _extract_template_parts(template_name)
    layout_uri = LAYOUT_URIS.get(template_name, "")

    if hierarchy is not None:
        data_xml = _build_data_model_hierarchy(parts['data'], hierarchy, layout_uri)
    else:
        data_xml = _build_data_model_flat(parts['data'], items, layout_uri)

    _inject_smartart(doc, template_name, data_xml, width_emu, height_emu)


# ─── Public API ─────────────────────────────────────────────────────────────

class SmartArt:
    """
    Static methods to add SmartArt diagrams to Word documents.

    Each method creates a native, editable SmartArt object - not an image.
    The diagrams will render and be editable in Microsoft Word.

    Templates must be pre-generated (run generate_templates.py once with Word).
    """

    @staticmethod
    def add_basic_list(doc: Document, title: str, items: list[str],
                       width_emu: int = 5486400, height_emu: int = 3200400):
        """
        Add a Basic Block List SmartArt diagram.

        Best for: Ungrouped or non-sequential information, simple lists of items.

        Args:
            doc: python-docx Document object
            title: Heading text above the diagram
            items: List of strings, one per block
            width_emu: Diagram width in EMUs (default ~5.7")
            height_emu: Diagram height in EMUs (default ~3.3")

        Example:
            SmartArt.add_basic_list(doc, "Key Features", [
                "Fast Performance",
                "Easy to Use",
                "Secure by Default",
                "Open Source",
            ])
        """
        _add_smartart(doc, "basic_list", items, title=title,
                      width_emu=width_emu, height_emu=height_emu)

    @staticmethod
    def add_basic_process(doc: Document, title: str, items: list[str],
                          width_emu: int = 5486400, height_emu: int = 3200400):
        """
        Add a Basic Process (linear flow) SmartArt diagram.

        Best for: Sequential steps, workflows, timelines, procedures.

        Args:
            doc: python-docx Document object
            title: Heading text above the diagram
            items: List of strings representing sequential steps
            width_emu: Diagram width in EMUs (default ~5.7")
            height_emu: Diagram height in EMUs (default ~3.3")

        Example:
            SmartArt.add_basic_process(doc, "Development Lifecycle", [
                "Requirements",
                "Design",
                "Implementation",
                "Testing",
                "Deployment",
            ])
        """
        _add_smartart(doc, "basic_process", items, title=title,
                      width_emu=width_emu, height_emu=height_emu)

    @staticmethod
    def add_hierarchy(doc: Document, title: str, hierarchy: dict,
                      width_emu: int = 5486400, height_emu: int = 4000000):
        """
        Add a Hierarchy (org chart / tree) SmartArt diagram.

        Best for: Organizational charts, tree structures, classification.

        Args:
            doc: python-docx Document object
            title: Heading text above the diagram
            hierarchy: Nested dict where keys are labels and values are child dicts.
            width_emu: Diagram width in EMUs (default ~5.7")
            height_emu: Diagram height in EMUs (default ~4.2")

        Example:
            SmartArt.add_hierarchy(doc, "Organization", {
                "CEO": {
                    "VP Engineering": {
                        "Frontend Lead": {},
                        "Backend Lead": {},
                    },
                    "VP Marketing": {
                        "Brand Manager": {},
                    },
                }
            })
        """
        _add_smartart(doc, "hierarchy", [], title=title, hierarchy=hierarchy,
                      width_emu=width_emu, height_emu=height_emu)

    @staticmethod
    def add_cycle(doc: Document, title: str, items: list[str],
                  width_emu: int = 5486400, height_emu: int = 4000000):
        """
        Add a Cycle SmartArt diagram.

        Best for: Repeating processes, continuous workflows, feedback loops.

        Args:
            doc: python-docx Document object
            title: Heading text above the diagram
            items: List of strings representing stages in the cycle
            width_emu: Diagram width in EMUs (default ~5.7")
            height_emu: Diagram height in EMUs (default ~4.2")

        Example:
            SmartArt.add_cycle(doc, "Agile Sprint Cycle", [
                "Plan",
                "Design",
                "Develop",
                "Test",
                "Review",
            ])
        """
        _add_smartart(doc, "cycle", items, title=title,
                      width_emu=width_emu, height_emu=height_emu)

    @staticmethod
    def add_pyramid(doc: Document, title: str, items: list[str],
                    width_emu: int = 5486400, height_emu: int = 4000000):
        """
        Add a Pyramid SmartArt diagram.

        Best for: Proportional or hierarchical relationships, foundation concepts,
        priority levels (top = most important/smallest, bottom = foundation/largest).

        Args:
            doc: python-docx Document object
            title: Heading text above the diagram
            items: List of strings from top (apex) to bottom (base)
            width_emu: Diagram width in EMUs (default ~5.7")
            height_emu: Diagram height in EMUs (default ~4.2")

        Example:
            SmartArt.add_pyramid(doc, "Maslow's Hierarchy", [
                "Self-Actualization",
                "Esteem",
                "Love/Belonging",
                "Safety",
                "Physiological",
            ])
        """
        _add_smartart(doc, "pyramid", items, title=title,
                      width_emu=width_emu, height_emu=height_emu)

    @staticmethod
    def add_radial(doc: Document, title: str, center: str, items: list[str],
                   width_emu: int = 5486400, height_emu: int = 4000000):
        """
        Add a Radial SmartArt diagram with a central item and surrounding items.

        Best for: Relationships to a central concept, components of a system,
        hub-and-spoke architectures.

        Args:
            doc: python-docx Document object
            title: Heading text above the diagram
            center: Text for the central element
            items: List of strings for the surrounding elements
            width_emu: Diagram width in EMUs (default ~5.7")
            height_emu: Diagram height in EMUs (default ~4.2")

        Example:
            SmartArt.add_radial(doc, "Microservices", "API Gateway", [
                "Auth Service",
                "User Service",
                "Payment Service",
                "Notification Service",
            ])
        """
        all_items = [center] + items
        _add_smartart(doc, "radial", all_items, title=title,
                      width_emu=width_emu, height_emu=height_emu)
