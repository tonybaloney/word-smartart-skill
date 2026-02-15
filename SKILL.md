---
name: word-smartart
description: Create native, editable SmartArt diagrams in Microsoft Word documents using Python. Use this skill when asked to create diagrams, flowcharts, org charts, process flows, or any visual diagram in a Word (.docx) document. Do NOT use images or mermaid — this creates real SmartArt objects.
license: MIT
---

# SmartArt in Word Documents

## When to Use This Skill

Use this when the user asks you to:
- Create diagrams in Word documents
- Add flowcharts, org charts, process diagrams to .docx files
- Generate visual representations of data in Word
- Create architecture diagrams, lifecycle diagrams, or any structured visual

**Do NOT** use mermaid-to-PNG or any image-based approach. This creates native SmartArt that is editable in Word.

## Prerequisites

```bash
pip install python-docx lxml
```

The `templates/` directory must contain pre-generated SmartArt template files (`.docx`). These are created once using `generate_templates.py` (requires Microsoft Word installed). The templates are committed to the repository and do not need to be regenerated.

## Quick Start

```python
from docx import Document
from smartart import SmartArt

doc = Document()
doc.add_heading("Project Overview", 0)

# Add a process flow
SmartArt.add_basic_process(doc, "Development Lifecycle", [
    "Requirements",
    "Design",
    "Implementation",
    "Testing",
    "Deployment",
])

# Add an org chart
SmartArt.add_hierarchy(doc, "Team Structure", {
    "CTO": {
        "Engineering Manager": {
            "Frontend Lead": {},
            "Backend Lead": {},
        },
        "QA Manager": {
            "Test Lead": {},
        },
    }
})

SmartArt.save(doc, "project_overview.docx")
```

## Available Diagram Types

### 1. Basic Block List — `SmartArt.add_basic_list()`

**Best for:** Ungrouped information, feature lists, simple collections of items.

```python
SmartArt.add_basic_list(doc, "Key Features", [
    "Fast Performance",
    "Easy to Use",
    "Secure by Default",
    "Open Source",
])
```

**Renders as:** A grid of colored blocks, each containing one item.

### 2. Basic Process — `SmartArt.add_basic_process()`

**Best for:** Sequential steps, workflows, timelines, pipelines, procedures.

```python
SmartArt.add_basic_process(doc, "CI/CD Pipeline", [
    "Commit",
    "Build",
    "Test",
    "Deploy",
    "Monitor",
])
```

**Renders as:** Horizontal boxes connected by arrows showing flow direction.

### 3. Hierarchy — `SmartArt.add_hierarchy()`

**Best for:** Org charts, tree structures, classification, file system structures, inheritance.

```python
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
```

**Input format:** Nested Python dict. Keys are labels, values are dicts of children. Empty `{}` means leaf node.

**Renders as:** Top-down tree with connecting lines.

### 4. Cycle — `SmartArt.add_cycle()`

**Best for:** Repeating processes, feedback loops, iterative workflows, continuous processes.

```python
SmartArt.add_cycle(doc, "Agile Sprint", [
    "Plan",
    "Design",
    "Develop",
    "Test",
    "Review",
    "Retrospective",
])
```

**Renders as:** Circular arrangement of items with arrows suggesting continuous flow.

### 5. Pyramid — `SmartArt.add_pyramid()`

**Best for:** Proportional relationships, priority levels, layered architectures, hierarchical concepts.

```python
SmartArt.add_pyramid(doc, "Testing Pyramid", [
    "E2E Tests",        # Top (smallest/fewest)
    "Integration Tests", # Middle
    "Unit Tests",        # Bottom (largest/most)
])
```

**Item order:** Top of pyramid (apex) to bottom (base). First item = top.

**Renders as:** Pyramid shape with horizontal layers.

### 6. Radial — `SmartArt.add_radial()`

**Best for:** Hub-and-spoke relationships, central concept with related items, system components around a core.

```python
SmartArt.add_radial(doc, "Microservices", "API Gateway", [
    "Auth Service",
    "User Service",
    "Payment Service",
    "Notification Service",
])
```

**Note:** First argument after title is the **center** element, second is a list of **surrounding** elements.

**Renders as:** Central circle with surrounding circles connected by lines.

## Choosing the Right Diagram Type

| User's Intent | Diagram Type | Method |
|---|---|---|
| Show steps in order | Process | `add_basic_process()` |
| List items (no order) | List | `add_basic_list()` |
| Show org chart / tree | Hierarchy | `add_hierarchy()` |
| Show repeating process | Cycle | `add_cycle()` |
| Show priority / layers | Pyramid | `add_pyramid()` |
| Show central concept | Radial | `add_radial()` |
| Show architecture | Radial or Hierarchy | depends on shape |
| Show dependencies | Hierarchy | `add_hierarchy()` |
| Show workflow | Process or Cycle | depends on if it repeats |

## API Reference

All methods share these parameters:

| Parameter | Type | Default | Description |
|---|---|---|---|
| `doc` | `Document` | required | python-docx Document object |
| `title` | `str` | required | Heading text above the diagram |
| `items` | `list[str]` | required | List of text labels (flat diagrams) |
| `hierarchy` | `dict` | required | Nested dict (hierarchy only) |
| `center` | `str` | required | Central element text (radial only) |
| `width_emu` | `int` | `5486400` | Width in EMUs (~5.7 inches) |
| `height_emu` | `int` | `3200400` | Height in EMUs (~3.3 inches) |

**EMU reference:** 1 inch = 914400 EMUs. Default width is 6 inches. To make a diagram wider: `width_emu=7315200` (8 inches).

## Multiple Diagrams in One Document

You can add multiple diagrams to the same document. Each gets its own unique XML parts:

```python
doc = Document()
doc.add_heading("Architecture Document", 0)

SmartArt.add_basic_process(doc, "Data Flow", [
    "Ingest", "Transform", "Store", "Query", "Visualize",
])

doc.add_paragraph("The system is organized as follows:")

SmartArt.add_hierarchy(doc, "System Components", {
    "Platform": {
        "Data Layer": {"PostgreSQL": {}, "Redis": {}},
        "API Layer": {"REST API": {}, "GraphQL": {}},
        "UI Layer": {"Web App": {}, "Mobile App": {}},
    }
})

SmartArt.add_radial(doc, "Core Services", "Event Bus", [
    "Auth", "Users", "Payments", "Notifications",
])

SmartArt.save(doc, "architecture.docx")
```

## How It Works (Technical Details)

### Architecture

The library uses a **template-based + ZIP post-processing approach**:

1. **Templates** (`templates/` directory) are Word documents containing one SmartArt diagram each, generated by Microsoft Word via COM automation. These contain the complete, valid OpenXML parts that Word expects.

2. **At runtime**, the library:
   - Adds placeholder paragraphs with marker rIds (e.g., `__SMARTART_1_DM__`) into the document body
   - Extracts the layout, style, colors, and drawing XML parts from the template
   - Modifies the template's data model XML in-place with the user's content (preserving presentation points)
   - Stores diagram specs on the document object

3. **At save time** (`SmartArt.save()`):
   - Saves the document normally with python-docx
   - Post-processes the `.docx` ZIP to inject diagram XML parts directly
   - Replaces placeholder rIds with real relationship IDs
   - Adds relationships and content types to the ZIP

### Critical Implementation Notes

**You MUST use `SmartArt.save(doc, path)` instead of `doc.save(path)`.** The SmartArt parts are injected via ZIP-level post-processing because python-docx's Part API does not correctly package SmartArt parts — Word silently renders them as zero-height rectangles even though the ZIP structure appears identical to a working file.

**Diagram placement:** SmartArt paragraphs must be inserted before `<w:sectPr>` in the document body. Using `body.append()` places them after section properties, causing all diagrams to appear at the end of the document instead of inline.

**Compatibility mode:** Documents must use `compatibilityMode=15` (Word 2013+). python-docx defaults to 14 (Word 2010), which prevents Word from regenerating SmartArt rendering. The library handles this automatically.

**Presentation points are required:** The data model must contain `type="pres"` points and `type="presOf"`/`type="presParOf"` connections in addition to data points. Without these, Word renders SmartArt as collapsed zero-height rectangles. This is not documented in the ECMA-376 spec. The library preserves these from templates.

### OpenXML Structure

A SmartArt diagram in a `.docx` file consists of 5 XML parts:

| Part | Content Type | Purpose |
|---|---|---|
| `data{N}.xml` | `drawingml.diagramData` | The data model: nodes, text, connections |
| `layout{N}.xml` | `drawingml.diagramLayout` | Layout algorithm (how to arrange shapes) |
| `quickStyle{N}.xml` | `drawingml.diagramStyle` | Visual style (3D, flat, etc.) |
| `colors{N}.xml` | `drawingml.diagramColors` | Color scheme |
| `drawing{N}.xml` | `drawingml.diagramDrawing` | Cached rendering (Word regenerates this) |

These are stored in `word/diagrams/` inside the `.docx` ZIP archive.

### Data Model Structure

The data model (`data{N}.xml`) contains:

```xml
<dgm:dataModel>
  <dgm:ptLst>
    <!-- Root document point -->
    <dgm:pt modelId="{GUID}" type="doc">
      <dgm:prSet loTypeId="urn:microsoft.com/.../layout/default" .../>
    </dgm:pt>
    
    <!-- Data node -->
    <dgm:pt modelId="{GUID}">
      <dgm:t><a:p><a:r><a:t>Item Text</a:t></a:r></a:p></dgm:t>
    </dgm:pt>
    
    <!-- Parent transition (required per node) -->
    <dgm:pt modelId="{GUID}" type="parTrans" cxnId="{CXN_GUID}"/>
    
    <!-- Sibling transition (required per node) -->
    <dgm:pt modelId="{GUID}" type="sibTrans" cxnId="{CXN_GUID}"/>
  </dgm:ptLst>
  
  <dgm:cxnLst>
    <!-- Connection: parent -> child -->
    <dgm:cxn modelId="{GUID}" type="parOf" 
             srcId="{PARENT}" destId="{CHILD}"
             srcOrd="0" destOrd="0"
             parTransId="{PAR_TRANS}" sibTransId="{SIB_TRANS}"/>
  </dgm:cxnLst>
</dgm:dataModel>
```

**Key rules:**
- Every data node needs a `parTrans` point and a `sibTrans` point
- Every connection needs `srcOrd`/`destOrd` ordering attributes
- The `parTransId`/`sibTransId` on connections must reference the transition point GUIDs
- The root `doc` point's `prSet` must reference the layout URI

### Relationship Types

```
diagramData:       http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData
diagramLayout:     http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout  
diagramQuickStyle: http://schemas.microsoft.com/office/2007/relationships/diagramStyle
diagramColors:     http://schemas.microsoft.com/office/2007/relationships/diagramColors
diagramDrawing:    http://schemas.microsoft.com/office/2007/relationships/diagramDrawing
```

### Document Body Reference

The diagram is embedded via an inline drawing in a paragraph:

```xml
<w:drawing>
  <wp:inline>
    <wp:extent cx="5486400" cy="3200400"/>
    <wp:docPr id="101" name="Diagram 1"/>
    <wp:cNvGraphicFramePr/>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
        <dgm:relIds r:dm="rId4" r:lo="rId5" r:qs="rId6" r:cs="rId7"/>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

## Combining with Markdown-to-Word

When generating complete documents from Markdown content that also need diagrams, use this pattern:

1. Create the document with python-docx (adding text, headings, etc.)
2. At each point where a diagram is needed, call the appropriate SmartArt method
3. Use `SmartArt.save()` instead of `doc.save()` at the end

```python
from docx import Document
from smartart import SmartArt

doc = Document()
doc.add_heading("Architecture Overview", 0)

doc.add_paragraph(
    "This document explains our system architecture."
)

# Insert a diagram between text sections
SmartArt.add_hierarchy(doc, "System Components", {
    "Platform": {
        "Frontend": {"React App": {}, "Mobile App": {}},
        "Backend": {"REST API": {}, "Worker Service": {}},
        "Data": {"PostgreSQL": {}, "Redis Cache": {}},
    }
})

doc.add_paragraph(
    "The deployment follows a standard CI/CD pipeline."
)

SmartArt.add_basic_process(doc, "Deployment Pipeline", [
    "Commit", "Build", "Test", "Stage", "Deploy",
])

# IMPORTANT: Always use SmartArt.save() when the doc contains diagrams
SmartArt.save(doc, "architecture.docx")
```

**Key rules for combining text and diagrams:**
- Use `SmartArt.save(doc, path)` instead of `doc.save(path)` — this performs ZIP-level post-processing required for SmartArt rendering
- Diagrams can be placed anywhere in the document flow (between paragraphs, after headings, etc.)
- Multiple diagrams of different types can coexist in one document
- If the document has NO SmartArt, `SmartArt.save()` still works (it just calls `doc.save()` internally)

## Regenerating Templates

If you need to regenerate the template files (e.g., after changing supported layouts):

```bash
# Requires Microsoft Word installed
pip install pywin32
python generate_templates.py
```

This creates one `.docx` template per SmartArt type in the `templates/` directory using Word COM automation.

## File Structure

```
├── smartart.py              # Main library - import this
├── generate_templates.py    # One-time template generator (needs Word)
├── generate_test_docs.py    # Test document generator
├── templates/               # Pre-generated SmartArt templates
│   ├── basic_list.docx
│   ├── basic_process.docx
│   ├── hierarchy.docx
│   ├── cycle.docx
│   ├── pyramid.docx
│   └── radial.docx
└── tests/                   # Generated test documents
    ├── test_basic_list.docx
    ├── test_basic_process.docx
    ├── test_hierarchy.docx
    ├── test_cycle.docx
    ├── test_pyramid.docx
    ├── test_radial.docx
    └── test_combined_all_types.docx
```
