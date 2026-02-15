# Word SmartArt for AI Coding Agents

Create native, editable SmartArt diagrams in Microsoft Word documents using Python. No images — real SmartArt objects that render and are fully editable in Word.

## Supported Diagram Types

| Type | Method | Use Case |
|------|--------|----------|
| Basic Block List | `add_basic_list()` | Feature lists, simple collections |
| Basic Process | `add_basic_process()` | Workflows, pipelines, sequential steps |
| Hierarchy | `add_hierarchy()` | Org charts, tree structures |
| Cycle | `add_cycle()` | Iterative processes, feedback loops |
| Pyramid | `add_pyramid()` | Priority levels, layered architectures |
| Radial | `add_radial()` | Hub-and-spoke, central concepts |

## Quick Start

```bash
pip install python-docx lxml
```

```python
from docx import Document
from smartart import SmartArt

doc = Document()
doc.add_heading("My Document", 0)

SmartArt.add_basic_process(doc, "Build Pipeline", [
    "Code", "Build", "Test", "Deploy",
])

SmartArt.add_hierarchy(doc, "Team", {
    "CTO": {
        "Eng Manager": {"Dev Lead": {}, "QA Lead": {}},
        "Product Manager": {},
    }
})

doc.save("output.docx")
```

## Installing This Skill into GitHub Copilot

This is a [GitHub Copilot Agent Skill](https://docs.github.com/en/copilot/concepts/agents/about-agent-skills). It works with Copilot coding agent, GitHub Copilot CLI, and agent mode in VS Code.

### Option 1: As a Project Skill (per repository)

Copy this skill into your repository's `.github/skills/` directory:

```bash
# From your project root
git submodule add https://github.com/tonybaloney/word-smartart-skill .github/skills/word-smartart
```

Or copy the files manually:

```bash
mkdir -p .github/skills/word-smartart
cp path/to/mermaid-to-word/SKILL.md .github/skills/word-smartart/
cp path/to/mermaid-to-word/smartart.py .github/skills/word-smartart/
cp -r path/to/mermaid-to-word/templates .github/skills/word-smartart/
```

Copilot will automatically discover the skill and use it when you ask for diagrams in Word documents.

### Option 2: As a Personal Skill (all projects)

Install once for all your repositories:

```bash
# macOS / Linux
cp -r path/to/mermaid-to-word ~/.copilot/skills/word-smartart

# Windows
xcopy path\to\mermaid-to-word %USERPROFILE%\.copilot\skills\word-smartart\ /E /I
```

### After Installing

Install the Python dependencies in your project environment:

```bash
pip install python-docx lxml
```

Then just ask Copilot to create a diagram in a Word document — it will find and use the skill automatically based on the description in `SKILL.md`.

### What Gets Installed

| File | Purpose |
|------|---------|
| `SKILL.md` | Frontmatter + instructions that teach Copilot when and how to use this skill |
| `smartart.py` | Python library the agent calls to create SmartArt |
| `templates/` | Pre-generated SmartArt XML templates (6 files, ~100KB total) |

## For AI Agents

See **[SKILL.md](SKILL.md)** for the complete skill document with:
- All diagram types with examples
- Decision guide for choosing diagram types
- Full API reference
- OpenXML technical details for understanding the internals

## How It Works

Uses pre-generated Word templates (in `templates/`) that contain valid SmartArt XML parts. At runtime, the library extracts the layout/style/color definitions from templates and injects a custom data model with your content. Word regenerates the visual rendering on open.

## Regenerating Templates

Only needed if modifying supported layouts:

```bash
pip install pywin32  # Requires Microsoft Word installed
python generate_templates.py
```

## License

MIT
