"""
Generate SmartArt template .docx files using Word COM automation.

This script creates one template document per SmartArt layout type, each with 
placeholder data nodes. These templates are used by smartart.py to create 
SmartArt in documents without needing Word COM at runtime.

Requires: Microsoft Word installed, pywin32
"""

import win32com.client
import os
import time

TEMPLATE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")

# Layout configurations: (name, layout_search_term, layout_id, node_count, node_labels)
LAYOUTS = {
    "basic_list": {
        "id": "urn:microsoft.com/office/officeart/2005/8/layout/default",
        "nodes": ["Item 1", "Item 2", "Item 3", "Item 4", "Item 5", "Item 6", "Item 7", "Item 8"],
    },
    "basic_process": {
        "id": "urn:microsoft.com/office/officeart/2005/8/layout/process1",
        "nodes": ["Step 1", "Step 2", "Step 3", "Step 4", "Step 5", "Step 6", "Step 7", "Step 8"],
    },
    "hierarchy": {
        "id": "urn:microsoft.com/office/officeart/2005/8/layout/hierarchy1",
        "hierarchy": True,
        # Will create: Root -> [Child1 -> [GC1, GC2], Child2 -> [GC3, GC4], Child3 -> [GC5]]
    },
    "cycle": {
        "id": "urn:microsoft.com/office/officeart/2005/8/layout/cycle1",
        "nodes": ["Stage 1", "Stage 2", "Stage 3", "Stage 4", "Stage 5", "Stage 6", "Stage 7", "Stage 8"],
    },
    "pyramid": {
        "id": "urn:microsoft.com/office/officeart/2005/8/layout/pyramid1",
        "nodes": ["Level 1", "Level 2", "Level 3", "Level 4", "Level 5", "Level 6"],
    },
    "radial": {
        "id": "urn:microsoft.com/office/officeart/2005/8/layout/radial1",
        "nodes": ["Center", "Spoke 1", "Spoke 2", "Spoke 3", "Spoke 4", "Spoke 5", "Spoke 6"],
    },
}


def find_layout(word, layout_id):
    """Find a SmartArt layout by its URI."""
    layouts = word.Application.SmartArtLayouts
    for i in range(1, layouts.Count + 1):
        layout = layouts.Item(i)
        if layout.Id == layout_id:
            return layout
    raise ValueError("Layout not found: {}".format(layout_id))


def create_flat_template(word, layout_id, nodes, output_path):
    """Create a template with a flat list of nodes."""
    word.Documents.Add()
    doc = word.ActiveDocument
    layout = find_layout(word, layout_id)
    rng = doc.Content
    shape = doc.InlineShapes.AddSmartArt(layout, rng)
    sa = shape.SmartArt

    all_nodes = sa.AllNodes
    # Remove excess default nodes
    while all_nodes.Count > len(nodes):
        all_nodes.Item(all_nodes.Count).Delete()
    # Set text for existing nodes
    for i in range(1, min(all_nodes.Count + 1, len(nodes) + 1)):
        all_nodes.Item(i).TextFrame2.TextRange.Text = nodes[i - 1]
    # Add missing nodes
    for i in range(all_nodes.Count, len(nodes)):
        new_node = all_nodes.Add()
        new_node.TextFrame2.TextRange.Text = nodes[i]

    doc.SaveAs2(os.path.abspath(output_path))
    doc.Close(False)
    print("  Created: {}".format(output_path))


def create_hierarchy_template(word, layout_id, output_path):
    """Create a hierarchy template with multiple levels."""
    msoSmartArtNodeBelow = 4
    msoSmartArtNodeAfter = 2
    msoSmartArtNodeDefault = 1

    word.Documents.Add()
    doc = word.ActiveDocument
    layout = find_layout(word, layout_id)
    rng = doc.Content
    shape = doc.InlineShapes.AddSmartArt(layout, rng)
    sa = shape.SmartArt
    nodes_obj = sa.AllNodes

    # The default hierarchy has some nodes already. Let's just use them and
    # add/remove as needed. The default typically has 5 nodes in a tree.
    # Let's set text on all existing nodes.
    labels = []
    for i in range(1, nodes_obj.Count + 1):
        labels.append(nodes_obj.Item(i).TextFrame2.TextRange.Text)

    # Set labels on existing nodes (Word creates a reasonable default hierarchy)
    target_labels = [
        "Root", "Child 1", "Grandchild 1", "Child 2", "Grandchild 2"
    ]
    for i, label in enumerate(target_labels):
        if i < nodes_obj.Count:
            nodes_obj.Item(i + 1).TextFrame2.TextRange.Text = label

    # Add additional nodes using the last child nodes
    # Add Grandchild 3 as child of Child 2
    child2 = nodes_obj.Item(4)  # "Child 2"
    gc3 = child2.AddNode(msoSmartArtNodeBelow, msoSmartArtNodeDefault)
    gc3.TextFrame2.TextRange.Text = "Grandchild 3"

    # Add Child 3 after Child 2 (from Root)
    root = nodes_obj.Item(1)
    child3 = root.AddNode(msoSmartArtNodeBelow, msoSmartArtNodeDefault)
    child3.TextFrame2.TextRange.Text = "Child 3"

    doc.SaveAs2(os.path.abspath(output_path))
    doc.Close(False)
    print("  Created: {}".format(output_path))


def main():
    os.makedirs(TEMPLATE_DIR, exist_ok=True)

    print("Generating SmartArt templates using Word COM automation...")
    print()

    word = win32com.client.DispatchEx('Word.Application')
    time.sleep(3)
    word.Visible = False

    try:
        for name, config in LAYOUTS.items():
            output_path = os.path.join(TEMPLATE_DIR, "{}.docx".format(name))
            print("  Generating: {}".format(name))
            if config.get("hierarchy"):
                create_hierarchy_template(word, config["id"], output_path)
            else:
                create_flat_template(word, config["id"], config["nodes"], output_path)
    finally:
        word.Quit()

    print()
    print("All templates generated in: {}".format(TEMPLATE_DIR))


if __name__ == "__main__":
    main()
