#!/usr/bin/env python3
"""
Tests for the diagram parser module.

Tests the parsing of Mermaid and PlantUML syntax into structured diagrams.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import unittest
from diagram_parser import (
    MermaidParser, PlantUMLParser, DiagramParser,
    NodeShape, Direction, EdgeType, EdgeStyle
)


class TestMermaidParser(unittest.TestCase):
    """Tests for the Mermaid parser."""
    
    def setUp(self):
        self.parser = MermaidParser()
    
    def test_parse_simple_flow(self):
        """Test parsing a simple top-down flow."""
        code = """
graph TD
    A[Start] --> B[End]
"""
        diagram = self.parser.parse(code)
        
        self.assertEqual(diagram.direction, Direction.TOP_DOWN)
        self.assertEqual(len(diagram.nodes), 2)
        self.assertEqual(len(diagram.edges), 1)
        
        # Check nodes
        node_ids = {n.id for n in diagram.nodes}
        self.assertIn('A', node_ids)
        self.assertIn('B', node_ids)
        
        # Check edge
        edge = diagram.edges[0]
        self.assertEqual(edge.source, 'A')
        self.assertEqual(edge.target, 'B')
    
    def test_parse_left_right_direction(self):
        """Test parsing left-to-right direction."""
        code = "graph LR\n    A --> B"
        diagram = self.parser.parse(code)
        self.assertEqual(diagram.direction, Direction.LEFT_RIGHT)
    
    def test_parse_right_left_direction(self):
        """Test parsing right-to-left direction."""
        code = "graph RL\n    A --> B"
        diagram = self.parser.parse(code)
        self.assertEqual(diagram.direction, Direction.RIGHT_LEFT)
    
    def test_parse_bottom_up_direction(self):
        """Test parsing bottom-up direction."""
        code = "graph BT\n    A --> B"
        diagram = self.parser.parse(code)
        self.assertEqual(diagram.direction, Direction.BOTTOM_UP)
    
    def test_parse_flowchart_keyword(self):
        """Test parsing with 'flowchart' keyword."""
        code = "flowchart TD\n    A --> B"
        diagram = self.parser.parse(code)
        self.assertEqual(diagram.direction, Direction.TOP_DOWN)
    
    def test_parse_rectangle_node(self):
        """Test parsing rectangle node shape."""
        code = "graph TD\n    A[Rectangle] --> B"
        diagram = self.parser.parse(code)
        
        node_a = next(n for n in diagram.nodes if n.id == 'A')
        self.assertEqual(node_a.shape, NodeShape.RECTANGLE)
        self.assertEqual(node_a.label, 'Rectangle')
    
    def test_parse_rounded_node(self):
        """Test parsing rounded rectangle node shape."""
        code = "graph TD\n    A(Rounded) --> B"
        diagram = self.parser.parse(code)
        
        node_a = next(n for n in diagram.nodes if n.id == 'A')
        self.assertEqual(node_a.shape, NodeShape.ROUNDED_RECTANGLE)
        self.assertEqual(node_a.label, 'Rounded')
    
    def test_parse_diamond_node(self):
        """Test parsing diamond node shape."""
        code = "graph TD\n    A{Decision} --> B"
        diagram = self.parser.parse(code)
        
        node_a = next(n for n in diagram.nodes if n.id == 'A')
        self.assertEqual(node_a.shape, NodeShape.DIAMOND)
        self.assertEqual(node_a.label, 'Decision')
    
    def test_parse_circle_node(self):
        """Test parsing circle node shape."""
        code = "graph TD\n    A((Circle)) --> B"
        diagram = self.parser.parse(code)
        
        node_a = next(n for n in diagram.nodes if n.id == 'A')
        self.assertEqual(node_a.shape, NodeShape.CIRCLE)
        self.assertEqual(node_a.label, 'Circle')
    
    def test_parse_database_node(self):
        """Test parsing database node shape."""
        code = "graph TD\n    A[[Database]] --> B"
        diagram = self.parser.parse(code)
        
        node_a = next(n for n in diagram.nodes if n.id == 'A')
        self.assertEqual(node_a.shape, NodeShape.DATABASE)
        self.assertEqual(node_a.label, 'Database')
    
    def test_parse_edge_with_label(self):
        """Test parsing edge with label."""
        code = "graph TD\n    A -->|Yes| B"
        diagram = self.parser.parse(code)
        
        edge = diagram.edges[0]
        self.assertEqual(edge.label, 'Yes')
        self.assertEqual(edge.edge_type, EdgeType.ARROW)
    
    def test_parse_dashed_edge(self):
        """Test parsing dashed edge."""
        code = "graph TD\n    A -.-> B"
        diagram = self.parser.parse(code)
        
        edge = diagram.edges[0]
        self.assertEqual(edge.style, EdgeStyle.DASHED)
    
    def test_parse_multiple_edges(self):
        """Test parsing multiple edges from same node."""
        code = """
graph TD
    A --> B
    A --> C
"""
        diagram = self.parser.parse(code)
        
        self.assertEqual(len(diagram.edges), 2)
        sources = {e.source for e in diagram.edges}
        targets = {e.target for e in diagram.edges}
        self.assertEqual(sources, {'A'})
        self.assertEqual(targets, {'B', 'C'})
    
    def test_parse_chain(self):
        """Test parsing node chains."""
        code = "graph TD\n    A --> B --> C"
        diagram = self.parser.parse(code)
        
        self.assertEqual(len(diagram.nodes), 3)
        self.assertEqual(len(diagram.edges), 2)
    
    def test_skip_comments(self):
        """Test that comments are skipped."""
        code = """
graph TD
    %% This is a comment
    A --> B
"""
        diagram = self.parser.parse(code)
        self.assertEqual(len(diagram.nodes), 2)


class TestPlantUMLParser(unittest.TestCase):
    """Tests for the PlantUML parser."""
    
    def setUp(self):
        self.parser = PlantUMLParser()
    
    def test_parse_simple_activity(self):
        """Test parsing a simple activity diagram."""
        code = """
@startuml
start
:Action 1;
stop
@enduml
"""
        diagram = self.parser.parse(code)
        
        self.assertEqual(diagram.direction, Direction.TOP_DOWN)
        self.assertGreaterEqual(len(diagram.nodes), 3)  # start, action, stop
        
        # Check that we have start and stop nodes
        node_ids = {n.id for n in diagram.nodes}
        self.assertIn('start', node_ids)
        self.assertIn('stop', node_ids)
    
    def test_parse_multiple_actions(self):
        """Test parsing multiple actions."""
        code = """
@startuml
start
:Step 1;
:Step 2;
:Step 3;
stop
@enduml
"""
        diagram = self.parser.parse(code)
        
        # Should have start + 3 actions + stop = 5 nodes
        self.assertEqual(len(diagram.nodes), 5)
        
        # Should have 4 edges connecting them
        self.assertEqual(len(diagram.edges), 4)
    
    def test_parse_start_node_shape(self):
        """Test that start node has circle shape."""
        code = "@startuml\nstart\nstop\n@enduml"
        diagram = self.parser.parse(code)
        
        start_node = next(n for n in diagram.nodes if n.id == 'start')
        self.assertEqual(start_node.shape, NodeShape.CIRCLE)
    
    def test_parse_action_label(self):
        """Test parsing action labels."""
        code = "@startuml\nstart\n:My Action Label;\nstop\n@enduml"
        diagram = self.parser.parse(code)
        
        action_nodes = [n for n in diagram.nodes if n.id not in ('start', 'stop')]
        self.assertEqual(len(action_nodes), 1)
        self.assertEqual(action_nodes[0].label, 'My Action Label')
    
    def test_parse_without_markers(self):
        """Test parsing without @startuml/@enduml markers."""
        code = "start\n:Action;\nstop"
        diagram = self.parser.parse(code)
        
        self.assertGreaterEqual(len(diagram.nodes), 3)
    
    def test_parse_explicit_arrows(self):
        """Test parsing explicit arrow syntax."""
        code = "@startuml\nA --> B\nB --> C\n@enduml"
        diagram = self.parser.parse(code)
        
        self.assertEqual(len(diagram.nodes), 3)
        self.assertEqual(len(diagram.edges), 2)


class TestDiagramParser(unittest.TestCase):
    """Tests for the unified diagram parser."""
    
    def setUp(self):
        self.parser = DiagramParser()
    
    def test_detect_mermaid_graph(self):
        """Test detection of Mermaid graph syntax."""
        code = "graph TD\n    A --> B"
        self.assertEqual(self.parser.detect_diagram_type(code), 'mermaid')
    
    def test_detect_mermaid_flowchart(self):
        """Test detection of Mermaid flowchart syntax."""
        code = "flowchart LR\n    A --> B"
        self.assertEqual(self.parser.detect_diagram_type(code), 'mermaid')
    
    def test_detect_plantuml_markers(self):
        """Test detection of PlantUML markers."""
        code = "@startuml\nstart\nstop\n@enduml"
        self.assertEqual(self.parser.detect_diagram_type(code), 'plantuml')
    
    def test_detect_plantuml_actions(self):
        """Test detection of PlantUML action syntax."""
        code = "start\n:Action;\nstop"
        self.assertEqual(self.parser.detect_diagram_type(code), 'plantuml')
    
    def test_auto_parse_mermaid(self):
        """Test auto-parsing Mermaid code."""
        code = "graph TD\n    A[Start] --> B[End]"
        diagram = self.parser.parse(code)
        
        self.assertEqual(len(diagram.nodes), 2)
        self.assertEqual(len(diagram.edges), 1)
    
    def test_auto_parse_plantuml(self):
        """Test auto-parsing PlantUML code."""
        code = "@startuml\nstart\n:Action;\nstop\n@enduml"
        diagram = self.parser.parse(code)
        
        self.assertGreaterEqual(len(diagram.nodes), 3)
    
    def test_explicit_mermaid_parse(self):
        """Test explicit Mermaid parsing."""
        code = "A --> B"  # Could be ambiguous
        diagram = self.parser.parse(code, diagram_type='mermaid')
        
        self.assertEqual(len(diagram.nodes), 2)
    
    def test_explicit_plantuml_parse(self):
        """Test explicit PlantUML parsing."""
        code = "start\nstop"
        diagram = self.parser.parse(code, diagram_type='plantuml')
        
        self.assertGreaterEqual(len(diagram.nodes), 2)
    
    def test_to_layout_elements(self):
        """Test conversion to layout elements."""
        code = "graph TD\n    A[Start] --> B[End]"
        diagram = self.parser.parse(code)
        
        elements, edges = self.parser.to_layout_elements(diagram)
        
        self.assertEqual(len(elements), 2)
        self.assertEqual(len(edges), 1)
        
        # Check element structure
        for element in elements:
            self.assertIn('content', element)
            self.assertIn('element_type', element)


class TestEdgeCases(unittest.TestCase):
    """Tests for edge cases and error handling."""
    
    def setUp(self):
        self.mermaid = MermaidParser()
        self.plantuml = PlantUMLParser()
    
    def test_empty_mermaid(self):
        """Test parsing empty Mermaid code."""
        with self.assertRaises(ValueError):
            self.mermaid.parse("")
    
    def test_mermaid_header_only(self):
        """Test parsing Mermaid with header only."""
        diagram = self.mermaid.parse("graph TD")
        self.assertEqual(len(diagram.nodes), 0)
    
    def test_plantuml_empty_markers(self):
        """Test parsing empty PlantUML markers."""
        diagram = self.plantuml.parse("@startuml\n@enduml")
        self.assertEqual(len(diagram.nodes), 0)
    
    def test_node_without_shape(self):
        """Test parsing node without explicit shape."""
        diagram = self.mermaid.parse("graph TD\n    A --> B")
        
        # Both nodes should default to rectangle
        for node in diagram.nodes:
            self.assertEqual(node.shape, NodeShape.RECTANGLE)
            # Label should be same as ID
            self.assertEqual(node.label, node.id)


if __name__ == '__main__':
    unittest.main()
