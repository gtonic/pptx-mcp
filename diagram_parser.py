"""
Diagram Parser Module

Parses text-based diagram languages (Mermaid, PlantUML) into structured representations
that can be rendered as editable PowerPoint vector shapes.

Supports:
- Mermaid flowcharts (graph TD/LR/BT/RL)
- PlantUML activity diagrams and basic flowcharts

Goal: Allow AI/LLM workflows to describe diagrams in familiar DSLs
instead of manually specifying individual shapes and connectors.
"""
from typing import Optional, Dict, Any, List, Tuple, Union
from dataclasses import dataclass, field
from enum import Enum
import re
import logging

logger = logging.getLogger(__name__)


class DiagramType(str, Enum):
    """Supported diagram types."""
    FLOWCHART = "flowchart"
    SEQUENCE = "sequence"
    HIERARCHY = "hierarchy"


class Direction(str, Enum):
    """Diagram flow direction."""
    TOP_DOWN = "TD"
    BOTTOM_UP = "BU"
    LEFT_RIGHT = "LR"
    RIGHT_LEFT = "RL"


class NodeShape(str, Enum):
    """Shape types for diagram nodes."""
    RECTANGLE = "rectangle"
    ROUNDED_RECTANGLE = "rounded_rectangle"
    DIAMOND = "diamond"
    CIRCLE = "circle"
    STADIUM = "stadium"  # Pill shape
    HEXAGON = "hexagon"
    PARALLELOGRAM = "parallelogram"
    TRAPEZOID = "trapezoid"
    DATABASE = "database"


class EdgeStyle(str, Enum):
    """Edge/connector styles."""
    SOLID = "solid"
    DASHED = "dashed"
    DOTTED = "dotted"


class EdgeType(str, Enum):
    """Edge/arrow types."""
    ARROW = "arrow"
    LINE = "line"
    THICK_ARROW = "thick_arrow"


@dataclass
class DiagramNode:
    """Represents a node in a diagram."""
    id: str
    label: str
    shape: NodeShape = NodeShape.RECTANGLE
    fill_color: Optional[List[int]] = None
    text_color: Optional[List[int]] = None
    line_color: Optional[List[int]] = None


@dataclass
class DiagramEdge:
    """Represents an edge/connection in a diagram."""
    source: str
    target: str
    label: Optional[str] = None
    style: EdgeStyle = EdgeStyle.SOLID
    edge_type: EdgeType = EdgeType.ARROW


@dataclass
class ParsedDiagram:
    """Represents a fully parsed diagram."""
    diagram_type: DiagramType
    direction: Direction
    nodes: List[DiagramNode]
    edges: List[DiagramEdge]
    title: Optional[str] = None
    subgraphs: List['ParsedDiagram'] = field(default_factory=list)


class MermaidParser:
    """
    Parser for Mermaid diagram syntax.
    
    Supports:
    - Flowcharts: graph TD/LR/BT/RL
    - Node shapes: [], (), {}, (()), [[]], [/\\]
    - Edge types: -->, ---, -.->
    - Edge labels: -->|label|
    
    Example:
        graph TD
            A[Start] --> B{Decision}
            B -->|Yes| C[Process]
            B -->|No| D[End]
    """
    
    # Node shape patterns (Mermaid syntax -> NodeShape)
    NODE_PATTERNS = [
        (r'\[\[([^\]]*)\]\]', NodeShape.DATABASE),       # [[Database]]
        (r'\(\(([^\)]*)\)\)', NodeShape.CIRCLE),         # ((Circle))
        (r'\(\[([^\]]*)\]\)', NodeShape.STADIUM),        # ([Stadium])
        (r'\{([^\}]*)\}', NodeShape.DIAMOND),            # {Diamond}
        (r'\[\/([^\/]*)\/\]', NodeShape.PARALLELOGRAM),  # [/Parallelogram/]
        (r'\[\\([^\\]*)\\]', NodeShape.TRAPEZOID),       # [\Trapezoid\]
        (r'\[([^\]]*)\]', NodeShape.RECTANGLE),          # [Rectangle]
        (r'\(([^\)]*)\)', NodeShape.ROUNDED_RECTANGLE),  # (Rounded)
    ]
    
    # Edge patterns (Mermaid syntax -> EdgeStyle, EdgeType)
    EDGE_PATTERNS = [
        (r'-->\|([^\|]*)\|', EdgeStyle.SOLID, EdgeType.ARROW, True),    # -->|label|
        (r'-->', EdgeStyle.SOLID, EdgeType.ARROW, False),               # -->
        (r'-\.->\|([^\|]*)\|', EdgeStyle.DASHED, EdgeType.ARROW, True),  # -.->|label|
        (r'-\.->',  EdgeStyle.DASHED, EdgeType.ARROW, False),           # -.->
        (r'===>', EdgeStyle.SOLID, EdgeType.THICK_ARROW, False),        # ===>
        (r'---\|([^\|]*)\|', EdgeStyle.SOLID, EdgeType.LINE, True),     # ---|label|
        (r'---', EdgeStyle.SOLID, EdgeType.LINE, False),                # ---
        (r'-\.-\|([^\|]*)\|', EdgeStyle.DASHED, EdgeType.LINE, True),   # -.-|label|
        (r'-\.-', EdgeStyle.DASHED, EdgeType.LINE, False),              # -.-
    ]
    
    def parse(self, mermaid_code: str) -> ParsedDiagram:
        """
        Parse Mermaid code into a structured diagram.
        
        Args:
            mermaid_code: Mermaid diagram code
            
        Returns:
            ParsedDiagram structure
        """
        lines = mermaid_code.strip().split('\n')
        lines = [line.strip() for line in lines if line.strip() and not line.strip().startswith('%%')]
        
        if not lines:
            raise ValueError("Empty diagram code")
        
        # Check if first line is a header or content
        first_line_lower = lines[0].lower()
        has_header = first_line_lower.startswith('graph ') or first_line_lower.startswith('flowchart ')
        
        # Parse header line to get direction
        direction = self._parse_header(lines[0]) if has_header else Direction.TOP_DOWN
        
        nodes: Dict[str, DiagramNode] = {}
        edges: List[DiagramEdge] = []
        
        # Parse content lines (skip header if present)
        start_idx = 1 if has_header else 0
        for line in lines[start_idx:]:
            # Skip subgraph declarations for now
            if line.startswith('subgraph') or line == 'end':
                continue
            
            # Try to parse as edge definition
            edge_result = self._parse_edge_line(line, nodes)
            if edge_result:
                edges.extend(edge_result)
        
        return ParsedDiagram(
            diagram_type=DiagramType.FLOWCHART,
            direction=direction,
            nodes=list(nodes.values()),
            edges=edges
        )
    
    def _parse_header(self, header: str) -> Direction:
        """Parse the diagram header to get direction."""
        header_lower = header.lower()
        
        if 'graph' in header_lower or 'flowchart' in header_lower:
            if 'lr' in header_lower:
                return Direction.LEFT_RIGHT
            elif 'rl' in header_lower:
                return Direction.RIGHT_LEFT
            elif 'bt' in header_lower or 'bu' in header_lower:
                return Direction.BOTTOM_UP
            else:  # TD or TB is default
                return Direction.TOP_DOWN
        
        return Direction.TOP_DOWN
    
    def _parse_node(self, node_str: str, nodes: Dict[str, DiagramNode]) -> str:
        """
        Parse a node definition and add to nodes dict.
        
        Args:
            node_str: Node string like "A[Label]" or just "A"
            nodes: Dict to add node to
            
        Returns:
            Node ID
        """
        node_str = node_str.strip()
        
        # Try each node pattern
        for pattern, shape in self.NODE_PATTERNS:
            match = re.match(r'(\w+)' + pattern, node_str)
            if match:
                node_id = match.group(1)
                label = match.group(2).strip()
                
                if node_id not in nodes:
                    nodes[node_id] = DiagramNode(
                        id=node_id,
                        label=label,
                        shape=shape
                    )
                return node_id
        
        # No shape specified - use node_str as both ID and label
        node_id = re.match(r'(\w+)', node_str)
        if node_id:
            node_id = node_id.group(1)
            if node_id not in nodes:
                nodes[node_id] = DiagramNode(
                    id=node_id,
                    label=node_id,
                    shape=NodeShape.RECTANGLE
                )
            return node_id
        
        return node_str
    
    def _parse_edge_line(self, line: str, nodes: Dict[str, DiagramNode]) -> List[DiagramEdge]:
        """
        Parse a line that may contain edge definitions, including chains like A --> B --> C.
        
        Args:
            line: Line to parse
            nodes: Dict to add discovered nodes to
            
        Returns:
            List of parsed edges
        """
        edges = []
        
        # Find all edge patterns and their positions
        edge_matches = []
        for pattern, style, edge_type, has_label in self.EDGE_PATTERNS:
            for match in re.finditer(pattern, line):
                edge_matches.append({
                    'start': match.start(),
                    'end': match.end(),
                    'style': style,
                    'edge_type': edge_type,
                    'has_label': has_label,
                    'match': match
                })
        
        if not edge_matches:
            # No edges found, check if there's a standalone node
            if line.strip():
                self._parse_node(line.strip(), nodes)
            return edges
        
        # Sort edge matches by position, then by length (longer matches preferred)
        edge_matches.sort(key=lambda x: (x['start'], -(x['end'] - x['start'])))
        
        # Remove overlapping matches (keep the longest one at each position)
        filtered_matches = []
        last_end = -1
        for em in edge_matches:
            if em['start'] >= last_end:
                filtered_matches.append(em)
                last_end = em['end']
        
        if not filtered_matches:
            return edges
        
        # Process each edge
        prev_end = 0
        for i, em in enumerate(filtered_matches):
            # Source is text between prev_end and this edge start
            source_str = line[prev_end:em['start']].strip()
            
            # Target is text between this edge end and next edge start (or end of line)
            if i + 1 < len(filtered_matches):
                target_end = filtered_matches[i + 1]['start']
            else:
                target_end = len(line)
            
            target_str = line[em['end']:target_end].strip()
            
            if source_str and target_str:
                source_id = self._parse_node(source_str, nodes)
                target_id = self._parse_node(target_str, nodes)
                
                label = None
                if em['has_label'] and em['match'].lastindex:
                    label = em['match'].group(1)
                
                edges.append(DiagramEdge(
                    source=source_id,
                    target=target_id,
                    label=label,
                    style=em['style'],
                    edge_type=em['edge_type']
                ))
            
            prev_end = em['end']
        
        return edges


class PlantUMLParser:
    """
    Parser for PlantUML diagram syntax.
    
    Supports:
    - Activity diagrams: start, stop, :action;
    - Basic flowcharts with if/then/else
    - Arrows: -->, ->, ->>
    
    Example:
        @startuml
        start
        :Initialize;
        if (Condition?) then (yes)
            :Process A;
        else (no)
            :Process B;
        endif
        stop
        @enduml
    """
    
    def parse(self, plantuml_code: str) -> ParsedDiagram:
        """
        Parse PlantUML code into a structured diagram.
        
        Args:
            plantuml_code: PlantUML diagram code
            
        Returns:
            ParsedDiagram structure
        """
        lines = plantuml_code.strip().split('\n')
        lines = [line.strip() for line in lines if line.strip()]
        
        # Remove @startuml/@enduml markers
        lines = [l for l in lines if not l.startswith('@')]
        
        nodes: Dict[str, DiagramNode] = {}
        edges: List[DiagramEdge] = []
        
        prev_node_id: Optional[str] = None
        node_counter = 0
        if_stack: List[Tuple[str, Optional[str]]] = []  # Stack of (if_node_id, else_node_id)
        
        for line in lines:
            line_lower = line.lower()
            
            # Handle start/stop
            if line_lower == 'start':
                node_id = 'start'
                nodes[node_id] = DiagramNode(
                    id=node_id,
                    label='Start',
                    shape=NodeShape.CIRCLE,
                    fill_color=[0, 176, 80]  # Green
                )
                prev_node_id = node_id
                continue
                
            if line_lower in ('stop', 'end'):
                node_id = 'stop'
                nodes[node_id] = DiagramNode(
                    id=node_id,
                    label='End',
                    shape=NodeShape.CIRCLE,
                    fill_color=[192, 0, 0]  # Red
                )
                if prev_node_id:
                    edges.append(DiagramEdge(source=prev_node_id, target=node_id))
                prev_node_id = node_id
                continue
            
            # Handle actions :action;
            action_match = re.match(r':([^;]+);', line)
            if action_match:
                node_counter += 1
                node_id = f'action_{node_counter}'
                label = action_match.group(1).strip()
                
                nodes[node_id] = DiagramNode(
                    id=node_id,
                    label=label,
                    shape=NodeShape.ROUNDED_RECTANGLE
                )
                
                if prev_node_id:
                    edges.append(DiagramEdge(source=prev_node_id, target=node_id))
                prev_node_id = node_id
                continue
            
            # Handle if conditions
            if_match = re.match(r'if\s*\(([^)]+)\)\s*then\s*\(([^)]*)\)', line)
            if if_match:
                node_counter += 1
                node_id = f'decision_{node_counter}'
                label = if_match.group(1).strip()
                yes_label = if_match.group(2).strip() or 'yes'
                
                nodes[node_id] = DiagramNode(
                    id=node_id,
                    label=label,
                    shape=NodeShape.DIAMOND,
                    fill_color=[255, 192, 0]  # Yellow
                )
                
                if prev_node_id:
                    edges.append(DiagramEdge(source=prev_node_id, target=node_id))
                
                if_stack.append((node_id, None))
                prev_node_id = node_id
                continue
            
            # Handle else
            else_match = re.match(r'else\s*\(([^)]*)\)', line)
            if else_match:
                if if_stack:
                    if_node_id, _ = if_stack[-1]
                    no_label = else_match.group(1).strip() or 'no'
                    
                    # Create a marker for the else branch
                    node_counter += 1
                    else_marker_id = f'else_marker_{node_counter}'
                    
                    if_stack[-1] = (if_node_id, else_marker_id)
                    prev_node_id = if_node_id  # Edges from else go from decision node
                continue
            
            # Handle endif
            if line_lower == 'endif':
                if if_stack:
                    if_stack.pop()
                continue
            
            # Handle arrows between explicit nodes: A --> B
            arrow_match = re.match(r'(\w+)\s*(-+>+)\s*(\w+)', line)
            if arrow_match:
                source = arrow_match.group(1)
                target = arrow_match.group(3)
                
                if source not in nodes:
                    nodes[source] = DiagramNode(id=source, label=source, shape=NodeShape.RECTANGLE)
                if target not in nodes:
                    nodes[target] = DiagramNode(id=target, label=target, shape=NodeShape.RECTANGLE)
                
                edges.append(DiagramEdge(source=source, target=target))
                prev_node_id = target
                continue
        
        return ParsedDiagram(
            diagram_type=DiagramType.FLOWCHART,
            direction=Direction.TOP_DOWN,
            nodes=list(nodes.values()),
            edges=edges
        )


class DiagramParser:
    """
    Unified diagram parser that automatically detects and parses diagram types.
    
    Usage:
        parser = DiagramParser()
        
        # Parse Mermaid
        diagram = parser.parse('''
            graph TD
            A[Start] --> B[End]
        ''')
        
        # Parse PlantUML
        diagram = parser.parse('''
            @startuml
            start
            :Process;
            stop
            @enduml
        ''')
    """
    
    def __init__(self):
        self._mermaid_parser = MermaidParser()
        self._plantuml_parser = PlantUMLParser()
    
    def detect_diagram_type(self, code: str) -> str:
        """
        Detect the diagram type from the code.
        
        Args:
            code: Diagram code
            
        Returns:
            'mermaid' or 'plantuml'
        """
        code_lower = code.lower().strip()
        
        # PlantUML markers
        if '@startuml' in code_lower or '@enduml' in code_lower:
            return 'plantuml'
        
        # PlantUML-style keywords
        if code_lower.startswith('start') or (':' in code and ';' in code):
            return 'plantuml'
        
        # Mermaid markers
        if code_lower.startswith('graph ') or code_lower.startswith('flowchart '):
            return 'mermaid'
        
        # Default to Mermaid for arrow-based syntax
        if '-->' in code or '---' in code:
            return 'mermaid'
        
        return 'mermaid'  # Default
    
    def parse(self, code: str, diagram_type: Optional[str] = None) -> ParsedDiagram:
        """
        Parse diagram code into a structured diagram.
        
        Args:
            code: Diagram code in Mermaid or PlantUML syntax
            diagram_type: Optional explicit type ('mermaid' or 'plantuml')
            
        Returns:
            ParsedDiagram structure
        """
        if diagram_type is None:
            diagram_type = self.detect_diagram_type(code)
        
        if diagram_type == 'plantuml':
            return self._plantuml_parser.parse(code)
        else:
            return self._mermaid_parser.parse(code)
    
    def to_layout_elements(
        self,
        diagram: ParsedDiagram,
        style_mapping: Optional[Dict[NodeShape, Dict[str, Any]]] = None
    ) -> Tuple[List[Dict[str, Any]], List[Tuple[str, str, Optional[str]]]]:
        """
        Convert a parsed diagram to layout engine format.
        
        Args:
            diagram: Parsed diagram
            style_mapping: Optional mapping from node shapes to styling
            
        Returns:
            Tuple of (elements list for layout engine, edges list for connectors)
        """
        # Default style mapping
        default_styles = {
            NodeShape.RECTANGLE: {"shape_type": "rectangle"},
            NodeShape.ROUNDED_RECTANGLE: {"shape_type": "rounded_rectangle"},
            NodeShape.DIAMOND: {"shape_type": "diamond"},
            NodeShape.CIRCLE: {"shape_type": "oval"},
            NodeShape.STADIUM: {"shape_type": "rounded_rectangle"},
            NodeShape.HEXAGON: {"shape_type": "hexagon"},
            NodeShape.PARALLELOGRAM: {"shape_type": "flowchart_data"},
            NodeShape.TRAPEZOID: {"shape_type": "rectangle"},
            NodeShape.DATABASE: {"shape_type": "flowchart_document"},
        }
        
        if style_mapping:
            default_styles.update(style_mapping)
        
        elements = []
        node_index_map = {}  # Map node IDs to their indices
        
        for idx, node in enumerate(diagram.nodes):
            style = default_styles.get(node.shape, {"shape_type": "rectangle"})
            
            element = {
                "content": node.label,
                "element_type": "shape",
                **style
            }
            
            if node.fill_color:
                element["fill_color"] = node.fill_color
            if node.text_color:
                element["text_color"] = node.text_color
            if node.line_color:
                element["line_color"] = node.line_color
            
            elements.append(element)
            node_index_map[node.id] = idx
        
        # Convert edges to tuples (source_id, target_id, label)
        edge_tuples = [
            (edge.source, edge.target, edge.label)
            for edge in diagram.edges
        ]
        
        return elements, edge_tuples


# Global instance
diagram_parser = DiagramParser()
