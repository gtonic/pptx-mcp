"""
Diagram Renderer Module

Renders parsed diagrams (from Mermaid/PlantUML) as editable PowerPoint vector shapes.
Integrates with the layout engine to automatically position diagram elements.

This module bridges the diagram_parser with the layout_manager and slide_manager
to produce native PowerPoint shapes (not images).
"""
from typing import Optional, Dict, Any, List, Tuple, Union
from dataclasses import dataclass
import logging

from diagram_parser import (
    DiagramParser, ParsedDiagram, DiagramNode, DiagramEdge,
    NodeShape, Direction, DiagramType, diagram_parser
)
from presentation_manager import presentation_manager
from layout_manager import LayoutBounds, layout_manager
from template_manager import template_manager
from performance_optimizer import performance_monitor
import ppt_utils

logger = logging.getLogger(__name__)


@dataclass
class DiagramStyle:
    """Styling configuration for diagram rendering."""
    # Node styles
    default_fill_color: Optional[List[int]] = None
    default_text_color: Optional[List[int]] = None
    default_line_color: Optional[List[int]] = None
    
    # Connector styles
    connector_color: Optional[List[int]] = None
    connector_width: float = 1.5
    
    # Font settings
    font_name: Optional[str] = None
    font_size: int = 14
    bold: bool = False


def get_default_diagram_style() -> DiagramStyle:
    """Get default diagram style based on template."""
    fonts = template_manager.get_default_font_settings()
    colors = template_manager.get_default_color_settings()
    
    return DiagramStyle(
        default_fill_color=list(colors.get("accent_1", (79, 129, 189))),
        default_text_color=list(colors.get("text_1", (255, 255, 255))),
        default_line_color=list(colors.get("text_1", (0, 0, 0))),
        connector_color=list(colors.get("text_1", (64, 64, 64))),
        font_name=fonts.get("body_font_name", "Calibri"),
        font_size=14
    )


class DiagramRenderer:
    """
    Renders diagrams as editable PowerPoint shapes.
    
    Converts parsed diagrams into native PowerPoint shapes with proper
    positioning, styling, and connectors.
    """
    
    def __init__(self):
        self._parser = diagram_parser
    
    @performance_monitor.track_operation("render_diagram")
    def render_mermaid(
        self,
        slide_index: int,
        mermaid_code: str,
        style: Optional[DiagramStyle] = None,
        bounds: Optional[LayoutBounds] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Render a Mermaid diagram as PowerPoint shapes.
        
        Args:
            slide_index: Target slide index
            mermaid_code: Mermaid diagram code
            style: Optional styling configuration
            bounds: Optional layout bounds
            presentation_id: Optional presentation ID
            
        Returns:
            Dict with rendering results
        """
        try:
            diagram = self._parser.parse(mermaid_code, 'mermaid')
            return self._render_diagram(
                slide_index, diagram, style, bounds, presentation_id
            )
        except Exception as e:
            logger.error(f"Failed to parse Mermaid diagram: {e}")
            return {"error": f"Failed to parse Mermaid diagram: {str(e)}"}
    
    @performance_monitor.track_operation("render_diagram")
    def render_plantuml(
        self,
        slide_index: int,
        plantuml_code: str,
        style: Optional[DiagramStyle] = None,
        bounds: Optional[LayoutBounds] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Render a PlantUML diagram as PowerPoint shapes.
        
        Args:
            slide_index: Target slide index
            plantuml_code: PlantUML diagram code
            style: Optional styling configuration
            bounds: Optional layout bounds
            presentation_id: Optional presentation ID
            
        Returns:
            Dict with rendering results
        """
        try:
            diagram = self._parser.parse(plantuml_code, 'plantuml')
            return self._render_diagram(
                slide_index, diagram, style, bounds, presentation_id
            )
        except Exception as e:
            logger.error(f"Failed to parse PlantUML diagram: {e}")
            return {"error": f"Failed to parse PlantUML diagram: {str(e)}"}
    
    @performance_monitor.track_operation("render_diagram")
    def render_auto(
        self,
        slide_index: int,
        diagram_code: str,
        style: Optional[DiagramStyle] = None,
        bounds: Optional[LayoutBounds] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Automatically detect diagram type and render.
        
        Args:
            slide_index: Target slide index
            diagram_code: Diagram code (Mermaid or PlantUML)
            style: Optional styling configuration
            bounds: Optional layout bounds
            presentation_id: Optional presentation ID
            
        Returns:
            Dict with rendering results
        """
        try:
            diagram_type = self._parser.detect_diagram_type(diagram_code)
            diagram = self._parser.parse(diagram_code)
            result = self._render_diagram(
                slide_index, diagram, style, bounds, presentation_id
            )
            result["detected_type"] = diagram_type
            return result
        except Exception as e:
            logger.error(f"Failed to render diagram: {e}")
            return {"error": f"Failed to render diagram: {str(e)}"}
    
    def _render_diagram(
        self,
        slide_index: int,
        diagram: ParsedDiagram,
        style: Optional[DiagramStyle] = None,
        bounds: Optional[LayoutBounds] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Internal method to render a parsed diagram.
        
        Uses different layout strategies based on diagram type and structure.
        """
        try:
            pres = presentation_manager.get_presentation(presentation_id)
            if not (0 <= slide_index < len(pres.slides)):
                return {"error": f"Invalid slide index: {slide_index}"}
            slide = pres.slides[slide_index]
            
            if style is None:
                style = get_default_diagram_style()
            
            if bounds is None:
                bounds = layout_manager._get_slide_bounds(presentation_id)
            
            # Determine rendering strategy based on diagram structure
            if self._is_linear_flow(diagram):
                return self._render_as_flow(slide, slide_index, diagram, style, bounds)
            elif self._is_hierarchical(diagram):
                return self._render_as_hierarchy(slide, slide_index, diagram, style, bounds)
            else:
                # Default to flow layout for general graphs
                return self._render_as_flow(slide, slide_index, diagram, style, bounds)
                
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def _is_linear_flow(self, diagram: ParsedDiagram) -> bool:
        """Check if diagram is a linear flow (no branching or merging)."""
        if not diagram.edges:
            return True
        
        # Count incoming and outgoing edges per node
        outgoing = {}
        incoming = {}
        
        for edge in diagram.edges:
            outgoing[edge.source] = outgoing.get(edge.source, 0) + 1
            incoming[edge.target] = incoming.get(edge.target, 0) + 1
        
        # Linear if no node has more than one outgoing edge
        return all(count <= 1 for count in outgoing.values())
    
    def _is_hierarchical(self, diagram: ParsedDiagram) -> bool:
        """Check if diagram has hierarchical structure (tree-like)."""
        if not diagram.edges:
            return False
        
        # Count incoming edges per node
        incoming = {}
        for edge in diagram.edges:
            incoming[edge.target] = incoming.get(edge.target, 0) + 1
        
        # Check if there's a root (node with no incoming edges)
        node_ids = {node.id for node in diagram.nodes}
        targets = {edge.target for edge in diagram.edges}
        sources = {edge.source for edge in diagram.edges}
        
        roots = sources - targets
        
        # Hierarchical if exactly one root and each node has at most one parent
        return len(roots) == 1 and all(count <= 1 for count in incoming.values())
    
    def _render_as_flow(
        self,
        slide,
        slide_index: int,
        diagram: ParsedDiagram,
        style: DiagramStyle,
        bounds: LayoutBounds
    ) -> Dict[str, Any]:
        """Render diagram as a flow layout."""
        # Determine direction
        is_horizontal = diagram.direction in [Direction.LEFT_RIGHT, Direction.RIGHT_LEFT]
        direction = "horizontal" if is_horizontal else "vertical"
        
        # Order nodes by their position in the flow
        ordered_nodes = self._order_nodes_for_flow(diagram)
        
        # Create elements for flow layout
        elements = []
        for node in ordered_nodes:
            element = self._node_to_element(node, style)
            elements.append(element)
        
        # Use the flow layout engine
        result = layout_manager.create_flow_layout(
            slide_index=slide_index,
            steps=elements,
            direction=direction,
            gap=0.4,
            show_connectors=True,
            connector_style="arrow",
            bounds=bounds,
            presentation_id=None  # Already got slide
        )
        
        result["diagram_type"] = "flow"
        result["node_count"] = len(ordered_nodes)
        result["edge_count"] = len(diagram.edges)
        
        return result
    
    def _render_as_hierarchy(
        self,
        slide,
        slide_index: int,
        diagram: ParsedDiagram,
        style: DiagramStyle,
        bounds: LayoutBounds
    ) -> Dict[str, Any]:
        """Render diagram as a hierarchy layout."""
        # Build hierarchy tree from nodes and edges
        root_element = self._build_hierarchy_tree(diagram, style)
        
        if root_element is None:
            # Fallback to flow if hierarchy building fails
            return self._render_as_flow(slide, slide_index, diagram, style, bounds)
        
        # Use the hierarchy layout engine
        result = layout_manager.create_hierarchy_layout(
            slide_index=slide_index,
            root=root_element,
            level_gap=0.8,
            sibling_gap=0.3,
            show_connectors=True,
            bounds=bounds,
            presentation_id=None
        )
        
        result["diagram_type"] = "hierarchy"
        result["node_count"] = len(diagram.nodes)
        result["edge_count"] = len(diagram.edges)
        
        return result
    
    def _order_nodes_for_flow(self, diagram: ParsedDiagram) -> List[DiagramNode]:
        """Order nodes in flow sequence based on edges."""
        if not diagram.edges:
            return diagram.nodes
        
        # Build adjacency info
        outgoing: Dict[str, List[str]] = {}
        incoming: Dict[str, List[str]] = {}
        
        for edge in diagram.edges:
            if edge.source not in outgoing:
                outgoing[edge.source] = []
            outgoing[edge.source].append(edge.target)
            
            if edge.target not in incoming:
                incoming[edge.target] = []
            incoming[edge.target].append(edge.source)
        
        # Find start nodes (no incoming edges)
        node_ids = {node.id for node in diagram.nodes}
        node_map = {node.id: node for node in diagram.nodes}
        
        start_nodes = [nid for nid in node_ids if nid not in incoming]
        
        if not start_nodes:
            # No clear start, use first node
            start_nodes = [diagram.nodes[0].id] if diagram.nodes else []
        
        # BFS to order nodes
        ordered = []
        visited = set()
        queue = list(start_nodes)
        
        while queue:
            node_id = queue.pop(0)
            if node_id in visited:
                continue
            
            visited.add(node_id)
            if node_id in node_map:
                ordered.append(node_map[node_id])
            
            # Add successors
            for target in outgoing.get(node_id, []):
                if target not in visited:
                    queue.append(target)
        
        # Add any unvisited nodes at the end
        for node in diagram.nodes:
            if node.id not in visited:
                ordered.append(node)
        
        return ordered
    
    def _build_hierarchy_tree(
        self,
        diagram: ParsedDiagram,
        style: DiagramStyle
    ) -> Optional[Dict[str, Any]]:
        """Build a hierarchy tree structure for the layout engine."""
        if not diagram.nodes:
            return None
        
        # Build child relationships
        children: Dict[str, List[str]] = {}
        for edge in diagram.edges:
            if edge.source not in children:
                children[edge.source] = []
            children[edge.source].append(edge.target)
        
        # Find root (node with no incoming edges)
        targets = {edge.target for edge in diagram.edges}
        node_map = {node.id: node for node in diagram.nodes}
        
        roots = [nid for nid in node_map.keys() if nid not in targets]
        
        if not roots:
            return None
        
        root_id = roots[0]
        
        def build_tree(node_id: str) -> Dict[str, Any]:
            node = node_map.get(node_id)
            if not node:
                return {"content": node_id}
            
            element = self._node_to_element(node, style)
            
            child_ids = children.get(node_id, [])
            if child_ids:
                element["children"] = [build_tree(cid) for cid in child_ids]
            
            return element
        
        return build_tree(root_id)
    
    def _node_to_element(self, node: DiagramNode, style: DiagramStyle) -> Dict[str, Any]:
        """Convert a diagram node to a layout element dict."""
        # Map node shapes to PowerPoint shape types
        shape_map = {
            NodeShape.RECTANGLE: "rectangle",
            NodeShape.ROUNDED_RECTANGLE: "rounded_rectangle",
            NodeShape.DIAMOND: "diamond",
            NodeShape.CIRCLE: "oval",
            NodeShape.STADIUM: "rounded_rectangle",
            NodeShape.HEXAGON: "hexagon",
            NodeShape.PARALLELOGRAM: "flowchart_data",
            NodeShape.TRAPEZOID: "rectangle",
            NodeShape.DATABASE: "flowchart_document",
        }
        
        element = {
            "content": node.label,
            "element_type": "shape",
            "shape_type": shape_map.get(node.shape, "rounded_rectangle"),
            "font_size": style.font_size,
        }
        
        if style.bold:
            element["bold"] = True
        if style.font_name:
            element["font_name"] = style.font_name
        
        # Use node-specific colors if set, otherwise use style defaults
        if node.fill_color:
            element["fill_color"] = node.fill_color
        elif style.default_fill_color:
            element["fill_color"] = style.default_fill_color
        
        if node.text_color:
            element["text_color"] = node.text_color
        elif style.default_text_color:
            element["text_color"] = style.default_text_color
        
        if node.line_color:
            element["line_color"] = node.line_color
        elif style.default_line_color:
            element["line_color"] = style.default_line_color
        
        return element


# Global instance
diagram_renderer = DiagramRenderer()
