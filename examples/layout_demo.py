#!/usr/bin/env python3
"""
Layout Engine Demo

This script demonstrates the High-Level Layout Engine for AI-generated PowerPoint diagrams.
The layout engine eliminates the need for explicit coordinates by using structural descriptions.
"""

import os
import sys

# Add the current directory to the Python path to import our modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from presentation_manager import presentation_manager
from layout_manager import layout_manager
from slide_manager import slide_manager

import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_layout_demo():
    """Create a demonstration presentation showing all layout types."""
    
    # Create output directory
    data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
    os.makedirs(data_dir, exist_ok=True)
    
    logger.info("Creating Layout Engine Demo Presentation...")
    
    # 1. Create a new presentation
    result = presentation_manager.create_presentation("layout_demo")
    logger.info(f"Created presentation: {result}")
    
    # Set properties
    presentation_manager.set_core_properties(
        title="Layout Engine Demo",
        subject="High-Level Layout System for AI-Generated Diagrams",
        author="PowerPoint MCP Server",
        keywords="layout, grid, flow, hierarchy, AI",
        comments="Generated using the High-Level Layout Engine"
    )
    
    # ==========================================================================
    # SLIDE 1: Title Slide
    # ==========================================================================
    slide_manager.add_slide(layout_index=0, title="High-Level Layout Engine")
    logger.info("Added title slide")
    
    # ==========================================================================
    # SLIDE 2: Grid Layout Demo
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Grid Layout - No Coordinates Needed")
    
    # Create a 2x2 grid of quarterly data - NO COORDINATES SPECIFIED!
    grid_result = layout_manager.create_grid_layout(
        slide_index=1,
        elements=[
            {
                "content": "Q1 Revenue\n$2.5M",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [79, 129, 189],
                "text_color": [255, 255, 255],
                "font_size": 18,
                "bold": True
            },
            {
                "content": "Q2 Revenue\n$3.2M", 
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [192, 80, 77],
                "text_color": [255, 255, 255],
                "font_size": 18,
                "bold": True
            },
            {
                "content": "Q3 Revenue\n$3.8M",
                "element_type": "shape", 
                "shape_type": "rounded_rectangle",
                "fill_color": [155, 187, 89],
                "text_color": [255, 255, 255],
                "font_size": 18,
                "bold": True
            },
            {
                "content": "Q4 Revenue\n$4.4M",
                "element_type": "shape",
                "shape_type": "rounded_rectangle", 
                "fill_color": [128, 100, 162],
                "text_color": [255, 255, 255],
                "font_size": 18,
                "bold": True
            }
        ],
        rows=2,
        cols=2,
        gap=0.3
    )
    logger.info(f"Grid layout result: {grid_result['message']}")
    
    # ==========================================================================
    # SLIDE 3: List Layout Demo (Vertical)
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="List Layout - Vertical")
    
    # Create a vertical list - NO COORDINATES SPECIFIED!
    list_result = layout_manager.create_list_layout(
        slide_index=2,
        elements=[
            {
                "content": "✓ Feature A: 40% Performance Improvement",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [70, 130, 180],
                "text_color": [255, 255, 255],
                "font_size": 16
            },
            {
                "content": "✓ Feature B: Modern User Interface",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [70, 130, 180],
                "text_color": [255, 255, 255],
                "font_size": 16
            },
            {
                "content": "✓ Feature C: Enhanced Security",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [70, 130, 180],
                "text_color": [255, 255, 255],
                "font_size": 16
            },
            {
                "content": "✓ Feature D: Cloud Integration",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [70, 130, 180],
                "text_color": [255, 255, 255],
                "font_size": 16
            }
        ],
        direction="vertical",
        alignment="center",
        gap=0.2
    )
    logger.info(f"Vertical list result: {list_result['message']}")
    
    # ==========================================================================
    # SLIDE 4: List Layout Demo (Horizontal)
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="List Layout - Horizontal Comparison")
    
    # Create a horizontal list - NO COORDINATES SPECIFIED!
    horiz_list_result = layout_manager.create_list_layout(
        slide_index=3,
        elements=[
            {
                "content": "Basic\n$9/mo",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [100, 149, 237],
                "text_color": [255, 255, 255],
                "font_size": 20,
                "bold": True
            },
            {
                "content": "Pro\n$29/mo",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [65, 105, 225],
                "text_color": [255, 255, 255],
                "font_size": 20,
                "bold": True
            },
            {
                "content": "Enterprise\n$99/mo",
                "element_type": "shape",
                "shape_type": "rounded_rectangle",
                "fill_color": [25, 25, 112],
                "text_color": [255, 255, 255],
                "font_size": 20,
                "bold": True
            }
        ],
        direction="horizontal",
        alignment="middle",
        gap=0.4
    )
    logger.info(f"Horizontal list result: {horiz_list_result['message']}")
    
    # ==========================================================================
    # SLIDE 5: Flow Layout Demo
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Flow Layout - Process Diagram")
    
    # Create a horizontal process flow - NO COORDINATES SPECIFIED!
    flow_result = layout_manager.create_flow_layout(
        slide_index=4,
        steps=[
            {
                "content": "Ideation",
                "fill_color": [46, 139, 87],
                "text_color": [255, 255, 255]
            },
            {
                "content": "Design",
                "fill_color": [60, 179, 113],
                "text_color": [255, 255, 255]
            },
            {
                "content": "Develop",
                "fill_color": [144, 238, 144],
                "text_color": [0, 0, 0]
            },
            {
                "content": "Test",
                "fill_color": [60, 179, 113],
                "text_color": [255, 255, 255]
            },
            {
                "content": "Deploy",
                "fill_color": [46, 139, 87],
                "text_color": [255, 255, 255]
            }
        ],
        direction="horizontal",
        gap=0.5,
        show_connectors=True,
        connector_style="arrow"
    )
    logger.info(f"Flow layout result: {flow_result['message']}")
    
    # ==========================================================================
    # SLIDE 6: Hierarchy Layout Demo
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Hierarchy Layout - Organization Chart")
    
    # Create an org chart - NO COORDINATES SPECIFIED!
    hierarchy_result = layout_manager.create_hierarchy_layout(
        slide_index=5,
        root={
            "content": "CEO",
            "element_type": "shape",
            "shape_type": "rounded_rectangle",
            "fill_color": [47, 79, 79],
            "text_color": [255, 255, 255],
            "children": [
                {
                    "content": "VP Sales",
                    "fill_color": [70, 130, 180],
                    "text_color": [255, 255, 255],
                    "children": [
                        {"content": "Sales East", "fill_color": [135, 206, 235]},
                        {"content": "Sales West", "fill_color": [135, 206, 235]}
                    ]
                },
                {
                    "content": "VP Engineering",
                    "fill_color": [70, 130, 180],
                    "text_color": [255, 255, 255],
                    "children": [
                        {"content": "Frontend", "fill_color": [135, 206, 235]},
                        {"content": "Backend", "fill_color": [135, 206, 235]},
                        {"content": "DevOps", "fill_color": [135, 206, 235]}
                    ]
                },
                {
                    "content": "VP Marketing",
                    "fill_color": [70, 130, 180],
                    "text_color": [255, 255, 255],
                    "children": [
                        {"content": "Brand", "fill_color": [135, 206, 235]},
                        {"content": "Digital", "fill_color": [135, 206, 235]}
                    ]
                }
            ]
        },
        level_gap=0.9,
        sibling_gap=0.25,
        show_connectors=True
    )
    logger.info(f"Hierarchy layout result: {hierarchy_result['message']}")
    
    # ==========================================================================
    # SLIDE 7: Complex Grid Layout
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="3x3 Grid - Dashboard Style")
    
    # Create a 3x3 dashboard grid - NO COORDINATES SPECIFIED!
    dashboard_result = layout_manager.create_grid_layout(
        slide_index=6,
        elements=[
            {"content": "Users\n12,458", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [52, 152, 219]},
            {"content": "Revenue\n$847K", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [46, 204, 113]},
            {"content": "Growth\n+23%", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [155, 89, 182]},
            {"content": "Orders\n3,291", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [231, 76, 60]},
            {"content": "Avg Order\n$257", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [241, 196, 15]},
            {"content": "Rating\n4.8★", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [26, 188, 156]},
            {"content": "Support\n98%", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [52, 73, 94]},
            {"content": "Uptime\n99.9%", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [230, 126, 34]},
            {"content": "NPS\n72", "element_type": "shape", "shape_type": "rounded_rectangle", "fill_color": [149, 165, 166]}
        ],
        rows=3,
        cols=3,
        gap=0.2
    )
    logger.info(f"Dashboard grid result: {dashboard_result['message']}")
    
    # ==========================================================================
    # Save the presentation
    # ==========================================================================
    save_path = os.path.join(data_dir, "layout_engine_demo.pptx")
    save_result = presentation_manager.save_presentation(save_path)
    logger.info(f"Saved presentation: {save_result}")
    
    # Get final info
    info_result = presentation_manager.get_presentation_info()
    
    print("\n" + "="*70)
    print("🎉 LAYOUT ENGINE DEMO PRESENTATION CREATED SUCCESSFULLY!")
    print("="*70)
    print(f"📁 File saved: {save_path}")
    print(f"📊 Total slides: {info_result.get('slide_count', 'Unknown')}")
    print("\n✨ Layout Types Demonstrated:")
    print("   • Grid Layout (2x2 and 3x3)")
    print("   • List Layout (Vertical and Horizontal)")
    print("   • Flow Layout (Process Diagram with Arrows)")
    print("   • Hierarchy Layout (Organization Chart)")
    print("\n🤖 Key Benefits for AI/LLM:")
    print("   • No explicit coordinates required")
    print("   • Structural descriptions instead of pixel positions")
    print("   • Automatic positioning and sizing")
    print("   • Consistent professional appearance")
    print("="*70)


if __name__ == "__main__":
    try:
        create_layout_demo()
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
