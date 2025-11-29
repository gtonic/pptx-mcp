#!/usr/bin/env python3
"""
Diagram Support Demo

This script demonstrates the Mermaid/PlantUML diagram support for AI-generated
PowerPoint diagrams. The parser converts text-based diagram descriptions into
editable PowerPoint vector shapes (not images!).
"""

import os
import sys

# Add the current directory to the Python path to import our modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from presentation_manager import presentation_manager
from diagram_renderer import diagram_renderer
from slide_manager import slide_manager

import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_diagram_demo():
    """Create a demonstration presentation showing diagram support."""
    
    # Create output directory
    data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
    os.makedirs(data_dir, exist_ok=True)
    
    logger.info("Creating Diagram Support Demo Presentation...")
    
    # 1. Create a new presentation
    result = presentation_manager.create_presentation("diagram_demo")
    logger.info(f"Created presentation: {result}")
    
    # Set properties
    presentation_manager.set_core_properties(
        title="Diagram Support Demo",
        subject="Mermaid and PlantUML to PowerPoint",
        author="PowerPoint MCP Server",
        keywords="mermaid, plantuml, diagram, flowchart, AI",
        comments="Generated from text-based diagram DSLs"
    )
    
    # ==========================================================================
    # SLIDE 1: Title Slide
    # ==========================================================================
    slide_manager.add_slide(layout_index=0, title="Mermaid & PlantUML Diagram Support")
    logger.info("Added title slide")
    
    # ==========================================================================
    # SLIDE 2: Simple Mermaid Flowchart
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Mermaid Flowchart - Simple")
    
    mermaid_simple = """
graph TD
    A[Start] --> B[Process Data]
    B --> C[Generate Report]
    C --> D[End]
"""
    
    result = diagram_renderer.render_mermaid(
        slide_index=1,
        mermaid_code=mermaid_simple
    )
    logger.info(f"Simple Mermaid result: {result.get('message', result.get('error', 'Unknown'))}")
    
    # ==========================================================================
    # SLIDE 3: Mermaid with Decision Diamond
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Mermaid Flowchart - Decision")
    
    mermaid_decision = """
graph TD
    A[Start] --> B{Is Valid?}
    B -->|Yes| C[Process]
    B -->|No| D[Error Handler]
    C --> E[Save]
    D --> E
    E --> F[End]
"""
    
    result = diagram_renderer.render_mermaid(
        slide_index=2,
        mermaid_code=mermaid_decision
    )
    logger.info(f"Decision Mermaid result: {result.get('message', result.get('error', 'Unknown'))}")
    
    # ==========================================================================
    # SLIDE 4: Mermaid Left-to-Right Flow
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Mermaid Flowchart - Left to Right")
    
    mermaid_lr = """
graph LR
    A[Input] --> B[Validate]
    B --> C[Transform]
    C --> D[Output]
"""
    
    result = diagram_renderer.render_mermaid(
        slide_index=3,
        mermaid_code=mermaid_lr
    )
    logger.info(f"LR Mermaid result: {result.get('message', result.get('error', 'Unknown'))}")
    
    # ==========================================================================
    # SLIDE 5: PlantUML Activity Diagram
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="PlantUML Activity Diagram")
    
    plantuml_activity = """
@startuml
start
:Initialize System;
:Load Configuration;
:Connect to Database;
:Ready for Requests;
stop
@enduml
"""
    
    result = diagram_renderer.render_plantuml(
        slide_index=4,
        plantuml_code=plantuml_activity
    )
    logger.info(f"PlantUML activity result: {result.get('message', result.get('error', 'Unknown'))}")
    
    # ==========================================================================
    # SLIDE 6: PlantUML with Conditions
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="PlantUML with Conditions")
    
    plantuml_condition = """
@startuml
start
:Receive Request;
if (Authenticated?) then (yes)
    :Process Request;
    :Send Response;
else (no)
    :Return 401;
endif
stop
@enduml
"""
    
    result = diagram_renderer.render_plantuml(
        slide_index=5,
        plantuml_code=plantuml_condition
    )
    logger.info(f"PlantUML condition result: {result.get('message', result.get('error', 'Unknown'))}")
    
    # ==========================================================================
    # SLIDE 7: Auto-detect Mermaid
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Auto-Detect (Mermaid)")
    
    auto_mermaid = """
graph TD
    User[User] --> API[API Gateway]
    API --> Auth[Auth Service]
    API --> Data[Data Service]
"""
    
    result = diagram_renderer.render_auto(
        slide_index=6,
        diagram_code=auto_mermaid
    )
    logger.info(f"Auto-detect Mermaid result: detected_type={result.get('detected_type', 'unknown')}")
    
    # ==========================================================================
    # SLIDE 8: Auto-detect PlantUML
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Auto-Detect (PlantUML)")
    
    auto_plantuml = """
@startuml
start
:Open App;
:Login;
:Browse;
:Checkout;
stop
@enduml
"""
    
    result = diagram_renderer.render_auto(
        slide_index=7,
        diagram_code=auto_plantuml
    )
    logger.info(f"Auto-detect PlantUML result: detected_type={result.get('detected_type', 'unknown')}")
    
    # ==========================================================================
    # Save the presentation
    # ==========================================================================
    save_path = os.path.join(data_dir, "diagram_demo.pptx")
    save_result = presentation_manager.save_presentation(save_path)
    logger.info(f"Saved presentation: {save_result}")
    
    # Get final info
    info_result = presentation_manager.get_presentation_info()
    
    print("\n" + "="*70)
    print("🎉 DIAGRAM SUPPORT DEMO PRESENTATION CREATED SUCCESSFULLY!")
    print("="*70)
    print(f"📁 File saved: {save_path}")
    print(f"📊 Total slides: {info_result.get('slide_count', 'Unknown')}")
    print("\n✨ Diagram Types Demonstrated:")
    print("   • Mermaid Flowchart (Top-Down)")
    print("   • Mermaid Flowchart (Left-Right)")
    print("   • Mermaid with Decision Diamonds")
    print("   • PlantUML Activity Diagram")
    print("   • PlantUML with Conditionals")
    print("   • Auto-detection of diagram type")
    print("\n🤖 Key Benefits for AI/LLM:")
    print("   • Write diagrams in familiar Mermaid/PlantUML syntax")
    print("   • No need to specify individual shapes and connectors")
    print("   • Output is editable PowerPoint shapes (not images!)")
    print("   • Automatic layout and positioning")
    print("="*70)


if __name__ == "__main__":
    try:
        create_diagram_demo()
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
