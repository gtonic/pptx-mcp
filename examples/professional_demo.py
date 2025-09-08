#!/usr/bin/env python3
"""
Professional PowerPoint Generation Example

This script demonstrates the enhanced capabilities of the PowerPoint MCP Server
for creating professional presentations with charts, tables, and styling.
"""

import os
import sys
import logging

# Add the current directory to the Python path to import our modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from presentation_manager import presentation_manager
from template_manager import template_manager
from slide_manager import slide_manager

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_professional_presentation():
    """Create a professional presentation demonstrating all features."""
    
    # Create a local data directory
    data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
    os.makedirs(data_dir, exist_ok=True)
    
    logger.info("Creating professional PowerPoint presentation...")
    
    # 1. Create a new presentation
    result = presentation_manager.create_presentation("demo_presentation")
    logger.info(f"Created presentation: {result}")
    
    # 2. Set professional core properties
    props_result = presentation_manager.set_core_properties(
        title="Professional Business Report 2024",
        subject="Q4 Sales Analysis and Strategic Planning",
        author="PowerPoint MCP Server",
        keywords="business, analytics, strategy, Q4, sales",
        comments="Generated using enhanced PowerPoint MCP Server with professional features"
    )
    logger.info(f"Set core properties: {props_result}")
    
    # 3. Add title slide
    title_slide = slide_manager.add_slide(layout_index=0, title="Q4 Business Review")
    logger.info(f"Added title slide: {title_slide}")
    
    # 4. Add overview slide with bullet points
    overview_slide = slide_manager.add_slide(layout_index=1, title="Executive Summary")
    logger.info(f"Added overview slide: {overview_slide}")
    
    # Add bullet points to the content placeholder
    bullet_result = slide_manager.add_bullet_points(
        slide_index=1,
        placeholder_idx=1,
        bullet_points=[
            "Revenue increased 15% over Q3",
            "Customer satisfaction reached 94%",
            "Successful launch of 3 new products",
            "Team expanded to 150 employees",
            "Market share grew to 23%"
        ],
        font_size=18
    )
    logger.info(f"Added bullet points: {bullet_result}")
    
    # 5. Add chart slide
    chart_slide = slide_manager.add_slide(layout_index=6, title="Quarterly Revenue Growth")
    logger.info(f"Added chart slide: {chart_slide}")
    
    # Add a professional chart
    chart_data = {
        'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
        'series': [
            {
                'name': 'Revenue (in millions)',
                'values': [2.5, 3.2, 3.8, 4.4]
            },
            {
                'name': 'Profit (in millions)', 
                'values': [0.5, 0.8, 1.1, 1.5]
            }
        ]
    }
    
    chart_result = slide_manager.add_chart(
        slide_index=2,
        chart_type='column',
        left=1.0,
        top=1.5,
        width=8.0,
        height=5.0,
        data=chart_data
    )
    logger.info(f"Added chart: {chart_result}")
    
    # 6. Add table slide
    table_slide = slide_manager.add_slide(layout_index=6, title="Regional Performance")
    logger.info(f"Added table slide: {table_slide}")
    
    # Add a professional table
    table_data = [
        ['Region', 'Q3 Sales', 'Q4 Sales', 'Growth %'],
        ['North America', '$1.2M', '$1.5M', '25%'],
        ['Europe', '$0.8M', '$1.0M', '25%'],
        ['Asia Pacific', '$0.6M', '$0.9M', '50%'],
        ['Latin America', '$0.2M', '$0.3M', '50%']
    ]
    
    table_result = slide_manager.add_table(
        slide_index=3,
        left=1.0,
        top=1.5,
        rows=5,
        cols=4,
        data=table_data
    )
    logger.info(f"Added table: {table_result}")
    
    # 7. Add shapes and text for visual appeal
    visual_slide = slide_manager.add_slide(layout_index=6, title="Key Initiatives for 2025")
    logger.info(f"Added visual slide: {visual_slide}")
    
    # Add some professional shapes
    shape_result = slide_manager.add_shape(
        slide_index=4,
        shape_type='rounded_rectangle',
        left=1.0,
        top=2.0,
        width=3.0,
        height=1.5,
        fill_color=[79, 129, 189],  # Professional blue
        line_color=[0, 0, 0],
        line_width=2
    )
    logger.info(f"Added shape: {shape_result}")
    
    # Add text over the shape
    text_result = slide_manager.add_textbox(
        slide_index=4,
        left=1.2,
        top=2.2,
        width=2.6,
        height=1.1,
        text="Digital\nTransformation",
        font_size=18,
        font_name="Calibri",
        bold=True,
        color=[255, 255, 255],  # White text
        alignment="center"
    )
    logger.info(f"Added textbox: {text_result}")
    
    # 8. Save the presentation
    save_path = os.path.join(data_dir, "professional_demo.pptx")
    save_result = presentation_manager.save_presentation(save_path)
    logger.info(f"Saved presentation: {save_result}")
    
    # 9. Get final presentation info
    info_result = presentation_manager.get_presentation_info()
    logger.info(f"Final presentation info: {info_result}")
    
    print("\n" + "="*60)
    print("🎉 PROFESSIONAL PRESENTATION CREATED SUCCESSFULLY!")
    print("="*60)
    print(f"📁 File saved: {save_path}")
    print(f"📊 Total slides: {info_result.get('slide_count', 'Unknown')}")
    print(f"📋 Title: {info_result.get('core_properties', {}).get('title', 'Unknown')}")
    print(f"👤 Author: {info_result.get('core_properties', {}).get('author', 'Unknown')}")
    print("\n✨ Features demonstrated:")
    print("   • Professional slide layouts")
    print("   • Executive bullet points")
    print("   • Business charts and graphs") 
    print("   • Data tables with styling")
    print("   • Shapes and professional text")
    print("   • Corporate metadata")
    print("="*60)

if __name__ == "__main__":
    try:
        create_professional_presentation()
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")
        sys.exit(1)