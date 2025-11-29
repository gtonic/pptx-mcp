#!/usr/bin/env python3
"""
Auto-Fit Text Demo

This script demonstrates the Intelligent Text Auto-Fit feature for PowerPoint slides.
The auto-fit engine handles extensive AI-generated content by automatically:
- Adjusting font size for optimal readability
- Distributing content across multiple columns
- Splitting content across multiple slides when needed
"""

import os
import sys

# Add the current directory to the Python path to import our modules
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from presentation_manager import presentation_manager
from slide_manager import slide_manager

import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def create_autofit_demo():
    """Create a demonstration presentation showing auto-fit capabilities."""
    
    # Create output directory
    data_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data')
    os.makedirs(data_dir, exist_ok=True)
    
    logger.info("Creating Auto-Fit Text Demo Presentation...")
    
    # 1. Create a new presentation
    result = presentation_manager.create_presentation("autofit_demo")
    logger.info(f"Created presentation: {result}")
    
    # Set properties
    presentation_manager.set_core_properties(
        title="Intelligent Text Auto-Fit Demo",
        subject="Automatic text fitting for AI-generated content",
        author="PowerPoint MCP Server",
        keywords="auto-fit, text, AI, readability",
        comments="Generated using the Intelligent Text Auto-Fit Engine"
    )
    
    # ==========================================================================
    # SLIDE 1: Title Slide
    # ==========================================================================
    slide_manager.add_slide(layout_index=0, title="Intelligent Text Auto-Fit")
    logger.info("Added title slide")
    
    # ==========================================================================
    # SLIDE 2: Short Text - Optimal Font Size
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Short Text - Automatic Font Sizing")
    
    short_text = """Key Benefits:

• Maximum readability
• Smart content handling
• Professional appearance"""

    result = slide_manager.add_auto_fit_text(
        slide_index=1,
        left=0.5, top=1.5, width=9.0, height=5.0,
        text=short_text,
        strategy="smart"
    )
    logger.info(f"Short text result: {result.get('strategy_used')} at {result.get('font_size')}pt")
    
    # ==========================================================================
    # SLIDE 3: Medium Text - Font Size Adjustment
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Medium Text - Font Adjustment")
    
    medium_text = """Executive Summary: Q4 Performance Analysis

Our Q4 results demonstrate strong performance across all key metrics. Revenue increased by 23% year-over-year, driven by robust growth in our enterprise segment.

Key Highlights:
• Enterprise sales grew 35% compared to Q3
• Customer retention rate improved to 94%
• New product launches exceeded targets
• International expansion progressing well

Looking ahead, we expect continued momentum in Q1 as we capitalize on seasonal trends and new market opportunities. Our strategic investments in technology and talent are yielding positive results.

Recommendations:
1. Continue investment in product development
2. Expand sales team in growing markets
3. Maintain focus on customer success"""

    result = slide_manager.add_auto_fit_text(
        slide_index=2,
        left=0.5, top=1.5, width=9.0, height=5.0,
        text=medium_text,
        strategy="smart"
    )
    logger.info(f"Medium text result: {result.get('strategy_used')} at {result.get('font_size')}pt")
    
    # ==========================================================================
    # SLIDE 4: Multi-Column Layout
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Long Text - Multi-Column Layout")
    
    long_text = """Feature Overview

Our product offers a comprehensive suite of features designed for enterprise use. Each feature has been carefully developed to meet the demanding requirements of modern businesses.

Core Capabilities:
• Advanced analytics dashboard with real-time metrics
• Seamless integration with existing systems
• Enterprise-grade security and compliance
• Customizable workflows and automation

Technical Specifications:
• Cloud-native architecture for scalability
• RESTful API for easy integration
• Support for multiple authentication providers
• Data encryption at rest and in transit

Support and Services:
• 24/7 technical support
• Dedicated account management
• Regular training sessions
• Custom development options"""

    result = slide_manager.add_auto_fit_text(
        slide_index=3,
        left=0.5, top=1.5, width=9.0, height=5.0,
        text=long_text,
        strategy="multi_column"  # Force multi-column for demonstration
    )
    logger.info(f"Multi-column result: {result.get('strategy_used')} with {result.get('columns')} columns")
    
    # ==========================================================================
    # SLIDE 5+: Very Long Text - Slide Splitting
    # ==========================================================================
    slide_manager.add_slide(layout_index=6, title="Very Long Content - Slide Split")
    
    very_long_text = """Chapter 1: Introduction to Machine Learning

Machine learning is a branch of artificial intelligence that focuses on building systems that can learn from data and improve their performance over time. Unlike traditional programming, where explicit rules are coded, machine learning allows computers to discover patterns and make decisions based on examples.

The field has seen tremendous growth in recent years, driven by advances in computing power, availability of large datasets, and breakthroughs in algorithmic research. Today, machine learning powers many applications we use daily, from email spam filters to voice assistants.

Chapter 2: Types of Machine Learning

Supervised Learning: In supervised learning, the algorithm learns from labeled training data. The model is trained on input-output pairs and learns to map inputs to outputs. Common applications include image classification, spam detection, and price prediction.

Unsupervised Learning: Unlike supervised learning, unsupervised learning works with unlabeled data. The algorithm tries to find hidden patterns or structures in the data. Examples include customer segmentation, anomaly detection, and dimensionality reduction.

Reinforcement Learning: This type of learning involves an agent that learns to make decisions by interacting with an environment. The agent receives rewards or penalties based on its actions and learns to maximize cumulative rewards over time.

Chapter 3: Deep Learning and Neural Networks

Deep learning is a subset of machine learning that uses neural networks with multiple layers. These deep networks can automatically learn hierarchical representations of data, making them particularly effective for tasks like image recognition, natural language processing, and speech synthesis.

Convolutional Neural Networks (CNNs) are specialized for processing grid-like data such as images. They use convolutional layers to automatically detect features at different scales and positions.

Recurrent Neural Networks (RNNs) are designed for sequential data processing. They maintain a hidden state that allows them to remember information from previous time steps, making them suitable for tasks like language modeling and time series prediction.

Chapter 4: Practical Applications

Machine learning has transformed numerous industries. In healthcare, ML algorithms help diagnose diseases and predict patient outcomes. In finance, they detect fraudulent transactions and optimize trading strategies. In retail, they power recommendation systems and demand forecasting.

Chapter 5: Future Trends

The future of machine learning includes advances in few-shot learning, transfer learning, and explainable AI. As the field matures, we can expect more robust and interpretable models that can be deployed safely in critical applications."""

    result = slide_manager.add_auto_fit_text(
        slide_index=4,
        left=0.5, top=1.2, width=9.0, height=5.5,
        text=very_long_text,
        strategy="split_slides",  # Force slide splitting for demonstration
        create_new_slides=True,
        slide_title_template="ML Guide (Page {page})"
    )
    logger.info(f"Slide split result: {result.get('strategy_used')} using {result.get('slides_used')} slides")
    logger.info(f"New slides created: {result.get('new_slides_created')}")
    
    # ==========================================================================
    # Save the presentation
    # ==========================================================================
    save_path = os.path.join(data_dir, "autofit_demo.pptx")
    save_result = presentation_manager.save_presentation(save_path)
    logger.info(f"Saved presentation: {save_result}")
    
    # Get final info
    info_result = presentation_manager.get_presentation_info()
    
    print("\n" + "="*70)
    print("🎉 AUTO-FIT TEXT DEMO PRESENTATION CREATED SUCCESSFULLY!")
    print("="*70)
    print(f"📁 File saved: {save_path}")
    print(f"📊 Total slides: {info_result.get('slide_count', 'Unknown')}")
    print("\n✨ Auto-Fit Strategies Demonstrated:")
    print("   • Smart: Automatically chooses best approach")
    print("   • Shrink Font: Reduces font size for short-medium content")
    print("   • Multi-Column: Distributes text across columns")
    print("   • Split Slides: Divides content across multiple slides")
    print("\n🤖 Benefits for AI-Generated Content:")
    print("   • Handles extensive content automatically")
    print("   • Maintains maximum readability")
    print("   • Creates sensible slide divisions")
    print("   • No manual formatting required")
    print("="*70)


if __name__ == "__main__":
    try:
        create_autofit_demo()
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
