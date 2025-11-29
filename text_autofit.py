"""
Text Auto-Fit Module

Provides intelligent text auto-fit logic for PowerPoint slides.
When AI-generated content is extensive, this module automatically adjusts:
- Font size based on text length and container dimensions
- Multi-column layouts for better readability
- Slide splitting for very long content

Goal: Maximum readability and sensible slide division for large data sets.
"""
from typing import Optional, Dict, Any, List, Tuple
from dataclasses import dataclass, field
from enum import Enum
import math
import logging

logger = logging.getLogger(__name__)


class AutoFitStrategy(str, Enum):
    """Available auto-fit strategies for text content."""
    SHRINK_FONT = "shrink_font"  # Reduce font size to fit
    MULTI_COLUMN = "multi_column"  # Split into multiple columns
    SPLIT_SLIDES = "split_slides"  # Split across multiple slides
    SMART = "smart"  # Automatically choose best strategy


@dataclass
class TextMetrics:
    """Metrics for text content to determine optimal layout."""
    text: str
    char_count: int
    word_count: int
    line_count: int
    paragraph_count: int
    avg_word_length: float
    has_bullets: bool
    has_newlines: bool


@dataclass
class ContainerDimensions:
    """Dimensions of the text container in inches."""
    width: float
    height: float
    # Standard PowerPoint slide dimensions
    slide_width: float = 10.0
    slide_height: float = 7.5


@dataclass
class AutoFitConfig:
    """Configuration for auto-fit behavior."""
    min_font_size: int = 10  # Minimum readable font size in points
    max_font_size: int = 44  # Maximum font size in points
    default_font_size: int = 18  # Default font size in points
    target_chars_per_line: int = 50  # Ideal characters per line for readability
    max_lines_per_slide: int = 12  # Maximum lines before considering split
    column_gap: float = 0.3  # Gap between columns in inches
    min_column_width: float = 2.0  # Minimum column width in inches
    bullet_indent: float = 0.25  # Bullet indentation in inches
    chars_per_inch: float = 7.0  # Approximate characters per inch at 18pt
    points_per_inch: float = 72.0  # Points per inch
    stacking_gap: float = 0.2  # Gap between stacked elements in inches


@dataclass
class AutoFitResult:
    """Result of auto-fit calculation."""
    strategy: AutoFitStrategy
    font_size: int
    columns: int
    slides_needed: int
    text_segments: List[str]
    column_width: float
    recommendation: str
    container_adjustments: Optional[Dict[str, float]] = None


class TextAutoFitEngine:
    """
    Intelligent text auto-fit engine for PowerPoint slides.
    
    This engine analyzes text content and container dimensions to determine
    the optimal layout strategy for maximum readability.
    """
    
    def __init__(self, config: Optional[AutoFitConfig] = None):
        self.config = config or AutoFitConfig()
    
    def analyze_text(self, text: str) -> TextMetrics:
        """
        Analyze text content to extract metrics.
        
        Args:
            text: The text content to analyze
            
        Returns:
            TextMetrics object with content metrics
        """
        lines = text.split('\n')
        words = text.split()
        paragraphs = [p for p in text.split('\n\n') if p.strip()]
        
        has_bullets = any(
            line.strip().startswith(('•', '-', '*', '●', '○', '▪', '▸'))
            for line in lines
        )
        
        avg_word_length = (
            sum(len(w) for w in words) / len(words)
            if words else 0
        )
        
        return TextMetrics(
            text=text,
            char_count=len(text),
            word_count=len(words),
            line_count=len(lines),
            paragraph_count=len(paragraphs) or 1,
            avg_word_length=avg_word_length,
            has_bullets=has_bullets,
            has_newlines='\n' in text
        )
    
    def estimate_lines_needed(
        self,
        text: str,
        container_width: float,
        font_size: int
    ) -> int:
        """
        Estimate the number of lines needed to display text.
        
        Args:
            text: The text content
            container_width: Width of container in inches
            font_size: Font size in points
            
        Returns:
            Estimated number of lines needed
        """
        # Validate font_size to prevent division by zero
        if font_size <= 0:
            font_size = self.config.default_font_size
        
        # Adjust chars per inch based on font size relative to default
        scale_factor = self.config.default_font_size / font_size
        chars_per_inch = self.config.chars_per_inch * scale_factor
        chars_per_line = int(container_width * chars_per_inch)
        
        # Account for explicit line breaks
        explicit_lines = text.split('\n')
        total_lines = 0
        
        for line in explicit_lines:
            if not line.strip():
                total_lines += 1  # Empty line
            else:
                # Calculate wrapped lines for this content
                line_chars = len(line)
                wrapped_lines = max(1, math.ceil(line_chars / chars_per_line))
                total_lines += wrapped_lines
        
        return total_lines
    
    def calculate_optimal_font_size(
        self,
        metrics: TextMetrics,
        container: ContainerDimensions,
        target_lines: Optional[int] = None
    ) -> int:
        """
        Calculate the optimal font size for the given text and container.
        
        Args:
            metrics: Text metrics
            container: Container dimensions
            target_lines: Optional target number of lines
            
        Returns:
            Optimal font size in points
        """
        if target_lines is None:
            target_lines = self.config.max_lines_per_slide
        
        # Start with default font size and adjust
        font_size = self.config.default_font_size
        
        # Estimate lines at default size
        lines_needed = self.estimate_lines_needed(
            metrics.text, container.width, font_size
        )
        
        if lines_needed <= target_lines:
            # Text fits at default size, try to increase if possible
            while font_size < self.config.max_font_size:
                test_size = font_size + 2
                test_lines = self.estimate_lines_needed(
                    metrics.text, container.width, test_size
                )
                if test_lines <= target_lines:
                    font_size = test_size
                else:
                    break
        else:
            # Text doesn't fit, reduce font size
            while font_size > self.config.min_font_size and lines_needed > target_lines:
                font_size -= 1
                lines_needed = self.estimate_lines_needed(
                    metrics.text, container.width, font_size
                )
        
        return font_size
    
    def split_into_columns(
        self,
        text: str,
        num_columns: int,
        preserve_paragraphs: bool = True
    ) -> List[str]:
        """
        Split text into multiple columns for multi-column layout.
        
        Args:
            text: The text to split
            num_columns: Number of columns to create
            preserve_paragraphs: Whether to keep paragraphs together
            
        Returns:
            List of text segments, one per column
        """
        if num_columns <= 1:
            return [text]
        
        if preserve_paragraphs:
            # Split by paragraphs first
            paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
            if not paragraphs:
                paragraphs = [p.strip() for p in text.split('\n') if p.strip()]
            
            if len(paragraphs) >= num_columns:
                # Distribute paragraphs evenly
                per_column = len(paragraphs) // num_columns
                remainder = len(paragraphs) % num_columns
                
                columns = []
                idx = 0
                for col in range(num_columns):
                    col_paras = per_column + (1 if col < remainder else 0)
                    col_text = '\n\n'.join(paragraphs[idx:idx + col_paras])
                    columns.append(col_text)
                    idx += col_paras
                
                return columns
        
        # Fallback: split by character count
        total_chars = len(text)
        chars_per_column = total_chars // num_columns
        
        columns = []
        start = 0
        
        for col in range(num_columns):
            if col == num_columns - 1:
                # Last column gets the rest
                columns.append(text[start:].strip())
            else:
                # Find a good split point near the target
                end = start + chars_per_column
                
                # Try to split at a paragraph break
                para_break = text.rfind('\n\n', start, end + 50)
                if para_break > start:
                    end = para_break
                else:
                    # Try to split at a line break
                    line_break = text.rfind('\n', start, end + 20)
                    if line_break > start:
                        end = line_break
                    else:
                        # Try to split at a space
                        space = text.rfind(' ', start, end + 10)
                        if space > start:
                            end = space
                
                columns.append(text[start:end].strip())
                start = end
        
        return columns
    
    def split_for_multiple_slides(
        self,
        text: str,
        container: ContainerDimensions,
        font_size: int,
        max_lines_per_slide: Optional[int] = None
    ) -> List[str]:
        """
        Split text across multiple slides when it's too long.
        
        Args:
            text: The text to split
            container: Container dimensions
            font_size: Font size to use
            max_lines_per_slide: Maximum lines per slide
            
        Returns:
            List of text segments, one per slide
        """
        if max_lines_per_slide is None:
            max_lines_per_slide = self.config.max_lines_per_slide
        
        # Try to split by paragraphs or logical sections
        paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
        
        if not paragraphs:
            # No paragraphs, split by lines
            paragraphs = [p.strip() for p in text.split('\n') if p.strip()]
        
        slides = []
        current_slide_text = []
        current_lines = 0
        
        for para in paragraphs:
            para_lines = self.estimate_lines_needed(para, container.width, font_size)
            
            if current_lines + para_lines > max_lines_per_slide:
                if current_slide_text:
                    # Save current slide and start new one
                    slides.append('\n\n'.join(current_slide_text))
                    current_slide_text = [para]
                    current_lines = para_lines
                else:
                    # Single paragraph too long, need to split it
                    # This is a fallback for very long paragraphs
                    slides.append(para)
                    current_lines = 0
            else:
                current_slide_text.append(para)
                current_lines += para_lines
        
        # Add remaining content
        if current_slide_text:
            slides.append('\n\n'.join(current_slide_text))
        
        return slides if slides else [text]
    
    def calculate_optimal_columns(
        self,
        metrics: TextMetrics,
        container: ContainerDimensions,
        font_size: int
    ) -> int:
        """
        Calculate the optimal number of columns for the content.
        
        Args:
            metrics: Text metrics
            container: Container dimensions
            font_size: Font size in points
            
        Returns:
            Optimal number of columns (1-3)
        """
        # Calculate how many lines needed in single column
        lines_needed = self.estimate_lines_needed(
            metrics.text, container.width, font_size
        )
        
        # If content fits in single column, use it
        if lines_needed <= self.config.max_lines_per_slide:
            return 1
        
        # Check if 2 columns would work
        col2_width = (container.width - self.config.column_gap) / 2
        if col2_width >= self.config.min_column_width:
            # Use actual split text for more accurate line estimation
            col2_segments = self.split_into_columns(metrics.text, 2)
            max_lines_col2 = max(
                self.estimate_lines_needed(seg, col2_width, font_size)
                for seg in col2_segments
            )
            
            if max_lines_col2 <= self.config.max_lines_per_slide:
                return 2
        
        # Check if 3 columns would work
        col3_width = (container.width - 2 * self.config.column_gap) / 3
        if col3_width >= self.config.min_column_width:
            # Use actual split text for more accurate line estimation
            col3_segments = self.split_into_columns(metrics.text, 3)
            max_lines_col3 = max(
                self.estimate_lines_needed(seg, col3_width, font_size)
                for seg in col3_segments
            )
            
            if max_lines_col3 <= self.config.max_lines_per_slide:
                return 3
        
        # Default to 2 columns if none fit perfectly
        if col2_width >= self.config.min_column_width:
            return 2
        
        return 1
    
    def auto_fit(
        self,
        text: str,
        container: ContainerDimensions,
        strategy: AutoFitStrategy = AutoFitStrategy.SMART,
        preferred_font_size: Optional[int] = None
    ) -> AutoFitResult:
        """
        Automatically fit text to container using the best strategy.
        
        Args:
            text: The text content to fit
            container: Container dimensions
            strategy: Auto-fit strategy to use
            preferred_font_size: Optional preferred font size
            
        Returns:
            AutoFitResult with optimal layout settings
        """
        metrics = self.analyze_text(text)
        
        if strategy == AutoFitStrategy.SMART:
            return self._smart_auto_fit(metrics, container, preferred_font_size)
        elif strategy == AutoFitStrategy.SHRINK_FONT:
            return self._shrink_font_fit(metrics, container, preferred_font_size)
        elif strategy == AutoFitStrategy.MULTI_COLUMN:
            return self._multi_column_fit(metrics, container, preferred_font_size)
        elif strategy == AutoFitStrategy.SPLIT_SLIDES:
            return self._split_slides_fit(metrics, container, preferred_font_size)
        else:
            return self._smart_auto_fit(metrics, container, preferred_font_size)
    
    def _smart_auto_fit(
        self,
        metrics: TextMetrics,
        container: ContainerDimensions,
        preferred_font_size: Optional[int] = None
    ) -> AutoFitResult:
        """
        Intelligently choose the best auto-fit strategy.
        
        Priority:
        1. If text fits at reasonable font size (≥14pt), use single column
        2. If text fits with slightly smaller font (≥12pt), shrink font
        3. If content has multiple sections, try multi-column
        4. If content is too long, split across slides
        """
        base_font_size = preferred_font_size or self.config.default_font_size
        
        # First, try at preferred/default font size
        lines_at_base = self.estimate_lines_needed(
            metrics.text, container.width, base_font_size
        )
        
        if lines_at_base <= self.config.max_lines_per_slide:
            # Text fits at base size, maybe we can increase font
            optimal_size = self.calculate_optimal_font_size(
                metrics, container, self.config.max_lines_per_slide
            )
            return AutoFitResult(
                strategy=AutoFitStrategy.SHRINK_FONT,
                font_size=optimal_size,
                columns=1,
                slides_needed=1,
                text_segments=[metrics.text],
                column_width=container.width,
                recommendation="Text fits well at optimal font size"
            )
        
        # Try with reduced font size
        reduced_font_size = self.calculate_optimal_font_size(
            metrics, container, self.config.max_lines_per_slide
        )
        
        if reduced_font_size >= 12:  # Still readable
            lines_reduced = self.estimate_lines_needed(
                metrics.text, container.width, reduced_font_size
            )
            
            if lines_reduced <= self.config.max_lines_per_slide:
                return AutoFitResult(
                    strategy=AutoFitStrategy.SHRINK_FONT,
                    font_size=reduced_font_size,
                    columns=1,
                    slides_needed=1,
                    text_segments=[metrics.text],
                    column_width=container.width,
                    recommendation=f"Font reduced from {base_font_size}pt to {reduced_font_size}pt for readability"
                )
        
        # Try multi-column layout
        if metrics.paragraph_count > 1 or metrics.line_count > 5:
            num_columns = self.calculate_optimal_columns(
                metrics, container, base_font_size
            )
            
            if num_columns > 1:
                column_width = (
                    container.width - (num_columns - 1) * self.config.column_gap
                ) / num_columns
                
                text_segments = self.split_into_columns(metrics.text, num_columns)
                
                # Verify columns fit
                max_column_lines = max(
                    self.estimate_lines_needed(seg, column_width, base_font_size)
                    for seg in text_segments
                )
                
                if max_column_lines <= self.config.max_lines_per_slide:
                    return AutoFitResult(
                        strategy=AutoFitStrategy.MULTI_COLUMN,
                        font_size=base_font_size,
                        columns=num_columns,
                        slides_needed=1,
                        text_segments=text_segments,
                        column_width=column_width,
                        recommendation=f"Content distributed across {num_columns} columns for better readability"
                    )
        
        # Last resort: split across slides
        font_for_split = max(reduced_font_size, self.config.min_font_size + 2)
        text_segments = self.split_for_multiple_slides(
            metrics.text, container, font_for_split
        )
        
        return AutoFitResult(
            strategy=AutoFitStrategy.SPLIT_SLIDES,
            font_size=font_for_split,
            columns=1,
            slides_needed=len(text_segments),
            text_segments=text_segments,
            column_width=container.width,
            recommendation=f"Content split across {len(text_segments)} slides for optimal readability"
        )
    
    def _shrink_font_fit(
        self,
        metrics: TextMetrics,
        container: ContainerDimensions,
        preferred_font_size: Optional[int] = None
    ) -> AutoFitResult:
        """Fit text by shrinking font size only."""
        optimal_size = self.calculate_optimal_font_size(metrics, container)
        
        return AutoFitResult(
            strategy=AutoFitStrategy.SHRINK_FONT,
            font_size=optimal_size,
            columns=1,
            slides_needed=1,
            text_segments=[metrics.text],
            column_width=container.width,
            recommendation=f"Font size adjusted to {optimal_size}pt"
        )
    
    def _multi_column_fit(
        self,
        metrics: TextMetrics,
        container: ContainerDimensions,
        preferred_font_size: Optional[int] = None
    ) -> AutoFitResult:
        """Fit text using multi-column layout."""
        base_font_size = preferred_font_size or self.config.default_font_size
        num_columns = self.calculate_optimal_columns(metrics, container, base_font_size)
        
        if num_columns == 1:
            num_columns = 2  # Force at least 2 columns for this strategy
        
        column_width = (
            container.width - (num_columns - 1) * self.config.column_gap
        ) / num_columns
        
        text_segments = self.split_into_columns(metrics.text, num_columns)
        
        return AutoFitResult(
            strategy=AutoFitStrategy.MULTI_COLUMN,
            font_size=base_font_size,
            columns=num_columns,
            slides_needed=1,
            text_segments=text_segments,
            column_width=column_width,
            recommendation=f"Content arranged in {num_columns} columns"
        )
    
    def _split_slides_fit(
        self,
        metrics: TextMetrics,
        container: ContainerDimensions,
        preferred_font_size: Optional[int] = None
    ) -> AutoFitResult:
        """Fit text by splitting across multiple slides."""
        font_size = preferred_font_size or self.config.default_font_size
        text_segments = self.split_for_multiple_slides(
            metrics.text, container, font_size
        )
        
        return AutoFitResult(
            strategy=AutoFitStrategy.SPLIT_SLIDES,
            font_size=font_size,
            columns=1,
            slides_needed=len(text_segments),
            text_segments=text_segments,
            column_width=container.width,
            recommendation=f"Content split across {len(text_segments)} slides"
        )


# Global instance
text_autofit_engine = TextAutoFitEngine()
