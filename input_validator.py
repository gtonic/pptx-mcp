"""
Input Validation Module

Provides validation utilities for PowerPoint MCP Server to ensure data integrity
and security of user inputs.
"""
from typing import Any, Dict, List, Optional, Tuple
import os
import re


class ValidationError(Exception):
    """Custom exception for validation errors."""
    pass


class InputValidator:
    """Validates various input types for PowerPoint operations."""
    
    # Allowed file extensions for security
    ALLOWED_EXTENSIONS = {'.pptx', '.png', '.jpg', '.jpeg', '.gif', '.bmp'}
    
    # Slide dimension constraints (in inches)
    MIN_DIMENSION = 0.1
    MAX_DIMENSION = 20.0
    
    # Text length constraints
    MAX_TEXT_LENGTH = 10000
    MAX_TITLE_LENGTH = 255
    
    # Color value constraints
    MIN_COLOR_VALUE = 0
    MAX_COLOR_VALUE = 255
    
    @staticmethod
    def validate_file_path(file_path: str, check_exists: bool = False) -> str:
        """
        Validate file path for security and format.
        
        Args:
            file_path: The file path to validate
            check_exists: Whether to check if the file exists
            
        Returns:
            Normalized file path
            
        Raises:
            ValidationError: If path is invalid or insecure
        """
        if not file_path or not isinstance(file_path, str):
            raise ValidationError("File path must be a non-empty string")
        
        # Normalize path to prevent directory traversal
        normalized_path = os.path.normpath(file_path)
        
        # Check for directory traversal attempts
        if '..' in normalized_path or normalized_path.startswith('/'):
            if not normalized_path.startswith('/data/') and not os.path.isabs(normalized_path):
                raise ValidationError("Invalid file path: directory traversal not allowed")
        
        # Check file extension
        _, ext = os.path.splitext(normalized_path.lower())
        if ext not in InputValidator.ALLOWED_EXTENSIONS:
            raise ValidationError(f"Invalid file extension: {ext}. Allowed: {', '.join(InputValidator.ALLOWED_EXTENSIONS)}")
        
        # Check if file exists when required
        if check_exists and not os.path.exists(normalized_path):
            raise ValidationError(f"File not found: {normalized_path}")
        
        return normalized_path
    
    @staticmethod
    def validate_dimensions(left: float, top: float, width: float, height: float) -> Tuple[float, float, float, float]:
        """
        Validate position and size dimensions.
        
        Args:
            left: Left position in inches
            top: Top position in inches
            width: Width in inches
            height: Height in inches
            
        Returns:
            Validated dimensions tuple
            
        Raises:
            ValidationError: If dimensions are invalid
        """
        try:
            left = float(left)
            top = float(top)
            width = float(width)
            height = float(height)
        except (ValueError, TypeError):
            raise ValidationError("Dimensions must be numeric values")
        
        if left < 0 or top < 0:
            raise ValidationError("Position values (left, top) cannot be negative")
        
        if width < InputValidator.MIN_DIMENSION or height < InputValidator.MIN_DIMENSION:
            raise ValidationError(f"Width and height must be at least {InputValidator.MIN_DIMENSION} inches")
        
        if width > InputValidator.MAX_DIMENSION or height > InputValidator.MAX_DIMENSION:
            raise ValidationError(f"Width and height cannot exceed {InputValidator.MAX_DIMENSION} inches")
        
        return left, top, width, height
    
    @staticmethod
    def validate_text(text: str, max_length: Optional[int] = None) -> str:
        """
        Validate text input for length and content.
        
        Args:
            text: Text to validate
            max_length: Maximum allowed length (default: MAX_TEXT_LENGTH)
            
        Returns:
            Validated text
            
        Raises:
            ValidationError: If text is invalid
        """
        if not isinstance(text, str):
            raise ValidationError("Text must be a string")
        
        max_len = max_length or InputValidator.MAX_TEXT_LENGTH
        if len(text) > max_len:
            raise ValidationError(f"Text length ({len(text)}) exceeds maximum ({max_len})")
        
        return text
    
    @staticmethod
    def validate_color(color: List[int]) -> List[int]:
        """
        Validate RGB color values.
        
        Args:
            color: RGB color as list [r, g, b]
            
        Returns:
            Validated color list
            
        Raises:
            ValidationError: If color is invalid
        """
        if not isinstance(color, list) or len(color) != 3:
            raise ValidationError("Color must be a list of 3 RGB values")
        
        try:
            validated_color = [int(c) for c in color]
        except (ValueError, TypeError):
            raise ValidationError("Color values must be integers")
        
        for i, c in enumerate(validated_color):
            if not (InputValidator.MIN_COLOR_VALUE <= c <= InputValidator.MAX_COLOR_VALUE):
                raise ValidationError(f"Color value at index {i} ({c}) must be between {InputValidator.MIN_COLOR_VALUE} and {InputValidator.MAX_COLOR_VALUE}")
        
        return validated_color
    
    @staticmethod
    def validate_slide_index(slide_index: int, max_slides: int) -> int:
        """
        Validate slide index bounds.
        
        Args:
            slide_index: Index to validate
            max_slides: Maximum number of slides
            
        Returns:
            Validated slide index
            
        Raises:
            ValidationError: If index is invalid
        """
        try:
            slide_index = int(slide_index)
        except (ValueError, TypeError):
            raise ValidationError("Slide index must be an integer")
        
        if slide_index < 0:
            raise ValidationError("Slide index cannot be negative")
        
        if slide_index >= max_slides:
            raise ValidationError(f"Slide index ({slide_index}) exceeds available slides (0-{max_slides-1})")
        
        return slide_index
    
    @staticmethod
    def validate_chart_data(data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Validate chart data structure and content.
        
        Args:
            data: Chart data dictionary
            
        Returns:
            Validated chart data
            
        Raises:
            ValidationError: If chart data is invalid
        """
        if not isinstance(data, dict):
            raise ValidationError("Chart data must be a dictionary")
        
        if 'categories' not in data:
            raise ValidationError("Chart data must include 'categories'")
        
        if 'series' not in data:
            raise ValidationError("Chart data must include 'series'")
        
        categories = data['categories']
        if not isinstance(categories, list) or len(categories) == 0:
            raise ValidationError("Categories must be a non-empty list")
        
        series = data['series']
        if not isinstance(series, list) or len(series) == 0:
            raise ValidationError("Series must be a non-empty list")
        
        for i, series_item in enumerate(series):
            if not isinstance(series_item, dict):
                raise ValidationError(f"Series item {i} must be a dictionary")
            
            if 'name' not in series_item or 'values' not in series_item:
                raise ValidationError(f"Series item {i} must have 'name' and 'values'")
            
            values = series_item['values']
            if not isinstance(values, list):
                raise ValidationError(f"Series item {i} values must be a list")
            
            if len(values) != len(categories):
                raise ValidationError(f"Series item {i} values count ({len(values)}) must match categories count ({len(categories)})")
            
            # Validate numeric values
            try:
                [float(v) for v in values]
            except (ValueError, TypeError):
                raise ValidationError(f"Series item {i} values must be numeric")
        
        return data
    
    @staticmethod
    def validate_table_data(data: List[List[str]], rows: int, cols: int) -> List[List[str]]:
        """
        Validate table data structure and dimensions.
        
        Args:
            data: Table data as list of lists
            rows: Expected number of rows
            cols: Expected number of columns
            
        Returns:
            Validated table data
            
        Raises:
            ValidationError: If table data is invalid
        """
        if not isinstance(data, list):
            raise ValidationError("Table data must be a list of lists")
        
        if len(data) != rows:
            raise ValidationError(f"Table data rows ({len(data)}) must match specified rows ({rows})")
        
        for i, row in enumerate(data):
            if not isinstance(row, list):
                raise ValidationError(f"Table row {i} must be a list")
            
            if len(row) != cols:
                raise ValidationError(f"Table row {i} columns ({len(row)}) must match specified columns ({cols})")
            
            # Convert all values to strings and validate length
            for j, cell in enumerate(row):
                cell_str = str(cell)
                if len(cell_str) > 1000:  # Reasonable limit for table cells
                    raise ValidationError(f"Table cell [{i}][{j}] text too long (max 1000 characters)")
        
        return data


# Global validator instance
validator = InputValidator()