#!/usr/bin/env python3
"""
Tests for the slide manager module.

Tests the color resolution helper function and related functionality.
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import unittest
from unittest.mock import patch, MagicMock

from slide_manager import SlideManager


class TestResolveColorWithDefault(unittest.TestCase):
    """Tests for the _resolve_color_with_default helper function."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.manager = SlideManager()
    
    @patch('slide_manager.template_manager')
    def test_returns_rgb_list_when_provided(self, mock_tm):
        """Test that RGB list is returned when provided."""
        mock_tm.resolve_color.return_value = [255, 0, 0]
        
        result = self.manager._resolve_color_with_default([255, 0, 0])
        
        self.assertEqual(result, [255, 0, 0])
    
    @patch('slide_manager.template_manager')
    def test_returns_resolved_semantic_color(self, mock_tm):
        """Test that semantic color tags are resolved."""
        mock_tm.resolve_color.return_value = [0, 128, 255]
        
        result = self.manager._resolve_color_with_default("accent")
        
        mock_tm.resolve_color.assert_called_once_with("accent")
        self.assertEqual(result, [0, 128, 255])
    
    @patch('slide_manager.template_manager')
    def test_uses_accent_1_default_when_color_is_none(self, mock_tm):
        """Test that accent_1 is used as default when no color is provided."""
        mock_tm.resolve_color.return_value = None
        mock_tm.get_default_color_settings.return_value = {
            "accent_1": (79, 129, 189),
            "text_1": (0, 0, 0)
        }
        
        result = self.manager._resolve_color_with_default(None)
        
        self.assertEqual(result, [79, 129, 189])
    
    @patch('slide_manager.template_manager')
    def test_uses_custom_default_key(self, mock_tm):
        """Test that custom default key is used when specified."""
        mock_tm.resolve_color.return_value = None
        mock_tm.get_default_color_settings.return_value = {
            "accent_1": (79, 129, 189),
            "text_1": (0, 0, 0)
        }
        
        result = self.manager._resolve_color_with_default(None, default_key="text_1")
        
        self.assertEqual(result, [0, 0, 0])
    
    @patch('slide_manager.template_manager')
    def test_returns_none_when_no_default_available(self, mock_tm):
        """Test that None is returned when no default is available."""
        mock_tm.resolve_color.return_value = None
        mock_tm.get_default_color_settings.return_value = {}
        
        result = self.manager._resolve_color_with_default(None)
        
        self.assertIsNone(result)
    
    @patch('slide_manager.template_manager')
    def test_returns_none_for_missing_default_key(self, mock_tm):
        """Test that None is returned when default key doesn't exist."""
        mock_tm.resolve_color.return_value = None
        mock_tm.get_default_color_settings.return_value = {
            "accent_1": (79, 129, 189)
        }
        
        result = self.manager._resolve_color_with_default(None, default_key="nonexistent_key")
        
        self.assertIsNone(result)
    
    @patch('slide_manager.template_manager')
    def test_converts_tuple_to_list(self, mock_tm):
        """Test that tuple color values are converted to lists."""
        mock_tm.resolve_color.return_value = None
        mock_tm.get_default_color_settings.return_value = {
            "accent_1": (100, 150, 200)  # Tuple, as returned by get_default_color_settings
        }
        
        result = self.manager._resolve_color_with_default(None)
        
        self.assertIsInstance(result, list)
        self.assertEqual(result, [100, 150, 200])
    
    @patch('slide_manager.template_manager')
    def test_does_not_use_default_when_color_resolved(self, mock_tm):
        """Test that default is not used when color is successfully resolved."""
        mock_tm.resolve_color.return_value = [255, 128, 0]
        mock_tm.get_default_color_settings.return_value = {
            "accent_1": (79, 129, 189)
        }
        
        result = self.manager._resolve_color_with_default("warning")
        
        mock_tm.get_default_color_settings.assert_not_called()
        self.assertEqual(result, [255, 128, 0])


class TestResolveColor(unittest.TestCase):
    """Tests for the _resolve_color method."""
    
    def setUp(self):
        """Set up test fixtures."""
        self.manager = SlideManager()
    
    @patch('slide_manager.template_manager')
    def test_delegates_to_template_manager(self, mock_tm):
        """Test that _resolve_color delegates to template_manager."""
        mock_tm.resolve_color.return_value = [128, 64, 32]
        
        result = self.manager._resolve_color("test_color")
        
        mock_tm.resolve_color.assert_called_once_with("test_color")
        self.assertEqual(result, [128, 64, 32])
    
    @patch('slide_manager.template_manager')
    def test_returns_none_for_none_input(self, mock_tm):
        """Test that None input returns None."""
        mock_tm.resolve_color.return_value = None
        
        result = self.manager._resolve_color(None)
        
        self.assertIsNone(result)


if __name__ == '__main__':
    unittest.main()
