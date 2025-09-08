"""
Presentation Management Module

Handles the core presentation lifecycle including creation, opening, saving,
and managing presentation state.
"""
from typing import Optional, Dict, Any
import os
from pptx import Presentation
import ppt_utils


class PresentationManager:
    """Manages PowerPoint presentations in memory."""
    
    def __init__(self):
        # In-memory storage for presentations
        self.presentations: Dict[str, Presentation] = {}
        self.current_presentation_id: Optional[str] = None
        
    def get_current_presentation(self) -> Presentation:
        """Get the current presentation object or raise an error if none is loaded."""
        if self.current_presentation_id is None or self.current_presentation_id not in self.presentations:
            raise ValueError("No presentation is currently loaded. Please create or open a presentation first.")
        return self.presentations[self.current_presentation_id]
    
    def create_presentation(self, id: Optional[str] = None) -> Dict[str, Any]:
        """Create a new PowerPoint presentation."""
        pres = ppt_utils.create_presentation()
        if id is None:
            id = f"presentation_{len(self.presentations) + 1}"
        self.presentations[id] = pres
        self.current_presentation_id = id
        return {
            "presentation_id": id,
            "message": f"Created new presentation with ID: {id}",
            "slide_count": len(pres.slides)
        }
    
    def open_presentation(self, file_path: str, id: Optional[str] = None) -> Dict[str, Any]:
        """Open an existing PowerPoint presentation from a file."""
        # Ensure file_path is in /data
        if not file_path.startswith("/data/"):
            file_path = os.path.join("/data", os.path.basename(file_path))
        if not os.path.exists(file_path):
            return {"error": f"File not found: {file_path}"}
        try:
            pres = ppt_utils.open_presentation(file_path)
        except Exception as e:
            return {"error": f"Failed to open presentation: {str(e)}"}
        if id is None:
            id = f"presentation_{len(self.presentations) + 1}"
        self.presentations[id] = pres
        self.current_presentation_id = id
        return {
            "presentation_id": id,
            "message": f"Opened presentation from {file_path} with ID: {id}",
            "slide_count": len(pres.slides)
        }
    
    def save_presentation(self, file_path: str, presentation_id: Optional[str] = None) -> Dict[str, Any]:
        """Save a presentation to a file."""
        # Allow absolute paths, otherwise default to /data
        if not os.path.isabs(file_path) and not file_path.startswith("/data/"):
            file_path = os.path.join("/data", os.path.basename(file_path))
        try:
            pres = self.get_current_presentation() if presentation_id is None else self.presentations[presentation_id]
            saved_path = ppt_utils.save_presentation(pres, file_path)
            return {"message": f"Presentation saved to {saved_path}", "file_path": saved_path}
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def get_presentation_info(self, presentation_id: Optional[str] = None) -> Dict[str, Any]:
        """Get information about a presentation."""
        try:
            pres = self.get_current_presentation() if presentation_id is None else self.presentations[presentation_id]
            info = ppt_utils.get_presentation_info(pres)
            info["presentation_id"] = self.current_presentation_id if presentation_id is None else presentation_id
            return info
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def set_core_properties(
        self,
        title: Optional[str] = None,
        subject: Optional[str] = None,
        author: Optional[str] = None,
        keywords: Optional[str] = None,
        comments: Optional[str] = None,
        presentation_id: Optional[str] = None
    ) -> Dict[str, Any]:
        """Set core document properties."""
        try:
            pres = self.get_current_presentation() if presentation_id is None else self.presentations[presentation_id]
            updated_props = ppt_utils.set_core_properties(
                pres, title=title, subject=subject, author=author, keywords=keywords, comments=comments
            )
            return {"message": "Core properties updated successfully", "core_properties": updated_props}
        except (ValueError, KeyError) as e:
            return {"error": str(e)}
    
    def get_presentation(self, presentation_id: Optional[str] = None) -> Presentation:
        """Get a presentation by ID or return current presentation."""
        if presentation_id is None:
            return self.get_current_presentation()
        if presentation_id not in self.presentations:
            raise KeyError(f"Presentation with ID '{presentation_id}' not found")
        return self.presentations[presentation_id]


# Global instance
presentation_manager = PresentationManager()