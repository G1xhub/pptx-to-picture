"""
LibreOffice Backend
Wrapper for LibreOffice headless mode for document/presentation conversion.
"""

import subprocess
import tempfile
from pathlib import Path
from typing import Optional
import logging
import shutil

from ..utils.dependency_checker import dependency_checker

logger = logging.getLogger(__name__)


class LibreOfficeBackend:
    """
    Backend for document and presentation conversion using LibreOffice headless mode.
    Cross-platform alternative to win32com.
    """
    
    # Output format mappings for LibreOffice
    FORMAT_FILTERS = {
        # Documents
        'pdf': 'pdf',
        'docx': 'docx',
        'doc': 'doc',
        'odt': 'odt',
        'rtf': 'rtf',
        'txt': 'txt',
        'html': 'html',
        # Presentations
        'pptx': 'pptx',
        'ppt': 'ppt',
        'odp': 'odp',
        # Images (for presentations)
        'png': 'png',
        'jpg': 'jpg',
        'jpeg': 'jpg',
    }
    
    def __init__(self):
        self._soffice_path: Optional[Path] = None
    
    @property
    def soffice_path(self) -> Optional[Path]:
        """Get LibreOffice soffice executable path."""
        if self._soffice_path is None:
            self._soffice_path = dependency_checker.get_libreoffice_path()
        return self._soffice_path
    
    def is_available(self) -> bool:
        """Check if LibreOffice is available."""
        return self.soffice_path is not None
    
    def convert(
        self,
        input_path: Path,
        output_format: str,
        output_dir: Optional[Path] = None
    ) -> tuple[bool, Optional[Path], str]:
        """
        Convert a document using LibreOffice.
        
        Args:
            input_path: Path to input file
            output_format: Target format (pdf, docx, png, etc.)
            output_dir: Output directory (default: same as input)
            
        Returns:
            Tuple of (success, output_path, message)
        """
        if not self.soffice_path:
            return False, None, "LibreOffice not available"
        
        output_format = output_format.lower().lstrip('.')
        output_dir = output_dir or input_path.parent
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # LibreOffice filter name
        filter_name = self._get_filter(input_path, output_format)
        
        # Build command
        cmd = [
            str(self.soffice_path),
            '--headless',
            '--convert-to', filter_name,
            '--outdir', str(output_dir),
            str(input_path)
        ]
        
        logger.info(f"Running LibreOffice: {' '.join(cmd)}")
        
        try:
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120  # 2 minute timeout
            )
            
            if result.returncode == 0:
                # Find output file
                expected_output = output_dir / f"{input_path.stem}.{output_format}"
                if expected_output.exists():
                    return True, expected_output, "Conversion successful"
                else:
                    # LibreOffice may use different extension
                    for ext in [output_format, 'pdf', 'png']:
                        alt_output = output_dir / f"{input_path.stem}.{ext}"
                        if alt_output.exists():
                            return True, alt_output, "Conversion successful"
                    return False, None, "Output file not found after conversion"
            else:
                error_msg = result.stderr or result.stdout or "Unknown error"
                return False, None, f"LibreOffice error: {error_msg}"
                
        except subprocess.TimeoutExpired:
            return False, None, "Conversion timed out (>120s)"
        except Exception as e:
            logger.exception("LibreOffice conversion failed")
            return False, None, str(e)
    
    def convert_presentation_to_images(
        self,
        input_path: Path,
        output_dir: Path,
        image_format: str = 'png'
    ) -> tuple[bool, list[Path], str]:
        """
        Convert presentation slides to individual images.
        
        Args:
            input_path: Path to presentation file (PPTX, ODP, etc.)
            output_dir: Directory to save images
            image_format: Output format (png, jpg)
            
        Returns:
            Tuple of (success, list_of_image_paths, message)
        """
        if not self.soffice_path:
            return False, [], "LibreOffice not available"
        
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # First convert to PDF, then to images
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            
            # Step 1: Convert to PDF
            success, pdf_path, msg = self.convert(input_path, 'pdf', temp_path)
            if not success:
                return False, [], f"Failed to convert to PDF: {msg}"
            
            # Step 2: Convert PDF pages to images using pdf2image
            try:
                from pdf2image import convert_from_path
                
                images = convert_from_path(pdf_path, dpi=150)
                output_paths = []
                
                for i, image in enumerate(images, start=1):
                    output_file = output_dir / f"{input_path.stem}_slide_{i}.{image_format}"
                    image.save(output_file, image_format.upper())
                    output_paths.append(output_file)
                
                return True, output_paths, f"Created {len(output_paths)} images"
                
            except ImportError:
                return False, [], "pdf2image not installed. Run: pip install pdf2image"
            except Exception as e:
                return False, [], f"Image conversion failed: {e}"
    
    def get_document_info(self, file_path: Path) -> dict:
        """Get basic document info (page count, etc.)."""
        # Convert to PDF in temp dir and count pages
        # This is a workaround since LibreOffice doesn't have a simple info command
        info = {
            'file_name': file_path.name,
            'file_size': file_path.stat().st_size if file_path.exists() else 0,
            'extension': file_path.suffix.lower(),
        }
        return info
    
    def _get_filter(self, input_path: Path, output_format: str) -> str:
        """Get LibreOffice output filter for format."""
        input_ext = input_path.suffix.lower().lstrip('.')
        
        # Special cases for image output from presentations
        if input_ext in ('pptx', 'ppt', 'odp', 'ppsx', 'pps'):
            if output_format in ('png', 'jpg', 'jpeg'):
                # Will be handled by convert_presentation_to_images
                return output_format
        
        # Default filter name is usually same as extension
        return self.FORMAT_FILTERS.get(output_format, output_format)


# Global instance
libreoffice_backend = LibreOfficeBackend()
