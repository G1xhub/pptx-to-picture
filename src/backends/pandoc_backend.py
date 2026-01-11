"""
Pandoc Backend
Wrapper for Pandoc document converter.
"""

import subprocess
from pathlib import Path
from typing import Optional
import logging

from ..utils.dependency_checker import dependency_checker

logger = logging.getLogger(__name__)


class PandocBackend:
    """
    Backend for document conversion using Pandoc.
    Supports Markdown, DOCX, PDF, HTML, LaTeX, EPUB, and more.
    """
    
    # Input format mappings for Pandoc
    INPUT_FORMATS = {
        'md': 'markdown',
        'markdown': 'markdown',
        'txt': 'plain',
        'html': 'html',
        'htm': 'html',
        'docx': 'docx',
        'odt': 'odt',
        'rtf': 'rtf',
        'epub': 'epub',
        'tex': 'latex',
        'latex': 'latex',
        'rst': 'rst',
        'org': 'org',
        'json': 'json',
    }
    
    # Output format mappings for Pandoc
    OUTPUT_FORMATS = {
        'md': 'markdown',
        'markdown': 'markdown',
        'txt': 'plain',
        'html': 'html',
        'docx': 'docx',
        'odt': 'odt',
        'rtf': 'rtf',
        'epub': 'epub',
        'epub3': 'epub3',
        'pdf': 'pdf',
        'tex': 'latex',
        'latex': 'latex',
        'rst': 'rst',
        'org': 'org',
        'json': 'json',
        'pptx': 'pptx',
    }
    
    def __init__(self):
        self._pandoc_path: Optional[Path] = None
    
    @property
    def pandoc_path(self) -> Optional[Path]:
        """Get Pandoc executable path."""
        if self._pandoc_path is None:
            self._pandoc_path = dependency_checker.get_pandoc_path()
        return self._pandoc_path
    
    def is_available(self) -> bool:
        """Check if Pandoc is available."""
        return self.pandoc_path is not None
    
    def get_version(self) -> Optional[str]:
        """Get Pandoc version string."""
        if not self.pandoc_path:
            return None
        
        try:
            result = subprocess.run(
                [str(self.pandoc_path), '--version'],
                capture_output=True,
                text=True,
                timeout=10
            )
            if result.returncode == 0:
                return result.stdout.split('\n')[0]
        except Exception:
            pass
        return None
    
    def convert(
        self,
        input_path: Path,
        output_path: Path,
        input_format: Optional[str] = None,
        output_format: Optional[str] = None,
        extra_args: Optional[list[str]] = None
    ) -> tuple[bool, str]:
        """
        Convert a document using Pandoc.
        
        Args:
            input_path: Path to input file
            output_path: Path to output file
            input_format: Override input format detection
            output_format: Override output format detection
            extra_args: Additional Pandoc arguments
            
        Returns:
            Tuple of (success, message)
        """
        if not self.pandoc_path:
            return False, "Pandoc not available"
        
        # Detect formats from file extensions if not provided
        if not input_format:
            ext = input_path.suffix.lower().lstrip('.')
            input_format = self.INPUT_FORMATS.get(ext, ext)
        
        if not output_format:
            ext = output_path.suffix.lower().lstrip('.')
            output_format = self.OUTPUT_FORMATS.get(ext, ext)
        
        # Build command
        cmd = [
            str(self.pandoc_path),
            str(input_path),
            '-f', input_format,
            '-t', output_format,
            '-o', str(output_path),
            '--standalone'  # Produce standalone document
        ]
        
        # Add extra arguments
        if extra_args:
            cmd.extend(extra_args)
        
        # PDF needs a PDF engine
        if output_format == 'pdf':
            cmd.extend(['--pdf-engine=pdflatex'])
        
        logger.info(f"Running Pandoc: {' '.join(cmd)}")
        
        try:
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120
            )
            
            if result.returncode == 0:
                return True, "Conversion successful"
            else:
                error_msg = result.stderr or "Unknown error"
                return False, f"Pandoc error: {error_msg}"
                
        except subprocess.TimeoutExpired:
            return False, "Conversion timed out (>120s)"
        except Exception as e:
            logger.exception("Pandoc conversion failed")
            return False, str(e)
    
    def convert_string(
        self,
        content: str,
        input_format: str,
        output_format: str
    ) -> tuple[bool, str]:
        """
        Convert a string of content.
        
        Args:
            content: String content to convert
            input_format: Input format (e.g., 'markdown')
            output_format: Output format (e.g., 'html')
            
        Returns:
            Tuple of (success, result_string_or_error)
        """
        if not self.pandoc_path:
            return False, "Pandoc not available"
        
        input_format = self.INPUT_FORMATS.get(input_format, input_format)
        output_format = self.OUTPUT_FORMATS.get(output_format, output_format)
        
        cmd = [
            str(self.pandoc_path),
            '-f', input_format,
            '-t', output_format
        ]
        
        try:
            result = subprocess.run(
                cmd,
                input=content,
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if result.returncode == 0:
                return True, result.stdout
            else:
                return False, result.stderr or "Unknown error"
                
        except Exception as e:
            return False, str(e)
    
    def get_supported_input_formats(self) -> list[str]:
        """Get list of supported input formats."""
        return list(self.INPUT_FORMATS.keys())
    
    def get_supported_output_formats(self) -> list[str]:
        """Get list of supported output formats."""
        return list(self.OUTPUT_FORMATS.keys())


# Global instance
pandoc_backend = PandocBackend()
