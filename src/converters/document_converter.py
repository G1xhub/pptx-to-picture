"""
Document Converter
Converts between document formats using Pandoc and LibreOffice.
"""

from pathlib import Path
from typing import Optional
import time
import logging

from .base import (
    BaseConverter,
    ConversionCategory,
    ConversionOptions,
    ConversionResult
)
from ..backends.pandoc_backend import pandoc_backend
from ..backends.libreoffice_backend import libreoffice_backend

logger = logging.getLogger(__name__)


class DocumentConverter(BaseConverter):
    """
    Converter for document files using Pandoc and LibreOffice.
    Supports Markdown, DOCX, PDF, HTML, TXT, RTF, and more.
    """
    
    # Formats that Pandoc handles best
    PANDOC_FORMATS = {'md', 'markdown', 'txt', 'html', 'htm', 'rst', 'org', 'tex', 'latex', 'epub'}
    
    # Formats that LibreOffice handles best
    LIBREOFFICE_FORMATS = {'docx', 'doc', 'odt', 'rtf'}
    
    @property
    def category(self) -> ConversionCategory:
        return ConversionCategory.DOCUMENT
    
    @property
    def name(self) -> str:
        return "Document Converter (Pandoc/LibreOffice)"
    
    @property
    def supported_input_formats(self) -> list[str]:
        return ['docx', 'doc', 'odt', 'txt', 'md', 'markdown', 'html', 'htm', 'rtf', 'epub', 'tex', 'rst']
    
    @property
    def supported_output_formats(self) -> list[str]:
        return ['docx', 'odt', 'pdf', 'txt', 'md', 'html', 'rtf', 'epub', 'tex']
    
    def validate_dependencies(self) -> tuple[bool, str]:
        """Check if Pandoc or LibreOffice is available."""
        pandoc_ok = pandoc_backend.is_available()
        libre_ok = libreoffice_backend.is_available()
        
        if pandoc_ok and libre_ok:
            return True, "Pandoc and LibreOffice available"
        elif pandoc_ok:
            return True, "Pandoc available (some formats may be limited)"
        elif libre_ok:
            return True, "LibreOffice available (some formats may be limited)"
        else:
            return False, "Neither Pandoc nor LibreOffice found"
    
    def convert(
        self,
        input_path: Path,
        output_format: str,
        options: Optional[ConversionOptions] = None
    ) -> ConversionResult:
        """Convert a document file."""
        start_time = time.time()
        options = options or ConversionOptions()
        output_format = output_format.lower().lstrip('.')
        input_ext = input_path.suffix.lower().lstrip('.')
        
        try:
            self._report_progress(0.1, f"Opening {input_path.name}")
            
            # Determine output path
            output_path = self.get_output_path(input_path, output_format, options)
            
            # Handle overwrite
            if output_path.exists() and not options.overwrite:
                return ConversionResult(
                    success=False,
                    input_path=input_path,
                    error_message=f"Output file already exists: {output_path}"
                )
            
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            self._report_progress(0.3, f"Converting to {output_format.upper()}")
            
            # Choose backend based on formats
            success, msg = self._convert_with_best_backend(
                input_path, output_path, input_ext, output_format
            )
            
            self._report_progress(1.0, "Complete")
            duration = time.time() - start_time
            
            if success:
                return ConversionResult(
                    success=True,
                    input_path=input_path,
                    output_path=output_path,
                    duration_seconds=duration
                )
            else:
                return ConversionResult(
                    success=False,
                    input_path=input_path,
                    error_message=msg,
                    duration_seconds=duration
                )
                
        except Exception as e:
            logger.exception(f"Document conversion failed: {input_path}")
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message=str(e),
                duration_seconds=time.time() - start_time
            )
    
    def _convert_with_best_backend(
        self,
        input_path: Path,
        output_path: Path,
        input_format: str,
        output_format: str
    ) -> tuple[bool, str]:
        """Choose and use the best backend for the conversion."""
        
        # Prefer Pandoc for text-based formats
        if input_format in self.PANDOC_FORMATS or output_format in self.PANDOC_FORMATS:
            if pandoc_backend.is_available():
                return pandoc_backend.convert(input_path, output_path)
        
        # Prefer LibreOffice for Office formats and PDF output
        if input_format in self.LIBREOFFICE_FORMATS or output_format in ('pdf', 'docx', 'doc', 'odt'):
            if libreoffice_backend.is_available():
                success, out_path, msg = libreoffice_backend.convert(
                    input_path, output_format, output_path.parent
                )
                return success, msg
        
        # Fallback: try Pandoc first, then LibreOffice
        if pandoc_backend.is_available():
            success, msg = pandoc_backend.convert(input_path, output_path)
            if success:
                return success, msg
        
        if libreoffice_backend.is_available():
            success, out_path, msg = libreoffice_backend.convert(
                input_path, output_format, output_path.parent
            )
            return success, msg
        
        return False, "No suitable backend available for this conversion"
