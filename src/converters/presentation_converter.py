"""
Presentation Converter
Converts presentations (PPTX, ODP) to images or PDF using LibreOffice.
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
from ..backends.libreoffice_backend import libreoffice_backend

logger = logging.getLogger(__name__)


class PresentationConverter(BaseConverter):
    """
    Converter for presentation files using LibreOffice headless mode.
    Cross-platform replacement for win32com.
    """
    
    @property
    def category(self) -> ConversionCategory:
        return ConversionCategory.PRESENTATION
    
    @property
    def name(self) -> str:
        return "Presentation Converter (LibreOffice)"
    
    @property
    def supported_input_formats(self) -> list[str]:
        return ['pptx', 'ppt', 'odp', 'ppsx', 'pps']
    
    @property
    def supported_output_formats(self) -> list[str]:
        return ['pdf', 'png', 'jpg', 'odp', 'pptx']
    
    def validate_dependencies(self) -> tuple[bool, str]:
        """Check if LibreOffice is available."""
        if libreoffice_backend.is_available():
            return True, "LibreOffice available"
        return False, "LibreOffice not found. Please install LibreOffice."
    
    def convert(
        self,
        input_path: Path,
        output_format: str,
        options: Optional[ConversionOptions] = None
    ) -> ConversionResult:
        """Convert a presentation file."""
        start_time = time.time()
        options = options or ConversionOptions()
        output_format = output_format.lower().lstrip('.')
        
        if not libreoffice_backend.is_available():
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message="LibreOffice not available"
            )
        
        try:
            self._report_progress(0.1, f"Opening {input_path.name}")
            
            # Determine output directory
            output_dir = options.output_dir or input_path.parent
            output_dir.mkdir(parents=True, exist_ok=True)
            
            # Handle image output (special case - generates multiple files)
            if output_format in ('png', 'jpg', 'jpeg'):
                return self._convert_to_images(input_path, output_dir, output_format, options, start_time)
            
            # Standard conversion (PDF, PPTX, ODP)
            self._report_progress(0.3, f"Converting to {output_format.upper()}")
            
            success, output_path, msg = libreoffice_backend.convert(
                input_path,
                output_format,
                output_dir
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
            logger.exception(f"Presentation conversion failed: {input_path}")
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message=str(e),
                duration_seconds=time.time() - start_time
            )
    
    def _convert_to_images(
        self,
        input_path: Path,
        output_dir: Path,
        output_format: str,
        options: ConversionOptions,
        start_time: float
    ) -> ConversionResult:
        """Convert presentation slides to individual images."""
        self._report_progress(0.2, "Converting slides to images")
        
        success, image_paths, msg = libreoffice_backend.convert_presentation_to_images(
            input_path,
            output_dir,
            output_format
        )
        
        self._report_progress(1.0, "Complete")
        duration = time.time() - start_time
        
        if success and image_paths:
            return ConversionResult(
                success=True,
                input_path=input_path,
                output_path=image_paths[0],  # Return first image as main output
                duration_seconds=duration
            )
        else:
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message=msg,
                duration_seconds=duration
            )
