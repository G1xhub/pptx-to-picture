"""
Image Converter
Converts between image formats using Pillow.
"""

from pathlib import Path
from typing import Optional
import time
import logging

from PIL import Image

from .base import (
    BaseConverter,
    ConversionCategory,
    ConversionOptions,
    ConversionResult
)

logger = logging.getLogger(__name__)


class ImageConverter(BaseConverter):
    """
    Converter for image files using Pillow.
    Supports PNG, JPG, WebP, BMP, GIF, TIFF, ICO.
    """
    
    # Format mappings for Pillow
    FORMAT_MAP = {
        'jpg': 'JPEG',
        'jpeg': 'JPEG',
        'png': 'PNG',
        'webp': 'WEBP',
        'bmp': 'BMP',
        'gif': 'GIF',
        'tiff': 'TIFF',
        'tif': 'TIFF',
        'ico': 'ICO',
    }
    
    @property
    def category(self) -> ConversionCategory:
        return ConversionCategory.IMAGE
    
    @property
    def name(self) -> str:
        return "Image Converter (Pillow)"
    
    @property
    def supported_input_formats(self) -> list[str]:
        return ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif', 'tiff', 'tif', 'ico']
    
    @property
    def supported_output_formats(self) -> list[str]:
        return ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'gif', 'tiff', 'ico']
    
    def validate_dependencies(self) -> tuple[bool, str]:
        """Pillow is a Python package, always available if imported."""
        try:
            from PIL import Image
            return True, f"Pillow {Image.__version__} available"
        except ImportError:
            return False, "Pillow not installed. Run: pip install pillow"
    
    def convert(
        self,
        input_path: Path,
        output_format: str,
        options: Optional[ConversionOptions] = None
    ) -> ConversionResult:
        """Convert an image to the specified format."""
        start_time = time.time()
        options = options or ConversionOptions()
        output_format = output_format.lower().lstrip('.')
        
        try:
            self._report_progress(0.1, f"Opening {input_path.name}")
            
            # Open image
            with Image.open(input_path) as img:
                # Convert to RGB if saving as JPEG (no alpha channel support)
                if output_format in ('jpg', 'jpeg') and img.mode in ('RGBA', 'P'):
                    self._report_progress(0.3, "Converting color mode")
                    img = img.convert('RGB')
                
                # Resize if dimensions specified
                if options.width or options.height:
                    self._report_progress(0.4, "Resizing image")
                    img = self._resize_image(img, options.width, options.height)
                
                # Determine output path
                output_path = self.get_output_path(input_path, output_format, options)
                
                # Handle overwrite
                if output_path.exists() and not options.overwrite:
                    return ConversionResult(
                        success=False,
                        input_path=input_path,
                        error_message=f"Output file already exists: {output_path}"
                    )
                
                # Create output directory if needed
                output_path.parent.mkdir(parents=True, exist_ok=True)
                
                self._report_progress(0.6, f"Saving as {output_format.upper()}")
                
                # Get Pillow format name
                pil_format = self.FORMAT_MAP.get(output_format, output_format.upper())
                
                # Save with appropriate options
                save_kwargs = self._get_save_kwargs(pil_format, options)
                img.save(output_path, pil_format, **save_kwargs)
                
                self._report_progress(1.0, "Complete")
                
                duration = time.time() - start_time
                return ConversionResult(
                    success=True,
                    input_path=input_path,
                    output_path=output_path,
                    duration_seconds=duration
                )
                
        except Exception as e:
            logger.exception(f"Image conversion failed: {input_path}")
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message=str(e),
                duration_seconds=time.time() - start_time
            )
    
    def _resize_image(
        self,
        img: Image.Image,
        width: Optional[int],
        height: Optional[int]
    ) -> Image.Image:
        """Resize image maintaining aspect ratio if only one dimension given."""
        orig_width, orig_height = img.size
        
        if width and height:
            new_size = (width, height)
        elif width:
            ratio = width / orig_width
            new_size = (width, int(orig_height * ratio))
        elif height:
            ratio = height / orig_height
            new_size = (int(orig_width * ratio), height)
        else:
            return img
        
        return img.resize(new_size, Image.Resampling.LANCZOS)
    
    def _get_save_kwargs(self, pil_format: str, options: ConversionOptions) -> dict:
        """Get format-specific save options."""
        kwargs = {}
        
        if pil_format == 'JPEG':
            kwargs['quality'] = options.quality
            kwargs['optimize'] = True
        elif pil_format == 'PNG':
            kwargs['optimize'] = True
        elif pil_format == 'WEBP':
            kwargs['quality'] = options.quality
            kwargs['method'] = 6  # Best compression
        elif pil_format == 'TIFF':
            kwargs['compression'] = 'tiff_lzw'
        
        return kwargs
