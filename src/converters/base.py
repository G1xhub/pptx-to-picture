"""
Base Converter Interface
Abstract base class that all converters must implement.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Optional, Callable
import logging

logger = logging.getLogger(__name__)


class ConversionCategory(Enum):
    """Categories of file conversions."""
    DOCUMENT = "document"
    IMAGE = "image"
    VIDEO = "video"
    AUDIO = "audio"
    PRESENTATION = "presentation"
    EBOOK = "ebook"


@dataclass
class ConversionOptions:
    """Options for file conversion."""
    # General
    output_dir: Optional[Path] = None
    overwrite: bool = False
    
    # Image options
    quality: int = 95  # 1-100 for lossy formats
    width: Optional[int] = None
    height: Optional[int] = None
    dpi: int = 300
    
    # Video options
    video_codec: Optional[str] = None
    video_bitrate: Optional[str] = None
    fps: Optional[int] = None
    
    # Audio options
    audio_codec: Optional[str] = None
    audio_bitrate: Optional[str] = None
    sample_rate: Optional[int] = None
    
    # Document options
    page_range: Optional[str] = None  # e.g., "1-5,7,10"
    
    # Extra options as dict
    extra: dict = field(default_factory=dict)


@dataclass
class ConversionResult:
    """Result of a conversion operation."""
    success: bool
    input_path: Path
    output_path: Optional[Path] = None
    error_message: Optional[str] = None
    duration_seconds: float = 0.0
    
    def __str__(self) -> str:
        if self.success:
            return f"✓ {self.input_path.name} → {self.output_path.name}"
        return f"✗ {self.input_path.name}: {self.error_message}"


class BaseConverter(ABC):
    """
    Abstract base class for all converters.
    
    Each converter handles a specific category of files (documents, images, etc.)
    and must implement the core conversion methods.
    """
    
    def __init__(self):
        self._progress_callback: Optional[Callable[[float, str], None]] = None
    
    @property
    @abstractmethod
    def category(self) -> ConversionCategory:
        """Return the category this converter handles."""
        pass
    
    @property
    @abstractmethod
    def name(self) -> str:
        """Human-readable name of the converter."""
        pass
    
    @property
    @abstractmethod
    def supported_input_formats(self) -> list[str]:
        """
        List of supported input file extensions (without dot).
        Example: ['png', 'jpg', 'webp']
        """
        pass
    
    @property
    @abstractmethod
    def supported_output_formats(self) -> list[str]:
        """
        List of supported output file extensions (without dot).
        Example: ['png', 'jpg', 'webp', 'gif']
        """
        pass
    
    @abstractmethod
    def convert(
        self,
        input_path: Path,
        output_format: str,
        options: Optional[ConversionOptions] = None
    ) -> ConversionResult:
        """
        Convert a file to the specified output format.
        
        Args:
            input_path: Path to the input file
            output_format: Target format (e.g., 'png', 'mp4')
            options: Optional conversion options
            
        Returns:
            ConversionResult with success status and output path
        """
        pass
    
    @abstractmethod
    def validate_dependencies(self) -> tuple[bool, str]:
        """
        Check if all required dependencies are available.
        
        Returns:
            Tuple of (is_valid, message)
        """
        pass
    
    def can_convert(self, input_format: str, output_format: str) -> bool:
        """Check if this converter can handle the given conversion."""
        input_fmt = input_format.lower().lstrip('.')
        output_fmt = output_format.lower().lstrip('.')
        return (
            input_fmt in self.supported_input_formats and
            output_fmt in self.supported_output_formats
        )
    
    def set_progress_callback(self, callback: Callable[[float, str], None]) -> None:
        """Set a callback for progress updates. Args: (progress 0-1, message)"""
        self._progress_callback = callback
    
    def _report_progress(self, progress: float, message: str = "") -> None:
        """Report progress to callback if set."""
        if self._progress_callback:
            self._progress_callback(min(1.0, max(0.0, progress)), message)
    
    def get_output_path(
        self,
        input_path: Path,
        output_format: str,
        options: Optional[ConversionOptions] = None
    ) -> Path:
        """Generate output path based on input path and format."""
        output_dir = options.output_dir if options and options.output_dir else input_path.parent
        output_name = f"{input_path.stem}.{output_format.lower().lstrip('.')}"
        return output_dir / output_name
    
    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(category={self.category.value})"
