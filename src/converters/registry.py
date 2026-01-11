"""
Converter Registry
Central registry for all available converters.
"""

from pathlib import Path
from typing import Optional
import logging

from .base import BaseConverter, ConversionCategory

logger = logging.getLogger(__name__)


class ConverterRegistry:
    """
    Registry that manages all converter instances.
    Provides methods to find the right converter for a given file type.
    """
    
    _instance: Optional['ConverterRegistry'] = None
    
    def __new__(cls):
        """Singleton pattern to ensure single registry instance."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._converters = []
            cls._instance._initialized = False
        return cls._instance
    
    def register(self, converter: BaseConverter) -> None:
        """Register a converter instance."""
        if converter not in self._converters:
            self._converters.append(converter)
            logger.info(f"Registered converter: {converter.name}")
    
    def unregister(self, converter: BaseConverter) -> None:
        """Unregister a converter instance."""
        if converter in self._converters:
            self._converters.remove(converter)
            logger.info(f"Unregistered converter: {converter.name}")
    
    def get_converters(self, category: Optional[ConversionCategory] = None) -> list[BaseConverter]:
        """Get all converters, optionally filtered by category."""
        if category is None:
            return list(self._converters)
        return [c for c in self._converters if c.category == category]
    
    def find_converter(
        self,
        input_format: str,
        output_format: str
    ) -> Optional[BaseConverter]:
        """Find a converter that can handle the given conversion."""
        for converter in self._converters:
            if converter.can_convert(input_format, output_format):
                return converter
        return None
    
    def find_converter_for_file(self, file_path: Path) -> Optional[BaseConverter]:
        """Find a converter that can handle the input file."""
        ext = file_path.suffix.lower().lstrip('.')
        for converter in self._converters:
            if ext in converter.supported_input_formats:
                return converter
        return None
    
    def get_supported_input_formats(self) -> set[str]:
        """Get all supported input formats across all converters."""
        formats = set()
        for converter in self._converters:
            formats.update(converter.supported_input_formats)
        return formats
    
    def get_supported_output_formats(self, input_format: str) -> set[str]:
        """Get all supported output formats for a given input format."""
        formats = set()
        input_fmt = input_format.lower().lstrip('.')
        for converter in self._converters:
            if input_fmt in converter.supported_input_formats:
                formats.update(converter.supported_output_formats)
        return formats
    
    def get_available_conversions(self) -> dict[str, set[str]]:
        """Get a mapping of input formats to possible output formats."""
        conversions = {}
        for converter in self._converters:
            for input_fmt in converter.supported_input_formats:
                if input_fmt not in conversions:
                    conversions[input_fmt] = set()
                conversions[input_fmt].update(converter.supported_output_formats)
        return conversions
    
    def validate_all_dependencies(self) -> dict[str, tuple[bool, str]]:
        """Validate dependencies for all registered converters."""
        results = {}
        for converter in self._converters:
            results[converter.name] = converter.validate_dependencies()
        return results
    
    def clear(self) -> None:
        """Clear all registered converters (mainly for testing)."""
        self._converters.clear()
    
    def __len__(self) -> int:
        return len(self._converters)
    
    def __repr__(self) -> str:
        return f"ConverterRegistry({len(self._converters)} converters)"


# Global registry instance
registry = ConverterRegistry()
