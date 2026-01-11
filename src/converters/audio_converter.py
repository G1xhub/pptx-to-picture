"""
Audio Converter
Converts between audio formats and extracts audio from video files using FFmpeg.
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
from ..backends.ffmpeg_backend import ffmpeg_backend

logger = logging.getLogger(__name__)


class AudioConverter(BaseConverter):
    """
    Converter for audio files using FFmpeg.
    Also supports extracting audio from video files.
    """
    
    # Audio codec recommendations
    FORMAT_CODECS = {
        'mp3': 'libmp3lame',
        'aac': 'aac',
        'm4a': 'aac',
        'ogg': 'libvorbis',
        'flac': 'flac',
        'wav': 'pcm_s16le',
        'opus': 'libopus',
    }
    
    # Default bitrates
    FORMAT_BITRATES = {
        'mp3': '192k',
        'aac': '192k',
        'm4a': '192k',
        'ogg': '192k',
        'opus': '128k',
    }
    
    @property
    def category(self) -> ConversionCategory:
        return ConversionCategory.AUDIO
    
    @property
    def name(self) -> str:
        return "Audio Converter (FFmpeg)"
    
    @property
    def supported_input_formats(self) -> list[str]:
        # Includes video formats for audio extraction
        return [
            # Audio formats
            'mp3', 'wav', 'flac', 'ogg', 'm4a', 'aac', 'wma', 'opus',
            # Video formats (for audio extraction)
            'mp4', 'mkv', 'avi', 'mov', 'webm', 'flv'
        ]
    
    @property
    def supported_output_formats(self) -> list[str]:
        return ['mp3', 'wav', 'flac', 'ogg', 'm4a', 'aac', 'opus']
    
    def validate_dependencies(self) -> tuple[bool, str]:
        """Check if FFmpeg is available."""
        if ffmpeg_backend.is_available():
            return True, "FFmpeg available"
        return False, "FFmpeg not found. Please install FFmpeg."
    
    def convert(
        self,
        input_path: Path,
        output_format: str,
        options: Optional[ConversionOptions] = None
    ) -> ConversionResult:
        """Convert an audio file or extract audio from video."""
        start_time = time.time()
        options = options or ConversionOptions()
        output_format = output_format.lower().lstrip('.')
        
        if not ffmpeg_backend.is_available():
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message="FFmpeg not available"
            )
        
        try:
            self._report_progress(0.1, f"Preparing {input_path.name}")
            
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
            
            self._report_progress(0.2, f"Converting to {output_format.upper()}")
            
            # Get codec settings
            audio_codec = options.audio_codec or self.FORMAT_CODECS.get(output_format)
            bitrate = options.audio_bitrate or self.FORMAT_BITRATES.get(output_format)
            
            # Progress callback wrapper
            def progress_wrapper(progress: float):
                self._report_progress(0.2 + progress * 0.75, "Converting...")
            
            success, msg = ffmpeg_backend.convert_audio(
                input_path,
                output_path,
                audio_codec=audio_codec,
                bitrate=bitrate,
                sample_rate=options.sample_rate,
                progress_callback=progress_wrapper
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
            logger.exception(f"Audio conversion failed: {input_path}")
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message=str(e),
                duration_seconds=time.time() - start_time
            )
    
    def is_video_file(self, file_path: Path) -> bool:
        """Check if file is a video (for audio extraction context)."""
        video_extensions = {'mp4', 'mkv', 'avi', 'mov', 'webm', 'flv', 'wmv'}
        return file_path.suffix.lower().lstrip('.') in video_extensions
