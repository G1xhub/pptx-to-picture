"""
Video Converter
Converts between video formats using FFmpeg.
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


class VideoConverter(BaseConverter):
    """
    Converter for video files using FFmpeg.
    Supports MP4, MKV, AVI, MOV, WebM, and GIF output.
    """
    
    # Codec recommendations for formats
    FORMAT_CODECS = {
        'mp4': {'video': 'libx264', 'audio': 'aac'},
        'webm': {'video': 'libvpx-vp9', 'audio': 'libopus'},
        'mkv': {'video': 'libx264', 'audio': 'aac'},
        'avi': {'video': 'mpeg4', 'audio': 'mp3'},
        'mov': {'video': 'libx264', 'audio': 'aac'},
        'gif': {'video': None, 'audio': None},
    }
    
    @property
    def category(self) -> ConversionCategory:
        return ConversionCategory.VIDEO
    
    @property
    def name(self) -> str:
        return "Video Converter (FFmpeg)"
    
    @property
    def supported_input_formats(self) -> list[str]:
        return ['mp4', 'mkv', 'avi', 'mov', 'webm', 'flv', 'wmv', 'm4v', '3gp']
    
    @property
    def supported_output_formats(self) -> list[str]:
        return ['mp4', 'webm', 'mkv', 'avi', 'mov', 'gif']
    
    def validate_dependencies(self) -> tuple[bool, str]:
        """Check if FFmpeg is available."""
        if ffmpeg_backend.is_available():
            info = ffmpeg_backend.get_media_info(Path('.'))  # Dummy call
            return True, "FFmpeg available"
        return False, "FFmpeg not found. Please install FFmpeg."
    
    def convert(
        self,
        input_path: Path,
        output_format: str,
        options: Optional[ConversionOptions] = None
    ) -> ConversionResult:
        """Convert a video file to the specified format."""
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
            
            # Handle GIF specially
            if output_format == 'gif':
                self._report_progress(0.2, "Converting to GIF")
                success, msg = ffmpeg_backend.video_to_gif(
                    input_path,
                    output_path,
                    fps=options.fps or 10,
                    width=options.width or 480
                )
            else:
                self._report_progress(0.2, f"Converting to {output_format.upper()}")
                
                # Get codec settings
                codecs = self.FORMAT_CODECS.get(output_format, {})
                video_codec = options.video_codec or codecs.get('video')
                audio_codec = options.audio_codec or codecs.get('audio')
                
                # Resolution
                resolution = None
                if options.width and options.height:
                    resolution = (options.width, options.height)
                
                # Progress callback wrapper
                def progress_wrapper(progress: float):
                    self._report_progress(0.2 + progress * 0.75, "Converting...")
                
                success, msg = ffmpeg_backend.convert_video(
                    input_path,
                    output_path,
                    video_codec=video_codec,
                    audio_codec=audio_codec,
                    video_bitrate=options.video_bitrate,
                    audio_bitrate=options.audio_bitrate,
                    resolution=resolution,
                    fps=options.fps,
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
            logger.exception(f"Video conversion failed: {input_path}")
            return ConversionResult(
                success=False,
                input_path=input_path,
                error_message=str(e),
                duration_seconds=time.time() - start_time
            )
