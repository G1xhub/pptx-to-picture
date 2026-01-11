"""
FFmpeg Backend
Wrapper for FFmpeg command-line tool for video/audio conversion.
"""

import subprocess
import re
from pathlib import Path
from typing import Optional, Callable
import logging
import json

from ..utils.dependency_checker import dependency_checker

logger = logging.getLogger(__name__)


class FFmpegBackend:
    """
    Backend for video and audio conversion using FFmpeg.
    Provides high-level methods for common conversion tasks.
    """
    
    def __init__(self):
        self._ffmpeg_path: Optional[Path] = None
        self._ffprobe_path: Optional[Path] = None
    
    @property
    def ffmpeg_path(self) -> Optional[Path]:
        """Get FFmpeg executable path."""
        if self._ffmpeg_path is None:
            self._ffmpeg_path = dependency_checker.get_ffmpeg_path()
        return self._ffmpeg_path
    
    @property
    def ffprobe_path(self) -> Optional[Path]:
        """Get FFprobe executable path (same directory as FFmpeg)."""
        if self._ffprobe_path is None and self.ffmpeg_path:
            ffprobe = self.ffmpeg_path.parent / self.ffmpeg_path.name.replace('ffmpeg', 'ffprobe')
            if ffprobe.exists():
                self._ffprobe_path = ffprobe
        return self._ffprobe_path
    
    def is_available(self) -> bool:
        """Check if FFmpeg is available."""
        return self.ffmpeg_path is not None
    
    def get_media_info(self, file_path: Path) -> Optional[dict]:
        """Get media file information using ffprobe."""
        if not self.ffprobe_path:
            logger.warning("FFprobe not available")
            return None
        
        try:
            result = subprocess.run(
                [
                    str(self.ffprobe_path),
                    '-v', 'quiet',
                    '-print_format', 'json',
                    '-show_format',
                    '-show_streams',
                    str(file_path)
                ],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if result.returncode == 0:
                return json.loads(result.stdout)
        except Exception as e:
            logger.error(f"Failed to get media info: {e}")
        
        return None
    
    def convert_video(
        self,
        input_path: Path,
        output_path: Path,
        video_codec: Optional[str] = None,
        audio_codec: Optional[str] = None,
        video_bitrate: Optional[str] = None,
        audio_bitrate: Optional[str] = None,
        resolution: Optional[tuple[int, int]] = None,
        fps: Optional[int] = None,
        progress_callback: Optional[Callable[[float], None]] = None
    ) -> tuple[bool, str]:
        """
        Convert a video file.
        
        Returns:
            Tuple of (success, message)
        """
        if not self.ffmpeg_path:
            return False, "FFmpeg not available"
        
        # Build FFmpeg command
        cmd = [str(self.ffmpeg_path), '-i', str(input_path), '-y']
        
        # Video codec
        if video_codec:
            cmd.extend(['-c:v', video_codec])
        
        # Audio codec
        if audio_codec:
            cmd.extend(['-c:a', audio_codec])
        
        # Video bitrate
        if video_bitrate:
            cmd.extend(['-b:v', video_bitrate])
        
        # Audio bitrate
        if audio_bitrate:
            cmd.extend(['-b:a', audio_bitrate])
        
        # Resolution
        if resolution:
            cmd.extend(['-vf', f'scale={resolution[0]}:{resolution[1]}'])
        
        # FPS
        if fps:
            cmd.extend(['-r', str(fps)])
        
        # Output
        cmd.append(str(output_path))
        
        logger.info(f"Running FFmpeg: {' '.join(cmd)}")
        
        try:
            # Get duration for progress calculation
            duration = self._get_duration(input_path)
            
            # Run conversion
            process = subprocess.Popen(
                cmd,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            
            # Parse progress from stderr
            for line in process.stderr:
                if progress_callback and duration:
                    time_match = re.search(r'time=(\d+):(\d+):(\d+)', line)
                    if time_match:
                        h, m, s = map(int, time_match.groups())
                        current_time = h * 3600 + m * 60 + s
                        progress = min(current_time / duration, 1.0)
                        progress_callback(progress)
            
            process.wait()
            
            if process.returncode == 0:
                return True, "Conversion successful"
            else:
                return False, f"FFmpeg exited with code {process.returncode}"
                
        except Exception as e:
            logger.exception("FFmpeg conversion failed")
            return False, str(e)
    
    def convert_audio(
        self,
        input_path: Path,
        output_path: Path,
        audio_codec: Optional[str] = None,
        bitrate: Optional[str] = None,
        sample_rate: Optional[int] = None,
        progress_callback: Optional[Callable[[float], None]] = None
    ) -> tuple[bool, str]:
        """
        Convert an audio file.
        
        Returns:
            Tuple of (success, message)
        """
        if not self.ffmpeg_path:
            return False, "FFmpeg not available"
        
        cmd = [str(self.ffmpeg_path), '-i', str(input_path), '-y', '-vn']  # -vn removes video
        
        if audio_codec:
            cmd.extend(['-c:a', audio_codec])
        
        if bitrate:
            cmd.extend(['-b:a', bitrate])
        
        if sample_rate:
            cmd.extend(['-ar', str(sample_rate)])
        
        cmd.append(str(output_path))
        
        logger.info(f"Running FFmpeg (audio): {' '.join(cmd)}")
        
        try:
            duration = self._get_duration(input_path)
            
            process = subprocess.Popen(
                cmd,
                stderr=subprocess.PIPE,
                universal_newlines=True
            )
            
            for line in process.stderr:
                if progress_callback and duration:
                    time_match = re.search(r'time=(\d+):(\d+):(\d+)', line)
                    if time_match:
                        h, m, s = map(int, time_match.groups())
                        current_time = h * 3600 + m * 60 + s
                        progress = min(current_time / duration, 1.0)
                        progress_callback(progress)
            
            process.wait()
            
            if process.returncode == 0:
                return True, "Conversion successful"
            else:
                return False, f"FFmpeg exited with code {process.returncode}"
                
        except Exception as e:
            logger.exception("FFmpeg audio conversion failed")
            return False, str(e)
    
    def extract_audio(
        self,
        input_path: Path,
        output_path: Path,
        audio_format: str = 'mp3',
        bitrate: str = '192k'
    ) -> tuple[bool, str]:
        """Extract audio track from a video file."""
        return self.convert_audio(
            input_path,
            output_path,
            audio_codec=self._get_audio_codec(audio_format),
            bitrate=bitrate
        )
    
    def video_to_gif(
        self,
        input_path: Path,
        output_path: Path,
        fps: int = 10,
        width: int = 480
    ) -> tuple[bool, str]:
        """Convert video to animated GIF."""
        if not self.ffmpeg_path:
            return False, "FFmpeg not available"
        
        # Two-pass approach for better quality
        cmd = [
            str(self.ffmpeg_path),
            '-i', str(input_path),
            '-y',
            '-vf', f'fps={fps},scale={width}:-1:flags=lanczos',
            '-loop', '0',
            str(output_path)
        ]
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if result.returncode == 0:
                return True, "GIF created successfully"
            return False, result.stderr
        except Exception as e:
            return False, str(e)
    
    def _get_duration(self, file_path: Path) -> Optional[float]:
        """Get media duration in seconds."""
        info = self.get_media_info(file_path)
        if info and 'format' in info and 'duration' in info['format']:
            return float(info['format']['duration'])
        return None
    
    def _get_audio_codec(self, format_ext: str) -> str:
        """Get appropriate audio codec for format."""
        codec_map = {
            'mp3': 'libmp3lame',
            'aac': 'aac',
            'm4a': 'aac',
            'ogg': 'libvorbis',
            'flac': 'flac',
            'wav': 'pcm_s16le',
        }
        return codec_map.get(format_ext.lower(), 'copy')


# Global instance
ffmpeg_backend = FFmpegBackend()
