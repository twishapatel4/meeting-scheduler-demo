import { TranscriptSegment } from '@modules/booking/meeting-transcript.entity';

/**
 * Parses WebVTT content returned by Microsoft Teams / Graph into structured segments and plain text.
 * Example WebVTT:
 * WEBVTT
 *
 * 00:00:02.000 --> 00:00:05.000
 * <v Speaker 1>Hello, how are you?
 *
 * 00:00:06.000 --> 00:00:10.000
 * <v Speaker 2>I'm good, thank you.
 */
export function parseWebVttTranscript(vtt: string): {
  segments: TranscriptSegment[];
  plainText: string;
} {
  if (!vtt || typeof vtt !== 'string') {
    return { segments: [], plainText: '' };
  }

  const lines = vtt.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const segments: TranscriptSegment[] = [];
  const textBlocks: string[] = [];

  const timePattern = /(\d{2}:\d{2}:\d{2}\.\d{3}|\d{2}:\d{2}\.\d{3})\s*-->\s*(\d{2}:\d{2}:\d{2}\.\d{3}|\d{2}:\d{2}\.\d{3})/;
  const voiceTagPattern = /<v\s+([^>]+)>(.*)/s;

  let currentStart = '';
  let currentEnd = '';
  let currentTextLines: string[] = [];

  const flushSegment = () => {
    if (currentStart && currentEnd && currentTextLines.length > 0) {
      const combined = currentTextLines.join(' ').trim();
      let speaker = 'Unknown Speaker';
      let text = combined;

      const vMatch = combined.match(voiceTagPattern);
      if (vMatch) {
        speaker = vMatch[1].trim();
        text = vMatch[2].replace(/<\/v>/g, '').trim();
      } else {
        // Look for "Speaker Name: text" format if no <v> tag
        const colonMatch = combined.match(/^([^:]+):\s*(.+)$/);
        if (colonMatch) {
          speaker = colonMatch[1].trim();
          text = colonMatch[2].trim();
        }
      }

      // Clean remaining html tags
      text = text.replace(/<[^>]+>/g, '').trim();

      if (text) {
        segments.push({
          startTime: currentStart,
          endTime: currentEnd,
          speaker,
          text,
        });
        textBlocks.push(`${speaker}: ${text}`);
      }
    }
    currentStart = '';
    currentEnd = '';
    currentTextLines = [];
  };

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();

    // Ignore headers and identifiers
    if (line === 'WEBVTT' || line.startsWith('NOTE') || line.startsWith('STYLE')) {
      continue;
    }

    const timeMatch = line.match(timePattern);
    if (timeMatch) {
      flushSegment();
      currentStart = timeMatch[1];
      currentEnd = timeMatch[2];
    } else if (line === '') {
      flushSegment();
    } else if (currentStart) {
      currentTextLines.push(line);
    }
  }

  flushSegment();

  return {
    segments,
    plainText: textBlocks.join('\n'),
  };
}
