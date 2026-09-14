import { parseWebVttTranscript } from './vtt-parser';

describe('parseWebVttTranscript', () => {
  it('parses standard WebVTT with voice tags', () => {
    const vtt = `WEBVTT

00:00:02.000 --> 00:00:05.000
<v Speaker 1>Hello, how are you?</v>

00:00:06.000 --> 00:00:10.000
<v Speaker 2>I'm good, thank you.</v>`;

    const result = parseWebVttTranscript(vtt);
    expect(result.segments).toHaveLength(2);
    expect(result.segments[0]).toEqual({
      startTime: '00:00:02.000',
      endTime: '00:00:05.000',
      speaker: 'Speaker 1',
      text: 'Hello, how are you?',
    });
    expect(result.segments[1]).toEqual({
      startTime: '00:00:06.000',
      endTime: '00:00:10.000',
      speaker: 'Speaker 2',
      text: "I'm good, thank you.",
    });
    expect(result.plainText).toContain('Speaker 1: Hello, how are you?');
    expect(result.plainText).toContain("Speaker 2: I'm good, thank you.");
  });

  it('handles empty or invalid string gracefully', () => {
    const result = parseWebVttTranscript('');
    expect(result.segments).toEqual([]);
    expect(result.plainText).toBe('');
  });
});
