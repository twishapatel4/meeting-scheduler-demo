import { Service } from 'typedi';
import { Repository } from 'typeorm';
import { AppDataSource } from '@config/database';
import { createGraphClient } from '@shared/graph-client';
import { TokenService } from '@modules/auth/token.service';
import { Booking } from './booking.entity';
import { MeetingTranscript, TranscriptStatus } from './meeting-transcript.entity';
import { MeetingAiInsight, InsightStatus, ActionItem, MeetingNote } from './meeting-ai-insight.entity';
import { parseWebVttTranscript } from '@shared/utils/vtt-parser';
import { logger } from '@shared/utils/logger';
import { NotFoundError } from '@shared/errors/NotFoundError';
import { BadRequestError } from '@shared/errors/BadRequestError';

@Service()
export class MeetingInsightService {
  constructor(private readonly tokenService: TokenService) {}

  private get transcriptRepo(): Repository<MeetingTranscript> {
    return AppDataSource.getRepository(MeetingTranscript);
  }

  private get insightRepo(): Repository<MeetingAiInsight> {
    return AppDataSource.getRepository(MeetingAiInsight);
  }

  private get bookingRepo(): Repository<Booking> {
    return AppDataSource.getRepository(Booking);
  }

  /**
   * Resolves onlineMeetingId if missing on the booking.
   * Query: GET /me/onlineMeetings?$filter=JoinWebUrl eq '{joinUrl}'
   */
  async resolveOnlineMeetingId(booking: Booking): Promise<string | null> {
    if (booking.onlineMeetingId) {
      return booking.onlineMeetingId;
    }

    if (!booking.staff || !booking.joinUrl) {
      return null;
    }

    try {
      const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
      const graph = createGraphClient(accessToken);

      const encodedJoinUrl = encodeURIComponent(booking.joinUrl);
      const res = await graph.get(`/me/onlineMeetings?$filter=JoinWebUrl eq '${encodedJoinUrl}'`);

      const meetingObj = res.data?.value?.[0];
      if (meetingObj?.id) {
        booking.onlineMeetingId = meetingObj.id;
        await this.bookingRepo.save(booking);
        return meetingObj.id;
      }
    } catch (err: any) {
      logger.warn('Could not resolve onlineMeetingId via filter JoinWebUrl', {
        bookingId: booking.id,
        error: err.response?.data || err.message,
      });
    }

    return null;
  }

  /**
   * Fetches the transcript from Microsoft Graph:
   * 1. GET /me/onlineMeetings/{onlineMeetingId}/transcripts
   * 2. GET /me/onlineMeetings/{onlineMeetingId}/transcripts/{transcriptId}/content
   */
  async fetchTranscriptFromGraph(booking: Booking): Promise<MeetingTranscript> {
    if (!booking.staff) {
      throw new BadRequestError('Booking has no assigned staff');
    }

    const onlineMeetingId = await this.resolveOnlineMeetingId(booking);
    if (!onlineMeetingId) {
      throw new BadRequestError('No onlineMeetingId available for this booking');
    }

    let transcriptRecord = await this.transcriptRepo.findOne({
      where: { bookingId: booking.id },
    });

    if (!transcriptRecord) {
      transcriptRecord = this.transcriptRepo.create({
        booking,
        bookingId: booking.id,
        onlineMeetingId,
        status: 'processing',
      });
      await this.transcriptRepo.save(transcriptRecord);
    } else {
      transcriptRecord.status = 'processing';
      transcriptRecord.errorMessage = null;
      await this.transcriptRepo.save(transcriptRecord);
    }

    try {
      const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
      const graph = createGraphClient(accessToken);

      // List transcripts
      const listRes = await graph.get(`/me/onlineMeetings/${onlineMeetingId}/transcripts`);
      const transcriptsList = listRes.data?.value;

      if (!transcriptsList || transcriptsList.length === 0) {
        transcriptRecord.status = 'not_found';
        transcriptRecord.errorMessage = 'No transcripts found yet. Ensure transcription was started in Teams.';
        return await this.transcriptRepo.save(transcriptRecord);
      }

      // Pick the latest transcript
      const latestTranscript = transcriptsList[0];
      const transcriptId = latestTranscript.id;

      // Fetch WebVTT content
      const contentRes = await graph.get(
        `/me/onlineMeetings/${onlineMeetingId}/transcripts/${transcriptId}/content`,
        {
          headers: {
            Accept: 'text/vtt',
          },
          responseType: 'text',
        },
      );

      const rawVtt = typeof contentRes.data === 'string' ? contentRes.data : JSON.stringify(contentRes.data);
      const { segments, plainText } = parseWebVttTranscript(rawVtt);

      transcriptRecord.graphTranscriptId = transcriptId;
      transcriptRecord.rawContent = rawVtt;
      transcriptRecord.plainText = plainText;
      transcriptRecord.segments = segments;
      transcriptRecord.status = 'available';
      transcriptRecord.availableAt = new Date();
      transcriptRecord.errorMessage = null;

      return await this.transcriptRepo.save(transcriptRecord);
    } catch (err: any) {
      const status = err.response?.status;
      const graphError = err.response?.data?.error?.message || err.message;
      logger.error('Failed to fetch transcript from Microsoft Graph', {
        bookingId: booking.id,
        status,
        graphError,
      });

      transcriptRecord.status = status === 404 ? 'not_found' : 'failed';
      transcriptRecord.errorMessage = `Graph error (${status || 'unknown'}): ${graphError}`;
      return await this.transcriptRepo.save(transcriptRecord);
    }
  }

  /**
   * Fetches native AI insights (callAiInsight) from Microsoft Graph:
   * 1. GET /copilot/users/{userId}/onlineMeetings/{onlineMeetingId}/aiInsights or /me/onlineMeetings/{onlineMeetingId}/aiInsights
   * 2. GET /copilot/users/{userId}/onlineMeetings/{onlineMeetingId}/aiInsights/{id} (two-call pattern)
   */
  async fetchNativeAiInsightsFromGraph(booking: Booking): Promise<MeetingAiInsight> {
    if (!booking.staff) {
      throw new BadRequestError('Booking has no assigned staff');
    }

    const onlineMeetingId = await this.resolveOnlineMeetingId(booking);
    if (!onlineMeetingId) {
      throw new BadRequestError('No onlineMeetingId available for this booking');
    }

    let insightRecord = await this.insightRepo.findOne({
      where: { bookingId: booking.id },
    });

    if (!insightRecord) {
      insightRecord = this.insightRepo.create({
        booking,
        bookingId: booking.id,
        provider: 'microsoft_graph',
        status: 'processing',
      });
      await this.insightRepo.save(insightRecord);
    } else {
      insightRecord.status = 'processing';
      insightRecord.errorMessage = null;
      await this.insightRepo.save(insightRecord);
    }

    try {
      const accessToken = await this.tokenService.getValidAccessToken(booking.staff.id);
      const graph = createGraphClient(accessToken);

      // Attempt 1: Check native Copilot / callAiInsight list endpoint
      let listRes;
      let usedEndpointPrefix = '';
      try {
        listRes = await graph.get(`/copilot/users/${booking.staff.email}/onlineMeetings/${onlineMeetingId}/aiInsights`);
        usedEndpointPrefix = `/copilot/users/${booking.staff.email}/onlineMeetings/${onlineMeetingId}/aiInsights`;
      } catch (copilotErr: any) {
        // Fallback to /me/onlineMeetings/{onlineMeetingId}/aiInsights if copilot root differs
        try {
          listRes = await graph.get(`/me/onlineMeetings/${onlineMeetingId}/aiInsights`);
          usedEndpointPrefix = `/me/onlineMeetings/${onlineMeetingId}/aiInsights`;
        } catch (meErr) {
          throw copilotErr;
        }
      }

      const insightsList = listRes.data?.value;
      if (!insightsList || insightsList.length === 0) {
        insightRecord.status = 'not_found';
        insightRecord.errorMessage = 'No AI insights available yet from Microsoft Graph Copilot.';
        return await this.insightRepo.save(insightRecord);
      }

      const latestInsight = insightsList[0];
      const insightId = latestInsight.id;

      // Two-call retrieval: fetch full details
      const detailRes = await graph.get(`${usedEndpointPrefix}/${insightId}`);
      const detail = detailRes.data;

      const notes: MeetingNote[] = (detail.meetingNotes || []).map((n: any) => ({
        title: n.title,
        text: n.text,
        subpoints: n.subpoints,
      }));

      const actionItems: ActionItem[] = (detail.actionItems || []).map((item: any) => ({
        id: item.id,
        title: item.title,
        text: item.text || item.title,
        owner: item.ownerDisplayName || item.owner?.displayName || null,
        dueDate: item.dueDateTime || null,
      }));

      insightRecord.graphInsightId = insightId;
      insightRecord.summary = detail.summary || detail.title || 'Meeting Summary';
      insightRecord.notes = notes;
      insightRecord.actionItems = actionItems;
      insightRecord.status = 'completed';
      insightRecord.generatedAt = new Date();
      insightRecord.errorMessage = null;

      return await this.insightRepo.save(insightRecord);
    } catch (err: any) {
      const status = err.response?.status;
      const graphError = err.response?.data?.error?.message || err.message;
      logger.error('Failed to fetch AI insights from Microsoft Graph', {
        bookingId: booking.id,
        status,
        graphError,
      });

      insightRecord.status = status === 404 ? 'not_found' : 'failed';
      insightRecord.errorMessage = `Graph error (${status || 'unknown'}): ${graphError}`;
      return await this.insightRepo.save(insightRecord);
    }
  }

  /**
   * Retrieves both stored transcript & AI insights for a booking
   */
  async getInsightsForBooking(bookingId: string) {
    const booking = await this.bookingRepo.findOne({
      where: { id: bookingId },
      relations: ['staff', 'staff.organization'],
    });
    if (!booking) throw new NotFoundError('Booking not found');

    const transcript = await this.transcriptRepo.findOne({ where: { bookingId } });
    const aiInsight = await this.insightRepo.findOne({ where: { bookingId } });

    return {
      booking: {
        id: booking.id,
        subject: booking.subject,
        status: booking.status,
        joinUrl: booking.joinUrl,
        onlineMeetingId: booking.onlineMeetingId,
        staff: booking.staff ? { id: booking.staff.id, email: booking.staff.email } : null,
      },
      transcript,
      aiInsight,
    };
  }

  /**
   * On-demand trigger: Resolves meeting, fetches both transcript and AI insights
   */
  async syncInsightsForBooking(bookingId: string) {
    const booking = await this.bookingRepo.findOne({
      where: { id: bookingId },
      relations: ['staff', 'staff.organization'],
    });
    if (!booking) throw new NotFoundError('Booking not found');

    if (!booking.staff) {
      throw new BadRequestError('No staff assigned to this booking');
    }

    const transcript = await this.fetchTranscriptFromGraph(booking);
    const aiInsight = await this.fetchNativeAiInsightsFromGraph(booking);

    return {
      booking: {
        id: booking.id,
        subject: booking.subject,
        status: booking.status,
        joinUrl: booking.joinUrl,
        onlineMeetingId: booking.onlineMeetingId,
        staff: { id: booking.staff.id, email: booking.staff.email },
      },
      transcript,
      aiInsight,
    };
  }
}
