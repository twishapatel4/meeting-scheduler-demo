import { useEffect, useState } from 'react';
import { apiGet, apiPost, apiPatch, apiDelete, staffLoginUrl } from '../api';

interface Staff {
  id: string;
  email: string;
  connected: boolean;
  accountType: string;
  organization: { id: string; domain: string };
}

interface TranscriptSegment {
  startTime: string;
  endTime: string;
  speaker: string;
  text: string;
}

interface MeetingTranscript {
  id: string;
  status: string;
  rawContent: string | null;
  plainText: string | null;
  segments: TranscriptSegment[] | null;
  errorMessage: string | null;
  availableAt: string | null;
}

interface ActionItem {
  id?: string;
  title?: string;
  text: string;
  owner?: string | null;
  dueDate?: string | null;
}

interface MeetingNote {
  title?: string;
  text?: string;
  subpoints?: string[];
}

interface MeetingAiInsight {
  id: string;
  status: string;
  summary: string | null;
  notes: MeetingNote[] | null;
  actionItems: ActionItem[] | null;
  errorMessage: string | null;
  generatedAt: string | null;
}

interface InsightsData {
  booking: {
    id: string;
    subject: string;
    status: string;
    joinUrl: string | null;
    onlineMeetingId: string | null;
    staff: { id: string; email: string } | null;
  };
  transcript: MeetingTranscript | null;
  aiInsight: MeetingAiInsight | null;
}

interface Booking {
  id: string;
  visitorName: string;
  visitorEmail: string;
  requestedStart: string;
  requestedEnd: string;
  status: string;
  staff: Staff | null;
  joinUrl: string | null;
  onlineMeetingId?: string | null;
  subject: string;
}

function formatLocal(iso: string): string {
  return new Date(iso).toLocaleString(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short'
  });
}

// Converts a UTC ISO string into a value usable by a `datetime-local` input
// (local wall-clock time, "YYYY-MM-DDTHH:mm"). Mirrors the reverse conversion
// used in VisitorBookingPage (`new Date(value).toISOString()`).
function toDatetimeLocal(iso: string): string {
  const date = new Date(iso);
  const offsetMs = date.getTimezoneOffset() * 60 * 1000;
  const local = new Date(date.getTime() - offsetMs);
  return local.toISOString().slice(0, 16);
}

export function AdminDashboardPage() {
  const [staff, setStaff] = useState<Staff[]>([]);
  const [bookings, setBookings] = useState<Booking[]>([]);
  const [rescheduleTarget, setRescheduleTarget] = useState<Booking | null>(null);
  const [rescheduleSubject, setRescheduleSubject] = useState('');
  const [rescheduleStart, setRescheduleStart] = useState('');
  const [rescheduleEnd, setRescheduleEnd] = useState('');
  const [swapTarget, setSwapTarget] = useState<Booking | null>(null);
  const [selectedSwapStaffId, setSelectedSwapStaffId] = useState('');
  const [scheduleTarget, setScheduleTarget] = useState<Booking | null>(null);
  const [selectedScheduleStaffId, setSelectedScheduleStaffId] = useState('');
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [isSavingSchedule, setIsSavingSchedule] = useState(false);
  const [isSavingReschedule, setIsSavingReschedule] = useState(false);
  const [isSavingSwap, setIsSavingSwap] = useState(false);

  // Insights Modal state
  const [insightsTarget, setInsightsTarget] = useState<Booking | null>(null);
  const [insightsData, setInsightsData] = useState<InsightsData | null>(null);
  const [isLoadingInsights, setIsLoadingInsights] = useState(false);
  const [isFetchingFromGraph, setIsFetchingFromGraph] = useState(false);
  const [insightsError, setInsightsError] = useState<string | null>(null);

  const openInsights = async (b: Booking) => {
    setInsightsTarget(b);
    setInsightsData(null);
    setInsightsError(null);
    setIsLoadingInsights(true);
    try {
      const res = await apiGet(`/api/v1/bookings/${b.id}/insights`);
      if (res.data) {
        setInsightsData(res.data);
      }
    } catch (err: any) {
      setInsightsError(err.message || 'Failed to load insights');
    } finally {
      setIsLoadingInsights(false);
    }
  };

  const syncInsightsFromGraph = async () => {
    if (!insightsTarget) return;
    setIsFetchingFromGraph(true);
    setInsightsError(null);
    try {
      const res = await apiPost(`/api/v1/bookings/${insightsTarget.id}/fetch-insights`, {});
      if (res.data) {
        setInsightsData(res.data);
      } else if (res.error) {
        setInsightsError(res.error.message || 'Error fetching insights from Microsoft Graph');
      }
    } catch (err: any) {
      setInsightsError(err.message || 'Failed to fetch insights from Microsoft Graph');
    } finally {
      setIsFetchingFromGraph(false);
    }
  };

  const refresh = async () => {
    const staffRes = await apiGet('/api/v1/staff');
    const bookingRes = await apiGet('/api/v1/bookings');
    setStaff(staffRes.data ?? []);
    setBookings(bookingRes.data ?? []);
  };

  useEffect(() => {
    refresh();
  }, []);

  const scheduleFor = async (bookingId: string, staffId: string): Promise<boolean> => {
    const result = await apiPost('/api/v1/bookings', { bookingId, staffId });
    if (!result.success) {
      setErrorMessage(result.message);
      return false;
    }
    setErrorMessage(null);
    refresh();
    return true;
  };

  const cancelBooking = async (bookingId: string) => {
    const result = await apiPost(`/api/v1/bookings/${bookingId}/cancel`, {
      comment: 'Apologies, this meeting has been cancelled.',
      cancelledBy: 'admin',
    });
    if (!result.success) {
      setErrorMessage(result.message);
      return;
    }
    setErrorMessage(null);
    refresh();
  };

  const swapHost = async (bookingId: string, newStaffId: string): Promise<boolean> => {
    const result = await apiPost(`/api/v1/bookings/${bookingId}/swap`, { newStaffId });
    if (!result.success) {
      setErrorMessage(result.message);
      return false;
    }
    setErrorMessage(null);
    refresh();
    return true;
  };

  const reschedule = (booking: Booking) => {
    setRescheduleTarget(booking);
    setRescheduleSubject(booking.subject ?? '');
    setRescheduleStart(toDatetimeLocal(booking.requestedStart));
    setRescheduleEnd(toDatetimeLocal(booking.requestedEnd));
  };

  const cancelReschedule = () => {
    setRescheduleTarget(null);
  };

  const swapCandidates = (booking: Booking) =>
    staff.filter(
      (s) =>
        s.id !== booking.staff?.id &&
        s.connected &&
        s.organization?.id === booking.staff?.organization?.id
    );

  const openSwap = (booking: Booking) => {
    const otherStaff = swapCandidates(booking);
    setSwapTarget(booking);
    setSelectedSwapStaffId(otherStaff[0]?.id ?? '');
  };

  const cancelSwap = () => {
    setSwapTarget(null);
  };

  const saveSwap = async () => {
    if (!swapTarget || !selectedSwapStaffId || isSavingSwap) return;
    setIsSavingSwap(true);
    try {
      const ok = await swapHost(swapTarget.id, selectedSwapStaffId);
      if (!ok) return;
      setSwapTarget(null);
    } finally {
      setIsSavingSwap(false);
    }
  };

  const saveReschedule = async () => {
    if (!rescheduleTarget || isSavingReschedule) return;
    setIsSavingReschedule(true);
    try {
      const result = await apiPatch(`/api/v1/bookings/${rescheduleTarget.id}`, {
        newStart: new Date(rescheduleStart).toISOString(),
        newEnd: new Date(rescheduleEnd).toISOString(),
        subject: rescheduleSubject,
      });
      if (!result.success) {
        setErrorMessage(result.message);
        return;
      }
      setErrorMessage(null);
      setRescheduleTarget(null);
      refresh();
    } finally {
      setIsSavingReschedule(false);
    }
  };

  const deleteBooking = async (bookingId: string) => {
    if (!confirm('Delete this booking? This cannot be undone.')) return;
    const result = await apiDelete(`/api/v1/bookings/${bookingId}`);
    if (!result.success) {
      setErrorMessage(result.message);
      return;
    }
    setErrorMessage(null);
    refresh();
  };

  return (
    <div className="page admin-page">
      <h1>Admin Dashboard</h1>

      <h2>Staff</h2>
      <a href={staffLoginUrl()}>+ Connect Staff Account</a>
      <ul>
        {staff.map((s) => (
          <li key={s.id}>
            {s.email} — {s.connected ? 'Connected' : 'Disconnected'} ({s.accountType}, org:{' '}
            {s.organization?.domain})
          </li>
        ))}
      </ul>

      <h2>Bookings</h2>
      {errorMessage && (
        <div className="status-message" style={{ background: '#fee2e2', color: '#991b1b' }}>
          {errorMessage}
        </div>
      )}
      <table className="bookings-table">
        <thead>
          <tr>
            <th>Visitor</th>
            <th>Time</th>
            <th>Status</th>
            <th>Staff</th>
            <th>Join URL</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody>
          {bookings.map((b) => (
            <tr key={b.id}>
              <td>
                {b.visitorName} ({b.visitorEmail})
              </td>
              <td>
                {formatLocal(b.requestedStart)} → {formatLocal(b.requestedEnd)}
              </td>
              <td>{b.status}</td>
              <td>{b.staff?.email ?? '-'}</td>
              <td>{b.joinUrl ? <a href={b.joinUrl}>Join</a> : '-'}</td>
              <td className="actions">
                {b.status === 'Requested' && staff.filter((s) => s.connected).length > 0 && (
                  <button
                    className="btn btn-schedule"
                    onClick={() => {
                      setScheduleTarget(b);
                      setSelectedScheduleStaffId(staff.filter((s) => s.connected)[0].id);
                    }}
                  >
                    Schedule
                  </button>
                )}
                {b.status !== 'Cancelled' && b.status !== 'Requested' && (
                  <>
                    <button className="btn btn-reschedule" onClick={() => reschedule(b)}>
                      Reschedule
                    </button>
                    <button className="btn btn-cancel" onClick={() => cancelBooking(b.id)}>
                      Cancel
                    </button>
                    {swapCandidates(b).length > 0 && (
                      <button className="btn btn-swap" onClick={() => openSwap(b)}>
                        Swap
                      </button>
                    )}
                    <button
                      className="btn btn-insights"
                      style={{ background: '#7c3aed', color: '#ffffff' }}
                      onClick={() => openInsights(b)}
                    >
                      Insights & Transcript
                    </button>
                  </>
                )}
                <button className="btn btn-delete" onClick={() => deleteBooking(b.id)}>
                  Delete
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>

      {rescheduleTarget && (
        <div className="modal-overlay">
          <div className="modal">
            <h3>Reschedule Meeting</h3>
            <div className="form-field">
              <label>Subject</label>
              <input
                value={rescheduleSubject}
                onChange={(e) => setRescheduleSubject(e.target.value)}
              />
            </div>
            <div className="form-field">
              <label>New Start</label>
              <input
                type="datetime-local"
                value={rescheduleStart}
                onChange={(e) => setRescheduleStart(e.target.value)}
              />
            </div>
            <div className="form-field">
              <label>New End</label>
              <input
                type="datetime-local"
                value={rescheduleEnd}
                onChange={(e) => setRescheduleEnd(e.target.value)}
              />
            </div>
            <div className="modal-actions">
              <button className="btn btn-cancel-modal" onClick={cancelReschedule}>
                Cancel
              </button>
              <button className="btn btn-save" onClick={saveReschedule} disabled={isSavingReschedule}>
                {isSavingReschedule ? 'Saving...' : 'Save'}
              </button>
            </div>
          </div>
        </div>
      )}

      {swapTarget && (
        <div className="modal-overlay">
          <div className="modal">
            <h3>Swap Host</h3>
            <div className="form-field">
              <label>New Staff</label>
              <select
                value={selectedSwapStaffId}
                onChange={(e) => setSelectedSwapStaffId(e.target.value)}
              >
                {swapCandidates(swapTarget).map((s) => (
                  <option key={s.id} value={s.id}>
                    {s.email}
                  </option>
                ))}
              </select>
              {swapCandidates(swapTarget).length === 0 && (
                <p className="form-hint">No other staff available in this organization.</p>
              )}
            </div>
            <div className="modal-actions">
              <button className="btn btn-cancel-modal" onClick={cancelSwap}>
                Cancel
              </button>
              <button
                className="btn btn-save"
                onClick={saveSwap}
                disabled={swapCandidates(swapTarget).length === 0 || isSavingSwap}
              >
                {isSavingSwap ? 'Saving...' : 'Save'}
              </button>
            </div>
          </div>
        </div>
      )}
      {scheduleTarget && (
        <div className="modal-overlay">
          <div className="modal">
            <h3>Schedule Meeting</h3>
            <div className="form-field">
              <label>Staff</label>
              <select
                value={selectedScheduleStaffId}
                onChange={(e) => setSelectedScheduleStaffId(e.target.value)}
              >
                {staff
                  .filter((s) => s.connected)
                  .map((s) => (
                    <option key={s.id} value={s.id}>
                      {s.email}
                    </option>
                  ))}
              </select>
            </div>
            <div className="modal-actions">
              <button className="btn btn-cancel-modal" onClick={() => setScheduleTarget(null)}>
                Cancel
              </button>
              <button
                className="btn btn-save"
                disabled={isSavingSchedule}
                onClick={async () => {
                  if (isSavingSchedule) return;
                  setIsSavingSchedule(true);
                  try {
                    const ok = await scheduleFor(scheduleTarget.id, selectedScheduleStaffId);
                    if (!ok) return;
                    setScheduleTarget(null);
                  } finally {
                    setIsSavingSchedule(false);
                  }
                }}
              >
                {isSavingSchedule ? 'Saving...' : 'Save'}
              </button>
            </div>
          </div>
        </div>
      )}
      {insightsTarget && (
        <div className="modal-overlay">
          <div className="modal" style={{ maxWidth: '750px', maxHeight: '85vh', overflowY: 'auto' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <h3>Meeting Insights & Transcript</h3>
              <button
                className="btn btn-cancel-modal"
                style={{ padding: '4px 8px' }}
                onClick={() => setInsightsTarget(null)}
              >
                ✕ Close
              </button>
            </div>

            <p style={{ margin: '4px 0 16px', color: '#4b5563' }}>
              <strong>Subject:</strong> {insightsTarget.subject} | <strong>Visitor:</strong>{' '}
              {insightsTarget.visitorName}
            </p>

            <div style={{ display: 'flex', gap: '12px', alignItems: 'center', marginBottom: '16px' }}>
              <button
                className="btn btn-save"
                style={{ background: '#2563eb' }}
                disabled={isFetchingFromGraph}
                onClick={syncInsightsFromGraph}
              >
                {isFetchingFromGraph ? 'Fetching from Graph...' : '↻ Fetch / Refresh from Microsoft Graph'}
              </button>
              {isFetchingFromGraph && <span style={{ fontSize: '13px', color: '#6b7280' }}>Contacting Microsoft Graph API...</span>}
            </div>

            {insightsError && (
              <div
                className="status-message"
                style={{ background: '#fee2e2', color: '#991b1b', marginBottom: '16px' }}
              >
                {insightsError}
              </div>
            )}

            {isLoadingInsights ? (
              <p>Loading insight details...</p>
            ) : (
              <div>
                {/* AI Insights Section */}
                <div
                  style={{
                    background: '#f9fafb',
                    border: '1px solid #e5e7eb',
                    borderRadius: '8px',
                    padding: '16px',
                    marginBottom: '16px',
                  }}
                >
                  <h4 style={{ margin: '0 0 8px', color: '#4338ca' }}>
                    🤖 Microsoft Native AI Insights (Copilot)
                  </h4>
                  {insightsData?.aiInsight ? (
                    <div>
                      <p style={{ margin: '0 0 8px' }}>
                        <strong>Status:</strong>{' '}
                        <span
                          style={{
                            display: 'inline-block',
                            padding: '2px 8px',
                            borderRadius: '4px',
                            fontSize: '12px',
                            background:
                              insightsData.aiInsight.status === 'completed' ? '#d1fae5' : '#fef3c7',
                            color:
                              insightsData.aiInsight.status === 'completed' ? '#065f46' : '#92400e',
                          }}
                        >
                          {insightsData.aiInsight.status.toUpperCase()}
                        </span>
                      </p>

                      {insightsData.aiInsight.errorMessage && (
                        <p style={{ color: '#dc2626', fontSize: '13px' }}>
                          {insightsData.aiInsight.errorMessage}
                        </p>
                      )}

                      {insightsData.aiInsight.summary && (
                        <div style={{ marginBottom: '12px' }}>
                          <strong>Summary:</strong>
                          <p style={{ margin: '4px 0', background: '#ffffff', padding: '8px', borderRadius: '4px', border: '1px solid #e5e7eb' }}>
                            {insightsData.aiInsight.summary}
                          </p>
                        </div>
                      )}

                      {insightsData.aiInsight.notes && insightsData.aiInsight.notes.length > 0 && (
                        <div style={{ marginBottom: '12px' }}>
                          <strong>Notes & Key Points:</strong>
                          <ul style={{ margin: '4px 0', paddingLeft: '20px' }}>
                            {insightsData.aiInsight.notes.map((n, idx) => (
                              <li key={idx}>
                                {n.title && <strong>{n.title}: </strong>}
                                {n.text}
                              </li>
                            ))}
                          </ul>
                        </div>
                      )}

                      {insightsData.aiInsight.actionItems && insightsData.aiInsight.actionItems.length > 0 && (
                        <div>
                          <strong>Action Items:</strong>
                          <ul style={{ margin: '4px 0', paddingLeft: '20px' }}>
                            {insightsData.aiInsight.actionItems.map((item, idx) => (
                              <li key={idx}>
                                <span>{item.text}</span>
                                {item.owner && (
                                  <span style={{ marginLeft: '8px', color: '#2563eb', fontWeight: 500 }}>
                                    (Owner: {item.owner})
                                  </span>
                                )}
                              </li>
                            ))}
                          </ul>
                        </div>
                      )}
                    </div>
                  ) : (
                    <p style={{ margin: 0, color: '#6b7280', fontSize: '14px' }}>
                      No AI insights fetched yet. Click "Fetch / Refresh from Microsoft Graph" above.
                    </p>
                  )}
                </div>

                {/* Transcript Section */}
                <div
                  style={{
                    background: '#f9fafb',
                    border: '1px solid #e5e7eb',
                    borderRadius: '8px',
                    padding: '16px',
                  }}
                >
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                    <h4 style={{ margin: '0 0 8px', color: '#1f2937' }}>📝 Meeting Transcript</h4>
                    {insightsData?.transcript?.rawContent && (
                      <button
                        className="btn"
                        style={{ fontSize: '12px', padding: '2px 8px', background: '#e5e7eb', color: '#374151' }}
                        onClick={() => {
                          const blob = new Blob([insightsData.transcript?.rawContent || ''], {
                            type: 'text/vtt',
                          });
                          const url = URL.createObjectURL(blob);
                          const a = document.createElement('a');
                          a.href = url;
                          a.download = `transcript-${insightsTarget.id}.vtt`;
                          a.click();
                        }}
                      >
                        ⬇ Download .VTT
                      </button>
                    )}
                  </div>

                  {insightsData?.transcript ? (
                    <div>
                      <p style={{ margin: '0 0 8px' }}>
                        <strong>Status:</strong>{' '}
                        <span
                          style={{
                            display: 'inline-block',
                            padding: '2px 8px',
                            borderRadius: '4px',
                            fontSize: '12px',
                            background:
                              insightsData.transcript.status === 'available' ? '#d1fae5' : '#fef3c7',
                            color:
                              insightsData.transcript.status === 'available' ? '#065f46' : '#92400e',
                          }}
                        >
                          {insightsData.transcript.status.toUpperCase()}
                        </span>
                      </p>

                      {insightsData.transcript.errorMessage && (
                        <p style={{ color: '#dc2626', fontSize: '13px' }}>
                          {insightsData.transcript.errorMessage}
                        </p>
                      )}

                      {insightsData.transcript.segments && insightsData.transcript.segments.length > 0 ? (
                        <div
                          style={{
                            maxHeight: '260px',
                            overflowY: 'auto',
                            background: '#ffffff',
                            padding: '12px',
                            borderRadius: '4px',
                            border: '1px solid #e5e7eb',
                            display: 'flex',
                            flexDirection: 'column',
                            gap: '8px',
                          }}
                        >
                          {insightsData.transcript.segments.map((seg, idx) => (
                            <div key={idx} style={{ fontSize: '13px' }}>
                              <span style={{ color: '#6b7280', fontSize: '11px', marginRight: '8px' }}>
                                [{seg.startTime} - {seg.endTime}]
                              </span>
                              <strong style={{ color: '#1f2937', marginRight: '6px' }}>
                                {seg.speaker}:
                              </strong>
                              <span>{seg.text}</span>
                            </div>
                          ))}
                        </div>
                      ) : (
                        insightsData.transcript.plainText && (
                          <pre
                            style={{
                              whiteSpace: 'pre-wrap',
                              fontSize: '13px',
                              background: '#ffffff',
                              padding: '12px',
                              borderRadius: '4px',
                              border: '1px solid #e5e7eb',
                            }}
                          >
                            {insightsData.transcript.plainText}
                          </pre>
                        )
                      )}
                    </div>
                  ) : (
                    <p style={{ margin: 0, color: '#6b7280', fontSize: '14px' }}>
                      No transcript fetched yet. Make sure transcription was started in Microsoft Teams, then click "Fetch / Refresh from Microsoft Graph".
                    </p>
                  )}
                </div>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}
