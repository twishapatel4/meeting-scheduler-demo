import { useEffect, useState } from 'react';
import { apiGet, apiPost, apiPatch, apiDelete, staffLoginUrl } from '../api';

interface Staff {
  id: string;
  email: string;
  connected: boolean;
  accountType: string;
  organization: { id: string; domain: string };
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
    </div>
  );
}
