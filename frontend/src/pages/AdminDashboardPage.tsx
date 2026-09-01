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

  const refresh = async () => {
    const staffRes = await apiGet('/api/v1/staff');
    const bookingRes = await apiGet('/api/v1/bookings');
    setStaff(staffRes.data ?? []);
    setBookings(bookingRes.data ?? []);
  };

  useEffect(() => {
    refresh();
  }, []);

  const scheduleFor = async (bookingId: string, staffId: string) => {
    await apiPost('/api/v1/bookings', { bookingId, staffId, subject: 'SaaS Intro Meeting' });
    refresh();
  };

  const cancelBooking = async (bookingId: string) => {
    await apiPost(`/api/v1/bookings/${bookingId}/cancel`, {
      comment: 'Apologies, this meeting has been cancelled.',
      cancelledBy: 'admin',
    });
    refresh();
  };

  const swapHost = async (bookingId: string, newStaffId: string) => {
    await apiPost(`/api/v1/bookings/${bookingId}/swap`, { newStaffId });
    refresh();
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

  const saveReschedule = async () => {
    if (!rescheduleTarget) return;
    await apiPatch(`/api/v1/bookings/${rescheduleTarget.id}`, {
      newStart: new Date(rescheduleStart).toISOString(),
      newEnd: new Date(rescheduleEnd).toISOString(),
      subject: rescheduleSubject,
    });
    setRescheduleTarget(null);
    refresh();
  };

  const deleteBooking = async (bookingId: string) => {
    if (!confirm('Delete this booking? This cannot be undone.')) return;
    await apiDelete(`/api/v1/bookings/${bookingId}`);
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
                {b.status === 'Requested' && staff.length > 0 && (
                  <button className="btn btn-schedule" onClick={() => scheduleFor(b.id, staff[0].id)}>
                    Schedule w/ {staff[0].email}
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
                    {staff.length > 1 && (
                      <button className="btn btn-swap" onClick={() => swapHost(b.id, staff[1].id)}>
                        Swap to {staff[1].email}
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
              <button className="btn btn-save" onClick={saveReschedule}>
                Save
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
