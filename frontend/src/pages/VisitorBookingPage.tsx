import { useState } from 'react';
import { apiPost } from '../api';

export function VisitorBookingPage() {
  const [visitorName, setVisitorName] = useState('');
  const [visitorEmail, setVisitorEmail] = useState('');
  const [requestedStart, setRequestedStart] = useState('');
  const [requestedEnd, setRequestedEnd] = useState('');
  const [subject, setSubject] = useState('');
  const [status, setStatus] = useState<string | null>(null);

  const submit = async (e: React.FormEvent) => {
    e.preventDefault();
    const result = await apiPost('/api/v1/bookings/visitor-request', {
      visitorName,
      visitorEmail,
      requestedStart: new Date(requestedStart).toISOString(),
      requestedEnd: new Date(requestedEnd).toISOString(),
      subject: subject.trim() || undefined,
    });
    setStatus(result.success ? 'Request submitted — we will confirm shortly.' : result.message);
  };

  return (
    <div className="page visitor-page">
      <h1>Book a Meeting</h1>
      <form className="booking-form" onSubmit={submit}>
        <div className="form-field">
          <label>Name</label>
          <input value={visitorName} onChange={(e) => setVisitorName(e.target.value)} required />
        </div>
        <div className="form-field">
          <label>Email</label>
          <input
            type="email"
            value={visitorEmail}
            onChange={(e) => setVisitorEmail(e.target.value)}
            required
          />
        </div>
        <div className="form-field">
          <label>Meeting subject</label>
          <input
            value={subject}
            onChange={(e) => setSubject(e.target.value)}
            placeholder="What's this about?"
          />
        </div>
        <div className="form-field">
          <label>Preferred start</label>
          <input
            type="datetime-local"
            value={requestedStart}
            onChange={(e) => setRequestedStart(e.target.value)}
            required
          />
        </div>
        <div className="form-field">
          <label>Preferred end</label>
          <input
            type="datetime-local"
            value={requestedEnd}
            onChange={(e) => setRequestedEnd(e.target.value)}
            required
          />
        </div>
        <button className="btn btn-schedule" type="submit">
          Request Meeting
        </button>
      </form>
      {status && <p className="status-message">{status}</p>}
    </div>
  );
}
