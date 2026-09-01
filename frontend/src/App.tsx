import { BrowserRouter, Routes, Route, Link } from 'react-router-dom';
import { VisitorBookingPage } from './pages/VisitorBookingPage';
import { AdminDashboardPage } from './pages/AdminDashboardPage';

export function App() {
  return (
    <BrowserRouter>
      <nav className="app-nav">
        <Link to="/">Visitor Booking</Link> | <Link to="/admin">Admin Dashboard</Link>
      </nav>
      <Routes>
        <Route path="/" element={<VisitorBookingPage />} />
        <Route path="/admin" element={<AdminDashboardPage />} />
      </Routes>
    </BrowserRouter>
  );
}
