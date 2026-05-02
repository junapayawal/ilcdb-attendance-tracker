import React, { useState, useEffect, useRef } from 'react';
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';

// ==========================================
// 1. CONSTANTS & INTERFACES
// ==========================================
const TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1iVx4Bv2uqfPwOWzCszPDZ327kSJrsp0xks3NRFb6iWk/edit?usp=sharing";
const ADMIN_PASSWORD = "Admin";
const MIN_STAY_MINUTES = 30; // Note: prompt mentioned 5, code logic used 1. Adjust here.

interface Participant {
  ParticipantID: string;
  TrainingStartDate: string;
  Name: string;
  Email: string;
  QRCode: string;
  TimeIn: number | null;
  TimeOut: number | null;
  TotalDuration: number | null;
  hasSentEmail: string;
}

// ==========================================
// 2. UTILITY FUNCTIONS (Pure Logic)
// ==========================================
// ==========================================
// 2. UTILITY FUNCTIONS (Pure Logic)
// ==========================================
const Utils = {
  formatTime: (timestamp: number | null): string => {
    if (!timestamp) return '--';
    return new Date(timestamp).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });
  },

  calculateDiffInMinutes: (start: number, end: number): number => {
    return Math.floor((end - start) / 60000);
  },

  parseTimeBackToTimestamp: (timeStr: string | null): number | null => {
    if (!timeStr || timeStr === '--') return null;
    const [time, modifier] = timeStr.split(' ');
    let [hours, minutes] = time.split(':').map(Number);
    if (modifier === 'PM' && hours < 12) hours += 12;
    if (modifier === 'AM' && hours === 12) hours = 0;

    const date = new Date();
    date.setHours(hours, minutes, 0, 0);
    return date.getTime();
  },

  generateQRCodeString: (first: string, last: string, mi: string, count: number): string => {
    const d = new Date();
    const yyyymm = `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}`;
    const initials = `${first.charAt(0)}${last.charAt(0)}${mi.charAt(0)}`.toUpperCase();
    const sequence = String(count + 1).padStart(3, '0');
    return `${yyyymm}-${initials}-${sequence}`;
  },

  isValidEmail: (email: string): boolean => {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  },

  // NEW: Generates a self-contained beep sound
  playBeep: () => {
    try {
      const AudioContext = window.AudioContext || (window as any).webkitAudioContext;
      if (!AudioContext) return;
      const ctx = new AudioContext();
      const osc = ctx.createOscillator();
      const gainNode = ctx.createGain();

      osc.type = 'sine'; // Classic clean beep tone
      osc.frequency.setValueAtTime(800, ctx.currentTime); // 800Hz pitch
      gainNode.gain.setValueAtTime(0.1, ctx.currentTime); // Volume (10%)

      osc.connect(gainNode);
      gainNode.connect(ctx.destination);

      osc.start();
      osc.stop(ctx.currentTime + 0.15); // Play for 150ms
    } catch (err) {
      console.error("Audio playback failed:", err);
    }
  }
};

// ==========================================
// 3. CUSTOM HOOKS
// ==========================================
const useAttendanceData = () => {
  const [participants, setParticipants] = useState<Participant[]>(() => {
    const saved = localStorage.getItem('attendance_data');
    return saved ? JSON.parse(saved) : [];
  });

  useEffect(() => {
    localStorage.setItem('attendance_data', JSON.stringify(participants));
  }, [participants]);

  const addParticipant = (newParticipant: Participant) => {
    setParticipants(prev => [...prev, newParticipant]);
  };

  const updateAttendance = async (scannedValue: string, nowTs: number): Promise<{ success: boolean; msg: string; color: string }> => {
    let result = { success: false, msg: "Not Found", color: "#ef4444" };

    setParticipants((prev) => {
      const pIndex = prev.findIndex(p => p.ParticipantID === scannedValue || p.QRCode.includes(`data=${scannedValue}`));
      if (pIndex === -1) return prev;

      const p = prev[pIndex];
      const updated = [...prev];

      // Clock In
      if (!p.TimeIn) {
        updated[pIndex] = { ...p, TimeIn: nowTs };
        result = { success: true, msg: `Welcome, ${p.Name}! 👋`, color: "#2563eb" };
        return updated;
      }

      // Clock Out
      if (!p.TimeOut) {
        const minutesPassed = Utils.calculateDiffInMinutes(p.TimeIn, nowTs);
        if (minutesPassed < MIN_STAY_MINUTES) {
          result = { success: false, msg: "Already Checked In", color: "#f59e0b" };
          return prev;
        }

        updated[pIndex] = {
          ...p,
          TimeOut: nowTs,
          TotalDuration: Utils.calculateDiffInMinutes(p.TimeIn, nowTs)
        };
        result = { success: true, msg: `Goodbye, ${p.Name}! 🚗`, color: "#059669" };
        return updated;
      }

      return prev; // Already checked out
    });

    return result;
  };

  const importFromExcel = (file: File, onSuccess: () => void) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      const wb = XLSX.read(evt.target?.result, { type: 'binary', cellDates: true });
      const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]) as any[];
      const formatted = data.map((item) => ({
        ParticipantID: String(item.ParticipantID || ''),
        TrainingStartDate: String(item['Training Start Date'] || ''),
        Name: item.Name || '',
        Email: item.Email || '',
        QRCode: String(item.QRCode || ''),
        TimeIn: Utils.parseTimeBackToTimestamp(item.TimeIn),
        TimeOut: Utils.parseTimeBackToTimestamp(item.TimeOut),
        TotalDuration: item.TotalDuration ? Number(item.TotalDuration) : null,
        hasSentEmail: String(item.hasSentEmail || ''),
      }));
      setParticipants(formatted);
      onSuccess();
    };
    reader.readAsBinaryString(file);
  };

  const exportToExcel = () => {
    const exportData = participants.map(p => ({
      'ParticipantID': p.ParticipantID,
      'Training Start Date': p.TrainingStartDate,
      'Name': p.Name,
      'Email': p.Email,
      'QRCode': p.QRCode,
      'TimeIn': Utils.formatTime(p.TimeIn),
      'TimeOut': Utils.formatTime(p.TimeOut),
      'TotalDuration': p.TotalDuration,
      'hasSentEmail': p.hasSentEmail
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exportData), "Attendance");
    XLSX.writeFile(wb, `Attendance_Report_${new Date().toLocaleDateString()}.xlsx`);
  };

  const clearData = () => {
    if (window.confirm("Delete all data? This cannot be undone.")) {
      setParticipants([]);
      localStorage.removeItem('attendance_data');
    }
  };

  return { participants, addParticipant, updateAttendance, importFromExcel, exportToExcel, clearData };
};

// ==========================================
// 4. SUB-COMPONENTS
// ==========================================

const PasswordModal = ({ isOpen, onClose, onSuccess }: any) => {
  const [password, setPassword] = useState("");

  if (!isOpen) return null;

  const handleVerify = () => {
    if (password === ADMIN_PASSWORD) {
      onSuccess();
      setPassword("");
    } else {
      alert("Unauthorized access. Incorrect password.");
      setPassword("");
    }
  };

  return (
    <div style={styles.modalOverlay}>
      <div style={styles.modalContent}>
        <h2 style={{ marginTop: 0, color: '#1f2937' }}>Security Check</h2>
        <p style={{ color: '#666' }}>Enter Admin Password to proceed:</p>
        <input
          style={styles.inputField} type="password" placeholder="Password" autoFocus
          value={password} onChange={(e) => setPassword(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && handleVerify()}
        />
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '20px' }}>
          <button onClick={() => { onClose(); setPassword(""); }} style={{ ...styles.btn, backgroundColor: '#6b7280' }}>Cancel</button>
          <button onClick={handleVerify} style={{ ...styles.btn, backgroundColor: '#2563eb' }}>Verify</button>
        </div>
      </div>
    </div>
  );
};

const WalkInModal = ({ isOpen, onClose, onAddParticipant, participantsCount }: any) => {
  const [form, setForm] = useState({ lastName: '', firstName: '', mi: '', email: '' });
  const [qrCode, setQrCode] = useState("");

  if (!isOpen) return null;

  const handleSubmit = () => {
    if (!form.lastName || !form.firstName || !form.email) return alert("First Name, Last Name, and Email are required.");
    if (!Utils.isValidEmail(form.email)) return alert("Please enter a valid email address.");

    const formattedMI = form.mi ? `${form.mi.charAt(0).toUpperCase()}.` : '';
    const fullName = `${form.lastName}, ${form.firstName} ${formattedMI}`.trim();
    const qrString = Utils.generateQRCodeString(form.firstName, form.lastName, form.mi, participantsCount);

    onAddParticipant({
      ParticipantID: qrString,
      TrainingStartDate: new Date().toLocaleDateString(),
      Name: fullName,
      Email: form.email,
      QRCode: `data=${qrString}`,
      TimeIn: null, TimeOut: null, TotalDuration: null, hasSentEmail: "No"
    });
    setQrCode(qrString);
  };

  const handleClose = () => {
    setForm({ lastName: '', firstName: '', mi: '', email: '' });
    setQrCode("");
    onClose();
  };

  return (
    <div style={styles.modalOverlay}>
      <div style={styles.modalContent}>
        {!qrCode ? (
          <>
            <h2 style={{ marginTop: 0, borderBottom: '1px solid #eee', paddingBottom: '10px', color: '#1f2937' }}>Add Walk-In</h2>
            <div style={{ display: 'flex', flexDirection: 'column', gap: '10px', marginBottom: '20px' }}>
              <input style={styles.inputField} type="text" placeholder="Last Name" value={form.lastName} onChange={e => setForm({ ...form, lastName: e.target.value })} />
              <input style={styles.inputField} type="text" placeholder="First Name" value={form.firstName} onChange={e => setForm({ ...form, firstName: e.target.value })} />
              <input style={styles.inputField} type="text" placeholder="Middle Initial (e.g., A)" maxLength={1} value={form.mi} onChange={e => setForm({ ...form, mi: e.target.value })} />
              <input style={styles.inputField} type="email" placeholder="Email Address" value={form.email} onChange={e => setForm({ ...form, email: e.target.value })} />
            </div>
            <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              <button onClick={handleClose} style={{ ...styles.btn, backgroundColor: '#6b7280' }}>Cancel</button>
              <button onClick={handleSubmit} style={{ ...styles.btn, backgroundColor: '#2563eb' }}>Okay</button>
            </div>
          </>
        ) : (
          <div style={{ textAlign: 'center' }}>
            <h2 style={{ marginTop: 0, color: '#1f2937' }}>QR Generated!</h2>
            <p style={{ color: '#666', fontSize: '14px' }}>Please have the participant take a photo.</p>
            <img src={`https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=${encodeURIComponent(qrCode)}`} alt="QR" style={{ margin: '20px auto', display: 'block', border: '1px solid #ccc', padding: '10px', borderRadius: '8px' }} />
            <p style={{ fontWeight: 'bold' }}>{qrCode}</p>
            <button onClick={handleClose} style={{ ...styles.btn, backgroundColor: '#059669', width: '100%', marginTop: '20px' }}>Done</button>
          </div>
        )}
      </div>
    </div>
  );
};

// ==========================================
// 5. MAIN ORCHESTRATOR COMPONENT
// ==========================================
const AttendanceTracker = () => {
  const { participants, addParticipant, updateAttendance, importFromExcel, exportToExcel, clearData } = useAttendanceData();

  // UI State
  const [searchTerm, setSearchTerm] = useState("");
  const [statusMsg, setStatusMsg] = useState({ text: "Ready to Scan", color: "#666" });
  const [selectedAction, setSelectedAction] = useState("import");
  const [modals, setModals] = useState({ password: false, walkIn: false });

  // Refs
  const scannerRef = useRef<Html5QrcodeScanner | null>(null);
  const isProcessing = useRef(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Initialize Scanner
  useEffect(() => {
    const startScanner = () => {
      if (scannerRef.current) return;
      const scanner = new Html5QrcodeScanner(
        "reader",
        {
          fps: 10,
          qrbox: { width: 250, height: 250 },
          videoConstraints: {
            facingMode: "user" // 👈 front camera
          }
        },
        false
      );
      scanner.render((decodedText) => handleScan(decodedText.trim()), () => { });
      scannerRef.current = scanner;
    };
    const timer = setTimeout(startScanner, 500);
    return () => { clearTimeout(timer); scannerRef.current?.clear().catch(console.error); };
  }, []);

  const showStatus = (msg: string, color: string) => {
    setStatusMsg({ text: msg, color });
    setTimeout(() => setStatusMsg({ text: "Ready to Scan", color: "#666" }), 4000);
  };

  const handleScan = async (scannedValue: string) => {
    if (isProcessing.current) return;

    isProcessing.current = true;

    // Optional: Pause camera to give visual "snap" feedback
    try {
      scannerRef.current?.pause(true);
      updateAttendance(scannedValue, Date.now());
      Utils.playBeep();
    } catch (e) { }

    // Wait 3 seconds before allowing the next scan
    setTimeout(() => {
      try { scannerRef.current?.resume(); } catch (e) { }
      isProcessing.current = false;
    }, 3000);
  };

  const handleExecuteAction = () => {
    setModals(m => ({ ...m, password: false }));
    switch (selectedAction) {
      case 'add': setModals(m => ({ ...m, walkIn: true })); break;
      case 'template': window.open(TEMPLATE_URL, '_blank'); break;
      case 'import': fileInputRef.current?.click(); break;
      case 'export': exportToExcel(); break;
      case 'reset': clearData(); break;
      default: break;
    }
    setSelectedAction("import"); // Reset dropdown
  };

  const filteredParticipants = participants.filter(p =>
    p.Name?.toLowerCase().includes(searchTerm.toLowerCase()) || p.ParticipantID?.includes(searchTerm)
  );

  return (
    <div style={styles.container}>
      <PasswordModal
        isOpen={modals.password}
        onClose={() => setModals(m => ({ ...m, password: false }))}
        onSuccess={handleExecuteAction}
      />

      <WalkInModal
        isOpen={modals.walkIn}
        onClose={() => setModals(m => ({ ...m, walkIn: false }))}
        onAddParticipant={addParticipant}
        participantsCount={participants.length}
      />

      <header style={styles.header}>
        <div>
          <h1 style={{ margin: 0, fontSize: '1.4rem' }}>QR Attendance Scanner</h1>
          <span style={{ fontSize: '12px', color: '#666' }}>Numeric Duration Engine Active</span>
        </div>

        <div style={styles.toolbar}>
          <select value={selectedAction} onChange={(e) => setSelectedAction(e.target.value)} style={styles.dropdown}>
            <option value="import">Import Excel</option>
            <option value="add">Add Participant (Walk-in)</option>
            <option value="template">Get Template</option>
            <option value="export">Export Results</option>
            <option value="reset">Reset Data</option>
          </select>

          <button onClick={() => setModals(m => ({ ...m, password: true }))} style={{ ...styles.btn, backgroundColor: '#2563eb' }}>
            Execute
          </button>

          <input type="file" accept=".xlsx, .xls" style={{ display: 'none' }} ref={fileInputRef}
            onChange={(e) => e.target.files?.[0] && importFromExcel(e.target.files[0], () => showStatus("Data Loaded", "#2563eb"))} />
        </div>
      </header>

      <main style={styles.main}>
        <div style={styles.scannerSection}>
          <div id="reader" style={{ width: '100%' }}></div>
          <div style={{ marginTop: '15px', textAlign: 'center' }}>
            <div style={{ ...styles.statusDisplay, color: statusMsg.color }}>{statusMsg.text}</div>
            <p style={{ fontSize: '11px', color: '#999', marginTop: '5px' }}>Min stay required for checkout: {MIN_STAY_MINUTES} mins</p>
          </div>
        </div>

        <div style={styles.listSection}>
          <input type="text" placeholder="Search by name or ID..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} style={styles.searchBar} />

          <table style={styles.table}>
            <thead>
              <tr style={styles.tableHeader}>
                <th>Participant ID</th><th>In</th><th>Out</th><th>Total (min)</th>
              </tr>
            </thead>
            <tbody>
              {filteredParticipants.map((p, i) => (
                <tr key={i} style={{ ...styles.row, backgroundColor: p.TimeOut ? '#dcfce7' : p.TimeIn ? '#fef9c3' : 'white' }}>
                  <td style={{ padding: '10px 5px' }}><small style={{ color: '#666' }}>{p.ParticipantID}</small><br /><strong>{p.Name}</strong></td>
                  <td>{Utils.formatTime(p.TimeIn)}</td>
                  <td>{Utils.formatTime(p.TimeOut)}</td>
                  <td style={{ fontWeight: 'bold' }}>{p.TotalDuration ?? '--'}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {filteredParticipants.length === 0 && <div style={{ textAlign: 'center', padding: '20px', color: '#999' }}>No participants found.</div>}
        </div>
      </main>
    </div>
  );
};

// ==========================================
// 6. STYLES
// ==========================================
const styles: { [key: string]: React.CSSProperties } = {
  container: { padding: '15px', fontFamily: 'system-ui, sans-serif', maxWidth: '1200px', margin: '0 auto' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #eee', paddingBottom: '15px', flexWrap: 'wrap', gap: '10px' },
  toolbar: { display: 'flex', gap: '8px', flexWrap: 'wrap', alignItems: 'center' },
  dropdown: { padding: '8px 12px', borderRadius: '6px', border: '1px solid #ccc', fontSize: '14px', cursor: 'pointer', backgroundColor: 'white', color: '#1f2937' },
  btn: { padding: '8px 16px', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer', fontSize: '13px', fontWeight: '600', minWidth: '100px' },
  main: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '20px', marginTop: '20px' },
  scannerSection: {
    background: '#fff',
    padding: '15px',
    borderRadius: '10px',
    border: '1px solid #ddd',
    boxShadow: '0 2px 4px rgba(0,0,0,0.05)',
    alignSelf: 'start'
  },
  statusDisplay: { fontSize: '1.1rem', fontWeight: 'bold', minHeight: '1.5rem' },
  listSection: { overflowX: 'auto', background: '#fff', padding: '15px', borderRadius: '10px', border: '1px solid #ddd' },
  searchBar: { width: '100%', boxSizing: 'border-box', padding: '10px', marginBottom: '15px', borderRadius: '6px', border: '1px solid #ccc' },
  table: { width: '100%', borderCollapse: 'collapse', fontSize: '14px' },
  tableHeader: { textAlign: 'left', borderBottom: '2px solid #eee', color: '#4b5563' },
  row: { borderBottom: '1px solid #f3f4f6' },
  modalOverlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.5)', zIndex: 1000, display: 'flex', justifyContent: 'center', alignItems: 'center' },
  modalContent: { backgroundColor: 'white', padding: '25px', borderRadius: '10px', width: '90%', maxWidth: '400px', boxShadow: '0 4px 6px rgba(0,0,0,0.1)' },
  inputField: { width: '100%', padding: '10px', border: '1px solid #ccc', borderRadius: '6px', boxSizing: 'border-box' }
};

export default AttendanceTracker;