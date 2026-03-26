import React, { useState, useEffect, useRef } from 'react';
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';

interface Participant {
  ParticipantID: string;
  Name: string;
  Email: string;
  QRCode: string;
  TimeIn: string;
  TimeOut: string;
  TotalDuration: string;
}

const AttendanceTracker = () => {
  // --- 1. STATE & PERSISTENCE ---
  const [participants, setParticipants] = useState<Participant[]>(() => {
    const saved = localStorage.getItem('attendance_data_v2');
    return saved ? JSON.parse(saved) : [];
  });
  
  const [searchTerm, setSearchTerm] = useState("");
  const scannerRef = useRef<Html5QrcodeScanner | null>(null);

  // Auto-save to local storage whenever participants list changes
  useEffect(() => {
    localStorage.setItem('attendance_data_v2', JSON.stringify(participants));
  }, [participants]);

  // --- 2. CAMERA INITIALIZATION (BUG-FREE VERSION) ---
  useEffect(() => {
    const startScanner = () => {
      if (scannerRef.current) return; // Don't start if already running

      const scanner = new Html5QrcodeScanner(
        "reader", 
        { 
          fps: 10, 
          qrbox: { width: 250, height: 250 },
          rememberLastUsedCamera: true
        }, 
        false
      );

      scanner.render(
        (decodedText) => handleAttendance(decodedText.trim()),
        () => { /* Quiet scan errors */ }
      );

      scannerRef.current = scanner;
    };

    const timer = setTimeout(startScanner, 500);

    return () => {
      clearTimeout(timer);
      if (scannerRef.current) {
        scannerRef.current.clear().then(() => {
          scannerRef.current = null;
        }).catch(err => console.error("Scanner cleanup error:", err));
      }
    };
  }, []);

  // --- 3. TIME CALCULATIONS ---
  const getMinutes = (timeStr: string) => {
    const [time, modifier] = timeStr.split(' ');
    let [hrs, mins] = time.split(':').map(Number);
    if (modifier === 'PM' && hrs < 12) hrs += 12;
    if (modifier === 'AM' && hrs === 12) hrs = 0;
    return hrs * 60 + mins;
  };

  const calculateDiff = (start: string, end: string): string => {
    const diff = getMinutes(end) - getMinutes(start);
    if (diff <= 0) return "0m";
    const h = Math.floor(diff / 60);
    const m = diff % 60;
    return h > 0 ? `${h}h ${m}m` : `${m}m`;
  };

  // --- 4. ATTENDANCE LOGIC WITH 10-MIN GUARD ---
  const handleAttendance = (id: string) => {
    const now = new Date();
    const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });

    setParticipants((prev) => {
      const target = prev.find(p => p.ParticipantID === id || p.QRCode === id);

      if (!target) return prev; // ID not found in list

      // Scenario A: First time scanning (Time In)
      if (!target.TimeIn) {
        return prev.map(p => (p.ParticipantID === id || p.QRCode === id) ? { ...p, TimeIn: timeStr } : p);
      }

      // Scenario B: Already Timed In, attempting Time Out
      if (!target.TimeOut) {
        const minutesPassed = getMinutes(timeStr) - getMinutes(target.TimeIn);

        // THE 10-MINUTE GUARD
        if (minutesPassed < 10) {
          alert(`🚫 Access Denied: ${target.Name} only checked in ${minutesPassed} mins ago. Minimum stay is 10 mins.`);
          return prev; 
        }

        // Proceed with Time Out
        return prev.map(p => (p.ParticipantID === id || p.QRCode === id) 
          ? { ...p, TimeOut: timeStr, TotalDuration: calculateDiff(p.TimeIn, timeStr) } 
          : p
        );
      }

      // Scenario C: Already has both In and Out
      return prev;
    });
  };

  // --- 5. IMPORT/EXPORT ---
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws) as any[];
      
      const formattedData: Participant[] = data.map((item) => ({
        ParticipantID: String(item.ParticipantID || ''),
        Name: item.Name || '',
        Email: item.Email || '',
        QRCode: String(item.QRCode || ''),
        TimeIn: item.TimeIn || '',
        TimeOut: item.TimeOut || '',
        TotalDuration: item.TotalDuration || '',
      }));
      setParticipants(formattedData);
    };
    reader.readAsBinaryString(file);
  };

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(participants);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance");
    XLSX.writeFile(wb, `Final_Attendance_${new Date().toLocaleDateString()}.xlsx`);
  };

  const clearAllData = () => {
    if (window.confirm("Danger! This will delete all attendance data. Have you exported your file yet?")) {
      setParticipants([]);
      localStorage.removeItem('attendance_data_v2');
    }
  };

  const filtered = participants.filter(p => 
    p.Name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    p.ParticipantID.includes(searchTerm)
  );

  return (
    <div style={styles.container}>
      <header style={styles.header}>
        <h1>Event Attendance Pro</h1>
        <div style={styles.toolbar}>
          <input type="file" accept=".xlsx" onChange={handleFileUpload} />
          <button onClick={exportToExcel} style={styles.exportBtn}>Export Result</button>
          <button onClick={clearAllData} style={styles.clearBtn}>Reset</button>
        </div>
      </header>

      <div style={styles.mainLayout}>
        <div style={styles.scannerBox}>
          <div id="reader"></div>
          <div style={styles.statusBox}>
            <p><strong>Status:</strong> Waiting for Scan...</p>
            <small>Note: 10 min minimum stay required to Time Out.</small>
          </div>
        </div>

        <div style={styles.listBox}>
          <input 
            type="text" 
            placeholder="Search by name..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={styles.search}
          />
          <div style={styles.tableScroll}>
            <table style={styles.table}>
              <thead>
                <tr style={styles.th}>
                  <th>Name</th>
                  <th>In</th>
                  <th>Out</th>
                  <th>Duration</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((p, i) => (
                  <tr key={i} style={{ 
                    ...styles.tr, 
                    backgroundColor: p.TimeOut ? '#dcfce7' : p.TimeIn ? '#fef9c3' : 'white' 
                  }}>
                    <td>{p.Name}</td>
                    <td>{p.TimeIn || '--'}</td>
                    <td>{p.TimeOut || '--'}</td>
                    <td>{p.TotalDuration || '--'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};

// --- STYLING (Internal) ---
const styles: { [key: string]: React.CSSProperties } = {
  container: { padding: '20px', fontFamily: 'sans-serif', backgroundColor: '#f4f4f7', minHeight: '100vh' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', backgroundColor: '#fff', padding: '15px', borderRadius: '8px' },
  toolbar: { display: 'flex', gap: '10px' },
  mainLayout: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '20px' },
  scannerBox: { backgroundColor: '#fff', padding: '15px', borderRadius: '12px', boxShadow: '0 2px 5px rgba(0,0,0,0.1)' },
  listBox: { backgroundColor: '#fff', padding: '15px', borderRadius: '12px', boxShadow: '0 2px 5px rgba(0,0,0,0.1)' },
  statusBox: { marginTop: '15px', padding: '10px', borderTop: '1px solid #eee', textAlign: 'center' },
  search: { width: '100%', padding: '10px', marginBottom: '10px', borderRadius: '6px', border: '1px solid #ddd', boxSizing: 'border-box' },
  tableScroll: { maxHeight: '500px', overflowY: 'auto' },
  table: { width: '100%', borderCollapse: 'collapse' },
  th: { textAlign: 'left', background: '#f8f8f8', padding: '10px' },
  tr: { borderBottom: '1px solid #eee' },
  exportBtn: { backgroundColor: '#2563eb', color: '#fff', border: 'none', padding: '10px 15px', borderRadius: '6px', cursor: 'pointer' },
  clearBtn: { backgroundColor: '#dc2626', color: '#fff', border: 'none', padding: '10px 15px', borderRadius: '6px', cursor: 'pointer' },
};

export default AttendanceTracker;