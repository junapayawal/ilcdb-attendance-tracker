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
  // --- 1. STATE WITH LOCAL STORAGE PERSISTENCE ---
  const [participants, setParticipants] = useState<Participant[]>(() => {
    const saved = localStorage.getItem('attendance_data');
    return saved ? JSON.parse(saved) : [];
  });
  
  const [searchTerm, setSearchTerm] = useState("");
  const scannerRef = useRef<Html5QrcodeScanner | null>(null);

  // Sync state to LocalStorage whenever it changes
  useEffect(() => {
    localStorage.setItem('attendance_data', JSON.stringify(participants));
  }, [participants]);

  // --- 2. ROBUST SCANNER INITIALIZATION ---
  useEffect(() => {
    const startScanner = () => {
      // Prevent double-initialization
      if (scannerRef.current) return;

      const scanner = new Html5QrcodeScanner(
        "reader", 
        { 
          fps: 10, 
          qrbox: { width: 250, height: 250 },
          rememberLastUsedCamera: true,
          aspectRatio: 1.0
        }, 
        false
      );

      scanner.render(
        (decodedText) => handleAttendance(decodedText.trim()),
        () => { /* Ignore scan errors */ }
      );

      scannerRef.current = scanner;
    };

    // Small delay to ensure the DOM element #reader is ready
    const timer = setTimeout(startScanner, 500);

    return () => {
      clearTimeout(timer);
      if (scannerRef.current) {
        scannerRef.current.clear().then(() => {
          scannerRef.current = null;
        }).catch(err => console.error("Scanner cleanup error:", err));
      }
    };
  }, []); // Only runs once on mount

  // --- 3. EXCEL HELPERS ---
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
    const worksheet = XLSX.utils.json_to_sheet(participants);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");
    XLSX.writeFile(workbook, `Attendance_${new Date().toLocaleDateString()}.xlsx`);
  };

  const clearData = () => {
    if (window.confirm("Are you sure? This will delete all current progress!")) {
      setParticipants([]);
      localStorage.removeItem('attendance_data');
    }
  };

  // --- 4. ATTENDANCE LOGIC ---
  const handleAttendance = (id: string) => {
    const now = new Date();
    const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });

    setParticipants((prev) => {
      const newList = prev.map((p) => {
        if (p.ParticipantID === id || p.QRCode === id) {
          if (!p.TimeIn) return { ...p, TimeIn: timeStr };
          if (!p.TimeOut) {
            const duration = calculateDiff(p.TimeIn, timeStr);
            return { ...p, TimeOut: timeStr, TotalDuration: duration };
          }
        }
        return p;
      });
      return [...newList]; // Trigger re-render
    });
  };

  const calculateDiff = (start: string, end: string): string => {
    const parseTime = (t: string) => {
      const [time, modifier] = t.split(' ');
      let [hrs, mins] = time.split(':').map(Number);
      if (modifier === 'PM' && hrs < 12) hrs += 12;
      if (modifier === 'AM' && hrs === 12) hrs = 0;
      return hrs * 60 + mins;
    };
    const diff = parseTime(end) - parseTime(start);
    return diff > 0 ? `${Math.floor(diff / 60)}h ${diff % 60}m` : "0m";
  };

  const filteredParticipants = participants.filter(p => 
    p.Name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    p.ParticipantID.includes(searchTerm)
  );

  return (
    <div style={styles.container}>
      <header style={styles.header}>
        <div>
          <h1>Event Check-In</h1>
          <p style={{color: '#666'}}>Vercel Production Version</p>
        </div>
        <div style={styles.toolbar}>
          <input type="file" accept=".xlsx" onChange={handleFileUpload} />
          <button onClick={exportToExcel} style={styles.exportBtn}>Export Excel</button>
          <button onClick={clearData} style={styles.clearBtn}>Reset List</button>
        </div>
      </header>

      <div style={styles.mainLayout}>
        <section style={styles.scannerCard}>
          <div id="reader" style={{ width: '100%' }}></div>
          <p style={styles.hint}>Point camera at Participant QR Code</p>
        </section>

        <section style={styles.listCard}>
          <input 
            type="text" 
            placeholder="Search participants..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={styles.searchBar}
          />

          <div style={styles.tableWrapper}>
            <table style={styles.table}>
              <thead>
                <tr style={styles.tableHead}>
                  <th>Name</th>
                  <th>ID</th>
                  <th>In</th>
                  <th>Out</th>
                  <th>Duration</th>
                </tr>
              </thead>
              <tbody>
                {filteredParticipants.map((p, i) => (
                  <tr key={i} style={{ 
                    ...styles.row, 
                    backgroundColor: p.TimeOut ? '#d1fae5' : p.TimeIn ? '#fef3c7' : 'transparent' 
                  }}>
                    <td><strong>{p.Name}</strong><br/><small>{p.Email}</small></td>
                    <td>{p.ParticipantID}</td>
                    <td>{p.TimeIn || '--'}</td>
                    <td>{p.TimeOut || '--'}</td>
                    <td>{p.TotalDuration || '--'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      </div>
    </div>
  );
};

// --- STYLING ---
const styles: { [key: string]: React.CSSProperties } = {
  container: { padding: '20px', maxWidth: '1200px', margin: '0 auto', fontFamily: 'sans-serif' },
  header: { display: 'flex', justifyContent: 'space-between', flexWrap: 'wrap', gap: '20px', borderBottom: '2px solid #eee', paddingBottom: '20px' },
  toolbar: { display: 'flex', gap: '10px', alignItems: 'center' },
  mainLayout: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(350px, 1fr))', gap: '20px', marginTop: '20px' },
  scannerCard: { background: '#fff', border: '1px solid #ddd', padding: '15px', borderRadius: '12px', boxShadow: '0 4px 6px rgba(0,0,0,0.05)' },
  listCard: { background: '#fff', border: '1px solid #ddd', padding: '15px', borderRadius: '12px', boxShadow: '0 4px 6px rgba(0,0,0,0.05)' },
  searchBar: { width: '100%', padding: '12px', marginBottom: '15px', boxSizing: 'border-box', borderRadius: '8px', border: '1px solid #ccc' },
  tableWrapper: { maxHeight: '60vh', overflowY: 'auto' },
  table: { width: '100%', borderCollapse: 'collapse' },
  tableHead: { textAlign: 'left', position: 'sticky', top: 0, background: '#f8f8f8' },
  row: { borderBottom: '1px solid #eee' },
  exportBtn: { backgroundColor: '#10b981', color: 'white', border: 'none', padding: '10px 15px', borderRadius: '6px', cursor: 'pointer' },
  clearBtn: { backgroundColor: '#ef4444', color: 'white', border: 'none', padding: '10px 15px', borderRadius: '6px', cursor: 'pointer' },
  hint: { textAlign: 'center', fontSize: '0.9rem', color: '#888', marginTop: '10px' }
};

export default AttendanceTracker;