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
  const [participants, setParticipants] = useState<Participant[]>(() => {
    const saved = localStorage.getItem('attendance_data');
    return saved ? JSON.parse(saved) : [];
  });
  
  const [searchTerm, setSearchTerm] = useState("");
  const scannerRef = useRef<Html5QrcodeScanner | null>(null);
  
  // LOCK: Prevents the "Infinite Alert" loop
  const isProcessing = useRef(false);

  useEffect(() => {
    localStorage.setItem('attendance_data', JSON.stringify(participants));
  }, [participants]);

  useEffect(() => {
    const startScanner = () => {
      if (scannerRef.current) return;
      const scanner = new Html5QrcodeScanner(
        "reader", 
        { fps: 10, qrbox: { width: 250, height: 250 } }, 
        false
      );
      scanner.render(
        (decodedText) => handleAttendance(decodedText.trim()),
        () => { /* scan errors */ }
      );
      scannerRef.current = scanner;
    };
    const timer = setTimeout(startScanner, 500);
    return () => {
      clearTimeout(timer);
      if (scannerRef.current) {
        scannerRef.current.clear().then(() => {
          scannerRef.current = null;
        }).catch(err => console.error(err));
      }
    };
  }, []);

  const getMinutes = (timeStr: string) => {
    const [time, modifier] = timeStr.split(' ');
    let [hrs, mins] = time.split(':').map(Number);
    if (modifier === 'PM' && hrs < 12) hrs += 12;
    if (modifier === 'AM' && hrs === 12) hrs = 0;
    return hrs * 60 + mins;
  };

  const calculateDiff = (start: string, end: string): string => {
    const diff = getMinutes(end) - getMinutes(start);
    const h = Math.floor(diff / 60);
    const m = diff % 60;
    return h > 0 ? `${h}h ${m}m` : `${m}m`;
  };

  const handleAttendance = (scannedValue: string) => {
    // If the "Lock" is on, ignore the camera completely
    if (isProcessing.current) return;

    const now = new Date();
    const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });

    setParticipants((prev) => {
      const pIndex = prev.findIndex(p => p.ParticipantID === scannedValue || p.QRCode === scannedValue);
      if (pIndex === -1) return prev;

      const p = prev[pIndex];

      // Time In
      if (!p.TimeIn) {
        return prev.map((item, idx) => idx === pIndex ? { ...item, TimeIn: timeStr } : item);
      } 
      
      // Time Out (With Anti-Loop Logic)
      if (!p.TimeOut) {
        const minutesPassed = getMinutes(timeStr) - getMinutes(p.TimeIn);
        
        if (minutesPassed < 10) {
          // 1. ENGAGE LOCK
          isProcessing.current = true;
          
          alert(`⚠️ Access Denied\n\n${p.Name} checked in only ${minutesPassed} mins ago.\nMinimum 10 mins required.`);
          
          // 2. RELEASE LOCK AFTER 2 SECONDS (Gives user time to move phone)
          setTimeout(() => {
            isProcessing.current = false;
          }, 2000);

          return prev;
        }

        return prev.map((item, idx) => idx === pIndex 
          ? { ...item, TimeOut: timeStr, TotalDuration: calculateDiff(item.TimeIn, timeStr) } 
          : item
        );
      }
      
      return prev;
    });
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws) as any[];
      const formatted: Participant[] = data.map((item) => ({
        ParticipantID: String(item.ParticipantID || ''),
        Name: item.Name || '',
        Email: item.Email || '',
        QRCode: String(item.QRCode || ''),
        TimeIn: item.TimeIn || '',
        TimeOut: item.TimeOut || '',
        TotalDuration: item.TotalDuration || '',
      }));
      setParticipants(formatted);
    };
    reader.readAsBinaryString(file);
  };

  const exportToExcel = () => {
    const ws = XLSX.utils.json_to_sheet(participants);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance");
    XLSX.writeFile(wb, `Attendance_Report.xlsx`);
  };

  const clearData = () => {
    if (window.confirm("Delete all data? This cannot be undone.")) {
      setParticipants([]);
      localStorage.removeItem('attendance_data');
    }
  };

  const filteredParticipants = participants.filter(p => 
    p.Name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    p.ParticipantID.includes(searchTerm)
  );

  return (
    <div style={styles.container}>
      <header style={styles.header}>
        <h1 style={{margin: 0, fontSize: '1.4rem'}}>QR Attendance Tracker</h1>
        <div style={styles.toolbar}>
          <label style={{...styles.btn, backgroundColor: '#4b5563'}}>
            Import Excel
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} style={{display: 'none'}} />
          </label>
          <button onClick={exportToExcel} style={{...styles.btn, backgroundColor: '#2563eb'}}>
            Export Results
          </button>
          <button onClick={clearData} style={{...styles.btn, backgroundColor: '#dc2626'}}>
            Reset
          </button>
        </div>
      </header>

      <main style={styles.main}>
        <div style={styles.scannerSection}>
          <div id="reader" style={{ width: '100%' }}></div>
          <div style={{marginTop: '15px', textAlign: 'center'}}>
             <p style={styles.hint}>Scan QR Code to Time In / Time Out</p>
             <p style={{fontSize: '11px', color: '#ef4444', fontWeight: 'bold'}}>Min Stay: 10 Minutes</p>
          </div>
        </div>

        <div style={styles.listSection}>
          <input 
            type="text" 
            placeholder="Search list..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={styles.searchBar}
          />

          <table style={styles.table}>
            <thead>
              <tr style={styles.tableHeader}>
                <th>Participant</th>
                <th>In</th>
                <th>Out</th>
                <th>Stay</th>
              </tr>
            </thead>
            <tbody>
              {filteredParticipants.map((p, i) => (
                <tr key={i} style={{ 
                  ...styles.row, 
                  backgroundColor: p.TimeOut ? '#dcfce7' : p.TimeIn ? '#fef9c3' : 'white' 
                }}>
                  <td style={{padding: '10px 5px'}}>
                    <strong>{p.Name}</strong><br/>
                    <small style={{color: '#666'}}>{p.ParticipantID}</small>
                  </td>
                  <td>{p.TimeIn || '--'}</td>
                  <td>{p.TimeOut || '--'}</td>
                  <td>{p.TotalDuration || '--'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </main>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  container: { padding: '15px', fontFamily: 'system-ui, sans-serif', maxWidth: '1200px', margin: '0 auto' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #eee', paddingBottom: '15px', flexWrap: 'wrap', gap: '10px' },
  toolbar: { display: 'flex', gap: '8px' },
  btn: { padding: '8px 16px', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer', fontSize: '13px', fontWeight: '600', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', minWidth: '110px' },
  main: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '20px', marginTop: '20px' },
  scannerSection: { background: '#fff', padding: '15px', borderRadius: '10px', border: '1px solid #ddd', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' },
  listSection: { overflowX: 'auto', background: '#fff', padding: '15px', borderRadius: '10px', border: '1px solid #ddd' },
  searchBar: { width: '100%', padding: '10px', marginBottom: '15px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' },
  table: { width: '100%', borderCollapse: 'collapse', fontSize: '14px' },
  tableHeader: { textAlign: 'left', borderBottom: '2px solid #eee', color: '#4b5563' },
  row: { borderBottom: '1px solid #f3f4f6' },
  hint: { fontSize: '0.9rem', color: '#374151', margin: '0' }
};

export default AttendanceTracker;