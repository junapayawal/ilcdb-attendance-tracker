import React, { useState, useEffect } from 'react';
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';

// 1. Define the Data Structure based on your Excel Columns
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
  const [participants, setParticipants] = useState<Participant[]>([]);
  const [searchTerm, setSearchTerm] = useState("");

  // --- 2. EXCEL IMPORT LOGIC ---
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws) as any[];
      
      // Map Excel columns to our State, ensuring empty strings for time fields
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

  // --- 3. QR SCANNER SETUP ---
  useEffect(() => {
    const scanner = new Html5QrcodeScanner(
      "reader", 
      { fps: 10, qrbox: { width: 250, height: 250 } }, 
      false
    );

    scanner.render(
      (decodedText) => handleAttendance(decodedText.trim()),
      (error) => { /* Ignore constant scanning noise */ }
    );

    return () => {
      scanner.clear().catch(err => console.error("Scanner cleanup failed", err));
    };
  }, [participants]); // Re-sync when list updates

  // --- 4. ATTENDANCE & DURATION LOGIC ---
  const handleAttendance = (id: string) => {
    const now = new Date();
    const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });

    setParticipants((prev) =>
      prev.map((p) => {
        // Match by ParticipantID or the QRCode string itself
        if (p.ParticipantID === id || p.QRCode === id) {
          if (!p.TimeIn) {
            return { ...p, TimeIn: timeStr };
          } else if (!p.TimeOut) {
            const duration = calculateDiff(p.TimeIn, timeStr);
            return { ...p, TimeOut: timeStr, TotalDuration: duration };
          }
        }
        return p;
      })
    );
  };

  const calculateDiff = (start: string, end: string): string => {
    const parseTime = (t: string) => {
      const d = new Date();
      const [time, modifier] = t.split(' ');
      let [hrs, mins] = time.split(':').map(Number);
      if (modifier === 'PM' && hrs < 12) hrs += 12;
      if (modifier === 'AM' && hrs === 12) hrs = 0;
      return hrs * 60 + mins;
    };
    const diff = parseTime(end) - parseTime(start);
    return diff > 0 ? `${Math.floor(diff / 60)}h ${diff % 60}m` : "0m";
  };

  // --- 5. EXCEL EXPORT LOGIC ---
  const exportToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(participants);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "AttendanceResults");
    XLSX.writeFile(workbook, `Attendance_Export_${new Date().toLocaleDateString()}.xlsx`);
  };

  // Filter for search bar
  const filteredParticipants = participants.filter(p => 
    p.Name.toLowerCase().includes(searchTerm.toLowerCase()) || 
    p.ParticipantID.includes(searchTerm)
  );

  return (
    <div style={styles.container}>
      <header style={styles.header}>
        <h1>QR Attendance Tracker</h1>
        <div style={styles.toolbar}>
          <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
          <button onClick={exportToExcel} style={styles.exportBtn}>Export to Excel</button>
        </div>
      </header>

      <main style={styles.main}>
        <div style={styles.scannerSection}>
          <div id="reader" style={{ width: '100%' }}></div>
          <p style={styles.hint}>Scan a Participant's QR code to Time In / Time Out</p>
        </div>

        <div style={styles.listSection}>
          <input 
            type="text" 
            placeholder="Search by name or ID..." 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            style={styles.searchBar}
          />

          <table style={styles.table}>
            <thead>
              <tr style={styles.tableHeader}>
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
                  backgroundColor: p.TimeOut ? '#d1fae5' : p.TimeIn ? '#fef3c7' : 'white' 
                }}>
                  <td>{p.Name}<br/><small>{p.Email}</small></td>
                  <td>{p.ParticipantID}</td>
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

// Basic CSS-in-JS
const styles: { [key: string]: React.CSSProperties } = {
  container: { padding: '20px', fontFamily: 'system-ui, sans-serif', maxWidth: '1200px', margin: '0 auto' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #eee', paddingBottom: '20px' },
  toolbar: { display: 'flex', gap: '10px' },
  main: { display: 'grid', gridTemplateColumns: '1fr 2fr', gap: '30px', marginTop: '20px' },
  scannerSection: { background: '#f9f9f9', padding: '20px', borderRadius: '8px' },
  listSection: { overflowX: 'auto' },
  searchBar: { width: '100%', padding: '10px', marginBottom: '15px', borderRadius: '5px', border: '1px solid #ccc' },
  table: { width: '100%', borderCollapse: 'collapse' },
  tableHeader: { textAlign: 'left', borderBottom: '2px solid #ddd' },
  row: { borderBottom: '1px solid #eee', transition: 'background 0.3s' },
  exportBtn: { backgroundColor: '#2563eb', color: 'white', border: 'none', padding: '10px 20px', borderRadius: '5px', cursor: 'pointer' },
  hint: { fontSize: '0.8rem', color: '#666', marginTop: '10px', textAlign: 'center' }
};

export default AttendanceTracker;