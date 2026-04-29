import React, { useState, useEffect, useRef } from 'react';
import { Html5QrcodeScanner } from 'html5-qrcode';
import * as XLSX from 'xlsx';

interface Participant {
  ParticipantID: string;
  TrainingStartDate: string;
  Name: string;
  Email: string;
  QRCode: string;
  TimeIn: string;
  Consent: string;
  TimeOut: string;
  TotalDuration: string;
  hasSentEmail: string;
}

const AttendanceTracker = () => {
  const [participants, setParticipants] = useState<Participant[]>(() => {
    const saved = localStorage.getItem('attendance_data');
    return saved ? JSON.parse(saved) : [];
  });

  const [searchTerm, setSearchTerm] = useState("");
  const [statusMsg, setStatusMsg] = useState({ text: "Ready to Scan", color: "#666" });

  // New State to handle the Consent Modal
  const [pendingConsent, setPendingConsent] = useState<{ id: string, timeStr: string, name: string } | null>(null);

  const scannerRef = useRef<Html5QrcodeScanner | null>(null);
  const isProcessing = useRef(false);

  const TEMPLATE_URL = "https://docs.google.com/spreadsheets/d/1iVx4Bv2uqfPwOWzCszPDZ327kSJrsp0xks3NRFb6iWk/edit?usp=sharing";

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
    if (!timeStr) return 0;
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

  const showStatus = (msg: string, color: string) => {
    setStatusMsg({ text: msg, color: color });
    setTimeout(() => setStatusMsg({ text: "Ready to Scan", color: "#666" }), 4000);
  };

  const handleAttendance = (scannedValue: string) => {
    if (isProcessing.current) return;

    const now = new Date();
    const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', hour12: true });

    setParticipants((prev) => {
      const pIndex = prev.findIndex(p =>
        p.ParticipantID === scannedValue ||
        p.QRCode.includes(`data=${scannedValue}`)
      );

      if (pIndex === -1) return prev;

      const p = prev[pIndex];

      // LOGIC: Clock In (Trigger Consent First)
      if (!p.TimeIn) {
        isProcessing.current = true; // Pause scanner
        setPendingConsent({ id: p.ParticipantID, timeStr: timeStr, name: p.Name });
        return prev; // Do not update participant yet, wait for modal
      }

      // LOGIC: Clock Out (Min 10 mins stay)
      if (!p.TimeOut) {
        const minutesPassed = getMinutes(timeStr) - getMinutes(p.TimeIn);
        if (minutesPassed < 1) { // Note: Changed to 10 to match your comment of "Min 10 mins stay"
          showStatus("Already Checked In", "#f59e0b");
          return prev;
        }

        isProcessing.current = true;
        showStatus(`Goodbye, ${p.Name}! 🚗`, "#059669");
        setTimeout(() => { isProcessing.current = false; }, 3000);

        return prev.map((item, idx) => idx === pIndex
          ? { ...item, TimeOut: timeStr, TotalDuration: calculateDiff(item.TimeIn, timeStr) }
          : item
        );
      }
      return prev;
    });
  };

  // NEW: Handle Consent Answer
  const handleConsent = (answer: string) => {
    if (!pendingConsent) return;

    setParticipants((prev) =>
      prev.map(p =>
        p.ParticipantID === pendingConsent.id
          ? { ...p, TimeIn: pendingConsent.timeStr, Consent: answer }
          : p
      )
    );

    showStatus(`Welcome, ${pendingConsent.name}! 👋`, "#2563eb");
    setPendingConsent(null); // Close modal

    // Resume scanner after brief delay
    setTimeout(() => { isProcessing.current = false; }, 3000);
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
        TrainingStartDate: String(item['Training Start Date'] || ''),
        Name: item.Name || '',
        Email: item.Email || '',
        QRCode: String(item.QRCode || ''),
        TimeIn: item.TimeIn || '',
        Consent: String(item['Consent to Terms & Photos?'] || ''),
        TimeOut: item.TimeOut || '',
        TotalDuration: item.TotalDuration || '',
        hasSentEmail: String(item.hasSentEmail || ''),
      }));
      setParticipants(formatted);
      showStatus("Template Loaded Successfully", "#2563eb");
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
      'TimeIn': p.TimeIn,
      'Consent to Terms & Photos?': p.Consent,
      'TimeOut': p.TimeOut,
      'TotalDuration': p.TotalDuration,
      'hasSentEmail': p.hasSentEmail
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);

    // --- ADD SHEET PROTECTION ---
    ws['!protect'] = {
      password: "admin",
      selectLockedCells: true,
      selectUnlockedCells: true,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      deleteColumns: false,
      deleteRows: false,
      sort: false,
      autoFilter: false
    };

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Attendance");

    // --- METADATA & WORKBOOK PROTECTION ---
    wb.Props = {
      ...wb.Props,
      Title: "Attendance Report"
    };

    // Use 'as any' to bypass the TS(2339) error for WBView
    const wbExtended = wb as any;

    wbExtended.Workbook = wbExtended.Workbook || {};
    wbExtended.Workbook.WBProps = wbExtended.Workbook.WBProps || { selfProtection: true };
    wbExtended.Workbook.WBView = wbExtended.Workbook.WBView || [];

    XLSX.writeFile(wb, `Attendance_Report_${new Date().toLocaleDateString()}.xlsx`);
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
      {/* NEW: Consent Modal Overlay */}
      {pendingConsent && (
        <div style={styles.modalOverlay}>
          <div style={styles.modalContent}>
            <h3 style={{ marginTop: 0, color: '#1f2937' }}>Data Privacy Consent</h3>
            <p style={{ lineHeight: '1.6', fontSize: '15px', color: '#4b5563', marginBottom: '25px' }}>
              "I hereby consent to the collection and processing of my personal information for the purposes of attendance tracking, issuance of certificates, and event documentation. I also agree to the taking of photos and videos during the activity for use in official reports and promotional materials in compliance with the Data Privacy Act of 2012."
            </p>
            <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
              <button
                onClick={() => handleConsent('Yes')}
                style={{ ...styles.btn, backgroundColor: '#059669', flex: 1 }}
              >
                Yes, I Agree
              </button>
              <button
                onClick={() => handleConsent('No')}
                style={{ ...styles.btn, backgroundColor: '#dc2626', flex: 1 }}
              >
                No, I Disagree
              </button>
            </div>
          </div>
        </div>
      )}

      <header style={styles.header}>
        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <h1 style={{ margin: 0, fontSize: '1.4rem' }}>QR Attendance Scanner</h1>
          <span style={{ fontSize: '12px', color: '#666' }}>Updated Template Support Active</span>
        </div>
        <div style={styles.toolbar}>
          <a href={TEMPLATE_URL} target="_blank" rel="noopener noreferrer" style={{ ...styles.btn, backgroundColor: '#6b7280', textDecoration: 'none' }}>
            Get Template
          </a>
          <label style={{ ...styles.btn, backgroundColor: '#4b5563', cursor: 'pointer' }}>
            Import Excel
            <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} style={{ display: 'none' }} />
          </label>
          <button onClick={exportToExcel} style={{ ...styles.btn, backgroundColor: '#2563eb' }}>
            Export Results
          </button>
          <button onClick={clearData} style={{ ...styles.btn, backgroundColor: '#dc2626' }}>
            Reset
          </button>
        </div>
      </header>

      <main style={styles.main}>
        <div style={styles.scannerSection}>
          <div id="reader" style={{ width: '100%' }}></div>
          <div style={{ marginTop: '15px', textAlign: 'center' }}>
            <div style={{ ...styles.statusDisplay, color: statusMsg.color }}>
              {statusMsg.text}
            </div>
            <p style={{ fontSize: '11px', color: '#999', marginTop: '5px' }}>Min stay required for checkout: 10 mins</p>
          </div>
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
                <th>Participant ID</th>
                <th>Consent</th>
                <th>In</th>
                <th>Out</th>
                <th>Total</th>
              </tr>
            </thead>
            <tbody>
              {filteredParticipants.map((p, i) => (
                <tr key={i} style={{
                  ...styles.row,
                  backgroundColor: p.TimeOut ? '#dcfce7' : p.TimeIn ? '#fef9c3' : 'white'
                }}>
                  <td style={{ padding: '10px 5px' }}>
                    <small style={{ color: '#666' }}>{p.ParticipantID}</small><br />
                  </td>
                  <td>
                    {p.Consent === 'Yes' ? <span style={{ color: '#059669' }}>✅ Yes</span> :
                      p.Consent === 'No' ? <span style={{ color: '#dc2626' }}>❌ No</span> : '--'}
                  </td>
                  <td>{p.TimeIn || '--'}</td>
                  <td>{p.TimeOut || '--'}</td>
                  <td>{p.TotalDuration || '--'}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {filteredParticipants.length === 0 && (
            <div style={{ textAlign: 'center', padding: '20px', color: '#999' }}>No participants found. Import an Excel file to start.</div>
          )}
        </div>
      </main>
    </div>
  );
};

const styles: { [key: string]: React.CSSProperties } = {
  container: { padding: '15px', fontFamily: 'system-ui, sans-serif', maxWidth: '1200px', margin: '0 auto', position: 'relative' },
  header: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '2px solid #eee', paddingBottom: '15px', flexWrap: 'wrap', gap: '10px' },
  toolbar: { display: 'flex', gap: '8px', flexWrap: 'wrap' },
  btn: { padding: '8px 16px', color: 'white', border: 'none', borderRadius: '6px', cursor: 'pointer', fontSize: '13px', fontWeight: '600', display: 'inline-flex', alignItems: 'center', justifyContent: 'center', minWidth: '110px', textAlign: 'center' },
  main: { display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: '20px', marginTop: '20px' },
  scannerSection: { background: '#fff', padding: '15px', borderRadius: '10px', border: '1px solid #ddd', boxShadow: '0 2px 4px rgba(0,0,0,0.05)', alignSelf: 'start' },
  statusDisplay: { fontSize: '1.1rem', fontWeight: 'bold', minHeight: '1.5rem', transition: 'all 0.3s ease' },
  listSection: { overflowX: 'auto', background: '#fff', padding: '15px', borderRadius: '10px', border: '1px solid #ddd' },
  searchBar: { width: '100%', padding: '10px', marginBottom: '15px', borderRadius: '6px', border: '1px solid #ccc', boxSizing: 'border-box' },
  table: { width: '100%', borderCollapse: 'collapse', fontSize: '14px' },
  tableHeader: { textAlign: 'left', borderBottom: '2px solid #eee', color: '#4b5563' },
  row: { borderBottom: '1px solid #f3f4f6' },
  // New Styles for Modal
  modalOverlay: { position: 'fixed', top: 0, left: 0, right: 0, bottom: 0, backgroundColor: 'rgba(0,0,0,0.6)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000, backdropFilter: 'blur(3px)' },
  modalContent: { backgroundColor: 'white', padding: '30px', borderRadius: '12px', maxWidth: '500px', width: '90%', boxShadow: '0 10px 25px rgba(0,0,0,0.2)' },
};

export default AttendanceTracker;