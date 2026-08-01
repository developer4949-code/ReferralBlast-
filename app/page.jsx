'use client';

import { useState, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';

export default function Home() {
  const [contacts, setContacts] = useState([]);
  const [attachments, setAttachments] = useState([]);
  
  const [config, setConfig] = useState({
    senderEmail: '',
    appPassword: '',
    subject: '{{company}} | Referral Request - {{role}}',
    template: `Hi {{first_name}} {{sir_maam}},

I'm Debi Prasad Das, a final-year B.Tech CSE student at IIIT Bhubaneswar (batch 2023–2027) — reaching out as a fellow IIIT BBSR alumnus. I'm applying for the {{role}} at {{company}} (Job ID: {{job_id}}) and would be very grateful if you'd consider referring me.

A few highlights:
• 8.7 CGPA at IIITBBSR
• SWE Intern at Yinolite — built REST APIs, auth systems, and worked in Agile sprints
• Projects in Java/Spring Boot, PostgreSQL, AWS, Docker, Redis, and FastAPI — including a cloud storage platform (CloudNest) and a distributed LLM memory system (Mnemosyne)
• 5 merged PRs to Checkstyle (open source)
• Finalist at DSCI/MeitY Cybersecurity Innovation Challenge

The role aligns closely with my interests in DSA, distributed systems, and backend engineering. I've attached my resume.

I completely understand if you're unable to refer — either way, thank you for your time.

Warm regards,
Debi Prasad Das

IIIT BHUBANESWAR  | 📧 debiprasaddas4949@gmail.com | 📞 +91-8260057716
🔗 <a href="{{github_link}}">GitHub</a> | <a href="{{linkedin_link}}">LinkedIn</a> | <a href="{{portfolio_link}}">Portfolio</a>`,
    portfolioLink: '',
    githubLink: '',
    linkedinLink: ''
  });

  const [isSending, setIsSending] = useState(false);
  const [progress, setProgress] = useState({ current: 0, total: 0, sent: 0, failed: 0 });
  const [quota, setQuota] = useState(null);
  const fileInputRef = useRef(null);

  useEffect(() => {
    const storedQuota = localStorage.getItem('emailQuota');
    const now = Date.now();
    if (storedQuota) {
      const parsed = JSON.parse(storedQuota);
      if (now > parsed.resetTime) {
        const newQuota = { remaining: 100, resetTime: now + 3600000 };
        setQuota(newQuota);
        localStorage.setItem('emailQuota', JSON.stringify(newQuota));
      } else {
        setQuota(parsed);
      }
    } else {
      const newQuota = { remaining: 100, resetTime: now + 3600000 };
      setQuota(newQuota);
      localStorage.setItem('emailQuota', JSON.stringify(newQuota));
    }
  }, []);

  const handleConfigChange = (e) => {
    setConfig({ ...config, [e.target.name]: e.target.value });
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
      
      const normalizedData = data.map(row => {
        const norm = {};
        for (const key in row) {
          norm[key.toLowerCase().trim()] = row[key];
        }
        // Ensure status column exists
        if (!('status' in norm)) norm.status = '';
        return norm;
      });
      
      setContacts(normalizedData);
    };
    reader.readAsBinaryString(file);
  };

  const handleAttachmentsUpload = (e) => {
    const files = Array.from(e.target.files);
    
    Promise.all(files.map(file => {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
          resolve({
            filename: file.name,
            content: reader.result.split(',')[1],
            contentType: file.type || 'application/octet-stream'
          });
        };
      });
    })).then(base64Files => {
      setAttachments(base64Files);
    });
  };

  const processTemplate = (row) => {
    let firstName = (row.first_name || row.name || '').toString().trim();
    if (firstName.toLowerCase() === 'nan') firstName = '';

    let gender = (row.gender || '').toString().trim().toUpperCase();
    let sirMaam = '';
    if (gender === 'M' || gender === 'MALE') sirMaam = 'Sir';
    else if (gender === 'F' || gender === 'FEMALE') sirMaam = "Ma'am";

    let text = config.template
      .replace(/{{first_name}}/g, firstName)
      .replace(/{{sir_maam}}/g, sirMaam)
      .replace(/{{portfolio_link}}/g, config.portfolioLink)
      .replace(/{{github_link}}/g, config.githubLink)
      .replace(/{{linkedin_link}}/g, config.linkedinLink);

    for (const key in row) {
      const val = (row[key] || '').toString().trim();
      if (val !== 'nan') {
        const regex = new RegExp(`{{${key}}}`, 'g');
        text = text.replace(regex, val);
      }
    }

    // Clean double spaces just in case
    text = text.replace(/  +/g, ' ').replace(/ ,/g, ',');
    return text;
  };


  const sendAll = async () => {
    if (!config.senderEmail || !config.appPassword) {
      alert("Please provide Sender Email and App Password.");
      return;
    }
    if (contacts.length === 0) {
      alert("Please upload an Excel file first.");
      return;
    }

    let currentQuota = { ...quota };
    const now = Date.now();
    if (now > currentQuota.resetTime) {
      currentQuota = { remaining: 100, resetTime: now + 3600000 };
      setQuota(currentQuota);
      localStorage.setItem('emailQuota', JSON.stringify(currentQuota));
    }

    if (currentQuota.remaining <= 0) {
      alert("You have reached your hourly limit of 100 emails. The counter will reset at " + new Date(currentQuota.resetTime).toLocaleTimeString());
      return;
    }

    setIsSending(true);
    let sentCount = 0;
    let failedCount = 0;
    let updatedContacts = [...contacts];

    setProgress({ current: 0, total: contacts.length, sent: 0, failed: 0 });

    for (let i = 0; i < updatedContacts.length; i++) {
      const contact = updatedContacts[i];
      const status = (contact.status || '').toString().trim().toLowerCase();
      
      if (status === 'sent') {
        setProgress(prev => ({ ...prev, current: i + 1, sent: prev.sent + 1 }));
        sentCount++;
        continue;
      }

      const toEmail = (contact.email || '').toString().trim();
      const ccEmail = (contact.cc_email || '').toString().trim();

      if (!toEmail || toEmail.toLowerCase() === 'nan') {
        updatedContacts[i].status = 'failed: missing email';
        failedCount++;
        setProgress(prev => ({ ...prev, current: i + 1, failed: prev.failed + 1 }));
        continue;
      }

      const body = processTemplate(contact);

      let finalSubject = config.subject;
      for (const key in contact) {
        const val = (contact[key] || '').toString().trim();
        if (val !== 'nan') {
          const regex = new RegExp(`{{${key}}}`, 'g');
          finalSubject = finalSubject.replace(regex, val);
        }
      }

      if (currentQuota.remaining <= 0) {
        alert("Hourly limit reached! Stopping email blast. It will reset at " + new Date(currentQuota.resetTime).toLocaleTimeString());
        break;
      }

      try {
        const res = await fetch('/api/send', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            senderEmail: config.senderEmail,
            appPassword: config.appPassword,
            to: toEmail,
            cc: ccEmail !== 'nan' ? ccEmail : '',
            subject: finalSubject,
            body: body,
            attachments: attachments
          })
        });

        if (res.ok) {
          updatedContacts[i].status = 'sent';
          sentCount++;
          
          currentQuota.remaining -= 1;
          setQuota({ ...currentQuota });
          localStorage.setItem('emailQuota', JSON.stringify(currentQuota));
        } else {
          const data = await res.json();
          updatedContacts[i].status = `failed: ${data.error || 'unknown error'}`;
          failedCount++;
        }
      } catch (err) {
        updatedContacts[i].status = `failed: ${err.message}`;
        failedCount++;
      }

      setProgress(prev => ({ 
        ...prev, 
        current: i + 1, 
        sent: sentCount, 
        failed: failedCount 
      }));
      setContacts([...updatedContacts]);

      // Small delay to prevent rate-limiting/spam flags
      await new Promise(r => setTimeout(r, 1500));
    }

    setIsSending(false);
    alert(`Done! Sent: ${sentCount}, Failed: ${failedCount}`);
  };

  return (
    <div className="container">
      {/* Sidebar Configuration */}
      <div className="sidebar">
        <div className="panel">
          <h2>⚙️ Configuration</h2>
          <div className="form-group">
            <label>Sender Email</label>
            <input 
              type="email" 
              name="senderEmail"
              placeholder="youremail@gmail.com"
              value={config.senderEmail}
              onChange={handleConfigChange}
            />
          </div>
          <div className="form-group">
            <label>App Password</label>
            <input 
              type="password" 
              name="appPassword"
              placeholder="16-character code"
              value={config.appPassword}
              onChange={handleConfigChange}
            />
          </div>

          <h2 style={{marginTop: '2rem'}}>🔗 Links</h2>
          <div className="form-group">
            <label>Portfolio Link</label>
            <input 
              type="text" 
              name="portfolioLink"
              placeholder="https://..."
              value={config.portfolioLink}
              onChange={handleConfigChange}
            />
          </div>
          <div className="form-group">
            <label>GitHub Link</label>
            <input 
              type="text" 
              name="githubLink"
              placeholder="https://github.com/..."
              value={config.githubLink}
              onChange={handleConfigChange}
            />
          </div>
          <div className="form-group">
            <label>LinkedIn Link</label>
            <input 
              type="text" 
              name="linkedinLink"
              placeholder="https://linkedin.com/in/..."
              value={config.linkedinLink}
              onChange={handleConfigChange}
            />
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="main-content">
        <div className="panel">
          <h1>🚀 ReferralBlast</h1>
          <p style={{marginBottom: '1.5rem', marginTop: '0.5rem'}}>Automate your personalized referral and application emails entirely in the browser. Zero databases.</p>
          
          <div className="alert-banner">
            <span style={{ fontSize: '1.5rem' }}>⚠️</span>
            <div>
              <strong style={{ display: 'block', marginBottom: '0.25rem' }}>Email Limit Notice</strong>
              <span style={{ fontSize: '0.85rem', color: 'var(--text-secondary)' }}>To comply with limits, you can send a maximum of 100 emails per hour. Your quota resets automatically.</span>
            </div>
          </div>

          <div className="metrics">
            <div className="metric-card">
              <div className="metric-value">{contacts.length}</div>
              <div className="metric-label">Total Contacts</div>
            </div>
            <div className="metric-card">
              <div className="metric-value">{progress.sent}</div>
              <div className="metric-label">Emails Sent</div>
            </div>
            <div className="metric-card">
              <div className="metric-value">{progress.failed}</div>
              <div className="metric-label">Failed</div>
            </div>
            <div className="metric-card" style={{ borderColor: quota?.remaining < 10 ? 'var(--danger)' : 'var(--border-color)' }}>
              <div className="metric-value" style={{ color: quota?.remaining < 10 ? 'var(--danger)' : 'var(--text-primary)' }}>{quota ? quota.remaining : 100}</div>
              <div className="metric-label">Quota Remaining</div>
              {quota && <div style={{fontSize: '0.7rem', color: 'var(--text-secondary)', marginTop: '0.25rem'}}>Resets: {new Date(quota.resetTime).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</div>}
            </div>
          </div>

          {isSending && (
             <div className="progress-container">
               <div 
                 className="progress-bar" 
                 style={{ width: `${contacts.length > 0 ? (progress.current / contacts.length) * 100 : 0}%` }}
               ></div>
             </div>
          )}
          {isSending && <div className="progress-text">Processing {progress.current} of {contacts.length}...</div>}

          <div style={{display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.5rem', marginTop: '1.5rem'}}>
            <div>
              <h2>1. Upload Data</h2>
              <div className="file-dropzone" onClick={() => fileInputRef.current.click()}>
                <p>Click to upload <b>Contacts Excel</b></p>
                <input 
                  type="file" 
                  accept=".xlsx, .xls"
                  ref={fileInputRef}
                  onChange={handleExcelUpload}
                />
              </div>
              {contacts.length > 0 && (
                <div style={{marginTop: '1rem'}}>
                  <p className="status-badge status-sent">Loaded {contacts.length} rows.</p>
                </div>
              )}
            </div>

            <div>
              <h2>2. Attachments</h2>
               <div className="file-dropzone" onClick={() => document.getElementById('attachmentsInput').click()}>
                <p>Click to add <b>Files (e.g. Resume)</b></p>
                <input 
                  id="attachmentsInput"
                  type="file" 
                  multiple
                  onChange={handleAttachmentsUpload}
                />
              </div>
              <div className="file-list">
                {attachments.map((att, idx) => (
                  <div key={idx} className="file-item">
                    <span>{att.filename}</span>
                    <span className="status-badge status-pending">Ready</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>

        <div className="panel">
          <h2>3. Email Content</h2>
          <div className="form-group">
            <label>Subject Line</label>
            <input 
              type="text" 
              name="subject"
              value={config.subject}
              onChange={handleConfigChange}
            />
          </div>
          <div className="form-group">
            <label>Template (Use {'{{first_name}}'}, {'{{sir_maam}}'})</label>
            <textarea 
              name="template"
              value={config.template}
              onChange={handleConfigChange}
            ></textarea>
          </div>
          
          <button 
            className="btn" 
            onClick={sendAll} 
            disabled={isSending || contacts.length === 0}
            style={{marginTop: '1rem'}}
          >
            {isSending ? 'Sending in progress...' : '🚀 Blast Emails Now'}
          </button>
        </div>

        {contacts.length > 0 && (
          <div className="panel">
            <h2>Data Preview</h2>
            <div className="table-wrapper">
              <table>
                <thead>
                  <tr>
                    <th>Name</th>
                    <th>Email</th>
                    <th>Gender</th>
                    <th>Status</th>
                  </tr>
                </thead>
                <tbody>
                  {contacts.map((row, idx) => (
                    <tr key={idx}>
                      <td>{row.first_name || row.name}</td>
                      <td>{row.email}</td>
                      <td>{row.gender}</td>
                      <td>
                        <span className={`status-badge ${
                          (row.status || '').includes('sent') ? 'status-sent' : 
                          (row.status || '').includes('failed') ? 'status-failed' : 'status-pending'
                        }`}>
                          {row.status || 'Pending'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
