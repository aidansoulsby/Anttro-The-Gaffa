const { google } = require('googleapis');

const SHEET_ID = '1X-1nLKpxBlFX636pSF-TI4SqNnC0DMSLt2eZjWGF-DA';

async function getSheetData(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: sheetName,
  });
  const [headers, ...rows] = res.data.values || [];
  if (!headers) return [];
  return rows.map(row =>
    headers.reduce((obj, key, i) => {
      obj[key.trim()] = row[i] || '';
      return obj;
    }, {})
  );
}

module.exports = async (req, res) => {
  // ── AUTH: shared secret for trusted backend callers (e.g. Make) ──
  const providedSecret = req.headers['x-email-api-secret'];
  if (!providedSecret || providedSecret !== process.env.EMAIL_API_SECRET) {
    return res.status(401).json({ error: 'Unauthorised.' });
  }

  const teamId = req.query.team;
  if (!teamId) {
    return res.status(400).json({ error: 'Missing team parameter.' });
  }

  try {
    // Auth with Google
    const raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    const credentials = JSON.parse(raw);
    if (credentials.private_key) {
      credentials.private_key = credentials.private_key.replace(/\\n/g, '\n');
    }
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    const sheets = google.sheets({ version: 'v4', auth });

    // Fetch all tabs in parallel
    const [snapshots, actionplans, reshots, teams] = await Promise.all([
      getSheetData(sheets, 'SnapShots'),
      getSheetData(sheets, 'ActionPlans'),
      getSheetData(sheets, 'ReShots'),
      getSheetData(sheets, 'Teams'),
    ]);

    // ── FILTER BY TEAM ──
    const snapshotsData = snapshots.filter(r => r['Team ID'] === teamId);
    const actionplansData = actionplans.filter(r => r['Team ID'] === teamId);
    const reshotsData = reshots.filter(r => r['Team ID'] === teamId);

    // ── SNAPSHOT STATS ── (matches data.js exactly)
    const snapshotOpen = snapshotsData.filter(r => r['Status'] === 'SnapShot Complete').length;

    // ── ACTIONPLAN STATS ── (matches data.js exactly)
    const actionPlanOpen = actionplansData.filter(r =>
      (r['Status'] === 'Active' || r['Status'] === 'In Progress') && r['Status'] !== 'Requires New ActionPlan'
    ).length;

    // ── LOOPS CLOSED ── (matches data.js exactly)
    const loopsClosed = reshotsData.filter(r => r['Problem Solved?'] === 'Yes').length;

    // ── NEEDS ATTENTION ── (matches data.js exactly)
    const today = new Date();
    const needsAttention = [];

    snapshotsData.forEach(r => {
      if (r['Status'] !== 'SnapShot Complete') return;
      if (!r['Date Submitted']) return;
      const submitted = new Date(r['Date Submitted'].split('/').reverse().join('-'));
      const daysAgo = Math.floor((today - submitted) / (1000 * 60 * 60 * 24));
      if (daysAgo >= 0) {
        needsAttention.push({
          title: r['Problem Title'],
          sub: `No ActionPlan · ${daysAgo} day${daysAgo !== 1 ? 's' : ''} since SnapShot`,
          type: 'no-actionplan',
        });
      }
    });

    actionplansData.forEach(r => {
      if (!r['Review Date'] || r['Status'] === 'Closed' || r['Status'] === 'Requires New ActionPlan' || r['Status'] === 'Loop Closed') return;
      const hasCompletedReshot = reshotsData.some(rs => rs['Action Plan ID'] === r['Action Plan ID'] && (rs['Problem Solved?'] === 'Yes' || rs['Status'] === 'Loop Closed'));
      if (hasCompletedReshot) return;
      const review = new Date(r['Review Date'].split('/').reverse().join('-'));
      const daysUntil = Math.floor((review - today) / (1000 * 60 * 60 * 24));
      const snap = snapshotsData.find(s => s['Snapshot ID'] === r['Snapshot ID']);
      const title = snap ? snap['Problem Title'] : r['Action Plan ID'];
      if (daysUntil < 0) {
        needsAttention.push({
          title,
          sub: `ReShot overdue · ${Math.abs(daysUntil)} day${Math.abs(daysUntil) !== 1 ? 's' : ''} past review date`,
          type: 'overdue',
        });
      }
    });

    // ── COMING UP (next 7 days, for the email digest specifically) ──
    const comingUp = [];
    actionplansData.forEach(r => {
      if (!r['Review Date'] || r['Status'] === 'Closed' || r['Status'] === 'Requires New ActionPlan') return;
      const review = new Date(r['Review Date'].split('/').reverse().join('-'));
      const daysUntil = Math.floor((review - today) / (1000 * 60 * 60 * 24));
      if (daysUntil >= 0 && daysUntil <= 7) {
        const snap = snapshotsData.find(s => s['Snapshot ID'] === r['Snapshot ID']);
        const title = snap ? snap['Problem Title'] : r['Action Plan ID'];
        comingUp.push({
          title,
          reviewDate: r['Review Date'],
          daysUntil,
        });
      }
    });
    comingUp.sort((a, b) => a.daysUntil - b.daysUntil);

    // ── TEAM NAME / MANAGER NAME ──
    const teamRow = teams.find(t => t['Team ID'] === teamId);
    const teamName = teamRow ? teamRow['Team Name'] : null;
    const managerName = teamRow ? teamRow['Contact Name'] : null;

    res.setHeader('Cache-Control', 'no-store, max-age=0');
    res.status(200).json({
      teamId,
      teamName,
      managerName,
      stats: {
        snapshotOpen,
        actionPlanOpen,
        loopsClosed,
      },
      needsAttention,
      comingUp,
    });

  } catch (err) {
    console.error('Email digest data error:', err);
    res.status(500).json({ error: err.message });
  }
};
