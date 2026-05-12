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
  // CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET');

  try {
    // Auth
    const raw = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    const credentials = JSON.parse(raw);
    // Vercel stores \n as literal \\n — fix the private key
    if (credentials.private_key) {
      credentials.private_key = credentials.private_key.replace(/\\n/g, '\n');
    }
    const auth = new google.auth.GoogleAuth({
      credentials,
      scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    const sheets = google.sheets({ version: 'v4', auth });

    // Fetch all tabs in parallel
    const [snapshots, actionplans, actions, reshots, teams] = await Promise.all([
      getSheetData(sheets, 'SnapShots'),
      getSheetData(sheets, 'ActionPlans'),
      getSheetData(sheets, 'Actions'),
      getSheetData(sheets, 'ReShots'),
      getSheetData(sheets, 'Teams'),
    ]);

    // ── SNAPSHOT STATS ──
    const totalSnapshots = snapshots.length;
    const newSnapshots = snapshots.filter(r => r['Status'] === 'SnapShot Complete');
    const snapshotOpen = newSnapshots.length;

    // ── ACTIONPLAN STATS ──
    const totalActionPlans = actionplans.length;
    const actionPlanOpen = actionplans.filter(r =>
      r['Status'] === 'Active' || r['Status'] === 'In Progress'
    ).length;

    // ── RESHOT STATS ──
    const totalReshots = reshots.length;
    const reshotOpen = reshots.filter(r => r['Status'] === 'ReShot Due').length;

    // ── LOOPS CLOSED ──
    const loopsClosed = reshots.filter(r => r['Problem Solved?'] === 'Yes').length;

    // ── AVG SNAPSHOT TO CLOSE (days) ──
    const closedWithDates = reshots
      .filter(r => r['Problem Solved?'] === 'Yes' && r['Date Completed'])
      .map(r => {
        const snap = snapshots.find(s => s['Snapshot ID'] === r['Snapshot ID']);
        if (!snap || !snap['Date Submitted']) return null;
        const submitted = new Date(snap['Date Submitted']);
        const completed = new Date(r['Date Completed']);
        const days = Math.round((completed - submitted) / (1000 * 60 * 60 * 24));
        return isNaN(days) ? null : days;
      })
      .filter(d => d !== null);

    const avgDaysToClose = closedWithDates.length
      ? Math.round(closedWithDates.reduce((a, b) => a + b, 0) / closedWithDates.length)
      : null;

    // ── LOOPS CLOSED THIS MONTH ──
    const now = new Date();
    const closedThisMonth = reshots.filter(r => {
      if (r['Problem Solved?'] !== 'Yes' || !r['Date Completed']) return false;
      const d = new Date(r['Date Completed']);
      return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    }).length;

    // ── CHART DATA — daily submissions this month ──
    const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const dailyChart = Array(daysInMonth).fill(0);
    snapshots.forEach(r => {
      if (!r['Date Submitted']) return;
      const d = new Date(r['Date Submitted']);
      if (d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear()) {
        const day = d.getDate() - 1;
        if (day >= 0 && day < daysInMonth) dailyChart[day]++;
      }
    });

    // ── CHART DATA — weekly (last 12 weeks) ──
    const weeklyChart = Array(12).fill(0);
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    snapshots.forEach(r => {
      if (!r['Date Submitted']) return;
      const d = new Date(r['Date Submitted']);
      const weeksAgo = Math.floor((now - d) / msPerWeek);
      if (weeksAgo >= 0 && weeksAgo < 12) {
        weeklyChart[11 - weeksAgo]++;
      }
    });

    // ── CHART DATA — monthly (last 12 months) ──
    const monthlyChart = Array(12).fill(0);
    snapshots.forEach(r => {
      if (!r['Date Submitted']) return;
      const d = new Date(r['Date Submitted']);
      const monthsAgo = (now.getFullYear() - d.getFullYear()) * 12 + (now.getMonth() - d.getMonth());
      if (monthsAgo >= 0 && monthsAgo < 12) {
        monthlyChart[11 - monthsAgo]++;
      }
    });

    // ── SNAPSHOT POPUP ROWS ──
    const snapshotRows = snapshots.map(r => {
      const ap = actionplans.find(a => a['Snapshot ID'] === r['Snapshot ID']);
      const rs = reshots.find(x => x['Snapshot ID'] === r['Snapshot ID']);
      let status, badge, dot, btn;
      if (rs && rs['Problem Solved?'] === 'Yes') {
        status = 'Loop Closed'; badge = 'badge-cl'; dot = 'dot-grey'; btn = 'View Slides';
      } else if (ap) {
        status = 'ActionPlan Active'; badge = 'badge-ap'; dot = 'dot-amber'; btn = 'View Slides';
      } else {
        status = 'SnapShot Complete'; badge = 'badge-ss'; dot = 'dot-green'; btn = 'Start ActionPlan';
      }
      return {
        id: r['Snapshot ID'],
        title: r['Problem Title'],
        by: r['Created By'],
        date: r['Date Submitted'],
        status, badge, dot, btn,
        slidesLink: r['Slides Link'] || null,
      };
    });

    // ── ACTIONPLAN POPUP ROWS ──
    const actionplanRows = actionplans.map(r => {
      const rs = reshots.find(x => x['Action Plan ID'] === r['Action Plan ID']);
      let status, badge, dot, btn;
      if (rs && rs['Problem Solved?'] === 'Yes') {
        status = 'Closed'; badge = 'badge-cl'; dot = 'dot-grey'; btn = 'View ActionPlan';
      } else if (r['Review Date']) {
        status = `Active · Review ${r['Review Date']}`; badge = 'badge-ap'; dot = 'dot-amber'; btn = 'View ActionPlan';
      } else {
        status = 'Active'; badge = 'badge-ap'; dot = 'dot-amber'; btn = 'View ActionPlan';
      }
      return {
        id: r['Action Plan ID'],
        title: snapshots.find(s => s['Snapshot ID'] === r['Snapshot ID'])?.['Problem Title'] || '',
        by: r['Created By Manager'],
        date: r['Date Created'],
        status, badge, dot, btn,
        slidesLink: r['Slides Link'] || null,
      };
    });

    // ── RESHOT POPUP ROWS ──
    const reshotRows = reshots.map(r => {
      const solved = r['Problem Solved?'] === 'Yes';
      return {
        id: r['ReShot ID'],
        title: snapshots.find(s => s['Snapshot ID'] === r['Snapshot ID'])?.['Problem Title'] || '',
        by: r['Completed By'],
        date: r['Date Completed'],
        status: solved ? 'Loop Closed' : 'ReShot Due',
        badge: solved ? 'badge-cl' : 'badge-rs',
        dot: solved ? 'dot-grey' : 'dot-blue',
        btn: solved ? 'View ReShot' : 'Start ReShot',
        slidesLink: r['Slides Link'] || null,
      };
    });

    // ── RESPONSE ──
    res.status(200).json({
      stats: {
        totalSnapshots,
        snapshotOpen,
        totalActionPlans,
        actionPlanOpen,
        totalReshots,
        reshotOpen,
        loopsClosed,
        avgDaysToClose,
        closedThisMonth,
      },
      chart: {
        daily: dailyChart,
        weekly: weeklyChart,
        monthly: monthlyChart,
      },
      rows: {
        snapshots: snapshotRows,
        actionplans: actionplanRows,
        reshots: reshotRows,
      },
      newSnapshots: snapshotRows.filter(r => r.status === 'SnapShot Complete'),
    });

  } catch (err) {
    console.error('Google Sheets error:', err);
    res.status(500).json({ error: err.message });
  }
};
