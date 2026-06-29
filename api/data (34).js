const { google } = require('googleapis');
const Clerk = require('@clerk/backend');

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

async function getTeamIdFromSession(req) {
  const authHeader = req.headers['authorization'];
  if (!authHeader || !authHeader.startsWith('Bearer ')) return null;

  const token = authHeader.replace('Bearer ', '');

  try {
    const clerkClient = Clerk.createClerkClient({
      secretKey: process.env.CLERK_SECRET_KEY,
    });

    const payload = await clerkClient.verifyToken(token);
    const userId = payload.sub;
    const user = await clerkClient.users.getUser(userId);
    return user.publicMetadata?.teamId || user.privateMetadata?.teamId || null;
  } catch (err) {
    console.error('Session verification failed:', err.message);
    return null;
  }
}

module.exports = async (req, res) => {
  // CORS — locked to your domain only
  res.setHeader('Access-Control-Allow-Origin', 'https://anttro.com');
  res.setHeader('Access-Control-Allow-Methods', 'GET');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // ── AUTH: get team ID from verified session ──
  const teamId = await getTeamIdFromSession(req);

  if (!teamId) {
    return res.status(401).json({ error: 'Unauthorised. Please sign in.' });
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
    const [snapshots, actionplans, actions, reshots, teams] = await Promise.all([
      getSheetData(sheets, 'SnapShots'),
      getSheetData(sheets, 'ActionPlans'),
      getSheetData(sheets, 'Actions'),
      getSheetData(sheets, 'ReShots'),
      getSheetData(sheets, 'Teams'),
    ]);

    // ── FILTER BY VERIFIED TEAM ID ──
    const filteredSnapshots = snapshots.filter(r => r['Team ID'] === teamId);
    const filteredActionPlans = actionplans.filter(r => r['Team ID'] === teamId);
    const filteredReshots = reshots.filter(r => r['Team ID'] === teamId);

    const snapshotsData = filteredSnapshots;
    const actionplansData = filteredActionPlans;
    const reshotsData = filteredReshots;

    // ── SNAPSHOT STATS ──
    const totalSnapshots = snapshotsData.length;
    const newSnapshots = snapshotsData.filter(r => r['Status'] === 'SnapShot Complete');
    const snapshotOpen = newSnapshots.length;

    // ── ACTIONPLAN STATS ──
    const totalActionPlans = actionplansData.length;
    const actionPlanOpen = actionplansData.filter(r =>
      (r['Status'] === 'Active' || r['Status'] === 'In Progress') && r['Status'] !== 'Requires New ActionPlan'
    ).length;

    // ── RESHOT STATS ──
    const totalReshots = reshotsData.length;
    const reshotOpen = reshotsData.filter(r => r['Status'] === 'ReShot Due').length;

    // ── LOOPS CLOSED ──
    const loopsClosed = reshotsData.filter(r => r['Problem Solved?'] === 'Yes').length;

    // ── AVG SNAPSHOT TO CLOSE (days) ──
    const closedWithDates = reshotsData
      .filter(r => r['Problem Solved?'] === 'Yes' && r['Date Completed'])
      .map(r => {
        const snap = snapshotsData.find(s => s['Snapshot ID'] === r['Snapshot ID']);
        if (!snap || !snap['Date Submitted']) return null;
        const submitted = new Date(snap['Date Submitted'].split('/').reverse().join('-'));
        const completed = new Date(r['Date Completed'].split('/').reverse().join('-'));
        const days = Math.round((completed - submitted) / (1000 * 60 * 60 * 24));
        return isNaN(days) ? null : days;
      })
      .filter(d => d !== null);

    const avgDaysToClose = closedWithDates.length
      ? Math.round(closedWithDates.reduce((a, b) => a + b, 0) / closedWithDates.length)
      : null;

    // ── LOOPS CLOSED THIS MONTH ──
    const now = new Date();
    const closedThisMonth = reshotsData.filter(r => {
      if (r['Problem Solved?'] !== 'Yes' || !r['Date Completed']) return false;
      const d = new Date(r['Date Completed'].split('/').reverse().join('-'));
      return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    }).length;

    // ── CHART DATA — daily submissions this month ──
    const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const dailyChart = Array(daysInMonth).fill(0);
    snapshotsData.forEach(r => {
      if (!r['Date Submitted']) return;
      const d = new Date(r['Date Submitted'].split('/').reverse().join('-'));
      if (d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear()) {
        const day = d.getDate() - 1;
        if (day >= 0 && day < daysInMonth) dailyChart[day]++;
      }
    });

    // ── CHART DATA — weekly (last 12 weeks) ──
    const weeklyChart = Array(12).fill(0);
    const msPerWeek = 7 * 24 * 60 * 60 * 1000;
    snapshotsData.forEach(r => {
      if (!r['Date Submitted']) return;
      const d = new Date(r['Date Submitted'].split('/').reverse().join('-'));
      const weeksAgo = Math.floor((now - d) / msPerWeek);
      if (weeksAgo >= 0 && weeksAgo < 12) weeklyChart[11 - weeksAgo]++;
    });

    // ── CHART DATA — monthly (last 12 months) ──
    const monthlyChart = Array(12).fill(0);
    snapshotsData.forEach(r => {
      if (!r['Date Submitted']) return;
      const d = new Date(r['Date Submitted'].split('/').reverse().join('-'));
      const monthsAgo = (now.getFullYear() - d.getFullYear()) * 12 + (now.getMonth() - d.getMonth());
      if (monthsAgo >= 0 && monthsAgo < 12) monthlyChart[11 - monthsAgo]++;
    });

    // ── SNAPSHOT POPUP ROWS ──
    const snapshotRows = snapshotsData.map(r => {
      const status = r['Status'] || 'SnapShot Complete';
      let badge, dot, btn;
      if (status === 'Loop Closed') {
        badge = 'badge-cl'; dot = 'dot-grey'; btn = 'View Slides';
      } else if (status === 'ActionPlan Active') {
        badge = 'badge-ap'; dot = 'dot-amber'; btn = 'View Slides';
      } else {
        badge = 'badge-ss'; dot = 'dot-red'; btn = 'Start ActionPlan';
      }
      let dateFormatted = r['Date Submitted'] || '';
      try {
        const d = new Date(r['Date Submitted'].split('/').reverse().join('-'));
        if (!isNaN(d)) dateFormatted = d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
      } catch(e) {}
      return {
        id: r['Snapshot ID'],
        title: r['Problem Title'],
        by: r['Created By'],
        date: dateFormatted,
        status, badge, dot, btn,
        slidesLink: r['Slides Link'] || null,
        currentState: r['Current State'] || '',
        ishikawa: r['Ishikawa Factors'] || '',
        waste: r['Waste Factors'] || '',
        whyThisMatters: r['Why This Matters'] || '',
        anttroInsight: r['Anttro Insight'] || '',
        estimatedTimeLostHours: r['Estimated Time Lost Hours'] ? parseFloat(r['Estimated Time Lost Hours']) || null : null,
        frequencyPerWeek: r['Frequency Per Week'] ? parseFloat(r['Frequency Per Week']) || null : null,
      };
    });

    // ── ACTIONPLAN POPUP ROWS ──
    const actionplanRows = actionplansData.map(r => {
      const status = r['Status'] || 'Active';
      let badge, dot, btn;
      if (status === 'Closed') {
        badge = 'badge-cl'; dot = 'dot-grey'; btn = 'View ActionPlan';
      } else if (status === 'Requires New ActionPlan') {
        badge = 'badge-rs'; dot = 'dot-blue'; btn = 'Start ActionPlan';
      } else {
        badge = 'badge-ap'; dot = 'dot-amber'; btn = 'View ActionPlan';
      }
      let apDateFormatted = r['Date Created'] || '';
      try {
        const d = new Date(r['Date Created'].split('/').reverse().join('-'));
        if (!isNaN(d)) apDateFormatted = d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
      } catch(e) {}
      return {
        id: r['Action Plan ID'],
        title: r['Problem Title'] || snapshotsData.find(s => s['Snapshot ID'] === r['Snapshot ID'])?.['Problem Title'] || '',
        by: r['Created By'],
        date: apDateFormatted,
        status, badge, dot, btn,
        slidesLink: r['Slides Link'] || null,
        reviewDate: r['Review Date'] || '',
        valueCategory: r['Value Category'] || '',
        snapshotId: r['Snapshot ID'] || '',
      };
    });

    // ── RESHOT POPUP ROWS ──
    const reshotRows = reshotsData.map(r => {
      const status = r['Status'] || 'ReShot Due';
      let badge, dot, btn;
      if (status === 'Loop Closed') {
        badge = 'badge-cl'; dot = 'dot-grey'; btn = 'View ReShot';
      } else if (status === 'Requires New ActionPlan') {
        badge = 'badge-rs'; dot = 'dot-blue'; btn = 'Start ActionPlan';
      } else {
        badge = 'badge-rs'; dot = 'dot-blue'; btn = 'Start ReShot';
      }
      return {
        id: r['ReShot ID'],
        title: snapshotsData.find(s => s['Snapshot ID'] === r['Snapshot ID'])?.['Problem Title'] || '',
        by: r['Completed By'],
        date: r['Date Completed'],
        status, badge, dot, btn,
        slidesLink: r['Slides Link'] || null,
        problemSolved: r['Problem Solved?'] || '',
        actionsSummary: r['Actions Summary'] || '',
        snapshotId: r['Snapshot ID'] || '',
        rawStatus: r['Status'] || '',
        actionPlanId: r['Action Plan ID'] || '',
      };
    });

    // ── NEEDS ATTENTION ──
    const today = new Date();
    const needsAttention = [];

    snapshotsData.forEach(r => {
      if (r['Status'] !== 'SnapShot Complete') return;
      if (!r['Date Submitted']) return;
      const submitted = new Date(r['Date Submitted'].split('/').reverse().join('-'));
      const daysAgo = Math.floor((today - submitted) / (1000 * 60 * 60 * 24));
      if (daysAgo >= 0) {
        needsAttention.push({
          dot: '#e05a5a',
          title: r['Problem Title'],
          sub: `No ActionPlan · ${daysAgo} day${daysAgo !== 1 ? 's' : ''} since SnapShot`,
          btn: 'Start ActionPlan',
          type: 'no-actionplan',
          id: r['Snapshot ID'],
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
          dot: '#f0a500',
          title,
          sub: `ReShot overdue · ${Math.abs(daysUntil)} day${Math.abs(daysUntil) !== 1 ? 's' : ''} past review date`,
          btn: 'Start ReShot',
          type: 'overdue',
          id: r['Action Plan ID'],
        });
      } else if (daysUntil <= 3) {
        needsAttention.push({
          dot: '#f0a500',
          title,
          sub: `Review due in ${daysUntil} day${daysUntil !== 1 ? 's' : ''}`,
          btn: 'View ActionPlan',
          type: 'upcoming',
          id: r['Action Plan ID'],
        });
      }
    });

    // ── COMING UP ──
    const comingUp = [];
    actionplansData.forEach(r => {
      if (!r['Review Date'] || r['Status'] === 'Closed' || r['Status'] === 'Requires New ActionPlan') return;
      const review = new Date(r['Review Date'].split('/').reverse().join('-'));
      const daysUntil = Math.floor((review - today) / (1000 * 60 * 60 * 24));
      if (daysUntil >= 0 && daysUntil <= 30) {
        const snap = snapshotsData.find(s => s['Snapshot ID'] === r['Snapshot ID']);
        const title = snap ? snap['Problem Title'] : r['Action Plan ID'];
        const actionCount = actionplansData.filter(a => a['Snapshot ID'] === r['Snapshot ID']).length;
        comingUp.push({
          dot: daysUntil <= 3 ? '#f0a500' : '#4a9b7f',
          title: `${title} — Review`,
          sub: daysUntil === 0 ? 'Today' : `In ${daysUntil} day${daysUntil !== 1 ? 's' : ''} · ${actionCount} action${actionCount !== 1 ? 's' : ''} to review`,
          btn: 'View ActionPlan',
          id: r['Action Plan ID'],
        });
      }
    });
    comingUp.sort((a, b) => a.sub.localeCompare(b.sub));

    // ── INSIGHT ──
    const noActionPlanCount = snapshotsData.filter(r => r['Status'] === 'SnapShot Complete').length;
    const oldestNoAP = snapshotsData
      .filter(r => r['Status'] === 'SnapShot Complete' && r['Date Submitted'])
      .map(r => ({ ...r, days: Math.floor((today - new Date(r['Date Submitted'].split('/').reverse().join('-'))) / (1000 * 60 * 60 * 24)) }))
      .sort((a, b) => b.days - a.days);
    const upcomingReviews = comingUp.filter(r => {
      const days = parseInt(r.sub);
      return !isNaN(days) && days <= 14;
    }).length;

    const insight = {
      loopHealth: noActionPlanCount > 0
        ? `Your team has <strong>${noActionPlanCount} open SnapShot${noActionPlanCount !== 1 ? 's' : ''}</strong> with no ActionPlan started.${avgDaysToClose ? ` The average time from SnapShot to close is <strong>${Math.abs(avgDaysToClose)} days</strong>.` : ''} ${upcomingReviews > 0 ? `${upcomingReviews} review date${upcomingReviews !== 1 ? 's are' : ' is'} coming up in the next 14 days.` : ''}`
        : `All SnapShots have ActionPlans in place. ${upcomingReviews > 0 ? `${upcomingReviews} review date${upcomingReviews !== 1 ? 's are' : ' is'} coming up in the next 14 days.` : 'Keep up the great work!'}`,
      trend: dailyChart.reduce((a, b) => a + b, 0) > 0
        ? `Your team has submitted <strong>${totalSnapshots} SnapShot${totalSnapshots !== 1 ? 's' : ''}</strong> in total. <strong>${loopsClosed} loop${loopsClosed !== 1 ? 's have' : ' has'} been closed</strong> so far.`
        : `No SnapShots submitted yet this month. Encourage your team to raise problems using the QR code.`,
      recommendation: oldestNoAP.length > 0
        ? `Prioritise starting an ActionPlan for <strong>${oldestNoAP[0]['Problem Title']}</strong> — it has been waiting <strong>${oldestNoAP[0].days} day${oldestNoAP[0].days !== 1 ? 's' : ''}</strong> with no action taken.`
        : loopsClosed > 0
        ? `Great work closing ${loopsClosed} loop${loopsClosed !== 1 ? 's' : ''}. Keep the momentum going by reviewing upcoming ActionPlans on time.`
        : `Start by creating ActionPlans for any open SnapShots to keep your continuous improvement loop moving.`,
    };

    // ── TEAM NAME ──
    const teamRow = teams.find(t => t['Team ID'] === teamId);
    const teamName = teamRow ? teamRow['Team Name'] : null;
    const managerName = teamRow ? teamRow['Contact Name'] : null;
    const hourlyWage = teamRow && teamRow['Average Hourly Wage'] ? parseFloat(teamRow['Average Hourly Wage']) || null : null;
    const businessContext = teamRow ? teamRow['Business Context'] || '' : '';

    res.setHeader('Cache-Control', 'no-store, max-age=0');
    res.status(200).json({
      teamName,
      managerName,
      hourlyWage,
      businessContext,
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
      actions: actions.filter(r => r['Team ID'] === teamId),
      newSnapshots: snapshotRows.filter(r => r.status === 'SnapShot Complete'),
      needsNewAP: reshotRows.filter(r => (r.problemSolved === 'No' || r.status === 'Requires New ActionPlan') && r.rawStatus !== 'Closed'),
      needsAttention,
      comingUp,
      insight,
    });

  } catch (err) {
    console.error('Google Sheets error:', err);
    res.status(500).json({ error: err.message });
  }
};
