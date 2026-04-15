'use strict';

const express = require('express');
const { v4: uuidv4 } = require('uuid');
const path = require('path');
const fs = require('fs');
const fsPromises = require('fs/promises');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_TOKEN = process.env.POLLS_ADMIN_TOKEN || '';
const DATA_DIR = path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'polls-data.json');
const TEMP_FILE = DATA_FILE + '.tmp';

// ─── File I/O ────────────────────────────────────────────────────────────────

async function initData() {
  try {
    await fsPromises.mkdir(DATA_DIR, { recursive: true });
  } catch (_) {}
  try {
    await fsPromises.access(DATA_FILE);
  } catch (_) {
    await fsPromises.writeFile(DATA_FILE, JSON.stringify({ polls: [] }, null, 2));
  }
}

async function readData() {
  const raw = await fsPromises.readFile(DATA_FILE, 'utf8');
  return JSON.parse(raw);
}

async function writeData(data) {
  await fsPromises.writeFile(TEMP_FILE, JSON.stringify(data, null, 2));
  await fsPromises.rename(TEMP_FILE, DATA_FILE);
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

function findPoll(data, pollId) {
  return data.polls.find(p => p.id === pollId);
}

function normalizeName(name) {
  return name.trim().replace(/\s+/g, ' ').toLowerCase();
}

function computeResults(poll) {
  const totalVotes = poll.votes.length;
  const options = poll.options.map(opt => {
    const count = poll.votes.filter(v => v.optionId === opt.id).length;
    return {
      id: opt.id,
      text: opt.text,
      votes: count,
      percentage: totalVotes > 0 ? Math.round((count / totalVotes) * 1000) / 10 : 0,
    };
  });
  const voters = [...poll.votes]
    .sort((a, b) => new Date(a.votedAt) - new Date(b.votedAt))
    .map(v => {
      const opt = poll.options.find(o => o.id === v.optionId);
      return {
        voterName: v.voterName,
        optionId: v.optionId,
        optionText: opt ? opt.text : '(deleted)',
        votedAt: v.votedAt,
      };
    });
  return { totalVotes, options, voters };
}

async function generateExcel(poll, res) {
  const results = computeResults(poll);
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Evote App';
  workbook.created = new Date();

  // ── Sheet 1: Summary ──────────────────────────────────────────────────────
  const summary = workbook.addWorksheet('Summary');
  summary.columns = [
    { header: 'Field', key: 'field', width: 22 },
    { header: 'Value', key: 'value', width: 50 },
  ];

  const metaRows = [
    ['Poll Title', poll.title],
    ['Description', poll.description || ''],
    ['Created By', poll.createdBy || ''],
    ['Created At', new Date(poll.createdAt).toLocaleString()],
    ['Status', poll.isOpen ? 'Open' : 'Closed'],
    ['Total Votes', results.totalVotes],
  ];
  metaRows.forEach(([field, value]) => summary.addRow({ field, value }));

  summary.addRow({});
  const optHeader = summary.addRow({ field: 'Option', value: 'Votes' });
  optHeader.getCell('field').font = { bold: true };
  optHeader.getCell('value').font = { bold: true };

  const pctCol = summary.getColumn('value');
  summary.getColumn('C') || summary.addColumn({ header: 'Percentage', key: 'pct', width: 15 });

  // Re-define columns to include percentage
  summary.columns = [
    { header: 'Field / Option', key: 'field', width: 28 },
    { header: 'Value / Votes', key: 'value', width: 20 },
    { header: 'Percentage', key: 'pct', width: 15 },
  ];
  // Clear and re-add all rows with the right column count
  summary.spliceRows(1, summary.rowCount);

  const titleRow = summary.addRow(['Poll Results Summary']);
  titleRow.font = { bold: true, size: 14 };
  summary.mergeCells(`A1:C1`);
  summary.addRow([]);
  summary.addRow(['Poll Title', poll.title]);
  summary.addRow(['Description', poll.description || '']);
  summary.addRow(['Created By', poll.createdBy || '']);
  summary.addRow(['Created At', new Date(poll.createdAt).toLocaleString()]);
  summary.addRow(['Status', poll.isOpen ? 'Open' : 'Closed']);
  summary.addRow(['Total Votes', results.totalVotes]);
  summary.addRow([]);

  const optionHeaderRow = summary.addRow(['Option', 'Votes', 'Percentage']);
  optionHeaderRow.eachCell(cell => {
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3F3F3F' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  });

  results.options.forEach(opt => {
    summary.addRow([opt.text, opt.votes, `${opt.percentage}%`]);
  });

  // ── Sheet 2: Voter Log ────────────────────────────────────────────────────
  const voterLog = workbook.addWorksheet('Voter Log');
  voterLog.columns = [
    { header: 'Voter Name', key: 'voterName', width: 30 },
    { header: 'Voted For', key: 'optionText', width: 40 },
    { header: 'Voted At', key: 'votedAt', width: 26 },
  ];

  const logHeaderRow = voterLog.getRow(1);
  logHeaderRow.eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3F3F3F' } };
  });

  results.voters.forEach(v => {
    voterLog.addRow({
      voterName: v.voterName,
      optionText: v.optionText,
      votedAt: new Date(v.votedAt).toLocaleString(),
    });
  });

  // Stream to response
  const safeTitle = poll.title.replace(/[^a-z0-9]/gi, '-').toLowerCase().slice(0, 40);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="poll-${safeTitle}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
}

// ─── Middleware ───────────────────────────────────────────────────────────────

app.use(express.json());
app.use(express.static(__dirname));

function requireAdmin(req, res, next) {
  if (!ADMIN_TOKEN) {
    return res.status(500).json({ error: 'Server has no POLLS_ADMIN_TOKEN set. Cannot perform admin actions.' });
  }
  const token = req.headers['x-admin-token'];
  if (!token || token !== ADMIN_TOKEN) {
    return res.status(403).json({ error: 'Invalid or missing admin token.' });
  }
  next();
}

function validateBody(rules) {
  return (req, res, next) => {
    for (const [field, { required, maxLength, type }] of Object.entries(rules)) {
      const val = req.body[field];
      if (required && (val === undefined || val === null || val === '')) {
        return res.status(400).json({ error: `Field "${field}" is required.` });
      }
      if (val !== undefined && maxLength && String(val).length > maxLength) {
        return res.status(400).json({ error: `Field "${field}" exceeds max length of ${maxLength}.` });
      }
    }
    next();
  };
}

// ─── Routes: Polls ────────────────────────────────────────────────────────────

// GET /api/polls — list all polls (no voter names)
app.get('/api/polls', async (req, res) => {
  try {
    const data = await readData();
    const polls = data.polls.map(p => ({
      id: p.id,
      title: p.title,
      description: p.description,
      createdBy: p.createdBy,
      createdAt: p.createdAt,
      isOpen: p.isOpen,
      optionCount: p.options.length,
      voteCount: p.votes.length,
    }));
    res.json({ polls });
  } catch (err) {
    res.status(500).json({ error: 'Failed to read polls.' });
  }
});

// POST /api/polls — create a poll (admin)
app.post(
  '/api/polls',
  requireAdmin,
  validateBody({
    title: { required: true, maxLength: 255 },
  }),
  async (req, res) => {
    const { title, description, createdBy, options: rawOptions } = req.body;

    if (!Array.isArray(rawOptions) || rawOptions.length < 2) {
      return res.status(400).json({ error: 'At least 2 options are required.' });
    }
    if (rawOptions.length > 20) {
      return res.status(400).json({ error: 'Maximum 20 options allowed.' });
    }
    const cleanOptions = rawOptions.map(o => String(o).trim()).filter(Boolean);
    if (cleanOptions.length < 2) {
      return res.status(400).json({ error: 'At least 2 non-empty options are required.' });
    }
    for (const opt of cleanOptions) {
      if (opt.length > 200) {
        return res.status(400).json({ error: `Option text exceeds 200 characters: "${opt.slice(0, 30)}..."` });
      }
    }

    const poll = {
      id: uuidv4(),
      title: String(title).trim(),
      description: description ? String(description).trim() : '',
      createdBy: createdBy ? String(createdBy).trim().slice(0, 100) : '',
      options: cleanOptions.map(text => ({ id: uuidv4(), text })),
      votes: [],
      createdAt: new Date().toISOString(),
      isOpen: true,
    };

    try {
      const data = await readData();
      data.polls.push(poll);
      await writeData(data);
      res.status(201).json({ poll });
    } catch (err) {
      res.status(500).json({ error: 'Failed to save poll.' });
    }
  }
);

// GET /api/polls/:pollId — get single poll
app.get('/api/polls/:pollId', async (req, res) => {
  try {
    const data = await readData();
    const poll = findPoll(data, req.params.pollId);
    if (!poll) return res.status(404).json({ error: 'Poll not found.' });
    res.json({
      id: poll.id,
      title: poll.title,
      description: poll.description,
      createdBy: poll.createdBy,
      createdAt: poll.createdAt,
      isOpen: poll.isOpen,
      options: poll.options,
      voteCount: poll.votes.length,
    });
  } catch (err) {
    res.status(500).json({ error: 'Failed to read poll.' });
  }
});

// PATCH /api/polls/:pollId — update poll (admin: close/reopen)
app.patch('/api/polls/:pollId', requireAdmin, async (req, res) => {
  try {
    const data = await readData();
    const poll = findPoll(data, req.params.pollId);
    if (!poll) return res.status(404).json({ error: 'Poll not found.' });

    if (typeof req.body.isOpen === 'boolean') poll.isOpen = req.body.isOpen;
    if (req.body.title && poll.votes.length === 0) {
      poll.title = String(req.body.title).trim().slice(0, 255);
    }

    await writeData(data);
    res.json({ poll: { id: poll.id, title: poll.title, isOpen: poll.isOpen } });
  } catch (err) {
    res.status(500).json({ error: 'Failed to update poll.' });
  }
});

// DELETE /api/polls/:pollId — delete poll (admin)
app.delete('/api/polls/:pollId', requireAdmin, async (req, res) => {
  try {
    const data = await readData();
    const idx = data.polls.findIndex(p => p.id === req.params.pollId);
    if (idx === -1) return res.status(404).json({ error: 'Poll not found.' });
    data.polls.splice(idx, 1);
    await writeData(data);
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: 'Failed to delete poll.' });
  }
});

// ─── Routes: Votes ────────────────────────────────────────────────────────────

// POST /api/polls/:pollId/votes — cast a vote
app.post(
  '/api/polls/:pollId/votes',
  validateBody({
    voterName: { required: true, maxLength: 100 },
    optionId: { required: true },
  }),
  async (req, res) => {
    const { voterName, optionId } = req.body;
    const normalized = normalizeName(String(voterName));

    try {
      const data = await readData();
      const poll = findPoll(data, req.params.pollId);
      if (!poll) return res.status(404).json({ error: 'Poll not found.' });
      if (!poll.isOpen) return res.status(409).json({ error: 'This poll is closed and no longer accepting votes.' });

      const existing = poll.votes.find(v => v.voterNameNormalized === normalized);
      if (existing) {
        const existingOption = poll.options.find(o => o.id === existing.optionId);
        return res.status(409).json({
          error: `"${existing.voterName}" has already voted in this poll.`,
          existingVote: {
            optionId: existing.optionId,
            optionText: existingOption ? existingOption.text : '(deleted)',
            votedAt: existing.votedAt,
          },
        });
      }

      const option = poll.options.find(o => o.id === String(optionId));
      if (!option) return res.status(400).json({ error: 'Invalid option.' });

      poll.votes.push({
        voterName: String(voterName).trim(),
        voterNameNormalized: normalized,
        optionId: option.id,
        votedAt: new Date().toISOString(),
      });

      await writeData(data);
      res.status(201).json({ success: true, optionText: option.text });
    } catch (err) {
      res.status(500).json({ error: 'Failed to save vote.' });
    }
  }
);

// ─── Routes: Results & Export ─────────────────────────────────────────────────

// GET /api/polls/:pollId/results
app.get('/api/polls/:pollId/results', async (req, res) => {
  try {
    const data = await readData();
    const poll = findPoll(data, req.params.pollId);
    if (!poll) return res.status(404).json({ error: 'Poll not found.' });
    const results = computeResults(poll);
    res.json({
      poll: { id: poll.id, title: poll.title, description: poll.description, isOpen: poll.isOpen },
      ...results,
    });
  } catch (err) {
    res.status(500).json({ error: 'Failed to compute results.' });
  }
});

// GET /api/polls/:pollId/export — download xlsx
app.get('/api/polls/:pollId/export', async (req, res) => {
  try {
    const data = await readData();
    const poll = findPoll(data, req.params.pollId);
    if (!poll) return res.status(404).json({ error: 'Poll not found.' });
    await generateExcel(poll, res);
  } catch (err) {
    console.error('Excel export error:', err);
    if (!res.headersSent) res.status(500).json({ error: 'Failed to generate export.' });
  }
});

// ─── Start ────────────────────────────────────────────────────────────────────

initData().then(() => {
  app.listen(PORT, () => {
    console.log(`Evote app running at http://localhost:${PORT}/polls.html`);
    if (!ADMIN_TOKEN) {
      console.warn('WARNING: POLLS_ADMIN_TOKEN is not set. Admin routes are disabled.');
    }
  });
});
