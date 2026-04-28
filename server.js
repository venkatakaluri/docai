require('dotenv').config();

const express = require('express');
const { buildDocx } = require('./buildDocx');

const app = express();
app.use(express.json({ limit: '5mb' }));
app.use(express.static('public'));

const GROQ_API_KEY = process.env.GROQ_API_KEY;
const GROQ_URL     = 'https://api.groq.com/openai/v1/chat/completions';
const MODEL        = 'llama-3.3-70b-versatile';

if (!GROQ_API_KEY || GROQ_API_KEY === 'YOUR_GROQ_API_KEY_HERE') {
  console.error('\n❌  ERROR: No Groq API key found in .env file\n');
  process.exit(1);
}

console.log('✅  Key loaded:', GROQ_API_KEY.substring(0, 12) + '...');

const SYSTEM = `You are a document formatting engine. The user gives you raw unformatted text. You must return ONLY a valid JSON object. No explanation. No markdown. No code fences. Just raw JSON.

JSON schema:
{
  "filename": "short_filename_no_spaces",
  "blocks": [ array of block objects ]
}

Block types:
{ "type": "title", "text": "..." }
{ "type": "subtitle", "text": "..." }
{ "type": "info_row", "fields": ["Name", "Date", "Class"] }
{ "type": "heading", "text": "..." }
{ "type": "subheading", "text": "..." }
{ "type": "paragraph", "text": "..." }
{ "type": "note", "label": "Instructions", "text": "..." }
{ "type": "section", "text": "Section A — MCQ (1 mark each)" }
{ "type": "question", "number": 1, "text": "Question here?", "marks": 2, "options": ["A) opt1", "B) opt2", "C) opt3", "D) opt4"], "answer_lines": 0 }
{ "type": "bullets", "items": ["point 1", "point 2"] }
{ "type": "numbered", "items": ["item 1", "item 2"] }
{ "type": "table", "headers": ["Col1", "Col2"], "rows": [["a", "b"], ["c", "d"]] }
{ "type": "space" }

Rules:
- Detect document type automatically from the content
- Question paper: title + subtitle (class, time, total marks) + info_row (Name, Roll No, Date) + instructions note + sections + questions
  - MCQ: fill options array, answer_lines = 0
  - Short answer: empty options [], answer_lines = 2
  - Long answer: empty options [], answer_lines = 6
  - Marks shown right-aligned per question
- Notes: heading per topic, bullets for key points, paragraph for explanation
- Report: heading per section, paragraphs, tables for data
- Resume: title = name, sections for Summary, Skills, Experience, Education
- Fix all spelling mistakes. Keep all original facts.
- Return ONLY the JSON. Nothing else.`;

app.post('/generate', async (req, res) => {
  const { text } = req.body;
  if (!text || !text.trim()) return res.status(400).json({ error: 'No text provided' });

  try {
    const groqRes = await fetch(GROQ_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${GROQ_API_KEY}`
      },
      body: JSON.stringify({
        model: MODEL,
        temperature: 0.1,
        max_tokens: 8000,
        messages: [
          { role: 'system', content: SYSTEM },
          { role: 'user',   content: text.trim() }
        ]
      })
    });

    const groqData = await groqRes.json();

    if (groqData.error) {
      console.error('Groq error:', groqData.error);
      return res.status(500).json({ error: groqData.error.message });
    }

    let raw = groqData.choices[0].message.content.trim();
    raw = raw.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/```\s*$/i, '').trim();

    const start = raw.indexOf('{');
    const end   = raw.lastIndexOf('}');
    if (start !== -1 && end !== -1) raw = raw.slice(start, end + 1);

    let structure;
    try {
      structure = JSON.parse(raw);
    } catch (e) {
      return res.status(500).json({ error: 'AI returned invalid format. Try again.' });
    }

    const buffer   = await buildDocx(structure);
    const filename = (structure.filename || 'document').replace(/[^a-zA-Z0-9_-]/g, '_') + '.docx';

    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);

  } catch (err) {
    console.error('Server error:', err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`\n🚀  Running at http://localhost:${PORT}\n`));
