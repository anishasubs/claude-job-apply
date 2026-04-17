# job-apply (public)

A Claude Code skill that helps you tailor resumes, write cover letters, and draft outreach for job applications. Works for any role type (PM, PMM, Growth, GTM, SWE, Marketing, Design, Data, etc.).

## What it does

When you share a job URL or paste a job description, Claude will:

1. **Analyze the role** — extract key requirements, keywords, culture signals.
2. **Pick the right resume starting point** from the variations you've uploaded.
3. **Tailor your resume** — restructure bullets, weave in keywords naturally, keep it one page.
4. **Draft a cover letter** — 3-4 tight paragraphs in your voice.
5. **Draft outreach** — LinkedIn, cold email, referral asks.
6. **Export polished files** — both `.docx` and `.pdf` for resume and cover letter.
7. **Update a tracker** — styled Excel spreadsheet with color-coded status flow.

Everything is saved under `~/.job-apply/applications/{company-slug}/`.

## Installation

### 1. Clone this repo into your Claude Code skills directory

**macOS / Linux:**
```bash
git clone https://github.com/anishasubs/claude-job-apply ~/.claude/skills/job-apply
```

**Windows (Git Bash or WSL):**
```bash
git clone https://github.com/anishasubs/claude-job-apply "$USERPROFILE/.claude/skills/job-apply"
```

**Windows (PowerShell):**
```powershell
git clone https://github.com/anishasubs/claude-job-apply "$env:USERPROFILE\.claude\skills\job-apply"
```

To update later: `cd ~/.claude/skills/job-apply && git pull`.

### 2. Install Python dependencies

```
pip install python-docx openpyxl docx2pdf lxml
```

### 3. Install Microsoft Word (or LibreOffice)

PDF conversion uses `docx2pdf`, which requires:
- **Windows / macOS**: Microsoft Word installed locally
- **Linux**: LibreOffice installed locally

### 4. Restart Claude Code

The skill will register automatically.

### 5. First-time setup

Type:

```
/job-apply
```

Claude will walk you through:

- Entering your name, contact info, and target role types
- Uploading 1–5 resume variations (to `~/.job-apply/resumes/`)
- Uploading 0–5 reference cover letters (to `~/.job-apply/cover-letters/`)
- Extracting a profile of your accomplishments from those files

## Resume template requirements

The resume generator works by **cloning your uploaded resume** and replacing content while preserving your formatting. For this to work, at least one of your `.docx` resumes must have:

- A paragraph with the literal text **`Experience`** as a section header
- A paragraph with the literal text **`Additional Information`** as a section header (marking the end of experience)
- Between them, at least one complete experience entry that includes:
  - A **bold** paragraph for the company name
  - An *italic* paragraph for a role description (optional but recommended)
  - A **bold** paragraph for the job title
  - Bulleted paragraphs for accomplishments

If your resumes don't match this structure, Claude will tell you during setup and offer to help restructure one into a usable template.

## Usage

After setup, just share a job:

```
/job-apply https://linkedin.com/jobs/view/1234567890
```

or paste a job description directly.

Claude will analyze the role and confirm before generating anything.

### Other commands

- `/job-apply tracker` — show current application tracker contents
- `/job-apply status` — same as above
- Mention a referral: *"I have a referral at Stripe, a former colleague named Alex Kim"* — Claude will ask follow-ups and generate referral-appropriate outreach.

## File layout

```
~/.job-apply/
  config.json           # your identity + preferences
  profile.md            # extracted accomplishments & themes
  tracker.xlsx          # styled application tracker
  resumes/              # your uploaded resume variations
  cover-letters/        # reference cover letters
  applications/
    {company-slug}/     # one folder per application
      role-analysis.md
      resume-content.json
      {FirstName}_{LastName}_Resume_{Company}.docx
      {FirstName}_{LastName}_Resume_{Company}.pdf
      cover-letter-content.json
      CoverLetter_{Company}.docx
      CoverLetter_{Company}.pdf
      outreach.md
      tracker-entry.json
```

## What it won't do

- Fabricate experience you don't have
- Exceed one page on the resume
- Stuff keywords unnaturally
- Store sensitive info (SSN, passport, salary details) unless you explicitly ask

## Scripts reference

All scripts are in `scripts/` and can be invoked directly:

- `generate_resume.py <content.json> <output.docx> --template <template.docx>`
- `generate_cover_letter.py <content.json> <output.docx> [--font NAME] [--size N]`
- `docx_to_pdf.py <file.docx> [file2.docx ...]`
- `update_tracker.py <tracker.xlsx> add <entry.json>` | `update <row#> <field> <value>` | `init`

Example JSON schemas live in `templates/`.

## License

MIT — use, modify, and redistribute freely.

---

Built by **[Anisha Subberwal](https://anishasubs.github.io/portfolio/)**. If this skill saves you time on your job hunt, a star on the repo is appreciated.
