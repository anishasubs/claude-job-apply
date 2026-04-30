![get-me-a-job](banner.png)

# get-me-a-job 💼☕

> **A Claude Code skill that tailors your resume, writes your cover letter, drafts your outreach, and actually updates your tracker — so you can stop pretending you'll "clean up later."**

## The problem 😩

Your job search folder looks like this:

```
Resume_FINAL.docx
Resume_FINAL_v2.docx
Resume_FINAL_v2_actually_final.docx
Resume_Stripe_edit_2.docx
Resume_use_this_one.docx
cover letter.docx
cover letter (1).docx
```

You opened your applications tracker spreadsheet twice this month. It's out of date. You meant to ask your friend for that referral three weeks ago.

Every new job post means re-uploading your resume to Claude, re-explaining your background, re-pasting the JD, and hoping the output is usable. Then you eyeball it for typos and send it off.

This skill fixes that. ✨

## What it does 🧠

You drop your existing resumes and cover letters in once. Claude:

- 📄 **Learns your story** — parses your name, contact info, accomplishments, projects, and the *layout* of your resume (section order, header style, bullet density)
- 🎯 **Asks sharp targeting questions** — roles, companies, what to emphasize, what to avoid, hard constraints, recent wins not on paper
- 🔗 **Takes a job URL** (LinkedIn, Greenhouse, Lever, Ashby, Workday — anything) and reads the JD
- 🔍 **Shows you a preview** of what it'll highlight + flags gaps before writing a single file
- ✍️ **Writes the resume** tailored to that role, as a one-page `.docx` + `.pdf` that matches YOUR formatting
- 💌 **Writes the cover letter** in your voice, based on past cover letters you uploaded
- 📬 **Drafts outreach** — LinkedIn connect, cold email, referral asks, follow-ups
- 📊 **Updates your tracker** automatically — a styled Excel spreadsheet with color-coded status flow, company, role, comp, notes, URL, the lot
- 🔄 **Invites revisions** at every step — change a bullet, a cover letter paragraph, an outreach message, a tracker field. Iteration is expected, not punished.

No more v1, v2, v2_final. One folder per application. Clean names. Pristine outputs. ✨

## Installation 🚀

### 1. Clone the repo into your Claude Code skills folder

**macOS / Linux:**
```bash
git clone https://github.com/anishasubs/get-me-a-job ~/.claude/skills/get-me-a-job
```

**Windows (Git Bash or WSL):**
```bash
git clone https://github.com/anishasubs/get-me-a-job "$USERPROFILE/.claude/skills/get-me-a-job"
```

**Windows (PowerShell):**
```powershell
git clone https://github.com/anishasubs/get-me-a-job "$env:USERPROFILE\.claude\skills\get-me-a-job"
```

To update later: `cd ~/.claude/skills/get-me-a-job && git pull` 🔁

### 2. Install Python dependencies 🐍

```bash
pip install python-docx openpyxl docx2pdf lxml
```

### 3. Install Microsoft Word (or LibreOffice)

PDF conversion uses `docx2pdf`, which requires:
- 🪟 **Windows / 🍎 macOS**: Microsoft Word installed locally
- 🐧 **Linux**: LibreOffice installed locally

### 4. Restart Claude Code

The skill auto-registers. ✅

> 💡 **Using Claude Cowork or another sandboxed environment?**
>
> Claude Cowork can only access folders that already exist — it can't create `~/.get-me-a-job/` for you. Do this once before your first `/get-me-a-job`:
>
> **macOS / Linux:**
> ```bash
> mkdir -p ~/.get-me-a-job/resumes ~/.get-me-a-job/cover-letters ~/.get-me-a-job/applications
> ```
>
> **Windows (PowerShell):**
> ```powershell
> mkdir $env:USERPROFILE\.get-me-a-job\resumes, $env:USERPROFILE\.get-me-a-job\cover-letters, $env:USERPROFILE\.get-me-a-job\applications
> ```
>
> Or skip the dotfolder entirely — pick any existing folder (e.g., `~/Documents/`) when Claude asks where to store your data during onboarding. Claude Code users can ignore this, the skill handles creation automatically.

### 5. First-time setup

Type:
```
/get-me-a-job
```

Claude will walk you through it:
- 📂 Pick where your data lives (default: `~/.get-me-a-job/` — or choose your own, like `~/job-hunt/`)
- 📥 Drop 1–5 of your existing resumes (`.docx` and `.pdf`) into `{your-folder}/resumes/`
- 💌 Drop 0–5 of your past cover letters into `{your-folder}/cover-letters/` (optional — for voice)
- 🤖 Claude reads them, extracts your identity + accomplishments + layout preferences
- 🗣️ Claude has a targeting conversation with you — roles, companies, constraints, narrative
- ✅ Done. You're ready to apply.

## Usage 💼

After setup, just share a job:

```
/get-me-a-job https://linkedin.com/jobs/view/1234567890
```

Or paste a JD directly:
```
/get-me-a-job
[paste job description]
```

Or just drop a URL in normal chat — the skill auto-triggers from the URL pattern.

### What happens next 🔄

1. Claude analyzes the role (requirements, keywords, culture)
2. Picks the best starting resume from your uploads
3. Shows you a **pre-generation preview**: what it'll highlight, what it'll lead with in the cover letter, keywords it'll weave in, gaps it'll flag, questions for you
4. Waits for your approval (tweak anything first)
5. Generates the tailored resume (`.docx` + `.pdf`), cover letter (`.docx` + `.pdf`), outreach messages, and updates the Excel tracker
6. Presents everything as a clean summary and invites revisions

### Other commands 🛠

- `/get-me-a-job tracker` — show your application tracker
- `/get-me-a-job status` — same
- Mention a referral in chat ("I have a referral at Stripe, my friend Alex") — Claude asks follow-ups and generates referral-specific outreach

## Resume template requirement ⚠️

The resume generator clones your uploaded `.docx` to preserve YOUR formatting. For this to work, at least one of your `.docx` resumes must contain:

- A paragraph with the literal text **`Experience`** as a section header
- A paragraph with the literal text **`Additional Information`** as a section header
- Between them, at least one complete experience entry with bold company, italic description, bold title, and bulleted accomplishments

If your resumes don't match, Claude will flag it during setup and offer to help restructure one into a usable template. 🛠️

## File layout 📁

```
~/.get-me-a-job/
  config.json                    # 🧾 you — name, email, preferences
  profile.md                     # 🧠 accomplishments, targeting, style
  tracker.xlsx                   # 📊 application tracker (auto-updated)
  resumes/                       # 📥 your uploaded base resumes
  cover-letters/                 # 💌 your past cover letters (reference)
  applications/
    {company-slug}/              # 🏢 one folder per application
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

No v1/v2/v3 chaos. No orphan drafts. ✨

## What it won't do 🚫

- ❌ Fabricate experience you don't have
- ❌ Exceed one page on the resume
- ❌ Stuff keywords unnaturally
- ❌ Store sensitive info (SSN, passport, comp details) unless you explicitly ask

## Scripts reference 📚

All scripts live in `scripts/` and can be invoked directly if you want:

- `generate_resume.py <content.json> <output.docx> --template <template.docx>`
- `generate_cover_letter.py <content.json> <output.docx> [--font NAME] [--size N]`
- `docx_to_pdf.py <file.docx> [file2.docx ...]`
- `update_tracker.py <tracker.xlsx> add <entry.json>` | `update <row#> <field> <value>` | `init`

Example JSON schemas live in `templates/`.

## License 📜

MIT — use, modify, and redistribute freely.

---

Built by **[Anisha Subberwal](https://anishasubs.github.io/portfolio/)** ☕

If this skill saves you from one more `Resume_FINAL_v3_actually_this_one.docx`, a ⭐ on the repo is appreciated.
