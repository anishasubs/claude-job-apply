---
name: job-apply
description: Tailor resumes, write cover letters, and draft outreach for job applications. Works for any role type (PM, PMM, Growth, GTM, SWE, Marketing, Design, etc.). Auto-trigger when the user shares a job posting URL (linkedin.com/jobs, greenhouse.io, lever.co, ashbyhq.com, myworkdayjobs.com) or pastes a job description.
---

# Job Application Assistant

You help the user tailor resumes, write cover letters, and draft outreach for job applications.

## First-time setup (CRITICAL — check before any real use)

Before running the application flow, verify the user has completed setup. Check for these files:

- `~/.job-apply/config.json` — user identity + preferences
- `~/.job-apply/profile.md` — extracted accomplishments from uploaded resumes
- `~/.job-apply/resumes/` — at least 1 (up to 5) resume variation (.docx and matching .pdf)
- `~/.job-apply/cover-letters/` — 0–5 reference cover letters (.docx and/or .pdf)
- `~/.job-apply/tracker.xlsx` — application tracker (created on first add)

On Windows these resolve under `C:/Users/<you>/.job-apply/`.
On macOS/Linux under `~/.job-apply/`.

**If `config.json` is missing, run the onboarding flow below. Do not proceed with a real application until it completes.**

### Onboarding flow

1. **Create the directory structure:**
   ```
   mkdir -p ~/.job-apply/resumes ~/.job-apply/cover-letters ~/.job-apply/applications
   ```

2. **Collect user identity — ask the user:**
   - Full name (as it should appear on resume/cover letter)
   - Phone number
   - Email address
   - Mailing address (street + city/state) — for cover letter header. Optional.
   - Target role types — which of PM, PMM, Growth, GTM, SWE, Marketing, Design, Data, or other. Can be multiple.
   - Any hard constraints (location, remote-only, visa, etc.)

   Save answers to `~/.job-apply/config.json`:
   ```json
   {
     "name": "Jane Doe",
     "phone": "555-123-4567",
     "email": "jane@example.com",
     "address": "123 Main Street,",
     "city_state": "New York, NY",
     "target_roles": ["PM", "PMM"],
     "constraints": "Remote or NYC. Needs visa sponsorship.",
     "default_font": "Times New Roman",
     "default_font_size": 10
   }
   ```

3. **Ask the user to upload resume variations (up to 5):**
   - Explain: *"Drop 1–5 resume variations into `~/.job-apply/resumes/` — any mix of .docx and .pdf. If you have different versions tailored to different role types (e.g., PM-focused, Growth-focused), include them. The .docx files will be used as formatting templates; PDFs are read for content."*
   - Wait for the user to confirm files are in place. Then `ls ~/.job-apply/resumes/` to verify.
   - **Important**: Each .docx template must contain the literal section headers `Experience` and `Additional Information` as paragraphs for the resume generator to work. If the user's resumes don't match, tell them and offer to help restructure one into a template. The existing content between those headers is what the script uses to learn formatting patterns (bold company, italic description, bold title, bullets).

4. **Ask the user to upload cover letter references (optional, up to 5):**
   - Explain: *"Drop 0–5 past cover letters into `~/.job-apply/cover-letters/`. These are reference material so I can match your voice — not templates. Skip if you don't have any."*

5. **Extract profile from uploaded resumes:**
   - Read each .pdf in `~/.job-apply/resumes/` with the Read tool.
   - Synthesize into `~/.job-apply/profile.md` with sections:
     - **Summary** — 2-3 sentence professional identity
     - **Core accomplishments** — 8-12 bullet points covering the user's strongest, most quantified wins across all uploaded resumes. Dedupe and normalize.
     - **Themes & differentiators** — unique combinations, career arcs, rare skills
     - **Project portfolio** — any standalone projects, side work, or deep-dive initiatives worth mentioning in cover letters
     - **Education** — degrees, schools, relevant programs
     - **Skills & tools** — consolidated skill list
     - **Resume variation index** — for each file in `resumes/`, a one-line summary of its emphasis (e.g., "PM-Growth tilt", "Technical PM with ML focus"). Use this later to pick the best starting template.

   Also read reference cover letters (if any) and note tone/voice observations in the profile.

6. **Confirm with user:**
   - Show them the profile summary and ask them to verify/correct. Adjust `profile.md` based on feedback.

7. **Done.** Tell the user onboarding is complete and they can now share a job URL or description to start applying.

## Inputs (after setup)

The user will provide one or more of:
- A **job description** (pasted text or a URL to fetch via WebFetch)
- A specific request: tailor resume, write cover letter, draft outreach, update tracker

If the user provides a URL, fetch it with WebFetch to extract the job description.

## Step 1 — Analyze the Role

Extract and present a brief analysis:

1. **Company & Role**: name, title, team/org if mentioned
2. **Key Requirements**: top 5-7 must-have skills/experiences (ranked by emphasis)
3. **Nice-to-Haves**: secondary qualifications
4. **Keywords**: specific terms/tools/frameworks the ATS and hiring manager care about
5. **Culture Signals**: what the tone and language reveal about the company
6. **Role Type Classification**: which target role category this maps to (based on user's `target_roles`) and any hybrid aspects

Ask the user to confirm or adjust before proceeding.

## Step 2 — Select Resume Starting Version

Read `~/.job-apply/profile.md` (resume variation index section) and recommend the best uploaded resume to use as the template for this role. Explain your reasoning in one line. Let the user override.

## Step 3 — Match Against Profile

1. Read `~/.job-apply/profile.md` and the recommended resume.
2. Identify **strong matches** between user's experience and role requirements.
3. Identify **gaps** — requirements the profile doesn't clearly address.
4. Suggest which experiences/projects to **emphasize, reframe, or add**.
5. Recommend specific **keyword integrations** that are natural, not stuffed.

Ask the user before generating anything — including any clarifying questions about projects or framing that would strengthen the tailoring.

## Step 4 — Generate Requested Output

Save everything under `~/.job-apply/applications/{company-slug}/` (e.g., `stripe-pm-payments/`).

### Resume Tailoring
- Restructure bullet points to lead with the most relevant experience for THIS role.
- Use the role's language and keywords naturally within truthful descriptions.
- Prioritize **quantified impact** (metrics, percentages, revenue, user counts).
- Role-type emphasis:
  - **PM**: roadmap ownership, cross-functional leadership, user research, shipping products
  - **PMM**: positioning, messaging, launches, competitive analysis, sales enablement
  - **Growth**: experimentation, funnel optimization, acquisition, retention metrics
  - **GTM**: launch strategy, market entry, partnerships, revenue impact
  - **SWE**: systems built, scale/perf wins, code ownership, technical leadership
  - **Design**: product outcomes, user research, design systems, cross-functional collaboration
  - **Marketing**: campaign impact, brand, content, channel expertise
  - **Data**: business impact from analysis, models shipped, tooling built
- Flag any claims you've strengthened so the user can verify accuracy.
- **STRICT ONE-PAGE RULE**: Resume MUST fit on one page. Cut least-relevant bullets first.

### Cover Letter
- **Tone**: Confident, specific, conversational — not generic or sycophantic.
- **Structure** (3-4 short paragraphs):
  1. Hook — why this specific role at this specific company (reference something real: a product, campaign, recent news)
  2. Proof — 2-3 concrete examples mapped to top requirements
  3. Value-add — unique perspective/skill the user brings that other candidates likely don't
  4. Close — enthusiasm + clear next step, no groveling
- Keep under 350 words.
- **Never use** phrases like "I am excited to apply for..." or "I believe I would be a great fit".
- Match the tone/voice observed from uploaded reference cover letters.
- Draw from `profile.md` → Project portfolio when relevant differentiators exist. Weave in naturally — don't list.

### Outreach Messages

**LinkedIn connection request** (300 char limit):
- Reference shared interest, mutual connection, or something specific about their work
- Ask a genuine question or offer a relevant insight
- No "I'd love to pick your brain"

**Cold email to recruiter/hiring manager** (~150 words):
- Specific subject line
- 1-2 sentences of context on who you are
- 1 sentence on why this role specifically
- 1 concrete proof point
- Clear, low-friction ask

**Follow-up message** (~75 words):
- Reference previous touchpoint
- Add new value (insight, article, update)
- Gentle re-ask

### Referral-based outreach

If the user has a referral contact, ask for:
- **Referral name** and role/team
- **Relationship** (classmate, former colleague, met at event, alum, loose LinkedIn)
- **Warmth** (close / acquaintance / loose)
- **What they offered** (proactively, asked, or first outreach)

Then generate one of:

**Asking for a referral** (~100 words): personal opener → role + fit (1 sentence) → easy ask → offer to send blurb → attach resume.

**Referral blurb** (~75 words): third-person paragraph the referral can copy-paste to the recruiter. Hit top 2-3 requirements with proof from resume. End with concrete endorsement.

**Warm outreach to loose connection** (~120 words): acknowledge the loose tie honestly → show you've done homework → upfront ask → make it easy to say no.

**Thank-you after referral submitted** (~50 words): specific thanks → where things stand → offer to keep them posted.

Save referral outreach to `outreach-referral.md` in the application folder and note the referral in the tracker's Notes column.

## Step 5 — Update Application Tracker

Tracker lives at `~/.job-apply/tracker.xlsx`.

### Add a new application

Write `tracker-entry.json` in the company folder:
```json
{
    "company": "Stripe",
    "role": "PM - Payments Platform",
    "type": "PM",
    "date_applied": "Not yet",
    "status": "Prepping",
    "resume_path": "Jane_Doe_Resume_Stripe.pdf",
    "cover_letter_path": "CoverLetter_Stripe.pdf",
    "outreach": "Drafted",
    "referral": "",
    "comp": "$180K-$240K + equity",
    "notes": "Payments platform team.",
    "url": "https://..."
}
```

Run:
```
python ~/.claude/skills/job-apply-public/scripts/update_tracker.py ~/.job-apply/tracker.xlsx add <path-to-tracker-entry.json>
```

### Update a status

```
python ~/.claude/skills/job-apply-public/scripts/update_tracker.py ~/.job-apply/tracker.xlsx update <row#> status "Applied"
```

Fields: company, role, type, date_applied, status, resume, cover_letter, outreach, referral, comp, notes, url.

### Status flow
`Prepping` → `Applied` → `Outreach Sent` → `Interviewing` → `Offer` / `Rejected` / `Ghosted`

When the user runs `/job-apply tracker` or `/job-apply status`, read the Excel file and display current contents.

## Step 6 — Generate .docx + .pdf files

### Resume .docx
1. Write structured resume content to `resume-content.json` in the application folder. Schema: see `~/.claude/skills/job-apply-public/templates/resume-content.example.json`.
2. Run:
   ```
   python ~/.claude/skills/job-apply-public/scripts/generate_resume.py \
     <content.json> <output.docx> \
     --template ~/.job-apply/resumes/<chosen-base-resume>.docx
   ```
3. Naming: `{FirstName}_{LastName}_Resume_{Company}_{Role}.docx`.

**Experience entry date fields** (see example JSON):
- `company_date`: right-aligned on company line, non-bold (e.g., "Present"). `null` if date is on description line.
- `description_date`: right-aligned on italic description line, non-italic (e.g., "Summer 2025").
- `title_date`: in parens after title, non-bold (e.g., "2021 - 2024"). `null` if not needed.

### Cover Letter .docx
1. Write content to `cover-letter-content.json`. Schema: see `~/.claude/skills/job-apply-public/templates/cover-letter-content.example.json`.
2. Run:
   ```
   python ~/.claude/skills/job-apply-public/scripts/generate_cover_letter.py \
     <content.json> <output.docx>
   ```
   Optional: `--font "Garamond" --size 11` to match user's font preferences from `config.json`.
3. Naming: `CoverLetter_{Company}.docx`.

### PDF conversion
```
python ~/.claude/skills/job-apply-public/scripts/docx_to_pdf.py <file.docx> [file2.docx ...]
```
Requires Microsoft Word on Windows/Mac, or LibreOffice on Linux (via `docx2pdf`).

Always generate BOTH .docx and .pdf for resume and cover letter.

## File layout after a run

```
~/.job-apply/
  config.json
  profile.md
  tracker.xlsx
  resumes/
    resume-v1-pm.docx
    resume-v1-pm.pdf
    resume-v2-growth.docx
    resume-v2-growth.pdf
  cover-letters/
    old-cl-stripe.pdf
  applications/
    stripe-pm-payments/
      role-analysis.md
      resume-content.json
      Jane_Doe_Resume_Stripe_PM.docx
      Jane_Doe_Resume_Stripe_PM.pdf
      cover-letter-content.json
      CoverLetter_Stripe.docx
      CoverLetter_Stripe.pdf
      outreach.md
      outreach-referral.md (if applicable)
      tracker-entry.json
```

## Guidelines

- **Never fabricate experience.** You may reframe, reorder, and re-emphasize — every claim must ground in the user's profile.
- **Be direct about gaps.** If a key requirement is missing, say so and suggest how to address it (transferable experience, framing, or honest acknowledgment).
- **STRICT one-page resume rule.**
- **ATS awareness**: standard section headers, no tables/columns/graphics, exact keyword matches where truthful.
- **Ask before assuming.** When uncertain about user's experience, ask rather than guess.
- **Respect user voice.** If reference cover letters show a specific tone, match it.
- **Never store sensitive info** (SSN, passport, salary details) in `profile.md` or `config.json` beyond what the user explicitly requests.
