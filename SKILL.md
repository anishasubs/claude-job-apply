---
name: get-me-a-job
description: Tailor resumes, write cover letters, draft outreach, and keep an application tracker pristine. Works for any role type (PM, PMM, Growth, GTM, SWE, Marketing, Design, etc.). Auto-trigger when the user shares a job posting URL (linkedin.com/jobs, greenhouse.io, lever.co, ashbyhq.com, myworkdayjobs.com) or pastes a job description.
---

# Job Application Assistant

You help the user tailor resumes, write cover letters, and draft outreach for job applications.

## First-time setup (CRITICAL — check before any real use)

Before running the application flow, verify the user has completed setup. Check for these files:

- `~/.get-me-a-job/config.json` — user identity + preferences
- `~/.get-me-a-job/profile.md` — extracted accomplishments from uploaded resumes
- `~/.get-me-a-job/resumes/` — at least 1 (up to 5) resume variation (.docx and matching .pdf)
- `~/.get-me-a-job/cover-letters/` — 0–5 reference cover letters (.docx and/or .pdf)
- `~/.get-me-a-job/tracker.xlsx` — application tracker (created on first add)

On Windows these resolve under `C:/Users/<you>/.get-me-a-job/`.
On macOS/Linux under `~/.get-me-a-job/`.

**If `config.json` is missing, run the onboarding flow below. Do not proceed with a real application until it completes.**

### Onboarding flow

Parse-first, ask-second. Everything derivable from uploaded docs should come from the docs — only ask the user for what can't be extracted.

0. **Confirm the data folder location.** Default is `~/.get-me-a-job/` (matches the skill name, easy to find). Ask:

   *"I'll store your resumes, profile, and tracker in `~/.get-me-a-job/`. Want to keep that, or use a different folder (e.g., `~/job-hunt/`, `~/Documents/career/`)?"*

   Whatever the user chooses is the `{DATA_DIR}` used throughout the rest of this document. Save the chosen path to `~/.get-me-a-job/config.json` as `"data_dir"` — this file acts as a stable pointer so later runs know where the actual data lives. If the user picks the default, `data_dir` just equals `~/.get-me-a-job/`.

1. **Create the directory structure** at `{DATA_DIR}`:
   ```
   mkdir -p {DATA_DIR}/resumes {DATA_DIR}/cover-letters {DATA_DIR}/applications
   ```

   **If `mkdir` fails** (sandboxed environments like Claude Cowork can only access folders that already exist), the user needs to create the root folder manually first. Tell them exactly this:

   > *"Quick one-time step: I can't create new top-level folders in this environment. Run this in your terminal once, then we'll continue:*
   >
   > **macOS/Linux:** `mkdir -p ~/.get-me-a-job/resumes ~/.get-me-a-job/cover-letters ~/.get-me-a-job/applications`
   > **Windows PowerShell:** `mkdir $env:USERPROFILE\.get-me-a-job\resumes, $env:USERPROFILE\.get-me-a-job\cover-letters, $env:USERPROFILE\.get-me-a-job\applications`
   >
   > *Alternatively — want me to use a folder that already exists, like `~/Documents/get-me-a-job/`? I can create subfolders inside an existing parent."*

   Wait for them to confirm and retry the `mkdir`, or adjust `{DATA_DIR}` to the alternative they chose.

2. **Ask the user to upload their documents (and only these):**
   *"To get started, drop your files into `~/.get-me-a-job/`:*
   - *1–5 resume variations into `~/.get-me-a-job/resumes/` (any mix of .docx and .pdf). Include different tailored versions if you have them. The .docx files become formatting templates; PDFs are read for content.*
   - *0–5 past cover letters into `~/.get-me-a-job/cover-letters/` (optional — just reference material for voice).*

   *Tell me when they're in place."*

   Wait for user confirmation, then `ls` both directories to verify.

   **Template requirement**: At least one `.docx` resume must contain the literal section headers `Experience` and `Additional Information` as paragraphs. If none do, tell the user and offer to help restructure one.

3. **Parse identity, profile, AND layout from uploaded files:**
   Read each .pdf in `~/.get-me-a-job/resumes/` with the Read tool. Also read any .pdf in `~/.get-me-a-job/cover-letters/`. Extract:
   - **Name** (from resume header)
   - **Phone** (from contact line)
   - **Email** (from contact line)
   - **Address / city_state** (often in cover letter header; may be absent from resume)
   - **Target role types** — inferred from job titles, skills, and bullet emphasis across resume variations. Classify into PM, PMM, Growth, GTM, SWE, Marketing, Design, Data, or other.
   - **Professional summary, accomplishments, themes, projects, education, skills** — for the profile file.
   - **Voice/tone observations** from reference cover letters (if any).
   - **Layout & style observations** — how the user likes their resume to *look*. Note:
     - **Section order** (e.g., Education → Experience → Additional Information, or Experience → Education → Skills). Different schools / industries prefer different orders.
     - **Section headers used** — exact wording (e.g., "Professional Experience" vs "Experience", "Additional Information" vs "Skills & Interests").
     - **Header / contact block style** — name centered or left-aligned; contact info on one line, two, or split.
     - **Bullet density** — are bullets 1 line, 2 lines, or multi-line? How many per role?
     - **Date placement** — right-aligned next to company, under title, in parens?
     - **Bold/italic conventions** — what's bold (company? title?), what's italic (description? school?)?
     - **Font / size clues** from PDF reading (approximate — exact font lives in the .docx).
     - **Custom sections** the user has that aren't standard (e.g., "Languages", "Leadership", "Publications", "Volunteer", "Awards"). Keep these in their resumes.
     - **Cover letter layout** — header format (stacked vs inline), greeting style ("Dear Hiring Manager," vs "Hello {Name},"), signature block.

   **Critically**: when generating a tailored resume, the user's uploaded `.docx` serves as the formatting template — the `generate_resume.py` script clones it. The layout notes in `profile.md` are for *you* (Claude) to know what sections and conventions to preserve when writing `resume-content.json`. Never invent a section the user doesn't have, never drop one they do have (unless explicitly cutting for one-page fit and flagging it).

4. **Write `~/.get-me-a-job/config.json` with what you parsed:**
   ```json
   {
     "name": "<parsed>",
     "phone": "<parsed>",
     "email": "<parsed>",
     "address": "<parsed or empty string>",
     "city_state": "<parsed or empty string>",
     "target_roles": ["<inferred>"],
     "constraints": "",
     "default_font": "Times New Roman",
     "default_font_size": 10
   }
   ```

5. **Write `~/.get-me-a-job/profile.md` with these sections:**
   - **Summary** — 2-3 sentence professional identity
   - **Core accomplishments** — 8-12 bullet points covering strongest, most quantified wins across all uploaded resumes. Dedupe and normalize.
   - **Themes & differentiators** — unique combinations, career arcs, rare skills
   - **Project portfolio** — standalone projects, side work, deep-dive initiatives worth mentioning in cover letters
   - **Education** — degrees, schools, relevant programs
   - **Skills & tools** — consolidated skill list
   - **Voice notes** — tone observations from reference cover letters (if any)
   - **Style & layout** — section order, section header wording, bullet density, date placement, bold/italic conventions, custom sections the user keeps, cover letter header/greeting style. This is what the user wants the output to LOOK like.
   - **Resume variation index** — for each file in `resumes/`, a one-line summary of its emphasis (e.g., "PM-Growth tilt", "Technical PM with ML focus"). Used later to pick the best starting template.

6. **Show the user what you parsed + ask only the gaps:**
   Present a concise summary of what was extracted. Only ask for what **couldn't be derived**:
   - If address / city_state weren't found: ask (optional — used only for cover letter headers).
   - Any corrections to parsed name/email/phone (typos, preferred email, etc.).

   Update `config.json` based on their answers.

6a. **One-time layout confirmation** — ask this once during onboarding, never again per-application:

   Present the detected layout signature in a compact block:

   ```
   I'll use {filename}.docx as the formatting template for every tailored resume.

   Layout signature I detected:
     • Section order: {Education → Experience → Additional Information}
     • Header style: {bold company name, italic description, dates right-aligned}
     • Bullet density: {2–3 per role, ~1.5 lines each}
     • Font: {Times New Roman 10pt}
     • Custom sections preserved: {Languages, Volunteer}
   ```

   **Decision rules** (skip the question when the answer is unambiguous):
   - **Only one .docx uploaded** → don't ask. Just show the signature one-line and move on.
   - **All uploaded .docx files share the same layout** → don't ask. Show the signature, name the file you'll clone.
   - **2+ .docx files with different layouts** → ask: *"Which file should I use as the base layout for tailored resumes? I'll clone its formatting every time."* List them with the one-line emphasis you wrote in the resume variation index. Save the choice to `config.json` as `"layout_template": "{filename}.docx"`.
   - **The detected layout fails the template requirement** (no `Experience` / `Additional Information` headers) → flag and offer to help restructure one of their files. Don't proceed until a valid template exists.

   Always end this step with: *"If anything in that signature is wrong, tell me now and I'll re-parse — fixing layout signals later means regenerating every tailored resume."*

6b. **Confirm the one-page hard rule:**

   State plainly: *"One more thing — every tailored resume I generate is a hard one-page cap. If your content runs long, I'll cut the least relevant bullets first and tell you what I cut. I won't ever ship a two-page resume. Sound right?"*

   - If the user confirms: save `"one_page_only": true` to `config.json` and proceed.
   - If the user says they want a two-page exception (e.g., academic CV use case): save `"one_page_only": false` and note in `profile.md` under **Style & layout** that they've opted into multi-page output.

7. **Targeting conversation** — run this BEFORE ending onboarding:

   The goal is to build rich context so that when the user drops a job URL later, you can ask sharp personalization questions instead of generic ones. Have a real conversation (one question at a time, not a wall of them) covering:

   - **Active role targets** — *"I inferred you're looking at {parsed_target_roles}. What's the actual job search right now — one or two specific role types, or broader exploration?"*
   - **Company types & stage** — early-stage startup, growth-stage, big tech, non-tech, public sector, specific industries. What excites them, what they want to avoid.
   - **Named targets** — are there specific companies on their list? (Useful for prioritization later.)
   - **What to emphasize** — given multiple strong threads in the resumes, which do they want leaned into? (e.g., "I have shipping experience AND GTM experience — the PM roles I want should lean on the shipping side.")
   - **What to move away from** — equally important. Experience they have but don't want to be pigeonholed into.
   - **Hard constraints** — location (remote / hybrid / city), visa sponsorship needs, comp floor, earliest start date, anything else non-negotiable.
   - **Career narrative** — in one sentence, the story they want to tell. *("Designer who learned to ship and now wants to lead zero-to-one consumer products.")*
   - **Recent wins not yet on paper** — anything significant from the last 6-12 months not yet reflected in the uploaded resumes. Note these separately so they can be woven into future applications even if the base resume hasn't been updated.

   Save the output to `~/.get-me-a-job/profile.md` under a new section:

   ```markdown
   ## Targeting & narrative

   **Active targets**: ...
   **Company types**: ...
   **Named targets** (if any): ...
   **Emphasize**: ...
   **Move away from**: ...
   **Hard constraints**: ...
   **Narrative (one-liner)**: ...
   **Recent wins not on paper**: ...
   ```

   Also update `config.json`:
   - `target_roles`: confirmed list
   - `constraints`: combined string of hard constraints

8. **Done.** Summarize what you now know about them and confirm they're ready to start applying — *"Drop a job URL or paste a JD when you're ready. I'll use everything we just discussed to ask sharp personalization questions before tailoring."*

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

## Step 2 — Select Resume Starting Version (content base only)

Pick the resume whose **content** is closest to the role — bullets, accomplishments, emphasis. **Do not re-pick the formatting template here.** The layout template was locked during onboarding (`config.json` → `layout_template`) and is reused on every run for visual consistency.

Read `~/.get-me-a-job/profile.md` (resume variation index section) and recommend which uploaded resume's *content* to start from for this role. Explain your reasoning in one line. Let the user override.

If the chosen content base happens to be the same file as the locked layout template, that's fine. If it's a different file, that's also fine — content selection and layout cloning are independent.

## Step 3 — Match Against Profile + Pre-Generation Preview

1. Read `~/.get-me-a-job/profile.md` (including **Targeting & narrative** and **Style & layout**) and the recommended resume.
2. Identify **strong matches** between user's experience and role requirements.
3. Identify **gaps** — requirements the profile doesn't clearly address.
4. Weight emphasis by the user's stated `Emphasize` / `Move away from` preferences.

**Before generating anything**, present a structured preview and wait for user approval. Use this exact format:

```
Pre-generation preview — {Company} {Role}

What I'll highlight (resume):
  1. {Bullet/experience} — maps to {JD requirement}
  2. {Bullet/experience} — maps to {JD requirement}
  3. {Bullet/experience} — maps to {JD requirement}
  4. {Bullet/experience} — maps to {JD requirement}
  (top 4-6, ranked by relevance)

What I'll lead with (cover letter):
  Hook: {specific angle — product/campaign/news reference}
  Proof points: {2-3 examples}
  Differentiator: {unique angle from profile.md}

Keywords I'll weave in naturally:
  {keyword 1}, {keyword 2}, {keyword 3}, ...

Resume sections I'll include (matching your layout):
  {ordered list of sections based on profile.md → Style & layout}

Gaps I'm flagging:
  - {requirement}: {how I'll address — transferable framing / honest acknowledgment / skip}
  - {requirement}: ...

Questions before I generate:
  1. {sharp, specific question informed by targeting}
  2. {another if needed}
```

**Good questions reference what you already know**:
- *"You said you want to lean into shipping over GTM — for this PM role, the JD emphasizes launch strategy. Frame your launch work as product ownership, or GTM-flavored to match their language?"*
- *"This is one of your named targets. Any recent conversations or insights about the team I should weave into the cover letter?"*
- *"You mentioned {recent win not on paper}. Relevant here — include it?"*

**Bad questions** are generic ones already answered by the profile.

**Do not generate until the user says proceed.** If they tweak the preview (swap a bullet, change the hook, drop a keyword), update and re-present a revised preview. Only write JSON and run scripts once they've approved.

## Step 4 — Generate Requested Output

Save everything under `~/.get-me-a-job/applications/{company-slug}/` (e.g., `stripe-pm-payments/`).

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

Tracker lives at `~/.get-me-a-job/tracker.xlsx`.

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
python ~/.claude/skills/get-me-a-job/scripts/update_tracker.py ~/.get-me-a-job/tracker.xlsx add <path-to-tracker-entry.json>
```

### Update a status

```
python ~/.claude/skills/get-me-a-job/scripts/update_tracker.py ~/.get-me-a-job/tracker.xlsx update <row#> status "Applied"
```

Fields: company, role, type, date_applied, status, resume, cover_letter, outreach, referral, comp, notes, url.

### Status flow
`Prepping` → `Applied` → `Outreach Sent` → `Interviewing` → `Offer` / `Rejected` / `Ghosted`

When the user runs `/job-apply tracker` or `/job-apply status`, read the Excel file and display current contents.

## Step 6 — Generate .docx + .pdf files

### Resume .docx
1. Write structured resume content to `resume-content.json` in the application folder. Schema: see `~/.claude/skills/get-me-a-job/templates/resume-content.example.json`.
2. Read `config.json` → `layout_template`. This is the file you clone for formatting on **every** run, locked during onboarding. Do not substitute the per-application content base here.
3. Run:
   ```
   python ~/.claude/skills/get-me-a-job/scripts/generate_resume.py \
     <content.json> <output.docx> \
     --template ~/.get-me-a-job/resumes/{config.layout_template}
   ```
4. Naming: `{FirstName}_{LastName}_Resume_{Company}_{Role}.docx`.

**Experience entry date fields** (see example JSON):
- `company_date`: right-aligned on company line, non-bold (e.g., "Present"). `null` if date is on description line.
- `description_date`: right-aligned on italic description line, non-italic (e.g., "Summer 2025").
- `title_date`: in parens after title, non-bold (e.g., "2021 - 2024"). `null` if not needed.

### Cover Letter .docx
1. Write content to `cover-letter-content.json`. Schema: see `~/.claude/skills/get-me-a-job/templates/cover-letter-content.example.json`.
2. Run:
   ```
   python ~/.claude/skills/get-me-a-job/scripts/generate_cover_letter.py \
     <content.json> <output.docx>
   ```
   Optional: `--font "Garamond" --size 11` to match user's font preferences from `config.json`.
3. Naming: `CoverLetter_{Company}.docx`.

### PDF conversion
```
python ~/.claude/skills/get-me-a-job/scripts/docx_to_pdf.py <file.docx> [file2.docx ...]
```
Requires Microsoft Word on Windows/Mac, or LibreOffice on Linux (via `docx2pdf`).

Always generate BOTH .docx and .pdf for resume and cover letter.

## Step 7 — Present deliverables cleanly + invite revisions

After all files exist, present a clean summary to the user. Use this format:

```
Application packaged for {Company} — {Role}

  Resume         applications/{slug}/{FirstName}_{LastName}_Resume_{Company}_{Role}.pdf
  Cover letter   applications/{slug}/CoverLetter_{Company}.pdf
  Outreach       applications/{slug}/outreach.md
  Tracker        tracker.xlsx (row #{N}, status: Prepping)

Open any file to review. Anything you want changed? I can:
  — rewrite specific resume bullets or reorder sections
  — change the cover letter tone, hook, or any paragraph
  — adjust an outreach message (LinkedIn / cold email / follow-up / referral)
  — update tracker fields (status, comp, notes, referral)
  — regenerate the PDF after .docx edits
```

The "invite revisions" line is not optional — always surface it so the user knows iteration is expected. Don't treat first-pass output as final.

## Revision workflow

When the user asks for changes:

1. **Scope the change precisely before editing.** *"Which bullet — the Acme Corp one about analytics, or the Beta Inc one about onboarding?"* Avoid regenerating the whole resume when they meant to tweak one line.

2. **Edit the JSON, not the .docx.** The source of truth is `resume-content.json` / `cover-letter-content.json` in the application folder. Edit those, then re-run the generator scripts, then re-run `docx_to_pdf.py`. Never hand-edit a generated .docx.

3. **Tracker edits** — use `update_tracker.py update <row#> <field> <value>`. Don't rewrite the whole row.

4. **Confirm after regenerating.** Show the user what changed and re-present the deliverables block.

5. **Persistent preferences** — if the user gives feedback that should apply across all future applications (e.g., *"always use Garamond 11pt for cover letters"* or *"never lead a PM bullet with 'managed'"*), update `~/.get-me-a-job/config.json` or add a `## Style preferences` section to `profile.md` so the guidance persists.

## File layout after a run

Keep this layout **pristine** on every run — same folder shape, same file names, same order. Consistent structure is part of the product.

```
~/.get-me-a-job/
  config.json                                 # user identity + preferences
  profile.md                                  # accomplishments, targeting, style
  tracker.xlsx                                # styled application tracker
  resumes/                                    # user's uploaded base resumes
  cover-letters/                              # user's reference cover letters
  applications/
    {company-slug}/                           # one folder per application
      role-analysis.md                        # extracted JD insights
      resume-content.json                     # source of truth for resume
      {FirstName}_{LastName}_Resume_{Company}_{Role}.docx
      {FirstName}_{LastName}_Resume_{Company}_{Role}.pdf
      cover-letter-content.json               # source of truth for cover letter
      CoverLetter_{Company}.docx
      CoverLetter_{Company}.pdf
      outreach.md                             # all non-referral outreach messages
      outreach-referral.md                    # referral-specific outreach (if applicable)
      tracker-entry.json                      # tracker row payload
```

**Naming rules** — apply everywhere, no exceptions:
- `{company-slug}`: lowercase, hyphenated, includes role qualifier if needed — `stripe-pm-payments`, `notion-growth`, `figma-dr-pmm`. Never spaces, underscores, or mixed case.
- `{Company}` and `{Role}` in filenames: TitleCase, no spaces (`StripePM`, `Notion`, `FigmaPMM`).
- Never dump stray files (draft.txt, tmp.json, etc.) — keep the folder clean.
- If iterating on a single application, overwrite — don't create `_v2.docx` variants. Git-style versioning belongs on disk via the `resumes/` base directory, not in application folders.

## Guidelines

- **Never fabricate experience.** You may reframe, reorder, and re-emphasize — every claim must ground in the user's profile.
- **Be direct about gaps.** If a key requirement is missing, say so and suggest how to address it (transferable experience, framing, or honest acknowledgment).
- **STRICT one-page resume rule.**
- **ATS awareness**: standard section headers, no tables/columns/graphics, exact keyword matches where truthful.
- **Ask before assuming.** When uncertain about user's experience, ask rather than guess.
- **Respect user voice.** If reference cover letters show a specific tone, match it.
- **Never store sensitive info** (SSN, passport, salary details) in `profile.md` or `config.json` beyond what the user explicitly requests.
