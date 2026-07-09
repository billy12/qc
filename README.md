# Queen Creek Interview Prep

Interview prep materials for the Recreation Coordinator – Special Events position
at the Town of Queen Creek Parks & Recreation.

---

## Getting Set Up on a New Computer

Follow these steps in order. Should take about 5 minutes.

---

### Step 1 — Make Sure Claude Code Is Installed

Open a terminal (Command Prompt or PowerShell on Windows, Terminal on Mac) and run:

```bash
claude --version
```

If you get a version number, skip to Step 2.

If you get an error, install Claude Code first:

```bash
npm install -g @anthropic-ai/claude-code
```

If `npm` isn't found either, you need to install Node.js first: https://nodejs.org
Then re-run the install command above.

---

### Step 2 — Clone This Repo

Pick a folder on your computer where you want the files to live, then run:

```bash
git clone https://github.com/billy12/qc.git
cd qc
```

---

### Step 3 — Install the Skills

This copies the resume and interview skills into Claude Code so they're available as slash commands.

**On Windows (PowerShell):**
```powershell
Copy-Item -Recurse "claude-skills\*" "$env:USERPROFILE\.claude\skills\"
```

**On Mac/Linux:**
```bash
cp -r claude-skills/* ~/.claude/skills/
```

---

### Step 4 — Open Claude Code in the Project Folder

```bash
claude
```

This opens Claude Code inside the `qc` folder.

---

### Step 5 — Load Your Context

Once Claude Code is open, paste this message to get up to speed instantly:

```
Please read HANDOVER.md and PROFILE.md — I'm continuing interview prep
for the Queen Creek Recreation Coordinator role and need to pick up where we left off.
```

Claude will read both files and be fully briefed on:
- The job, the homework, and the deadlines
- The Spring Into QC event and the recommendation we built
- Your background and how it maps to the role
- What files exist and what's left to do

---

### Step 6 — Regenerate the PowerPoints (if needed)

If the `.pptx` files didn't transfer or you want a fresh build:

```bash
pip install python-pptx
python build_presentation.py
python build_qa_deck.py
```

This rebuilds both decks in the same folder.

---

## What's in This Repo

| File / Folder | What It Is |
|---------------|------------|
| `HANDOVER.md` | Full session context — start here in a new chat |
| `PROFILE.md` | Your background, skills, and how they map to the QC role |
| `CLAUDE_SKILLS.md` | Reference guide for all available slash commands |
| `claude-skills/` | The actual skill files — installed in Step 3 |
| `SpringIntoQC_Presentation.pptx` | Main 9-slide executive recommendation deck |
| `SpringIntoQC_QA_Prep.pptx` | 9-slide Q&A prep deck with 7 challenges + responses |
| `build_presentation.py` | Script that builds the main deck |
| `build_qa_deck.py` | Script that builds the Q&A deck |
| `job details.pdf` | Original job listing |
| `interview email.pdf` | Email with homework instructions |
| `resume william.pdf` | Your resume |

---

## Key Dates

| | |
|-|-|
| Homework due | Monday, July 13, 2026 at noon |
| Submit to | mark.miller@queencreekaz.gov AND mckinna.evans@queencreekaz.gov |
| Interview | Wednesday, July 15, 2026 |
| Location | Queen Creek Recreation & Aquatic Center, 22343 Ryan Rd, Queen Creek, AZ 85142 |

---

## Useful Skills to Run at Home

Once set up, try these in Claude Code:

| What you want | What to say |
|---------------|-------------|
| Practice Q&A | `/interview-prep-generator` |
| Tailor resume to this job | `/resume-tailor` |
| Strengthen resume bullets | `/resume-bullet-writer` |
| Analyze the job posting | `/job-description-analyzer` |
| Write a cover letter | `/cover-letter-generator` |
| Frame nonprofit → municipal transition | `/career-changer-translator` |
