# Resume Creator

- [Python resume creator with JSON data source - ClaudeAI - PRIVATE](https://claude.ai/chat/ab05ad06-2bc9-40e8-b1d8-ca9ac2dd0667)

A small CLI that builds an **ATS-friendly Word (.docx) resume** from a JSON file.

## Why it's built this way (ATS notes)

Applicant Tracking Systems flatten a document into a single stream of text
before parsing it, so the layout deliberately:

- uses **one column, top-to-bottom** — the order you read is the order the ATS reads;
- uses **no tables** (the date on each entry is placed with a right-aligned *tab
  stop*, and the line under each section header is a *paragraph border* — neither
  is a table);
- uses **standard section names** (`Summary`, `Experience`, `Education`,
  `Certifications`, `Skills`) that parsers recognize;
- renders **skills as comma-separated rows**, which parsers split cleanly on the commas;
- makes **contact details clickable** (email → `mailto:`, phone → `tel:`,
  website / LinkedIn / GitHub → `https://`) while keeping the short display text
  from the JSON.

## Install

```
pip install python-docx
# or, with uv:
uv add python-docx
```

## Use

```
# Build from one of the sample files
uv run python resume_builder.py -i sample_one_page.json -o out/resume.docx
uv run python resume_builder.py -i sample_two_page.json -o out/resume.docx

# Force a fixed number of skills per row
uv run python resume_builder.py -i sample_one_page.json -o out/resume.docx --skills-per-row 10

# Render everything, ignoring "include": false flags
uv run python resume_builder.py -i sample_two_page.json -o out/resume_all.docx --all

# Your real resume (kept in a separate private repo)
uv run python resume_builder.py -i ../CAREER-REPO/bruce_stull_resume.json -o out/resume.docx
```

Start by copying `template.json`, filling it in, and pointing the tool at it.

## Options

| Flag | Default | Purpose |
|---|---|---|
| `-i, --input` | (required) | JSON data file |
| `-o, --output` | `resume.docx` | output path (parent folders are created if needed) |
| `--skills-per-row N` | off | force exactly N skills per row |
| `--skills-line-width N` | 90 | when packing by width, approx max chars per skills row |
| `--margin N` | 0.6 | page margin in inches (raise to shrink content, lower to fit more) |
| `--all` | off | ignore every `"include": false` flag and render all entries and skills |

By default skills are **width-packed** into balanced rows. Use `--skills-per-row`
if you'd rather have a fixed count per line.

## Hiding entries without deleting them

Everything is **included by default**. To keep an entry in your JSON but leave
it off the rendered resume, add `"include": false`:

- **Experience, education, and certification entries:** add `"include": false`
  to the entry object.
- **Skills:** write the skill as an object instead of a plain string:
  `{ "name": "Java", "include": false }`.

Plain-string skills are always rendered. `"include": true` is allowed and
behaves the same as leaving the flag off; it's handy for toggling. Use `--all`
to preview the full, unfiltered resume.

This lets one JSON file act as a master record of everything you've done, with
each application showing only what's relevant.

## Title options

`basics.title_options` is an optional list of alternate professional titles.
The builder **does not render it**; it's there so you can keep your vetted
titles in one place and copy the right one into `basics.title` when tailoring
a resume for a specific posting.

## Files

- `resume_builder.py` — the app
- `template.json` — empty schema to populate (shows every supported field)
- `sample_one_page.json` — dummy data that fills ~1 page (includes hidden entries/skills as examples)
- `sample_two_page.json` — dummy data that fills ~2 pages (includes hidden entries/skills/certs as examples)

## JSON schema

```
basics: name, title, title_options[], location, phone, email, website,
        linkedin, github, summary
experience[]:     include?, position, company, location, start_date, end_date, highlights[]
education[]:      include?, degree, field, institution, location, start_date, end_date, details[]
certifications[]: include?, name, issuer, date
skills[]:         mix of plain strings and { "name": ..., "include": bool } objects
```

`?` marks an optional field. Any field left blank is skipped, and `experience`,
`education`, `certifications`, and `skills` are all variable-length. Any
section that is missing or empty is omitted entirely, header and all.
