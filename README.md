# Resume Creator

- [Python resume creator with JSON data source - ClaudeAI](https://claude.ai/chat/ab05ad06-2bc9-40e8-b1d8-ca9ac2dd0667)

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
- renders **skills as comma-separated rows**, which parsers split cleanly on the commas.

## Install

```
pip install python-docx
# or, with uv:
uv add python-docx
```

## Use

```
# Build from one of the sample files
python resume_builder.py -i sample_one_page.json -o resume.docx

# Your real resume
python resume_builder.py -i my_resume.json -o resume.docx
```

Start by copying `template.json`, filling it in, and pointing the tool at it.

## Options

| Flag | Default | Purpose |
|---|---|---|
| `-i, --input` | (required) | JSON data file |
| `-o, --output` | `resume.docx` | output path |
| `--skills-per-row N` | off | force exactly N skills per row |
| `--skills-line-width N` | 90 | when packing by width, approx max chars per skills row |
| `--margin N` | 0.6 | page margin in inches (raise to shrink content, lower to fit more) |

By default skills are **width-packed** into balanced rows. Use `--skills-per-row`
if you'd rather have a fixed count per line.

## Files

- `resume_builder.py` — the app
- `template.json` — empty schema to populate
- `sample_one_page.json` — dummy data that fills ~1 page
- `sample_two_page.json` — dummy data that fills ~2 pages

## JSON schema

```
basics: name, title, location, phone, email, website, linkedin, github, summary
experience[]: position, company, location, start_date, end_date, highlights[]
education[]: degree, field, institution, location, start_date, end_date, details[]
certifications[]: name, issuer, date
skills[]: flat list of strings
```

Any field left blank is skipped; `experience`, `education`, `certifications`, and
`skills` are all variable-length.
