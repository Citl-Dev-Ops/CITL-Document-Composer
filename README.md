# CITL Document Composer v1.0

> **Shared Authorship** — CITL Development Team, Renton Technical College
> **Project Lead:** Doc McDowell

AI-powered professional document generator. Uses local Ollama LLMs to draft grant proposals, state policy briefs, technical training guides, and walkthrough documents — with full Word formatting: Georgia serif fonts, navy/maroon color scheme, shaded boxes, keyword blocks, and skills tables.

## What It Does

- **Template Library** — Grant/Proposal, State Policy Brief, Technical Training Guide, IT Walkthrough, Screenshot Walkthrough
- **Local LLM generation** — streams content from any Ollama model with per-template system prompts
- **Word output** — `.docx` files with professional formatting via `python-docx`
- **Grant-ready styling** — navy headers, maroon accents, shaded info boxes, resume keyword blocks
- **Screenshot-aware walkthroughs** — auto-inserts annotated screenshot evidence into technical docs

## Skills Trained (IT Workforce Portfolio)

| Skill | Job Board Keyword |
|-------|------------------|
| AI content generation | Generative AI, prompt engineering |
| Technical writing | Technical writer, documentation |
| Grant/proposal writing | Grant writing, RFP, state funding |
| Python document automation | python-docx, Word automation |
| LLM integration | Ollama, LLM API, AI tools |

## Portfolio Output

Publication-quality Word documents: grant proposals, training guides, or policy briefs — AI-drafted and formatted, ready to submit or present.

## Quick Start

```bash
pip install -r requirements.txt
python citl_doc_composer.py
```

Requires [Ollama](https://ollama.ai) running locally.


## Authors & Contributors

| Name | Role |
|------|------|
| **Doc McDowell** | Project Lead, CITL Director of Instructional Technology |
| **Abdo Mohammed** | Lead Developer — Factbook AI Engine & RAG Systems |
| **Wahaj Al Obid** | Lead Developer — Academic Advisor v2.0 |
| **Jerome Anti Porta** | Developer — UI/UX, App Integration |
| **Jonathan Reed** | Developer — LLMOps & Model Management |
| **Peter Anderson** | Developer — AV/IT Operations & Network Tools |
| **Will Cram** | Developer — Sync Systems & Portable Deployment |
| **William Grainger** | Developer — Technical Writing & Documentation Tools |
| **Mason Jones** | Developer — Staff Toolkit & Field Apps |

> Renton Technical College — Center for Instructional Technology & Learning (CITL)
> Department of IT & Cybersecurity Workforce Development
