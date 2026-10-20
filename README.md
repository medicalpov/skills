# medicalpov skills

Claude skills for clinicians, from [medicalpov.ai](https://medicalpov.ai).

| Skill | What it does |
|---|---|
| [`phi-scrub-slides`](phi-scrub-slides/) | Remove patient identifiers (PHI) from PowerPoint and Google Slides decks, including photo EXIF/GPS, notes, comments, hidden slides and metadata, with a per-image review checklist. Fail-closed; runs offline with plain Python. |

Each skill folder has a `SKILL.md` (instructions for Claude), a `README.md` (for people), `scripts/` and `tests/`.

## Install

Copy a skill folder into your Claude skills directory (Claude Code: `~/.claude/skills/`), or upload it
as a skill in the Claude app. Every skill here also documents how to run its scripts without AI.

## Important

These tools reduce risk; they do not certify compliance and are not legal advice. Follow your
institution's policies.

## License

MIT. See [LICENSE](LICENSE).
