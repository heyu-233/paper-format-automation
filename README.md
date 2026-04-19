# paper-format-automation

A Codex skill repository for template-driven Chinese journal manuscript formatting.

## Recommended install

Install the skill from the packaged subdirectory:

```powershell
python install-skill-from-github.py `
  --repo heyu-233/paper-format-automation `
  --path skills/paper-format-automation
```

If you use a GitHub URL form, point to the skill path explicitly:

```powershell
python install-skill-from-github.py `
  --url https://github.com/heyu-233/paper-format-automation/tree/main/skills/paper-format-automation
```

## Repository layout

```text
.
|-- LICENSE
|-- README.md
`-- skills/
    `-- paper-format-automation/
        |-- SKILL.md
        |-- README.md
        |-- CHANGELOG.md
        |-- ROADMAP.md
        |-- agents/
        |-- examples/
        |-- references/
        `-- scripts/
```

## Skill entry

The actual installable skill package lives at:

- `skills/paper-format-automation`

See `skills/paper-format-automation/README.md` for usage, workflow, and local tooling details.

## Development

Run the local regression checks with:

```powershell
python -m unittest discover -s tests -v
```

A matching GitHub Actions workflow runs the same test suite on pushes and pull requests.
