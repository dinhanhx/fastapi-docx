# FastAPI server doing .docx

This api server needs to be on a Windows machine with MS Word Office. The api server calls MS Word Office 2024 via `pywin32`, then process .docx file. This service is stateless, and only stores files temporarily during the process.
- `.doc` to `.docx`
- `.docx` to `.pdf`
- `.docx` to `.html`

## Setup Windows

- Install Windows 10
- Install MS Word Office 2024

(See Links section)

## Run with uv

All commands executed in wherever this repos is cloned to.

```bash
uv python pin 3.10
uv python install 3.10
```

## With uv

```bash
uv run -m fastapi-docx.main
```

### With uv but a bit old school or legacy

Make a local .venv folder
```bash
uv venv --seed
```

Activate it, install dependencies, and just run,
```bash
source .venv/bin/activate
pip install -e .
python -m fastapi-docx.main
```

Note: can use other .venv creators, just need to activate it properly.

## Links

https://massgrave.dev/