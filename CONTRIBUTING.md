# Contributing

This is a fork of [Aanerud/outlook-desktop-mcp](https://github.com/Aanerud/outlook-desktop-mcp)
(AGPL-3.0). `main` is the branch people install from: the README's setup clones it, and
"update it" pulls it. A push to `main` is therefore a release, so keep `main` working. The
offline unit tests run in GitHub Actions on every push (`.github/workflows/tests.yml`); the
PyPI workflow is manual-only and this fork does not publish there.

## Working on it

- Install the way a user would: open the folder in Claude Code and say "let's get started",
  or run `install.ps1`. The server is installed in editable mode, so a code change in this
  folder runs the next time Claude starts the server. Nothing needs reinstalling unless
  `pyproject.toml` gains a dependency; then run `install.ps1` again (it adds libraries to an
  existing `.venv`).
- Offline unit tests, no Outlook needed:

  ```powershell
  .\.venv\Scripts\python.exe -m unittest discover -s tests
  ```

- Live checks, with classic Outlook open: `tests\live_check.py` reads only and changes
  nothing; `tests\perf_check.py` times the searches that used to be slow. The
  `phase*`, `calendar*` and `extras*` scripts write a draft or an unsent meeting that you
  delete afterwards. Nothing in `tests/` sends.

## Making a change

1. Branch from `main`, make the change, open a pull request against `main`.
2. Keep the drafts-only rule: no tool may send, accept, decline or cancel anything.
3. When tools change, update the tool tables and the tool count in `README.md`, and the
   server instructions in `server.py` if a new category appears.
4. Bump the version in `pyproject.toml` and `src/outlook_desktop_mcp/__init__.py` and add
   a `CHANGELOG.md` entry.

## Adding a tool

1. Write the COM function that does the work. It receives `outlook, namespace` first and
   runs on the single COM thread.
2. Add an `@mcp.tool()` async handler in `server.py` that calls `bridge.call(your_function, ...)`.
3. Write the docstring for Claude, not for people: it is what the model reads when choosing
   and calling the tool. Say when to use it, when not to, and what the result contains.
4. Add a unit test for any pure logic (see `tests/test_*.py`) and, if it touches Outlook,
   a line in `tests/live_check.py` if it only reads.
