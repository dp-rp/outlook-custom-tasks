# Outlook Custom Tasks

## Development

### Prerequisites (for Developers)

#### Windows

```powershell
# NOTE: do not use an elevated (admin) terminal

git clone https://github.com/dp-rp/outlook-custom-tasks.git
cd outlook-custom-tasks

# Install pipx (used to install Poetry in it's own isolated venv)
py -m pip install --user pipx  # if you installed Python using Microsoft Store, replace `py` with `python3`
.\.util-scripts\configure_pipx_windows.ps1
```

Open a new terminal

```powershell
# NOTE: do not use an elevated (admin) terminal

# install Poetry (using Pipx)
pipx install poetry  # install Poetry
poetry -V  # test Poetry installed successfully
```

### Installation (for Developers)

**NOTE:** Ensure you have already installed the prerequisites by following [these instructions](#prerequisites-for-developers)

The following commands should be ran from the root of your copy of the project repository

#### Windows

```bash
# NOTE: do not use an elevated (admin) terminal

poetry install
```

## Usage

### 1. Configuration

Create your own `oct.settings.json` configuration file to define your rules.

You can use [`oct.settings.example.json`](./oct.settings.example.json) as an example.

To learn about how to write your own `oct.settings.json` file, see [Configuration](#configuration).

### 2. Run Configured Tasks

```bash
oct # if you're a developer, prefix this with `poetry run`
```

## Configuration

<!-- TODO: fill this out once an official json schema is implemented w/ validation -->

TODO

## Troubleshooting

### Everything suddenly slower when running rules??

If you're on a laptop or other portable system, your system might have switched over to a power-saving battery profile automatically - this often happens when you take your laptop/device off charge.

Try putting your device on charge to quickly check if that's the source of performance suddenly getting worse.

Depending on your system you should be able to change when different battery profiles are used, but only do this if you know what you're doing. If you really need to run some particularly demanding rules, you may be better off just putting your device on charge.

### Error: _"UnicodeEncodeError: 'charmap' codec can't encode characters in position 1-40: character maps to \<undefined\>"_

Unfortunately colorama (the library used to provide coloured output in the terminal) doesn't play nicely with GitBash on Windows.

Try running OCT again from Powershell to see if it resolves the issue.

## Notices

This project has no direct associations with Microsoft or the people responsible for the development of Microsoft's Outlook.
