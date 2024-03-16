# Outlook Custom Tasks <!-- omit in toc -->

## Table of Contents <!-- omit in toc -->

- [Installation](#%69%6E%73%74%61%6C%6C%61%74%69%6F%6E)
- [Usage](#%75%73%61%67%65)
  - [1. Configuration](#%31%2E%2D%63%6F%6E%66%69%67%75%72%61%74%69%6F%6E)
  - [2. Run Configured Tasks](#%32%2E%2D%72%75%6E%2D%63%6F%6E%66%69%67%75%72%65%64%2D%74%61%73%6B%73)
- [Configuration](#%63%6F%6E%66%69%67%75%72%61%74%69%6F%6E)
- [Troubleshooting](#%74%72%6F%75%62%6C%65%73%68%6F%6F%74%69%6E%67)
  - [Everything suddenly slower when running rules??](#%65%76%65%72%79%74%68%69%6E%67%2D%73%75%64%64%65%6E%6C%79%2D%73%6C%6F%77%65%72%2D%77%68%65%6E%2D%72%75%6E%6E%69%6E%67%2D%72%75%6C%65%73%3F%3F)
  - [Error: _"UnicodeEncodeError: 'charmap' codec can't encode characters in position 1-40: character maps to \<undefined\>"_](#%65%72%72%6F%72%3A%2D%5F%22%75%6E%69%63%6F%64%65%65%6E%63%6F%64%65%65%72%72%6F%72%3A%2D%27%63%68%61%72%6D%61%70%27%2D%63%6F%64%65%63%2D%63%61%6E%27%74%2D%65%6E%63%6F%64%65%2D%63%68%61%72%61%63%74%65%72%73%2D%69%6E%2D%70%6F%73%69%74%69%6F%6E%2D%31%2D%34%30%3A%2D%63%68%61%72%61%63%74%65%72%2D%6D%61%70%73%2D%74%6F%2D%5C%3C%75%6E%64%65%66%69%6E%65%64%5C%3E%22%5F)
  - [Error: _"RuntimeError: Failed to connect to Outlook locally"_](#%65%72%72%6F%72%3A%2D%5F%22%72%75%6E%74%69%6D%65%65%72%72%6F%72%3A%2D%66%61%69%6C%65%64%2D%74%6F%2D%63%6F%6E%6E%65%63%74%2D%74%6F%2D%6F%75%74%6C%6F%6F%6B%2D%6C%6F%63%61%6C%6C%79%22%5F)
- [Development](#%64%65%76%65%6C%6F%70%6D%65%6E%74)
  - [Prerequisites (for Developers)](#%70%72%65%72%65%71%75%69%73%69%74%65%73%2D%28%66%6F%72%2D%64%65%76%65%6C%6F%70%65%72%73%29)
    - [Windows](#%77%69%6E%64%6F%77%73)
  - [Installation (for Developers)](#%69%6E%73%74%61%6C%6C%61%74%69%6F%6E%2D%28%66%6F%72%2D%64%65%76%65%6C%6F%70%65%72%73%29)
    - [Windows](#%77%69%6E%64%6F%77%73-1)
- [Notices](#%6E%6F%74%69%63%65%73)
- [Special Thanks](#%73%70%65%63%69%61%6C%2D%74%68%61%6E%6B%73)

## Installation

<!-- TODO: add non-developer installation instructions -->

TODO

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

### Error: _"RuntimeError: Failed to connect to Outlook locally"_

OCT tried asking Outlook to connect locally but we didn't hear back from Outlook.

Try force closing Outlook from Task Manager, then try running OCT again.

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

## Notices

This project is not affiliated with Microsoft or Outlook.

## Special Thanks

- [Mark Hammond](https://github.com/mhammond) and [all the other contributors](https://github.com/mhammond/pywin32/graphs/contributors) to the [pywin32](https://github.com/mhammond/pywin32) library used to connect to Outlook locally
