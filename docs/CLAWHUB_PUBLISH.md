# ClawHub Publish Guide

This document is for repository maintainers who publish the `ppt2fig-export` skill to ClawHub.

## Prerequisites

- `clawhub` CLI installed
- Logged in with `clawhub login`
- A GitHub Release that already contains:

```text
ppt2fig-cli.exe
```

The skill metadata downloads this exact file from:

```text
https://github.com/elliottzheng/ppt2fig/releases/latest/download/ppt2fig-cli.exe
```

## Skill Source

The skill source directory is:

```text
skills/ppt2fig-export
```

## Publish Command

On Windows, use an absolute path when publishing:

```cmd
clawhub publish "E:\codes\ppt_figure_helper\ppt_figure_helper\skills\ppt2fig-export" --slug ppt2fig-export --name "ppt2fig Export" --version 1.1.0 --tags latest --changelog "Initial ClawHub release."
```

## Verification

Before publishing, you can confirm the CLI is logged in:

```cmd
clawhub whoami
```

After publishing, users can install the skill with:

```bash
clawhub install ppt2fig-export
```
