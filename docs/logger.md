# `src/logger.py` — Logging setup

## What it does
Configures one application-wide logger that writes to **both** the console and a
**timestamped file** under `logs/`. Provides two functions:
- `setup_logging(log_dir)` — call once at startup.
- `get_logger(name)` — call in every module to get a logger.

## Why it exists
This app runs unattended (Task Scheduler) on a remote PC. When something goes
wrong with row 37 at 1 a.m., you need a permanent record. `print()` disappears;
a log file does not. Logging also gives every message a timestamp, level and
source module automatically.

## Why it is built this way

### Console + file with different levels
```python
fh.setLevel(level)            # file: DEBUG and up (full detail)
ch.setLevel(logging.INFO)     # console: INFO and up (clean summary)
```
- The **console** stays readable (INFO and above) — you see progress.
- The **file** keeps everything (DEBUG and above) for later diagnosis.

### One timestamped file per run
```python
log_file = log_dir / f"run_{timestamp}.log"
```
Each run gets its own file (e.g. `run_20260630_013426.log`), so runs never
overwrite each other and you can compare them.

### Guard against duplicate handlers
```python
if root.handlers:
    return
```
If `setup_logging` is called twice (e.g. modules re-imported), you would
otherwise get every log line printed two or three times. This check prevents
that.

### A single named namespace `email_automation`
```python
return logging.getLogger(f"email_automation.{name}")
```
All loggers are children of one parent. That parent owns the handlers, so every
module's logs flow to the same console+file, and you can later filter or silence
the whole app by that one name.

## How other modules use it
```python
from src.logger import get_logger
log = get_logger(__name__)
log.info("Outlook connected")
```
`main.py` calls `setup_logging(cfg.log_dir)` once; everything else just calls
`get_logger`.
