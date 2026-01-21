# JobBOSS Material Quantity Update Tool

A two-phase workflow for updating material quantities in JobBOSS via the XML SDK.

## Requirements

- Windows with JobBOSS installed
- JobBOSS COM SDK registered (`JBRequestProcessor.RequestProcessor`)
- Python 3.10+

## Installation

```powershell
# Clone the repo
git clone <repo-url>
cd jobbossutil

# Create and activate virtual environment (recommended)
python -m venv venv
.\venv\Scripts\Activate

# Install dependencies
pip install -r requirements.txt
```

## Configuration

Set credentials via environment variables or pass them as CLI arguments:

```powershell
$env:JOBBOSS_USER = "your_username"
$env:JOBBOSS_PASSWORD = "your_password"
```

## Usage

### Step 1: Create input file

Create `material_ids.txt` with one material ID per line:

```
02390177
02390219
```

### Step 2: Generate XML files for review

```powershell
python xml_generator.py --input material_ids.txt --output ./pending_updates
```

This creates query and update XML files in `pending_updates/` along with a `manifest.json`.

### Step 3: Review generated files

Inspect the XML files in `pending_updates/` before executing.

### Step 4: Execute updates

```powershell
# Dry run (preview only)
python xml_executor.py --manifest ./pending_updates/manifest.json --dry-run

# Execute for real
python xml_executor.py --manifest ./pending_updates/manifest.json
```

## Local-only files

To exclude local documentation from git, add to `.git/info/exclude`:

```
JobBOSS XML SDK Developer's Guide.html
```
