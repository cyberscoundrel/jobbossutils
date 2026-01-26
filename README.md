# JobBOSS Utilities

Python tools for managing JobBOSS data via the XML SDK.

## Tools

This repository contains two sets of tools:

### 1. Material Inventory Management (`xml_generator.py` / `xml_executor.py`)
Adjust material on-hand quantities in inventory.

[See original README below](#material-inventory-management)

### 2. Job Material Management (`job_material_generator.py` / `job_material_executor.py`)
Add material requirements to existing jobs.

**[See detailed documentation →](README_JOB_MATERIALS.md)**

Quick example:
```powershell
# Add 10 units of material MAT-001 to job JOB-12345
python job_material_generator.py --job JOB-12345 --material MAT-001 --quantity 10
python job_material_executor.py --manifest ./pending_updates/manifest.json
```

---

## Material Inventory Management

A two-phase workflow for updating material quantities in JobBOSS via the XML SDK.

## Requirements

- Windows with JobBOSS installed
- JobBOSS COM SDK registered (`JBInterface.JBRequestProcessor`)
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


