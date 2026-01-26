# JobBOSS Job Material Management Tool

A two-phase workflow for adding materials to existing jobs in JobBOSS via the XML SDK.

## Overview

This tool allows you to add material requirements to existing JobBOSS jobs. Unlike the original `xml_generator.py` which adjusts material inventory quantities, this tool adds materials to a job's bill of materials.

## Requirements

- Windows with JobBOSS installed
- JobBOSS COM SDK registered (`JBInterface.JBRequestProcessor`)
- Python 3.10+
- pywin32 package

## Installation

```powershell
# Install dependencies
pip install pywin32
```

## Configuration

Set credentials via environment variables or pass them as CLI arguments:

```powershell
$env:JOBBOSS_USER = "your_username"
$env:JOBBOSS_PASSWORD = "your_password"
```

## Usage

### Step 1: Generate XML files for review

```powershell
python job_material_generator.py --job JOB-12345 --material MAT-001 --quantity 10
```

This creates:
- `query_job.xml` - Query to get job's LastUpdated timestamp
- `add_material.xml` - Add material requirement XML
- `manifest.json` - Summary for review

### Step 2: Review generated files

Inspect the XML files in `pending_updates/` before executing. Verify:
- Job ID is correct
- Material ID exists in JobBOSS
- Quantity is correct

### Step 3: Execute the operation

```powershell
# Dry run (preview only)
python job_material_executor.py --manifest ./pending_updates/manifest.json --dry-run

# Execute for real
python job_material_executor.py --manifest ./pending_updates/manifest.json
```

## XML Schema Explanation

### JobModRq with MaterialRequirementAdd

To add a material of quantity **n** to job **m**, use the following XML structure:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest Session="YOUR_SESSION_ID">
        <JobModRq>
            <JobMod>
                <ID>m</ID>
                <LastUpdated>TIMESTAMP_FROM_QUERY</LastUpdated>
            </JobMod>
            <MaterialRequirementAdd>
                <MaterialRef ID="MATERIAL_ID"/>
                <RequirementProperties>
                    <Quantity>n</Quantity>
                </RequirementProperties>
            </MaterialRequirementAdd>
        </JobModRq>
    </JBXMLRequest>
</JBXML>
```

### Key Elements

1. **JobMod** - Identifies which job to modify
   - `ID` - The job ID to add materials to
   - `LastUpdated` - Timestamp from job query (for optimistic locking)

2. **MaterialRequirementAdd** - Adds a material requirement
   - `MaterialRef` with `ID` attribute - The material to add
   - `RequirementProperties/Quantity` - How much to add

### Complete Example

To add **5 units of material "02390219" to job "12345"**:

**Step 1: Query the job**
```xml
<JBXML>
    <JBXMLRequest Session="abc123">
        <JobQueryRq>
            <JobQueryFilter>
                <ID>12345</ID>
                <IncludeJobOperations>false</IncludeJobOperations>
                <IncludeComponents>false</IncludeComponents>
                <IncludeMaterialRequirements>false</IncludeMaterialRequirements>
            </JobQueryFilter>
        </JobQueryRq>
    </JBXMLRequest>
</JBXML>
```

**Step 2: Add the material** (using LastUpdated from query response)
```xml
<JBXML>
    <JBXMLRequest Session="abc123">
        <JobModRq>
            <JobMod>
                <ID>12345</ID>
                <LastUpdated>2024-01-26T10:30:00</LastUpdated>
            </JobMod>
            <MaterialRequirementAdd>
                <MaterialRef ID="02390219"/>
                <RequirementProperties>
                    <Quantity>5</Quantity>
                </RequirementProperties>
            </MaterialRequirementAdd>
        </JobModRq>
    </JBXMLRequest>
</JBXML>
```

## Adding Multiple Materials

You can add multiple materials in one request by including multiple `MaterialRequirementAdd` elements:

```xml
<JobModRq>
    <JobMod>
        <ID>12345</ID>
        <LastUpdated>2024-01-26T10:30:00</LastUpdated>
    </JobMod>
    <MaterialRequirementAdd>
        <MaterialRef ID="MAT-001"/>
        <RequirementProperties>
            <Quantity>10</Quantity>
        </RequirementProperties>
    </MaterialRequirementAdd>
    <MaterialRequirementAdd>
        <MaterialRef ID="MAT-002"/>
        <RequirementProperties>
            <Quantity>5</Quantity>
        </RequirementProperties>
    </MaterialRequirementAdd>
</JobModRq>
```

## Examples

See the `examples/` directory for complete XML examples:
- `example_job_query.xml` - Query a job
- `example_add_material_to_job.xml` - Add single material
- `example_add_multiple_materials.xml` - Add multiple materials

## Command-Line Options

### job_material_generator.py

```
--job, -j       Job ID to add material to (required)
--material, -m  Material ID to add (required)
--quantity, -q  Quantity of material (required, must be positive)
--output-dir, -o Output directory (default: ./pending_updates)
```

### job_material_executor.py

```
--manifest, -m   Path to manifest.json (required)
--user, -u       JobBOSS username (or set JOBBOSS_USER env var)
--password, -p   JobBOSS password (or set JOBBOSS_PASSWORD env var)
--dry-run        Preview without executing
--verbose, -v    Show full XML requests/responses
--log-xml [DIR]  Save all XML to directory (default: ./xml_logs)
```

## Troubleshooting

### Job not found
- Verify the job ID exists in JobBOSS
- Ensure you have permission to view the job

### Material not found
- Verify the material ID exists in JobBOSS inventory
- Check for typos in the material ID

### LastUpdated mismatch
- The job was modified between query and update
- Re-run the generator to get a fresh timestamp

### Permission denied
- Ensure your JobBOSS user has permission to modify jobs
- Check that the job is not locked or in a status that prevents modification

## Differences from xml_generator.py

| Feature | xml_generator.py | job_material_generator.py |
|---------|-----------------|--------------------------|
| Purpose | Adjust material inventory | Add materials to jobs |
| Input | List of material IDs | Job ID + Material ID + Quantity |
| XML Request | MaterialModRq | JobModRq with MaterialRequirementAdd |
| Effect | Changes on-hand qty | Adds to job's BOM |

## Related Files

- `xml_generator.py` - For adjusting material inventory quantities
- `xml_executor.py` - For executing material inventory adjustments
- `JobBOSS XML SDK Developer's Guide.html` - Complete API documentation
