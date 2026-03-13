"""
XML Generator for JobBOSS Material Updates

Takes a list of material IDs (with duplicates representing individual pieces)
and generates XML documents for auditing before execution.

Supports two modes:
  adjust          - Adjust on-hand inventory quantities (AdjustOnHandQty).
                    Each occurrence = 1 piece to remove from inventory.
  job-requirement - Add materials as requirements on an existing job
                    (JobModRq / MaterialRequirementAdd).

Usage:
    # Adjust mode (default) - subtract from inventory
    python xml_generator.py --input material_ids.txt --output-dir ./pending_updates

    # Job-requirement mode - add materials to a job
    python xml_generator.py --mode job-requirement --job JOB-12345 --input material_ids.txt

Input file format (one material ID per line, duplicates = multiple pieces):
    MAT-001
    MAT-001
    MAT-002
    # This is a comment
    MAT-001
    
adjust mode generates:       MAT-001: -3, MAT-002: -1
job-requirement mode generates: MAT-001: 3, MAT-002: 1 (added to the job)
"""

import os
import sys
import argparse
import json
from datetime import datetime
from collections import Counter


# =============================================================================
# Material Counting
# =============================================================================

def count_materials(material_ids: list[str], negate: bool = True) -> dict[str, int]:
    """
    Count occurrences of each material ID.
    
    When negate=True (adjust mode), returns negative counts for inventory subtraction.
    When negate=False (job-requirement mode), returns positive counts.
    """
    counts = Counter(material_ids)
    if negate:
        return {mat_id: -count for mat_id, count in counts.items()}
    return dict(counts)


# =============================================================================
# XML Templates
# =============================================================================

def create_material_query_xml(session_id_placeholder: str, material_id: str) -> str:
    """Create XML to query a material and get its LastUpdated timestamp."""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest Session="{session_id_placeholder}">
        <MaterialQueryRq>
            <MaterialQueryFilter>
                <ID>{material_id}</ID>
                <IncludeMaterialLocations>false</IncludeMaterialLocations>
                <IncludeCustomerParts>false</IncludeCustomerParts>
                <IncludePriceBreaks>false</IncludePriceBreaks>
            </MaterialQueryFilter>
        </MaterialQueryRq>
    </JBXMLRequest>
</JBXML>'''


def create_material_mod_xml(session_id_placeholder: str, material_id: str,
                            last_updated_placeholder: str, quantity: int,
                            reason_id: str) -> str:
    """Create XML to modify a material's on-hand quantity."""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest Session="{session_id_placeholder}">
        <MaterialModRq>
            <MaterialMod>
                <ID>{material_id}</ID>
                <LastUpdated>{last_updated_placeholder}</LastUpdated>
            </MaterialMod>
            <AdjustOnHandQty>
                <ReasonRef ID="{reason_id}"/>
                <Quantity>{quantity}</Quantity>
            </AdjustOnHandQty>
        </MaterialModRq>
    </JBXMLRequest>
</JBXML>'''


def create_job_query_xml(session_id_placeholder: str, job_id: str) -> str:
    """Create XML to query a job and get its LastUpdated timestamp."""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest Session="{session_id_placeholder}">
        <JobQueryRq>
            <JobQueryFilter>
                <ID>{job_id}</ID>
                <IncludeAdditionalCharges>false</IncludeAdditionalCharges>
                <IncludeDeliveries>false</IncludeDeliveries>
                <IncludeRoutingLines>false</IncludeRoutingLines>
                <IncludeMaterialRequirements>false</IncludeMaterialRequirements>
                <IncludeComponents>false</IncludeComponents>
            </JobQueryFilter>
        </JobQueryRq>
    </JBXMLRequest>
</JBXML>'''


def create_job_material_add_xml(session_id_placeholder: str, job_id: str,
                                last_updated_placeholder: str, material_id: str,
                                quantity: int) -> str:
    """Create XML to add a material requirement to a job."""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest Session="{session_id_placeholder}">
        <JobModRq>
            <JobMod>
                <ID>{job_id}</ID>
                <LastUpdated>{last_updated_placeholder}</LastUpdated>
            </JobMod>
            <MaterialRequirementAdd>
                <MaterialRef ID="{material_id}"/>
                <RequirementProperties>
                    <Quantity>{quantity}</Quantity>
                </RequirementProperties>
            </MaterialRequirementAdd>
        </JobModRq>
    </JBXMLRequest>
</JBXML>'''


# =============================================================================
# File I/O
# =============================================================================

def load_material_ids(input_path: str) -> list[str]:
    """Load material IDs from a text file (one per line)."""
    material_ids = []
    with open(input_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            # Skip empty lines and comments
            if line and not line.startswith('#'):
                material_ids.append(line)
    return material_ids


def generate_update_package(material_ids: list[str], 
                            output_dir: str,
                            mode: str = "adjust",
                            reason_id: str = "",
                            job_id: str = "") -> dict:
    """
    Generate XML files and manifest for material updates.
    
    In 'adjust' mode (default):
    - query_<material_id>.xml: Query each material for LastUpdated
    - update_<material_id>.xml: AdjustOnHandQty to subtract from inventory
    
    In 'job-requirement' mode:
    - query_job.xml: Query the job for LastUpdated (shared across materials)
    - add_<material_id>.xml: MaterialRequirementAdd per material
    
    The XML files contain placeholders:
    - {{SESSION_ID}}: Replaced at execution time with actual session
    - {{LAST_UPDATED}}: Replaced with value from query response
    """
    os.makedirs(output_dir, exist_ok=True)
    
    quantity_changes = count_materials(material_ids, negate=(mode == "adjust"))
    
    manifest = {
        "generated_at": datetime.now().isoformat(),
        "mode": mode,
        "total_materials": len(quantity_changes),
        "total_pieces": sum(abs(q) for q in quantity_changes.values()),
        "materials": [],
        "input_ids": material_ids,
    }
    
    if mode == "adjust":
        manifest["reason_id"] = reason_id
        print(f"\nGenerating XML for {len(quantity_changes)} unique materials...")
        print(f"Total pieces to remove: {manifest['total_pieces']}")
    else:
        manifest["job_id"] = job_id
        print(f"\nGenerating XML for {len(quantity_changes)} unique materials...")
        print(f"Target job: {job_id}")
        print(f"Total pieces to add as requirements: {manifest['total_pieces']}")
    print()
    
    if mode == "job-requirement":
        # Single job query XML shared across all materials
        query_xml = create_job_query_xml("{{SESSION_ID}}", job_id)
        query_file = "query_job.xml"
        query_path = os.path.join(output_dir, query_file)
        with open(query_path, 'w', encoding='utf-8') as f:
            f.write(query_xml)
    
    for material_id, quantity in sorted(quantity_changes.items()):
        safe_id = "".join(c if c.isalnum() or c in '-_' else '_' for c in material_id)
        
        print(f"  {material_id}: {quantity:+d} pieces")
        
        if mode == "adjust":
            query_xml = create_material_query_xml("{{SESSION_ID}}", material_id)
            query_file = f"query_{safe_id}.xml"
            query_path = os.path.join(output_dir, query_file)
            with open(query_path, 'w', encoding='utf-8') as f:
                f.write(query_xml)
            
            update_xml = create_material_mod_xml(
                "{{SESSION_ID}}", material_id, "{{LAST_UPDATED}}",
                quantity, reason_id
            )
            update_file = f"update_{safe_id}.xml"
        else:
            # job-requirement: query_file is the shared query_job.xml
            query_file = "query_job.xml"
            update_xml = create_job_material_add_xml(
                "{{SESSION_ID}}", job_id, "{{LAST_UPDATED}}",
                material_id, quantity
            )
            update_file = f"add_{safe_id}.xml"
        
        update_path = os.path.join(output_dir, update_file)
        with open(update_path, 'w', encoding='utf-8') as f:
            f.write(update_xml)
        
        manifest["materials"].append({
            "material_id": material_id,
            "quantity_change": quantity,
            "occurrences": abs(quantity),
            "query_file": query_file,
            "update_file": update_file,
        })
    
    manifest_path = os.path.join(output_dir, "manifest.json")
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2)
    
    update_label = "update" if mode == "adjust" else "add"
    query_count = len(quantity_changes) if mode == "adjust" else 1
    print()
    print(f"Generated files in: {output_dir}")
    print(f"  - manifest.json (review this first!)")
    print(f"  - {query_count} query XML file{'s' if query_count > 1 else ''}")
    print(f"  - {len(quantity_changes)} {update_label} XML files")
    
    return manifest


# =============================================================================
# CLI
# =============================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description='Generate XML documents for JobBOSS material updates',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Modes:
    adjust (default)  Adjust material on-hand inventory via AdjustOnHandQty.
    job-requirement   Add materials as requirements on an existing job.

Examples:
    # Adjust mode (default) - subtract from inventory
    python xml_generator.py --input material_ids.txt --output-dir ./pending_updates
    python xml_generator.py -i used_materials.txt -o ./batch_001 --reason CONSUMED

    # Job-requirement mode - add materials to a job
    python xml_generator.py --mode job-requirement --job JOB-12345 -i material_ids.txt
    python xml_generator.py --mode job-requirement --job 99001 -i parts.txt -o ./job_updates

Input file format (one material ID per line):
    MAT-001
    MAT-001
    MAT-002
    # This is a comment (ignored)
    MAT-001
    
adjust mode:          MAT-001: -3 pieces, MAT-002: -1 piece
job-requirement mode: MAT-001: +3 pieces, MAT-002: +1 piece (added to job)
        '''
    )
    
    parser.add_argument(
        '--mode',
        choices=['adjust', 'job-requirement'],
        default='adjust',
        help='Operation mode (default: adjust)'
    )
    parser.add_argument(
        '--job', '-j',
        default='',
        help='Job ID to add material requirements to (required for job-requirement mode)'
    )
    parser.add_argument(
        '--input', '-i',
        required=True,
        help='Text file with material IDs (one per line, duplicates allowed)'
    )
    parser.add_argument(
        '--output-dir', '-o',
        default='./pending_updates',
        help='Directory to write XML files (default: ./pending_updates)'
    )
    parser.add_argument(
        '--reason', '-r',
        default='',
        help='Reason code for adjustment (adjust mode only, default: empty)'
    )
    
    return parser.parse_args()


def main():
    args = parse_args()
    
    print("=" * 60)
    print("JobBOSS XML Generator")
    print("=" * 60)
    print(f"Timestamp: {datetime.now().isoformat()}")
    print(f"Mode: {args.mode}")
    
    if args.mode == "job-requirement" and not args.job:
        print("\nERROR: --job is required when using --mode job-requirement")
        sys.exit(1)
    
    if not os.path.exists(args.input):
        print(f"\nERROR: Input file not found: {args.input}")
        sys.exit(1)
    
    material_ids = load_material_ids(args.input)
    
    if not material_ids:
        print("\nERROR: No material IDs found in input file")
        sys.exit(1)
    
    print(f"\nLoaded {len(material_ids)} material ID entries from: {args.input}")
    
    manifest = generate_update_package(
        material_ids, args.output_dir,
        mode=args.mode, reason_id=args.reason, job_id=args.job
    )
    
    manifest_rel = os.path.join(args.output_dir, 'manifest.json')
    
    print()
    print("=" * 60)
    print("NEXT STEPS")
    print("=" * 60)
    print()
    print("1. Review the manifest:")
    print(f"   {manifest_rel}")
    print()
    print("2. Inspect individual XML files if needed")
    print()
    print("3. When ready to execute, run:")
    print(f"   python xml_executor.py --manifest {manifest_rel} --user USERNAME --password PASSWORD")
    print()
    print("   Or with environment variables:")
    print("   set JOBBOSS_USER=your_username")
    print("   set JOBBOSS_PASSWORD=your_password")
    print(f"   python xml_executor.py --manifest {manifest_rel}")


if __name__ == "__main__":
    main()
