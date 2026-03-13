"""
XML Generator for Adding Materials to JobBOSS Jobs

Takes a job ID, material ID, and quantity to generate XML documents
for adding materials to existing jobs.

Usage:
    python job_material_generator.py --job JOB123 --material MAT001 --quantity 5 --output-dir ./pending_updates

This will generate XML to:
1. Query the job for its LastUpdated timestamp
2. Add the specified material with the specified quantity to the job
"""

import os
import sys
import argparse
import json
from datetime import datetime


# =============================================================================
# XML Templates
# =============================================================================

def create_job_query_xml(session_id_placeholder: str, job_id: str) -> str:
    """Create XML to query a job and get its LastUpdated timestamp."""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest Session="{session_id_placeholder}">
        <JobQueryRq>
            <JobQueryFilter>
                <ID>{job_id}</ID>
            </JobQueryFilter>
        </JobQueryRq>
    </JBXMLRequest>
</JBXML>'''


def create_job_material_add_xml(session_id_placeholder: str, job_id: str,
                                 last_updated_placeholder: str, material_id: str,
                                 quantity: float) -> str:
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

def generate_job_material_add_package(job_id: str, material_id: str,
                                       quantity: float, output_dir: str) -> dict:
    """
    Generate XML files and manifest for adding material to a job.
    
    Creates:
    - manifest.json: Summary of the operation for review
    - query_job.xml: Query XML to get job's LastUpdated timestamp
    - add_material.xml: Add material requirement XML template
    
    The XML files contain placeholders:
    - {{SESSION_ID}}: Replaced at execution time with actual session
    - {{LAST_UPDATED}}: Replaced with value from query response
    """
    os.makedirs(output_dir, exist_ok=True)
    
    manifest = {
        "generated_at": datetime.now().isoformat(),
        "operation": "add_material_to_job",
        "job_id": job_id,
        "material_id": material_id,
        "quantity": quantity,
    }
    
    print(f"\nGenerating XML to add material to job...")
    print(f"  Job ID: {job_id}")
    print(f"  Material ID: {material_id}")
    print(f"  Quantity: {quantity}")
    print()
    
    # Generate job query XML
    query_xml = create_job_query_xml("{{SESSION_ID}}", job_id)
    query_file = "query_job.xml"
    query_path = os.path.join(output_dir, query_file)
    with open(query_path, 'w', encoding='utf-8') as f:
        f.write(query_xml)
    
    # Generate add material XML (with placeholder for LastUpdated)
    add_xml = create_job_material_add_xml(
        "{{SESSION_ID}}", 
        job_id,
        "{{LAST_UPDATED}}",
        material_id,
        quantity
    )
    add_file = "add_material.xml"
    add_path = os.path.join(output_dir, add_file)
    with open(add_path, 'w', encoding='utf-8') as f:
        f.write(add_xml)
    
    manifest["query_file"] = query_file
    manifest["add_file"] = add_file
    
    # Write manifest
    manifest_path = os.path.join(output_dir, "manifest.json")
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2)
    
    print(f"Generated files in: {output_dir}")
    print(f"  - manifest.json (review this first!)")
    print(f"  - {query_file}")
    print(f"  - {add_file}")
    
    return manifest


# =============================================================================
# CLI
# =============================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description='Generate XML documents for adding materials to JobBOSS jobs',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
    # Add material MAT001 (quantity 5) to job JOB123
    python job_material_generator.py --job JOB123 --material MAT001 --quantity 5

    # Add material 02390219 (quantity 10) to job 12345
    python job_material_generator.py --job 12345 --material 02390219 --quantity 10 --output-dir ./job_updates

XML Schema:
    The tool generates JobModRq with MaterialRequirementAdd to add materials to jobs.
    Required: Job ID, Material ID, Quantity
    The material must already exist in JobBOSS inventory.
        '''
    )
    
    parser.add_argument(
        '--job', '-j',
        required=True,
        help='Job ID to add material to'
    )
    parser.add_argument(
        '--material', '-m',
        required=True,
        help='Material ID to add'
    )
    parser.add_argument(
        '--quantity', '-q',
        required=True,
        type=float,
        help='Quantity of material to add'
    )
    parser.add_argument(
        '--output-dir', '-o',
        default='./pending_updates',
        help='Directory to write XML files (default: ./pending_updates)'
    )
    
    return parser.parse_args()


def main():
    args = parse_args()
    
    print("=" * 60)
    print("JobBOSS Job Material XML Generator")
    print("=" * 60)
    print(f"Timestamp: {datetime.now().isoformat()}")
    
    # Validate quantity
    if args.quantity <= 0:
        print(f"\nERROR: Quantity must be positive, got: {args.quantity}")
        sys.exit(1)
    
    # Generate the XML package
    manifest = generate_job_material_add_package(
        args.job, 
        args.material, 
        args.quantity, 
        args.output_dir
    )
    
    print()
    print("=" * 60)
    print("NEXT STEPS")
    print("=" * 60)
    print()
    print("1. Review the manifest:")
    print(f"   {os.path.join(args.output_dir, 'manifest.json')}")
    print()
    print("2. Inspect the generated XML files if needed")
    print()
    print("3. When ready to execute, run:")
    print(f"   python job_material_executor.py --manifest {os.path.join(args.output_dir, 'manifest.json')} --user USERNAME --password PASSWORD")
    print()
    print("   Or with environment variables:")
    print("   set JOBBOSS_USER=your_username")
    print("   set JOBBOSS_PASSWORD=your_password")
    print(f"   python job_material_executor.py --manifest {os.path.join(args.output_dir, 'manifest.json')}")


if __name__ == "__main__":
    main()
