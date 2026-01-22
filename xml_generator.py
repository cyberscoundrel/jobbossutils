"""
XML Generator for JobBOSS Material Quantity Updates

Takes a list of material IDs (with duplicates representing individual pieces)
and generates XML documents for auditing before execution.

Each occurrence of a material ID in the input = 1 piece to remove from inventory.

Usage:
    python xml_generator.py --input material_ids.txt --output-dir ./pending_updates

Input file format (one material ID per line, duplicates = multiple pieces):
    MAT-001
    MAT-001
    MAT-002
    # This is a comment
    MAT-001
    
This would generate: MAT-001: -3, MAT-002: -1
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

def count_materials(material_ids: list[str]) -> dict[str, int]:
    """
    Count occurrences of each material ID.
    Each occurrence = 1 piece to remove from inventory.
    
    Returns dict of {material_id: negative_count} for subtraction
    """
    counts = Counter(material_ids)
    # Return negative counts to subtract from inventory
    return {mat_id: -count for mat_id, count in counts.items()}


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
                            reason_id: str = "") -> dict:
    """
    Generate XML files and manifest for material updates.
    
    Creates:
    - manifest.json: Summary of all updates for review
    - query_<material_id>.xml: Query XML for each material
    - update_<material_id>.xml: Update XML template for each material
    
    The XML files contain placeholders:
    - {{SESSION_ID}}: Replaced at execution time with actual session
    - {{LAST_UPDATED}}: Replaced with value from query response
    """
    os.makedirs(output_dir, exist_ok=True)
    
    # Count materials
    quantity_changes = count_materials(material_ids)
    
    manifest = {
        "generated_at": datetime.now().isoformat(),
        "reason_id": reason_id,
        "total_materials": len(quantity_changes),
        "total_pieces": sum(abs(q) for q in quantity_changes.values()),
        "materials": [],
        "input_ids": material_ids,  # Original list for reference/audit
    }
    
    print(f"\nGenerating XML for {len(quantity_changes)} unique materials...")
    print(f"Total pieces to remove: {sum(abs(q) for q in quantity_changes.values())}")
    print()
    
    for material_id, quantity in sorted(quantity_changes.items()):
        # Sanitize material ID for filename (replace unsafe chars)
        safe_id = "".join(c if c.isalnum() or c in '-_' else '_' for c in material_id)
        
        print(f"  {material_id}: {quantity} pieces")
        
        # Generate query XML
        query_xml = create_material_query_xml("{{SESSION_ID}}", material_id)
        query_file = f"query_{safe_id}.xml"
        query_path = os.path.join(output_dir, query_file)
        with open(query_path, 'w', encoding='utf-8') as f:
            f.write(query_xml)
        
        # Generate update XML (with placeholder for LastUpdated)
        update_xml = create_material_mod_xml(
            "{{SESSION_ID}}", 
            material_id,
            "{{LAST_UPDATED}}",
            quantity,
            reason_id
        )
        update_file = f"update_{safe_id}.xml"
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
    
    # Write manifest
    manifest_path = os.path.join(output_dir, "manifest.json")
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, indent=2)
    
    print()
    print(f"Generated files in: {output_dir}")
    print(f"  - manifest.json (review this first!)")
    print(f"  - {len(quantity_changes)} query XML files")
    print(f"  - {len(quantity_changes)} update XML files")
    
    return manifest


# =============================================================================
# CLI
# =============================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description='Generate XML documents for JobBOSS material quantity updates',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
    python xml_generator.py --input material_ids.txt --output-dir ./pending_updates
    python xml_generator.py -i used_materials.txt -o ./batch_001 --reason CONSUMED

Input file format (one material ID per line):
    MAT-001
    MAT-001
    MAT-002
    # This is a comment (ignored)
    MAT-001
    
Each line = 1 piece removed. Above example generates:
    MAT-001: -3 pieces
    MAT-002: -1 piece
        '''
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
        help='Reason code for adjustment (default: empty) - use a valid code from your JobBOSS system'
    )
    
    return parser.parse_args()


def main():
    args = parse_args()
    
    print("=" * 60)
    print("JobBOSS XML Generator")
    print("=" * 60)
    print(f"Timestamp: {datetime.now().isoformat()}")
    
    if not os.path.exists(args.input):
        print(f"\nERROR: Input file not found: {args.input}")
        sys.exit(1)
    
    # Load material IDs
    material_ids = load_material_ids(args.input)
    
    if not material_ids:
        print("\nERROR: No material IDs found in input file")
        sys.exit(1)
    
    print(f"\nLoaded {len(material_ids)} material ID entries from: {args.input}")
    
    # Generate the XML package
    manifest = generate_update_package(material_ids, args.output_dir, args.reason)
    
    print()
    print("=" * 60)
    print("NEXT STEPS")
    print("=" * 60)
    print()
    print("1. Review the manifest:")
    print(f"   {os.path.join(args.output_dir, 'manifest.json')}")
    print()
    print("2. Inspect individual XML files if needed")
    print()
    print("3. When ready to execute, run:")
    print(f"   python xml_executor.py --manifest {os.path.join(args.output_dir, 'manifest.json')} --user USERNAME --password PASSWORD")
    print()
    print("   Or with environment variables:")
    print("   set JOBBOSS_USER=your_username")
    print("   set JOBBOSS_PASSWORD=your_password")
    print(f"   python xml_executor.py --manifest {os.path.join(args.output_dir, 'manifest.json')}")


if __name__ == "__main__":
    main()
