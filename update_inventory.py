"""
JobBOSS Inventory Update - Master Script

A single command to update material quantities from an input file.
Uses temporary files internally - no manual review step.

Usage:
    python update_inventory.py --input material_ids.txt --user USERNAME --password PASSWORD

Each line in the input file = 1 piece removed from inventory.
Duplicate material IDs are counted and combined into a single update.

For the two-phase workflow with manual review, use:
    xml_generator.py -> review -> xml_executor.py
"""

import os
import sys
import re
import tempfile
import shutil
import argparse
from datetime import datetime
from collections import Counter

try:
    import win32com.client
except ImportError:
    print("ERROR: pywin32 is not installed. Run: pip install pywin32")
    sys.exit(1)


# =============================================================================
# Configuration
# =============================================================================

DEFAULT_REASON = "ADJUST"


# =============================================================================
# Input Processing
# =============================================================================

def load_material_ids(input_path: str) -> list[str]:
    """Load material IDs from a text file (one per line)."""
    material_ids = []
    with open(input_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith('#'):
                material_ids.append(line)
    return material_ids


def count_materials(material_ids: list[str]) -> dict[str, int]:
    """
    Count occurrences of each material ID.
    Each occurrence = 1 piece to remove from inventory.
    
    Returns dict of {material_id: negative_quantity}
    """
    counts = Counter(material_ids)
    return {mat_id: -count for mat_id, count in counts.items()}


# =============================================================================
# XML Generation
# =============================================================================

def create_query_xml(session_id: str, material_id: str) -> str:
    """Create XML to query a material."""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest>
        <SessionID>{session_id}</SessionID>
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


def create_update_xml(session_id: str, material_id: str, 
                      last_updated: str, quantity: int, reason_id: str) -> str:
    """Create XML to update a material's quantity."""
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRequest>
        <SessionID>{session_id}</SessionID>
        <MaterialModRq>
            <MaterialMod>
                <ID>{material_id}</ID>
                <LastUpdated>{last_updated}</LastUpdated>
            </MaterialMod>
            <AdjustOnHandQty>
                <ReasonRef ID="{reason_id}"/>
                <Quantity>{quantity}</Quantity>
            </AdjustOnHandQty>
        </MaterialModRq>
    </JBXMLRequest>
</JBXML>'''


# =============================================================================
# Response Parsing
# =============================================================================

def parse_last_updated(response: str) -> str | None:
    """Extract LastUpdated from query response."""
    match = re.search(r'<LastUpdated>([^<]+)</LastUpdated>', response)
    return match.group(1) if match else None


def parse_error(response: str) -> str | None:
    """Extract error message from response."""
    if '<StatusCode>0</StatusCode>' in response:
        return None
    
    for pattern in [r'<StatusMessage>([^<]+)</StatusMessage>',
                    r'<ErrorMessage>([^<]+)</ErrorMessage>']:
        match = re.search(pattern, response)
        if match:
            return match.group(1)
    
    match = re.search(r'<StatusCode>([^0][^<]*)</StatusCode>', response)
    if match:
        return f"Status code: {match.group(1)}"
    
    return None


# =============================================================================
# Main Execution
# =============================================================================

def run_updates(input_path: str, username: str, password: str, 
                reason_id: str, dry_run: bool = False) -> dict:
    """
    Run the complete update workflow.
    
    1. Load and count material IDs
    2. Connect to JobBOSS
    3. For each material: query -> update
    4. Report results
    """
    results = {"success": [], "failed": [], "errors": []}
    
    # Load input
    print(f"Loading: {input_path}")
    material_ids = load_material_ids(input_path)
    
    if not material_ids:
        print("ERROR: No material IDs found in input file")
        return results
    
    # Count materials
    quantity_changes = count_materials(material_ids)
    
    print(f"Loaded {len(material_ids)} entries -> {len(quantity_changes)} unique materials")
    print(f"Reason code: {reason_id}")
    print()
    print("Materials to update:")
    for mat_id, qty in sorted(quantity_changes.items()):
        print(f"  {mat_id}: {qty:+d}")
    print()
    
    # Dry run - stop here
    if dry_run:
        print("DRY RUN - No changes will be made")
        print("Remove --dry-run to execute.")
        return results
    
    # Connect to JobBOSS
    print("Connecting to JobBOSS...")
    try:
        jb = win32com.client.Dispatch("JBRequestProcessor.RequestProcessor")
    except Exception as e:
        results["errors"].append(f"Failed to create COM object: {e}")
        print(f"ERROR: {e}")
        return results
    
    # Create session
    print(f"Creating session for: {username}")
    try:
        error_msg = ""
        session_id = jb.CreateSession(username, password, error_msg)
        if not session_id:
            results["errors"].append(f"Session failed: {error_msg}")
            print(f"ERROR: {error_msg}")
            return results
        print("Session created")
    except Exception as e:
        results["errors"].append(f"Session failed: {e}")
        print(f"ERROR: {e}")
        return results
    
    try:
        # Process each material
        for material_id, quantity in sorted(quantity_changes.items()):
            print(f"\n[{material_id}] Quantity: {quantity:+d}")
            
            # Query for LastUpdated
            print("  Querying...")
            query_xml = create_query_xml(session_id, material_id)
            
            try:
                response = jb.ProcessRequest(query_xml)
            except Exception as e:
                print(f"  X Query failed: {e}")
                results["failed"].append({"id": material_id, "qty": quantity, "error": str(e)})
                continue
            
            error = parse_error(response)
            if error:
                print(f"  X {error}")
                results["failed"].append({"id": material_id, "qty": quantity, "error": error})
                continue
            
            last_updated = parse_last_updated(response)
            if not last_updated:
                print("  X Material not found")
                results["failed"].append({"id": material_id, "qty": quantity, "error": "Not found"})
                continue
            
            print(f"  LastUpdated: {last_updated}")
            
            # Execute update
            print("  Updating...")
            update_xml = create_update_xml(session_id, material_id, last_updated, quantity, reason_id)
            
            try:
                response = jb.ProcessRequest(update_xml)
            except Exception as e:
                print(f"  X Update failed: {e}")
                results["failed"].append({"id": material_id, "qty": quantity, "error": str(e)})
                continue
            
            error = parse_error(response)
            if error:
                print(f"  X {error}")
                results["failed"].append({"id": material_id, "qty": quantity, "error": error})
                continue
            
            print(f"  OK Success")
            results["success"].append({"id": material_id, "qty": quantity})
    
    finally:
        print("\nClosing session...")
        try:
            jb.CloseSession(session_id)
        except:
            pass
    
    return results


# =============================================================================
# CLI
# =============================================================================

def parse_args():
    parser = argparse.ArgumentParser(
        description='Update JobBOSS material quantities from an input file',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
    python update_inventory.py --input material_ids.txt --user myuser --password mypass
    python update_inventory.py -i used_parts.txt -u admin -p secret --reason CONSUMED
    python update_inventory.py -i parts.txt -u admin -p secret --dry-run

Input file format (one material ID per line):
    MAT-001
    MAT-001
    MAT-002
    # comments ignored
    
Each line = 1 piece removed. Duplicates are counted automatically.
        '''
    )
    
    parser.add_argument('--input', '-i', required=True,
                        help='Text file with material IDs (one per line)')
    parser.add_argument('--user', '-u', default=os.environ.get('JOBBOSS_USER'),
                        help='JobBOSS username (or JOBBOSS_USER env var)')
    parser.add_argument('--password', '-p', default=os.environ.get('JOBBOSS_PASSWORD'),
                        help='JobBOSS password (or JOBBOSS_PASSWORD env var)')
    parser.add_argument('--reason', '-r', default=DEFAULT_REASON,
                        help=f'Reason code (default: {DEFAULT_REASON})')
    parser.add_argument('--dry-run', action='store_true',
                        help='Show what would be done without executing')
    
    return parser.parse_args()


def main():
    args = parse_args()
    
    print("=" * 50)
    print("JobBOSS Inventory Update")
    print("=" * 50)
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    # Validate
    if not args.user:
        print("ERROR: Username required (--user or JOBBOSS_USER)")
        sys.exit(1)
    if not args.password:
        print("ERROR: Password required (--password or JOBBOSS_PASSWORD)")
        sys.exit(1)
    if not os.path.exists(args.input):
        print(f"ERROR: File not found: {args.input}")
        sys.exit(1)
    
    # Run
    results = run_updates(args.input, args.user, args.password, args.reason, args.dry_run)
    
    # Summary
    print()
    print("=" * 50)
    print("SUMMARY")
    print("=" * 50)
    print(f"Success: {len(results['success'])}")
    print(f"Failed:  {len(results['failed'])}")
    
    if results['success']:
        total = sum(r['qty'] for r in results['success'])
        print(f"Total adjusted: {total:+d} pieces")
    
    if results['failed']:
        print("\nFailed:")
        for r in results['failed']:
            print(f"  {r['id']}: {r['error']}")
    
    if results['errors']:
        print("\nErrors:")
        for e in results['errors']:
            print(f"  {e}")
    
    sys.exit(1 if results['failed'] or results['errors'] else 0)


if __name__ == "__main__":
    main()
