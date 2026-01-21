"""
Test Workflow for JobBOSS Material Update Scripts

Tests both the generator and executor using mock data.
No actual JobBOSS connection required.

Usage:
    python test_workflow.py
"""

import os
import sys
import json
import tempfile
import shutil
from datetime import datetime

# Install mock BEFORE importing the other modules
from jobboss_mock import install_mock, get_mock_instance

# Install mock (this patches win32com.client.Dispatch)
mock = install_mock()


def test_generator():
    """Test the XML generator with mock data."""
    print()
    print("=" * 60)
    print("TEST 1: XML Generator")
    print("=" * 60)
    
    # Create temp directory for test
    test_dir = tempfile.mkdtemp(prefix="jobboss_test_")
    input_file = os.path.join(test_dir, "test_input.txt")
    output_dir = os.path.join(test_dir, "output")
    
    try:
        # Create test input file with duplicate IDs
        test_ids = [
            "TEST-MAT-001",
            "TEST-MAT-001",
            "TEST-MAT-001",  # 3 occurrences = -3
            "TEST-MAT-002",
            "TEST-MAT-002",  # 2 occurrences = -2
            "TEST-MAT-003",  # 1 occurrence = -1
        ]
        
        with open(input_file, 'w') as f:
            f.write("# Test input file\n")
            for mat_id in test_ids:
                f.write(f"{mat_id}\n")
        
        print(f"\nTest input file: {input_file}")
        print(f"Material IDs: {test_ids}")
        print()
        
        # Import and run generator
        from xml_generator import load_material_ids, count_materials, generate_update_package
        
        # Test loading
        loaded_ids = load_material_ids(input_file)
        print(f"Loaded {len(loaded_ids)} IDs")
        assert len(loaded_ids) == 6, f"Expected 6 IDs, got {len(loaded_ids)}"
        
        # Test counting
        counts = count_materials(loaded_ids)
        print(f"Counts: {counts}")
        assert counts["TEST-MAT-001"] == -3, f"Expected -3 for MAT-001, got {counts['TEST-MAT-001']}"
        assert counts["TEST-MAT-002"] == -2, f"Expected -2 for MAT-002, got {counts['TEST-MAT-002']}"
        assert counts["TEST-MAT-003"] == -1, f"Expected -1 for MAT-003, got {counts['TEST-MAT-003']}"
        
        # Test package generation
        manifest = generate_update_package(loaded_ids, output_dir, "TEST-REASON")
        
        # Verify manifest
        assert manifest["total_materials"] == 3
        assert manifest["total_pieces"] == 6
        assert manifest["reason_id"] == "TEST-REASON"
        assert len(manifest["materials"]) == 3
        
        # Verify files exist
        manifest_path = os.path.join(output_dir, "manifest.json")
        assert os.path.exists(manifest_path), "manifest.json not created"
        
        for item in manifest["materials"]:
            query_path = os.path.join(output_dir, item["query_file"])
            update_path = os.path.join(output_dir, item["update_file"])
            assert os.path.exists(query_path), f"{item['query_file']} not created"
            assert os.path.exists(update_path), f"{item['update_file']} not created"
        
        # Show manifest contents
        print("\nGenerated manifest.json:")
        with open(manifest_path) as f:
            print(json.dumps(json.load(f), indent=2))
        
        print("\n[PASS] Generator test PASSED")
        return output_dir, manifest_path
        
    except Exception as e:
        print(f"\n[FAIL] Generator test FAILED: {e}")
        shutil.rmtree(test_dir, ignore_errors=True)
        raise


def test_executor(manifest_path: str):
    """Test the XML executor with mock JobBOSS."""
    print()
    print("=" * 60)
    print("TEST 2: XML Executor (with mock)")
    print("=" * 60)
    
    # Get initial mock state
    mock = get_mock_instance()
    print("\nMock material database (before):")
    for mat_id, mat in mock.materials.items():
        print(f"  {mat_id}: on_hand={mat['on_hand']}")
    
    # Import executor
    from xml_executor import execute_updates
    
    print(f"\nExecuting from manifest: {manifest_path}")
    print()
    
    # Run executor
    results = execute_updates(
        manifest_path=manifest_path,
        username="test_user",
        password="test_pass",
        dry_run=False
    )
    
    # Check results
    print(f"\nResults: {len(results['success'])} success, {len(results['failed'])} failed")
    
    assert len(results['success']) == 3, f"Expected 3 successes, got {len(results['success'])}"
    assert len(results['failed']) == 0, f"Expected 0 failures, got {len(results['failed'])}"
    
    # Verify mock state was updated
    print("\nMock material database (after):")
    for mat_id, mat in mock.materials.items():
        print(f"  {mat_id}: on_hand={mat['on_hand']}")
    
    # Verify quantities changed correctly
    assert mock.materials["TEST-MAT-001"]["on_hand"] == 97, "MAT-001 should be 100-3=97"
    assert mock.materials["TEST-MAT-002"]["on_hand"] == 48, "MAT-002 should be 50-2=48"
    assert mock.materials["TEST-MAT-003"]["on_hand"] == 24, "MAT-003 should be 25-1=24"
    
    print("\n[PASS] Executor test PASSED")


def test_master_script():
    """Test the master update_inventory.py script with mock."""
    print()
    print("=" * 60)
    print("TEST 3: Master Script (update_inventory.py)")
    print("=" * 60)
    
    # Reset mock to fresh state
    mock = get_mock_instance()
    mock.materials = {
        "TEST-MAT-001": {"id": "TEST-MAT-001", "description": "Test 1", "on_hand": 100, "last_updated": "2025-06-15T10:30:00"},
        "TEST-MAT-002": {"id": "TEST-MAT-002", "description": "Test 2", "on_hand": 50, "last_updated": "2025-08-22T14:15:30"},
    }
    
    # Create temp input file
    test_dir = tempfile.mkdtemp(prefix="jobboss_master_test_")
    input_file = os.path.join(test_dir, "input.txt")
    
    try:
        with open(input_file, 'w') as f:
            f.write("TEST-MAT-001\n")
            f.write("TEST-MAT-001\n")
            f.write("TEST-MAT-002\n")
        
        print(f"\nTest input: {input_file}")
        print("Materials: TEST-MAT-001 x2, TEST-MAT-002 x1")
        print()
        
        # Import and run
        from update_inventory import run_updates
        
        print("Mock state (before):")
        for mat_id, mat in mock.materials.items():
            print(f"  {mat_id}: {mat['on_hand']}")
        print()
        
        results = run_updates(
            input_path=input_file,
            username="test_user",
            password="test_pass",
            reason_id="TEST",
            dry_run=False
        )
        
        print("\nMock state (after):")
        for mat_id, mat in mock.materials.items():
            print(f"  {mat_id}: {mat['on_hand']}")
        
        assert len(results['success']) == 2
        assert mock.materials["TEST-MAT-001"]["on_hand"] == 98  # 100 - 2
        assert mock.materials["TEST-MAT-002"]["on_hand"] == 49  # 50 - 1
        
        print("\n[PASS] Master script test PASSED")
        
    finally:
        shutil.rmtree(test_dir, ignore_errors=True)


def test_nonexistent_material():
    """Test handling of non-existent material."""
    print()
    print("=" * 60)
    print("TEST 4: Non-existent Material Handling")
    print("=" * 60)
    
    # Reset mock
    mock = get_mock_instance()
    mock.materials = {
        "EXISTING-001": {"id": "EXISTING-001", "description": "Exists", "on_hand": 10, "last_updated": "2025-01-01T00:00:00"},
    }
    
    test_dir = tempfile.mkdtemp(prefix="jobboss_error_test_")
    input_file = os.path.join(test_dir, "input.txt")
    
    try:
        with open(input_file, 'w') as f:
            f.write("EXISTING-001\n")
            f.write("NONEXISTENT-999\n")  # This doesn't exist
        
        from update_inventory import run_updates
        
        print("\nTesting with one valid, one invalid material...")
        print()
        
        results = run_updates(
            input_path=input_file,
            username="test",
            password="test",
            reason_id="TEST",
            dry_run=False
        )
        
        assert len(results['success']) == 1, "Should have 1 success"
        assert len(results['failed']) == 1, "Should have 1 failure"
        assert results['failed'][0]['id'] == "NONEXISTENT-999"
        
        print("\n[PASS] Error handling test PASSED")
        
    finally:
        shutil.rmtree(test_dir, ignore_errors=True)


def main():
    print("=" * 60)
    print("JobBOSS Material Update - Test Suite")
    print("=" * 60)
    print(f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("Using MOCK JobBOSS (no real connection)")
    
    passed = 0
    failed = 0
    
    try:
        # Test 1: Generator
        output_dir, manifest_path = test_generator()
        passed += 1
        
        # Test 2: Executor
        test_executor(manifest_path)
        passed += 1
        
        # Cleanup test 1 & 2 files
        shutil.rmtree(os.path.dirname(manifest_path), ignore_errors=True)
        
        # Test 3: Master script
        test_master_script()
        passed += 1
        
        # Test 4: Error handling
        test_nonexistent_material()
        passed += 1
        
    except Exception as e:
        failed += 1
        print(f"\n[FAIL] Test failed with error: {e}")
        import traceback
        traceback.print_exc()
    
    # Summary
    print()
    print("=" * 60)
    print("TEST SUMMARY")
    print("=" * 60)
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")
    
    if failed == 0:
        print("\n[PASS] All tests passed!")
        return 0
    else:
        print("\n[FAIL] Some tests failed")
        return 1


if __name__ == "__main__":
    sys.exit(main())
