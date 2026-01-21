"""
Mock JobBOSS COM Object for Testing

Simulates the JBRequestProcessor COM object so tests can run
without an actual JobBOSS installation.

Usage:
    from jobboss_mock import MockJobBOSS, install_mock, uninstall_mock
    
    # Install the mock
    install_mock()
    
    # Now any code that calls win32com.client.Dispatch("JBRequestProcessor.RequestProcessor")
    # will get a MockJobBOSS instance instead
"""

import re
from datetime import datetime
from typing import Optional


class MockJobBOSS:
    """
    Mock implementation of JBRequestProcessor.RequestProcessor
    
    Simulates:
    - CreateSession: Returns a fake session ID
    - ProcessRequest: Parses XML and returns mock responses
    - CloseSession: No-op
    """
    
    def __init__(self):
        self.session_active = False
        self.session_id = None
        self.request_log = []  # Log of all requests for inspection
        
        # Simulated material database
        self.materials = {
            "TEST-MAT-001": {
                "id": "TEST-MAT-001",
                "description": "Test Material 1",
                "on_hand": 100,
                "last_updated": "2025-06-15T10:30:00",
            },
            "TEST-MAT-002": {
                "id": "TEST-MAT-002",
                "description": "Test Material 2",
                "on_hand": 50,
                "last_updated": "2025-08-22T14:15:30",
            },
            "TEST-MAT-003": {
                "id": "TEST-MAT-003",
                "description": "Test Material 3",
                "on_hand": 25,
                "last_updated": "2025-12-01T09:00:00",
            },
        }
    
    def CreateSession(self, username: str, password: str, error_msg: str = "") -> str:
        """Create a mock session."""
        print(f"  [MOCK] CreateSession(user={username})")
        
        # Simulate authentication
        if username and password:
            self.session_active = True
            self.session_id = f"MOCK-SESSION-{datetime.now().strftime('%Y%m%d%H%M%S')}"
            return self.session_id
        else:
            return ""
    
    def CloseSession(self, session_id: str) -> None:
        """Close the mock session."""
        print(f"  [MOCK] CloseSession()")
        self.session_active = False
        self.session_id = None
    
    def ProcessRequest(self, xml_request: str) -> str:
        """Process an XML request and return a mock response."""
        self.request_log.append(xml_request)
        
        # Determine request type
        if "<MaterialQueryRq>" in xml_request:
            return self._handle_material_query(xml_request)
        elif "<MaterialModRq>" in xml_request:
            return self._handle_material_mod(xml_request)
        else:
            return self._error_response("Unknown request type")
    
    def _handle_material_query(self, xml_request: str) -> str:
        """Handle MaterialQueryRq - return material info."""
        # Extract material ID from request
        match = re.search(r'<ID>([^<]+)</ID>', xml_request)
        if not match:
            return self._error_response("No material ID in query")
        
        material_id = match.group(1)
        print(f"  [MOCK] Query: {material_id}")
        
        if material_id in self.materials:
            mat = self.materials[material_id]
            return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRespond>
        <MaterialQueryRs>
            <StatusCode>0</StatusCode>
            <StatusMessage>Success</StatusMessage>
            <Material>
                <ID>{mat['id']}</ID>
                <Description>{mat['description']}</Description>
                <OnHand>{mat['on_hand']}</OnHand>
                <LastUpdated>{mat['last_updated']}</LastUpdated>
            </Material>
        </MaterialQueryRs>
    </JBXMLRespond>
</JBXML>'''
        else:
            return self._error_response(f"Material not found: {material_id}")
    
    def _handle_material_mod(self, xml_request: str) -> str:
        """Handle MaterialModRq - simulate quantity update."""
        # Extract material ID
        match = re.search(r'<MaterialMod>\s*<ID>([^<]+)</ID>', xml_request, re.DOTALL)
        if not match:
            return self._error_response("No material ID in update")
        
        material_id = match.group(1)
        
        # Extract LastUpdated from request
        match = re.search(r'<LastUpdated>([^<]+)</LastUpdated>', xml_request)
        if not match:
            return self._error_response("No LastUpdated in update")
        
        submitted_last_updated = match.group(1)
        
        # Extract quantity
        match = re.search(r'<Quantity>([^<]+)</Quantity>', xml_request)
        if not match:
            return self._error_response("No Quantity in update")
        
        quantity = float(match.group(1))
        
        print(f"  [MOCK] Update: {material_id} qty={quantity:+.0f}")
        
        # Check if material exists
        if material_id not in self.materials:
            return self._error_response(f"Material not found: {material_id}")
        
        mat = self.materials[material_id]
        
        # Validate LastUpdated matches (simulating optimistic locking)
        if submitted_last_updated != mat['last_updated']:
            return self._error_response(
                f"LastUpdated mismatch. Expected: {mat['last_updated']}, Got: {submitted_last_updated}"
            )
        
        # Apply the update
        old_qty = mat['on_hand']
        mat['on_hand'] += quantity
        mat['last_updated'] = datetime.now().isoformat()
        
        print(f"  [MOCK] {material_id}: {old_qty} -> {mat['on_hand']}")
        
        return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRespond>
        <MaterialModRs>
            <StatusCode>0</StatusCode>
            <StatusMessage>Success</StatusMessage>
            <MaterialRet>
                <ID>{material_id}</ID>
                <OnHand>{mat['on_hand']}</OnHand>
                <LastUpdated>{mat['last_updated']}</LastUpdated>
            </MaterialRet>
        </MaterialModRs>
    </JBXMLRespond>
</JBXML>'''
    
    def _error_response(self, message: str) -> str:
        """Generate an error response."""
        print(f"  [MOCK] Error: {message}")
        return f'''<?xml version="1.0" encoding="UTF-8"?>
<JBXML>
    <JBXMLRespond>
        <StatusCode>1</StatusCode>
        <StatusMessage>{message}</StatusMessage>
    </JBXMLRespond>
</JBXML>'''


# =============================================================================
# Mock Installation (patches win32com.client.Dispatch)
# =============================================================================

_original_dispatch = None
_mock_instance = None


def _mock_dispatch(prog_id: str):
    """Replacement for win32com.client.Dispatch that returns our mock."""
    global _mock_instance
    
    if prog_id == "JBRequestProcessor.RequestProcessor":
        if _mock_instance is None:
            _mock_instance = MockJobBOSS()
        return _mock_instance
    else:
        # For other COM objects, use the original if available
        if _original_dispatch:
            return _original_dispatch(prog_id)
        raise Exception(f"Mock: Unknown COM object: {prog_id}")


def install_mock():
    """Install the mock - patches win32com.client.Dispatch."""
    global _original_dispatch, _mock_instance
    
    try:
        import win32com.client
        _original_dispatch = win32com.client.Dispatch
        win32com.client.Dispatch = _mock_dispatch
        _mock_instance = MockJobBOSS()
        print("[MOCK] JobBOSS mock installed")
        return _mock_instance
    except ImportError:
        # win32com not available, create a fake module
        import sys
        from types import ModuleType
        
        # Create fake win32com.client module
        win32com = ModuleType('win32com')
        win32com.client = ModuleType('win32com.client')
        win32com.client.Dispatch = _mock_dispatch
        
        sys.modules['win32com'] = win32com
        sys.modules['win32com.client'] = win32com.client
        
        _mock_instance = MockJobBOSS()
        print("[MOCK] JobBOSS mock installed (fake win32com)")
        return _mock_instance


def uninstall_mock():
    """Uninstall the mock - restore original Dispatch."""
    global _original_dispatch, _mock_instance
    
    if _original_dispatch:
        import win32com.client
        win32com.client.Dispatch = _original_dispatch
        _original_dispatch = None
    
    _mock_instance = None
    print("[MOCK] JobBOSS mock uninstalled")


def get_mock_instance() -> Optional[MockJobBOSS]:
    """Get the current mock instance (for inspecting state)."""
    return _mock_instance
