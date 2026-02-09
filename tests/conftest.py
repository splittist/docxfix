"""Test configuration for pytest."""

import pytest


@pytest.fixture
def sample_xml_string() -> str:
    """Fixture providing a sample XML string for testing."""
    return """<?xml version="1.0"?>
<root>
    <child>Test content</child>
</root>"""
