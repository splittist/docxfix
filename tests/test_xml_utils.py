"""Tests for XML utilities using syrupy for snapshot testing."""

from syrupy.assertion import SnapshotAssertion

from docxfix.xml_utils import (
    add_child,
    create_simple_xml,
    parse_xml_string,
    xml_to_string,
)


def test_create_simple_xml():
    """Test creating a simple XML element."""
    element = create_simple_xml("root")
    assert element.tag == "root"
    assert element.text is None


def test_create_simple_xml_with_content():
    """Test creating a simple XML element with content."""
    element = create_simple_xml("root", "Hello, World!")
    assert element.tag == "root"
    assert element.text == "Hello, World!"


def test_add_child():
    """Test adding a child element."""
    parent = create_simple_xml("parent")
    child = add_child(parent, "child", "Child text")

    assert len(parent) == 1
    assert child.tag == "child"
    assert child.text == "Child text"


def test_add_child_with_attributes():
    """Test adding a child element with attributes."""
    parent = create_simple_xml("parent")
    child = add_child(parent, "child", "Text", {"id": "123", "type": "test"})

    assert child.get("id") == "123"
    assert child.get("type") == "test"


def test_xml_to_string():
    """Test converting XML to string."""
    element = create_simple_xml("root", "Content")
    xml_string = xml_to_string(element)

    assert "<root>Content</root>" in xml_string


def test_xml_to_string_snapshot(snapshot: SnapshotAssertion):
    """Test XML to string conversion with snapshot."""
    root = create_simple_xml("document")
    add_child(root, "title", "Test Document")
    add_child(root, "author", "John Doe")

    paragraph = add_child(root, "paragraph")
    add_child(paragraph, "text", "First sentence.")
    add_child(paragraph, "text", "Second sentence.")

    xml_string = xml_to_string(root)
    assert xml_string == snapshot


def test_parse_xml_string(sample_xml_string: str):
    """Test parsing an XML string."""
    element = parse_xml_string(sample_xml_string)
    assert element.tag == "root"

    children = list(element)
    assert len(children) == 1
    assert children[0].tag == "child"
    assert children[0].text == "Test content"


def test_round_trip_xml():
    """Test that XML can be converted to string and back."""
    original = create_simple_xml("root")
    add_child(original, "child1", "Text 1")
    add_child(original, "child2", "Text 2")

    xml_string = xml_to_string(original, pretty_print=False)
    parsed = parse_xml_string(xml_string)

    assert parsed.tag == original.tag
    assert len(parsed) == len(original)
