"""XML manipulation utilities using lxml."""

from lxml import etree

# Type alias for XML elements
# Using the private _Element type is acceptable for type hints
# as this is the canonical way with lxml until proper types are available
type XMLElement = etree._Element  # type: ignore[name-defined]


def create_simple_xml(root_tag: str, content: str | None = None) -> XMLElement:
    """
    Create a simple XML element.

    Args:
        root_tag: The tag name for the root element
        content: Optional text content for the element

    Returns:
        An lxml Element object
    """
    element = etree.Element(root_tag)
    if content is not None:
        element.text = content
    return element


def add_child(
    parent: XMLElement,
    tag: str,
    text: str | None = None,
    attributes: dict[str, str] | None = None,
) -> XMLElement:
    """
    Add a child element to a parent element.

    Args:
        parent: The parent element
        tag: The tag name for the child element
        text: Optional text content
        attributes: Optional dictionary of attributes

    Returns:
        The created child element
    """
    child = etree.SubElement(parent, tag)
    if text is not None:
        child.text = text
    if attributes is not None:
        for key, value in attributes.items():
            child.set(key, value)
    return child


def xml_to_string(element: XMLElement, pretty_print: bool = True) -> str:
    """
    Convert an XML element to a string.

    Args:
        element: The XML element to convert
        pretty_print: Whether to format the output with indentation

    Returns:
        The XML as a string
    """
    return etree.tostring(
        element,
        encoding="unicode",
        pretty_print=pretty_print,
    )


def parse_xml_string(xml_string: str) -> XMLElement:
    """
    Parse an XML string into an element.

    Args:
        xml_string: The XML string to parse

    Returns:
        The parsed XML element
    """
    return etree.fromstring(xml_string.encode("utf-8"))
