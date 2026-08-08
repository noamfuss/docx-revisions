"""Shared helpers for matching comment ranges to elements in the document body.

Provides the core logic for associating ``w:commentRangeStart`` /
``w:commentRangeEnd`` pairs with tracked changes (``w:ins`` / ``w:del``)
and paragraphs (``w:p``), used by the ``.comments`` properties on
``TrackedChange`` and ``RevisionParagraph``.
"""

from __future__ import annotations

from typing import Dict, List, Tuple

from docx.oxml.ns import qn
from lxml import etree

_COMMENT_START_TAG = qn("w:commentRangeStart")
_COMMENT_END_TAG = qn("w:commentRangeEnd")


def _parse_comment_ranges(parent: etree._Element) -> Dict[int, Tuple[int, int]]:
    """Build a mapping of ``comment_id → (child_start, child_end)`` for *parent*.

    Scans *parent*'s direct children for ``w:commentRangeStart`` and
    ``w:commentRangeEnd`` elements and records the child-index span for
    each comment ID found.

    Comment ranges that have a start but no matching end within *parent*'s
    children are recorded with ``end = -1`` (open-ended — the range extends
    beyond this parent).

    Args:
        parent: An XML element whose children are scanned (typically a
            ``<w:p>`` element).

    Returns:
        A dict mapping ``comment_id (int) → (child_start, child_end)``.
    """
    ranges: Dict[int, Tuple[int, int]] = {}
    children = list(parent)

    for idx, child in enumerate(children):
        tag = child.tag
        if tag == _COMMENT_START_TAG:
            cid = _comment_id(child)
            if cid is not None:
                ranges[cid] = (idx, -1)  # tentatively open-ended
        elif tag == _COMMENT_END_TAG:
            cid = _comment_id(child)
            if cid is not None and cid in ranges:
                ranges[cid] = (ranges[cid][0], idx)

    return ranges


def _comment_id(element: etree._Element) -> int | None:
    """Read the ``w:id`` attribute as an int, or return ``None``."""
    id_attr = element.get(qn("w:id"))
    if id_attr is None:
        return None
    try:
        return int(id_attr)
    except (ValueError, TypeError):
        return None


def element_within_comment_ranges(
    element: etree._Element, parent: etree._Element
) -> List[int]:
    """Return comment IDs whose range *parent* includes *element*.

    Examines the direct children of *parent* — finding ``w:commentRangeStart``
    and ``w:commentRangeEnd`` markers — and checks whether *element*'s position
    among those children falls between the start and end of any comment range.

    Args:
        element: The child element to locate (e.g. a ``<w:ins>`` or ``<w:del>``).
        parent: The parent element whose children are scanned (e.g. ``<w:p>``).

    Returns:
        A list of comment IDs whose range includes *element*.
    """
    ranges = _parse_comment_ranges(parent)
    if not ranges:
        return []

    try:
        elem_idx = list(parent).index(element)
    except ValueError:
        return []

    result: List[int] = []
    for cid, (start_idx, end_idx) in ranges.items():
        if end_idx < 0:
            if elem_idx >= start_idx:
                result.append(cid)
        else:
            if start_idx <= elem_idx <= end_idx:
                result.append(cid)
    return result


def comment_ids_in_paragraph(p_element: etree._Element) -> List[int]:
    """Return all unique comment IDs whose range starts in *p_element*.

    Scans the children of the ``<w:p>`` for ``w:commentRangeStart``
    markers and collects their IDs.

    Args:
        p_element: A ``<w:p>`` element.

    Returns:
        A list of unique comment IDs.
    """
    ids: List[int] = []
    seen: set[int] = set()
    for child in p_element:
        if child.tag == _COMMENT_START_TAG:
            cid = _comment_id(child)
            if cid is not None and cid not in seen:
                ids.append(cid)
                seen.add(cid)
    return ids
