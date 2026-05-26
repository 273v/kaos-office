"""OPC part-name resolver — relationship target → package-root part path.

ECMA-376 §9.3.2 defines three valid forms for a relationship's Target
attribute on an internal (``TargetMode="Internal"``) relationship:

1. **Absolute** — starts with ``/``. The path is relative to the
   package root. ``/xl/worksheets/sheet1.xml`` resolves to the
   package-root part name ``xl/worksheets/sheet1.xml``.
2. **Parent-relative** — starts with ``../``. The path walks up from
   the source part's directory. ``../media/image1.png`` from
   ``word/document.xml`` (source dir ``word``) resolves to
   ``media/image1.png``.
3. **Source-relative** — no leading ``/`` or ``../``. The path is
   joined to the source part's directory.
   ``worksheets/sheet1.xml`` from ``xl/workbook.xml`` (source dir
   ``xl``) resolves to ``xl/worksheets/sheet1.xml``.

The kaos-office native XLSX reader had a naive ``f"xl/{rel.target}"``
that broke on form 1: ``openpyxl``-produced workbooks (very common
in practice) emit absolute targets, and the naive concat produced
paths like ``xl//xl/worksheets/sheet1.xml`` which never matched
package parts. The sheet was silently dropped, returning
``tables=[]`` on a valid workbook — see corpus-stress S19 / S16 and
plan ``2026-05-26-corpus-stress-5-failure-resolution.md`` Failure 1.

This module centralizes the resolver so every OOXML reader uses the
same logic. The pptx/smartart resolver at
``kaos_office/pptx/smartart.py`` re-implemented the same logic
inline; once this helper ships it should call here too.
"""

from __future__ import annotations


class InvalidRelationshipTarget(ValueError):
    """The relationship target cannot be resolved against the source part.

    Raised when a parent-relative target walks above the package root
    or when the inputs are otherwise malformed. The package itself is
    the structural boundary — there is no concept of "outside" the
    package.
    """


def _source_dir(source_part: str) -> str:
    """Return the source part's directory, without trailing slash.

    ``xl/workbook.xml`` → ``xl``.
    ``word/document.xml`` → ``word``.
    ``[Content_Types].xml`` → ``""`` (package root).
    """
    if "/" not in source_part:
        return ""
    return source_part.rsplit("/", 1)[0]


def resolve_part_target(source_part: str, target: str) -> str:
    """Resolve an OPC relationship target to a package-root part name.

    Args:
        source_part: The package-root part name of the source part
            (the part that owns the ``.rels`` file containing the
            relationship). E.g. ``"xl/workbook.xml"`` for a sheet
            relationship, ``"word/document.xml"`` for a hyperlink
            from the main document, ``""`` (empty string) when the
            relationship lives in the package-level
            ``_rels/.rels``.
        target: The relationship's ``Target`` attribute value, taken
            verbatim from the ``.rels`` XML. Internal relationships
            only — external relationships (``TargetMode="External"``)
            are URIs and not resolvable to part names.

    Returns:
        The resolved part name in package-root form, with no leading
        slash (i.e. exactly the form ``OPCPackage.has_part`` expects).

    Raises:
        InvalidRelationshipTarget: If ``target`` is empty, or if a
            parent-relative target walks above the package root.

    Examples:
        >>> resolve_part_target("xl/workbook.xml", "worksheets/sheet1.xml")
        'xl/worksheets/sheet1.xml'
        >>> resolve_part_target("xl/workbook.xml", "/xl/worksheets/sheet1.xml")
        'xl/worksheets/sheet1.xml'
        >>> resolve_part_target("word/document.xml", "../media/image1.png")
        'media/image1.png'
        >>> resolve_part_target("ppt/slides/slide1.xml", "../media/image2.png")
        'ppt/media/image2.png'
    """
    if not target:
        msg = (
            "OPC relationship target is empty; cannot resolve. "
            "Fix: ensure the relationship's Target attribute is set "
            "in the source .rels XML."
        )
        raise InvalidRelationshipTarget(msg)

    # Form 1: absolute — strip leading slashes, return as-is.
    if target.startswith("/"):
        return target.lstrip("/")

    src_dir = _source_dir(source_part)

    # Form 2: parent-relative — walk up segment by segment.
    if target.startswith("../"):
        parts = [p for p in src_dir.split("/") if p]
        for seg in target.split("/"):
            if seg == "..":
                if not parts:
                    msg = (
                        f"OPC relationship target {target!r} walks above "
                        f"the package root from source {source_part!r}. "
                        "Fix: correct the .rels XML so the target stays "
                        "within the package, or use an absolute target."
                    )
                    raise InvalidRelationshipTarget(msg)
                parts.pop()
            elif seg == "" or seg == ".":
                continue
            else:
                parts.append(seg)
        return "/".join(parts)

    # Form 3: source-relative — join to source directory. Empty src_dir
    # (top-level relationship) means the target IS the package-root
    # part name.
    if not src_dir:
        return target
    return f"{src_dir}/{target}"


__all__ = ["InvalidRelationshipTarget", "resolve_part_target"]
