"""Shared OOXML constants and utilities.

This subpackage holds the namespace URIs, Clark-notation qualified
names, and small helpers that the DOCX / PPTX / XLSX subpackages
import directly. The catalogue is large (hundreds of qualified names
in :mod:`kaos_office.ooxml.namespace`); to keep the public surface
honest, callers import the specific symbol they need from
``kaos_office.ooxml.namespace`` rather than relying on package-level
re-exports. ``__all__`` is therefore intentionally empty here.
"""

__all__: tuple[str, ...] = ()
