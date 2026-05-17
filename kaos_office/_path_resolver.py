"""Internal adapter — resolve agent-supplied office-file paths via kaos-core.

Wraps :func:`kaos_core.path_resolver.resolve_input_path` with:

* The right mime allowlist per office format (DOCX / PPTX / XLSX).
* A lightweight context shim so CLI / standalone callers (no
  :class:`~kaos_core.base.context.KaosContext`) keep working — they
  fall through to the resolver's absolute-filesystem branch.

Every :mod:`kaos_office.tools` entry point that accepts a ``path``
parameter from agent input goes through this. Without it, files
uploaded into ``KaosRuntime.vfs`` by a UI host (e.g. ``kaos-ui``'s
single-user-chat SPA) are invisible to the tools — the agent sees
an unbroken sequence of "File not found" errors and is at risk of
hallucinating answers from zero successful reads. See
``kaos-modules/docs/plans/vfs-blind-tools-audit-and-fix-plan.md``
for the production post-mortem.
"""

from __future__ import annotations

from contextlib import asynccontextmanager
from typing import TYPE_CHECKING, Literal

from kaos_core.base.context import KaosContext
from kaos_core.path_resolver import (
    InputPathResolutionError,
    ResolvedInput,
    resolve_input_path,
)

if TYPE_CHECKING:
    from collections.abc import AsyncIterator

OfficeFormat = Literal["docx", "pptx", "xlsx"]

# Mime allowlists per format. The legacy ``application/msword`` /
# ``application/vnd.ms-powerpoint`` / ``application/vnd.ms-excel``
# entries cover artifacts whose mime type was inferred from the
# `.doc` / `.ppt` / `.xls` suffix on the original upload — the
# readers themselves still require true OOXML bytes, so a mismatch
# raises further down the pipeline. Keeping the legacy types in
# the allowlist means the resolver doesn't pre-empt that real
# error with a confusing mime-mismatch message of its own.
_DOCX_MIMES: tuple[str, ...] = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/msword",
)
_PPTX_MIMES: tuple[str, ...] = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "application/vnd.ms-powerpoint",
)
_XLSX_MIMES: tuple[str, ...] = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
)

_MIMES_BY_FORMAT: dict[OfficeFormat, tuple[str, ...]] = {
    "docx": _DOCX_MIMES,
    "pptx": _PPTX_MIMES,
    "xlsx": _XLSX_MIMES,
}


@asynccontextmanager
async def resolve_office_input(
    path_or_uri: str,
    context: KaosContext | None,
    *,
    format: OfficeFormat,
) -> AsyncIterator[ResolvedInput]:
    """Resolve an agent-supplied office file path to a real on-disk file.

    Yields a :class:`~kaos_core.path_resolver.ResolvedInput` whose
    ``path`` is always a real ``pathlib.Path`` the tool can hand to
    any third-party reader (python-docx, python-pptx, openpyxl,
    calamine, lxml). For artifact / VFS inputs the bytes are
    streamed to a private temp file that the context manager owns;
    on exit the temp file is unlinked. Filesystem inputs return the
    original path unchanged (no copy, no cleanup).

    Parameters
    ----------
    path_or_uri
        The path / URI / artifact reference the agent supplied.
    context
        ``KaosContext`` for the current turn (or ``None`` for CLI /
        standalone callers). Provides ``session_id`` and ``runtime``
        for VFS / artifact-store reads; without a runtime, only
        absolute filesystem paths resolve.
    format
        Office format the caller expects. Selects the mime allowlist
        that gets passed to the resolver so mismatched artifacts
        (e.g. a PDF passed to a DOCX tool) fail fast with an
        agent-friendly hint rather than a downstream parser exception.

    Raises
    ------
    InputPathResolutionError
        On empty input, missing file, unsupported scheme, mime
        mismatch, or runtime/context unavailability. Catch sites
        should compose the error via ``exc.to_agent_message()``.
    """
    mimes = _MIMES_BY_FORMAT[format]
    ctx = context if context is not None else _cli_context()
    async with resolve_input_path(
        path_or_uri,
        context=ctx,
        allowed_mime_types=mimes,
    ) as resolved:
        yield resolved


def _cli_context() -> KaosContext:
    """Build a minimal runtime-free context for CLI / standalone use.

    The resolver's filesystem branch only needs ``session_id`` (it
    never touches the VFS or artifact store when the input is an
    absolute path), so a one-off context with no runtime is enough
    to satisfy the helper's signature without forcing every CLI
    call site to construct one.
    """
    return KaosContext(session_id="cli")


__all__ = [
    "InputPathResolutionError",
    "OfficeFormat",
    "ResolvedInput",
    "resolve_office_input",
]
