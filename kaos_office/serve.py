"""Run the KAOS MCP server with Office tools.

Usage:
    # stdio (for Claude Code / Claude Desktop)
    kaos-office-serve

    # streamable HTTP
    kaos-office-serve --http --port 8000

    # with debug logging
    kaos-office-serve --debug
"""

from __future__ import annotations

import argparse
import sys


def main(argv: list[str] | None = None) -> None:
    """Entry point for the MCP server."""
    parser = argparse.ArgumentParser(description="KAOS MCP Server with Office tools")
    parser.add_argument("--http", action="store_true", help="Use streamable HTTP transport")
    parser.add_argument("--host", default="127.0.0.1", help="HTTP host (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000, help="HTTP port (default: 8000)")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    args = parser.parse_args(argv)

    try:
        from kaos_core import KaosRuntime

        # `kaos_mcp` resolves only when the `[mcp]` extra is installed.
        # The extra is intentionally absent at 0.1.0a1 because `kaos-mcp`
        # is not yet on PyPI, so ty cannot statically resolve the import
        # in the per-module repo. The runtime ImportError below handles
        # the missing-package case at call time.
        from kaos_mcp import KaosMCPServer, KaosMCPSettings  # ty: ignore[unresolved-import]
    except ImportError:
        print(
            "Error: MCP server requires the 'mcp' extra.\n"
            "Install with: pip install 'kaos-office[mcp]'",
            file=sys.stderr,
        )
        sys.exit(1)

    from kaos_office import register_office_tools

    # Create runtime and register Office tools
    runtime = KaosRuntime()
    n_tools = register_office_tools(runtime)
    print(f"Registered {n_tools} Office tools", file=sys.stderr)

    # Configure server
    settings = KaosMCPSettings(
        name="kaos-office-server",
        transport="streamable-http" if args.http else "stdio",
        host=args.host,
        port=args.port,
        debug=args.debug,
    )

    server = KaosMCPServer(runtime=runtime, settings=settings)

    if args.http:
        print(f"Starting HTTP server on {args.host}:{args.port}/mcp", file=sys.stderr)
        server.run_streamable_http()
    else:
        print("Starting stdio server", file=sys.stderr)
        server.run_stdio()


if __name__ == "__main__":
    main()
