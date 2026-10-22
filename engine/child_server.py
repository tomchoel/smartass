"""The real smartview MCP server (stdio). Started and restarted by run_server.py; can also be run directly."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from smartview_mcp.server import main  # noqa: E402

if __name__ == "__main__":
    main()
