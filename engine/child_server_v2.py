# engine/child_server_v2.py
"""The v2 MCP server (stdio), started by run_server_v2.py."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from smartass.server import main  # noqa: E402

if __name__ == "__main__":
    main()
