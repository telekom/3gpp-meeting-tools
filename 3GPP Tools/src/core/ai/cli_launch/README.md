# External LLM CLI Launch Package (`core.ai.cli_launch`)

Acyclic modular package extracted from `core/ai/cli_launcher.py` (thin backward-compatible re-export shim):

- `context_writer.py`: Standing context (`cli_user_context.md`), workspace/per-TDoc `_USER_NOTES.md`, task instructions, and `_CLI_CONTEXT.md` generation.
- `terminal_launch.py`: PowerShell command quoting, UTF-16LE Base64 `-EncodedCommand` encoding, and `wt.exe`/`powershell.exe` terminal launcher.
- `menu_builder.py`: Qt `QMenu` agent launcher, nested model/effort selector submenu, and MCP Cockpit link.
