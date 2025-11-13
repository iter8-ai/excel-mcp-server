# Setup Instructions

This folder contains installation scripts and configuration examples for setting up the Excel MCP Server.

## Files

- **`install_to_claude.py`** - Automated installation script for Claude Desktop
- **`claude_desktop_config.example.json`** - Example Claude Desktop configuration file

## Quick Start

Run the installation script:

```bash
python setup/install_to_claude.py
```

Or using uv:

```bash
uv run python setup/install_to_claude.py
```

The script will:
- Detect your operating system and Claude Desktop config location
- Create a virtual environment if needed (`uv sync`)
- Prompt you for environment variables (with sensible defaults)
- Automatically configure Claude Desktop with the best available method
- Handle backups if needed

## Manual Configuration

See the main [README.md](../README.md) for manual configuration instructions.

