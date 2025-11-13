#!/usr/bin/env python3
"""Install script for Excel MCP Server to Claude Desktop"""

import json
import os
import platform
import shutil
import sys
from pathlib import Path


def get_claude_config_path() -> Path:
    """Get the Claude Desktop configuration file path based on OS"""
    system = platform.system()
    
    if system == "Darwin":  # macOS
        config_dir = Path.home() / "Library" / "Application Support" / "Claude"
    elif system == "Windows":
        config_dir = Path(os.environ.get("APPDATA", "")) / "Claude"
    elif system == "Linux":
        config_dir = Path.home() / ".config" / "Claude"
    else:
        raise OSError(f"Unsupported operating system: {system}")
    
    config_file = config_dir / "claude_desktop_config.json"
    return config_file, config_dir


def get_project_path() -> Path:
    """Get the absolute path to the current project directory"""
    # Get the directory where this script is located
    script_path = Path(__file__).resolve()
    # Go up one level from setup/ to project root
    return script_path.parent.parent


def load_config(config_file: Path) -> dict:
    """Load existing Claude Desktop config or create new one"""
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            print(f"⚠️  Warning: Existing config file has invalid JSON: {e}")
            print("   Creating backup and starting fresh...")
            backup_path = config_file.with_suffix(".json.backup")
            shutil.copy2(config_file, backup_path)
            print(f"   Backup saved to: {backup_path}")
            return {}
    return {}


def get_env_vars() -> dict:
    """Prompt user for environment variables or use defaults"""
    print("\n📝 Environment Variables Configuration")
    print("   Press Enter to use defaults (shown in parentheses)")
    print("   Note: These are optional for stdio mode")
    
    excel_files_path = input(
        "   EXCEL_FILES_PATH "
        "(optional, for SSE/HTTP modes): "
    ).strip()
    
    fastmcp_host = input(
        "   FASTMCP_HOST "
        "(0.0.0.0): "
    ).strip()
    if not fastmcp_host:
        fastmcp_host = "0.0.0.0"
    
    fastmcp_port = input(
        "   FASTMCP_PORT "
        "(8017): "
    ).strip()
    if not fastmcp_port:
        fastmcp_port = "8017"
    
    env_vars = {
        "FASTMCP_HOST": fastmcp_host,
        "FASTMCP_PORT": fastmcp_port,
    }
    
    if excel_files_path:
        env_vars["EXCEL_FILES_PATH"] = excel_files_path
    
    return env_vars


def create_mcp_config(project_path: Path, env_vars: dict) -> dict:
    """Create MCP server configuration"""
    # Use absolute path
    project_path_str = str(project_path)
    
    # Check for virtual environment first (most reliable)
    venv_python = project_path / ".venv" / "bin" / "python"
    if venv_python.exists():
        print("   ✓ Using project virtual environment")
        config = {
            "command": str(venv_python),
            "args": ["-m", "excel_mcp.__main__", "stdio"],
            "env": env_vars,
        }
        return config
    
    # Check for uv and use it with the project
    uv_available = shutil.which("uv") is not None
    if uv_available:
        print("   ✓ Using uv (will sync dependencies automatically)")
        # Use uv run from the project directory
        # This will automatically sync and run the module
        config = {
            "command": "uv",
            "args": ["run", "--directory", project_path_str, "python", "-m", "excel_mcp.__main__", "stdio"],
            "env": env_vars,
        }
        return config
    
    # Fallback to system Python (user must have package installed)
    print("   ⚠️  Warning: No virtual environment or uv found.")
    print("      Using system Python. Make sure excel-mcp-server is installed.")
    config = {
        "command": "python",
        "args": ["-m", "excel_mcp.__main__", "stdio"],
        "env": env_vars,
    }
    return config


def ensure_venv(project_path: Path) -> bool:
    """Ensure virtual environment exists, create if needed"""
    venv_path = project_path / ".venv"
    if venv_path.exists():
        return True
    
    # Check if uv is available
    if shutil.which("uv") is None:
        return False
    
    print("\n📦 Virtual environment not found. Creating one...")
    import subprocess
    try:
        subprocess.run(
            ["uv", "sync"],
            cwd=project_path,
            check=True,
            capture_output=True,
        )
        print("   ✓ Virtual environment created successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"   ❌ Failed to create virtual environment: {e}")
        return False


def install(config_file: Path, config_dir: Path, project_path: Path) -> bool:
    """Install MCP server configuration to Claude Desktop"""
    print("🚀 Installing Excel MCP Server to Claude Desktop\n")
    
    # Ensure virtual environment exists
    if not ensure_venv(project_path):
        print("\n⚠️  Warning: Could not create virtual environment.")
        print("   Continuing with available options...")
    
    # Get environment variables
    env_vars = get_env_vars()
    
    # Load existing config
    config = load_config(config_file)
    
    # Ensure mcpServers exists
    if "mcpServers" not in config:
        config["mcpServers"] = {}
    
    # Check if already configured
    server_name = "excel-mcp-server"
    if server_name in config["mcpServers"]:
        print(f"\n⚠️  MCP server '{server_name}' already exists in config.")
        overwrite = input("   Overwrite existing configuration? (y/N): ").strip().lower()
        if overwrite != "y":
            print("   Installation cancelled.")
            return False
    
    # Create MCP server configuration
    mcp_config = create_mcp_config(project_path, env_vars)
    config["mcpServers"][server_name] = mcp_config
    
    # Ensure config directory exists
    config_dir.mkdir(parents=True, exist_ok=True)
    
    # Write config file
    try:
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        print(f"\n✅ Successfully installed MCP server configuration!")
        print(f"   Config file: {config_file}")
        print(f"   Server name: {server_name}")
        print(f"\n📋 Next steps:")
        print("   1. Restart Claude Desktop")
        print("   2. Check MCP server status in Claude Desktop settings")
        print("   3. Try asking Claude: 'Can you create an Excel workbook?'")
        return True
    except Exception as e:
        print(f"\n❌ Error writing config file: {e}")
        return False


def main():
    """Main installation function"""
    try:
        config_file, config_dir = get_claude_config_path()
        project_path = get_project_path()
        
        print(f"📁 Claude Desktop config: {config_file}")
        print(f"📁 Project directory: {project_path}")
        
        if not install(config_file, config_dir, project_path):
            sys.exit(1)
            
    except OSError as e:
        print(f"❌ Error: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n\n⚠️  Installation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

