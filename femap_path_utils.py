"""Utilities for locating the Femap type library file (.tlb)."""

import os
import glob
from pathlib import Path
from typing import Optional

# Cache file location for last user-selected .tlb path
_CACHE_FILE = Path(os.environ.get('TEMP', os.environ.get('TMP', '.'))) / '.femap_tlb_cache'


def _load_cached_path() -> Optional[str]:
    """Load cached .tlb path from temp directory."""
    try:
        if _CACHE_FILE.exists():
            cached_path = _CACHE_FILE.read_text(encoding='utf-8').strip()
            if cached_path and os.path.exists(cached_path):
                return cached_path
    except Exception:
        pass
    return None


def _save_cached_path(tlb_path: str) -> None:
    """Save .tlb path to cache file in temp directory."""
    try:
        _CACHE_FILE.write_text(tlb_path, encoding='utf-8')
    except Exception:
        pass  # Silent failure - caching is optional


def find_femap_install_dir() -> Optional[Path]:
    """
    Search common installation paths for Femap directory.

    Returns:
        Path to Femap installation directory, or None if not found.
    """
    search_paths = [
        r'C:\Program Files\Siemens\Femap *',
        r'C:\Program Files (x86)\Siemens\Femap *',
    ]

    for pattern in search_paths:
        matches = glob.glob(pattern)
        if matches:
            # Return the most recent version (sorted alphabetically, highest last)
            return Path(sorted(matches)[-1])

    return None


def prompt_for_tlb_file(initial_dir: Optional[Path] = None) -> Optional[str]:
    """
    Show a file dialog to select the femap.tlb file.

    Args:
        initial_dir: Directory to start the file dialog in.

    Returns:
        Path to selected .tlb file, or None if user cancels.
    """
    try:
        import tkinter as tk
        from tkinter import filedialog
    except ImportError:
        print("ERROR: tkinter not available for file dialog")
        return None

    # Create root window and hide it
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # Determine initial directory
    if initial_dir and initial_dir.exists():
        initial = str(initial_dir)
    else:
        # Fallback to Siemens directory
        siemens_dir = Path(r'C:\Program Files\Siemens')
        initial = str(siemens_dir) if siemens_dir.exists() else None

    # Show file dialog
    tlb_path = filedialog.askopenfilename(
        title='Select Femap Type Library (femap.tlb)',
        initialdir=initial,
        filetypes=[
            ('Type Library Files', '*.tlb'),
            ('All Files', '*.*')
        ]
    )

    root.destroy()

    return tlb_path if tlb_path else None


def get_tlb_path(cli_arg: Optional[str] = None) -> Optional[str]:
    """
    Resolve the path to femap.tlb using multiple strategies.

    Resolution order:
    1. Command-line argument (if provided)
    2. Environment variable FEMAP_TLB_PATH
    3. Cached path from last user selection
    4. Auto-detect in common installation paths
    5. Prompt user with file dialog (saves to cache)

    Args:
        cli_arg: Path from command-line --tlb argument, or None.

    Returns:
        Path to femap.tlb file, or None if not found/user cancelled.
    """
    # 1. Command-line argument takes precedence
    if cli_arg:
        if os.path.exists(cli_arg):
            return cli_arg
        else:
            print(f"WARNING: Specified .tlb file not found: {cli_arg}")
            # Continue to other methods instead of failing immediately

    # 2. Check environment variable
    env_path = os.getenv('FEMAP_TLB_PATH')
    if env_path and os.path.exists(env_path):
        print(f"Using FEMAP_TLB_PATH: {env_path}")
        return env_path

    # 3. Check cached path from previous user selection
    cached_path = _load_cached_path()
    if cached_path:
        print(f"Using cached path: {cached_path}")
        return cached_path

    # 4. Auto-detect in common installation paths
    install_dir = find_femap_install_dir()
    if install_dir:
        tlb_path = install_dir / 'femap.tlb'
        if tlb_path.exists():
            print(f"Auto-detected: {tlb_path}")
            return str(tlb_path)

    # 5. Prompt user with file dialog
    print("Femap type library not found. Please select femap.tlb...")
    selected_path = prompt_for_tlb_file(install_dir)

    if selected_path:
        print(f"Selected: {selected_path}")
        # Save to cache for next time (silent failure if unsuccessful)
        _save_cached_path(selected_path)
        return selected_path

    # User cancelled or no file found
    return None
