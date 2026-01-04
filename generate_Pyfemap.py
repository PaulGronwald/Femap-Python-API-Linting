"""Regenerate Pyfemap.py from Femap type library using makepy.

Based on : https://github.com/vsdsantos/PyFemap

This script generates the COM wrapper (Pyfemap.py) from the Femap type library.
It uses smart path resolution to find femap.tlb automatically or prompt the user.

Usage:
    python gen.py                    # Auto-detect or prompt for .tlb file
    python gen.py --tlb <path>       # Use specific .tlb file
"""

import sys
import argparse
from win32com.client import makepy
from femap_path_utils import get_tlb_path


def main():
    parser = argparse.ArgumentParser(
        description='Generate Pyfemap.py from Femap type library'
    )
    parser.add_argument(
        '--tlb',
        help='Path to femap.tlb file (auto-detects if not specified)'
    )

    args = parser.parse_args()

    # Resolve .tlb file path
    tlb_path = get_tlb_path(args.tlb)

    if not tlb_path:
        print("ERROR: Could not locate femap.tlb file")
        sys.exit(1)

    print(f"\nGenerating Pyfemap.py from: {tlb_path}")
    print("This may take a minute...\n")

    # Run makepy with resolved path
    sys.argv = ["makepy", "-o", "Pyfemap.py", tlb_path]
    makepy.main()

    print("\nSuccessfully generated Pyfemap.py")
    print("Note: This is an auto-generated file - do not edit manually")


if __name__ == '__main__':
    main() 