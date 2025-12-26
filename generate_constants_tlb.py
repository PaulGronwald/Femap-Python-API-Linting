#!/usr/bin/env python3
"""
Generate femap_constants.py from Femap type library (.tlb).

This script:
1. Reads enum definitions directly from the .tlb file (authoritative source)
2. Creates type-safe IntEnum classes with readable aliases
3. Generates femap_constants.py with proper type safety

Usage:
    python generate_constants_tlb.py [--tlb PATH] [--output PATH] [--list-enums]

Path Resolution:
    If --tlb is not specified, will use this order:
    1. Environment variable FEMAP_TLB_PATH
    2. Auto-detect in common Femap installation paths
    3. Prompt with file dialog
"""

import argparse
import pythoncom
from pathlib import Path
from collections import defaultdict
from typing import NamedTuple, Dict, List, Tuple
from femap_path_utils import get_tlb_path

# Type kind constants
TKIND_ENUM = 0


class ConstantInfo(NamedTuple):
    name: str
    value: int
    enum_name: str


# Configuration for alias class generation
# Maps enum name -> (AliasClassName, prefix_to_strip, use_nested_grouping)
# If use_nested_grouping is True, will create nested classes based on secondary prefix
ALIAS_CONFIG = {
    # Core return codes and messages
    'zReturnCode': ('ReturnCode', 'FE_', False),
    'zMessageColor': ('Message', 'FCM_', False),  # Note: actual enum name is zMessageColor

    # Entity types
    'zDataType': ('Entity', 'FT_', False),
    'zElementType': ('ElemType', 'FET_', False),  # Note: This is the element type enum
    'zTopologyType': ('Topo', 'FTO_', False),
    'zMaterialType': ('MatlType', 'FMT_', False),
    'zNodeType': ('NodeType', 'FNT_', False),

    # Analysis
    'zAnalysisType': ('Analysis', 'FAT_', False),
    'zAnalysisProgram': ('Solver', 'FAP_', False),

    # Coordinate systems
    'zCSysType': ('CSys', 'FCS_', False),

    # Groups - with nested grouping
    'zGroupBooleanOp': ('GroupOp', 'FGB_', False),
    'zGroupDataType': ('GroupType', 'FGR_', False),
    'zGroupDefinitionType': ('GroupDef', 'FGD_', True),  # Nested: GroupDef.Elem.BYCOLOR

    # Visual/Display - Note: FCL_, FPF_, FPL_ are all in zColor enum
    'zColor': ('Color', 'FCL_', False),  # Only FCL_ colors
    'zColor:FPF_': ('BrushPattern', 'FPF_', False),  # Virtual: FPF_ subset of zColor
    'zColor:FPL_': ('PenLineStyle', 'FPL_', False),  # Virtual: FPL_ subset of zColor
    'zViewMode': ('ViewMode', 'FVM_', False),
    'zViewOptions': ('ViewOptions', 'FVI_', False),  # FVI_ constants are in zViewOptions

    # Loads
    'zLoadType': ('LoadType', 'FLT_', False),
    'zLoadDirection': ('LoadDir', 'FLD_', False),
    'zLoadVariation': ('LoadVar', 'FLV_', False),

    # Results
    'zResultsLocation': ('ResultsLoc', 'FRL_', False),
    'zOutputType': ('OutputType', 'FOT_', False),
    'zOutputComplex': ('OutputComplex', 'FOC_', False),

    # Functions
    'zFunctionType': ('FuncType', 'FFT_', False),

    # Meshing
    'zMeshApproach': ('MeshApproach', 'FMA_', False),
    'zMesherType': ('MesherType', 'FME_', False),

    # Selection
    'zSelectorType': ('SelectorType', 'FST_', False),

    # Geometry
    'zCurveType': ('CurveType', 'FCT_', False),
    'zSurfaceType': ('SurfaceType', 'FSU_', False),
    'zPointType': ('PointType', 'FPT_', False),

    # Freebody
    'zFbdComponent': ('FbdComponent', 'FFBC_', False),
    'zFbdContribution': ('FbdContrib', 'FFBCN_', False),
    'zFbdDisplayMode': ('FbdDisplay', 'FFBD_', False),

    # Connections/Contacts
    'zConnectionRegionType': ('ConnRegion', 'FCR_', False),
    'zConnectionPropType': ('ConnProp', 'FCP_', False),

    # Charts
    'zChartStyle': ('ChartStyle', 'FCS_', False),  # Note: prefix collision with CSys
    'zChartSeriesType': ('ChartSeries', 'FCST_', False),

    # Library
    'zLibraryFile': ('LibFile', 'FLF_', False),

    # Alignment
    'zAlignment': ('Align', 'FAL_', False),

    # Feature type
    'zFeatureType': ('Feature', 'FFE_', False),

    # Output destination (heavily used - 32 occurrences)
    'zOutputDestination': ('OutputDest', 'FOD_', False),

    # Combined mode (21 occurrences)
    'zCombinedMode': ('CombinedMode', 'FCBM_', False),

    # Freebody vector mode (14 occurrences)
    'zFbdVecMode': ('FbdVecMode', 'FBD_', False),

    # Results conversion (9 occurrences)
    'zResultsConvert': ('ResultsConvert', 'FRC_', False),

    # Coordinate picking (9 occurrences)
    'zCoordPick': ('CoordPick', 'FCP_', False),

    # Analysis forms
    'zAnalysisAssignForm': ('AnalysisForm', 'FAF_', False),

    # Optimization
    'zOptBoundtype': ('OptBound', 'FOB_', False),

    # Vector/Plate results
    'zVecPlateResult': ('VecPlateResult', 'FVPR_', False),
    'zVecPlateType': ('VecPlateType', 'FVPT_', False),
    'zVecSolidLamLoc': ('VecSolidLoc', 'FVSL_', False),
    'zVecSolidLamResult': ('VecSolidResult', 'FVSR_', False),

    # GFX (Graphics)
    'zGFXEdgeFlags': ('GfxEdge', 'FGFX_', False),
    'zGFXPointSymbol': ('GfxSymbol', 'FGFXPS_', False),
    'zGFXArrowMode': ('GfxArrow', 'FGFXA_', False),

    # Visibility
    'zVisibilityType': ('Visibility', 'FVT_', False),

    # Shape evaluation
    'zShapeEvaluator': ('ShapeEval', 'FSE_', False),
    'zShapeOrient': ('ShapeOrient', 'FSO_', False),

    # Data conversion
    'zDataConvert': ('DataConvert', 'FDC_', False),

    # Chart axis
    'zChartAxisStyle': ('ChartAxis', 'FCAS_', False),
    'zChartNumberFormat': ('ChartNumFmt', 'FCNF_', False),
    'zChartMarkerStyle': ('ChartMarker', 'FCMS_', False),
    'zChartLegendLocation': ('ChartLegend', 'FCLL_', False),
    'zChartTextJustification': ('ChartJustify', 'FCTJ_', False),

    # Beam calculations
    'zBeamCalculatorStressComponent': ('BeamStress', 'FBCSC_', False),

    # Monitor points
    'zMptComponent': ('MptComponent', 'FMPC_', False),
    'zMptContribution': ('MptContrib', 'FMPCN_', False),
}


def parse_constants_from_tlb(tlb_path: str) -> Dict[str, List[ConstantInfo]]:
    """Parse constants directly from .tlb file (authoritative source)."""
    print(f"Loading type library: {tlb_path}")
    typelib = pythoncom.LoadTypeLib(tlb_path)

    constants: Dict[str, List[ConstantInfo]] = defaultdict(list)
    count = typelib.GetTypeInfoCount()

    for i in range(count):
        tinfo = typelib.GetTypeInfo(i)
        name = typelib.GetDocumentation(i)[0]
        attr = tinfo.GetTypeAttr()

        # Only process enums (typekind == 0)
        if attr.typekind != TKIND_ENUM:
            continue

        # Extract enum members
        for j in range(attr.cVars):
            try:
                vardesc = tinfo.GetVarDesc(j)
                member_name = tinfo.GetNames(vardesc.memid)[0]
                member_value = vardesc.value
                if isinstance(member_value, int):
                    constants[name].append(ConstantInfo(member_name, member_value, name))
            except Exception:
                continue

    return dict(constants)


def strip_prefix(name: str, prefix: str) -> str:
    """Strip prefix from constant name, handling edge cases."""
    if name.startswith(prefix):
        result = name[len(prefix):]
        # If result starts with digit, prefix with underscore
        if result and result[0].isdigit():
            result = '_' + result
        return result
    return name


def generate_nested_class(constants: List[ConstantInfo], prefix: str, enum_name: str) -> Tuple[List[str], List[str]]:
    """Generate nested class structure with IntEnum subclasses.

    E.g., FGD_ELEM_BYCOLOR -> GroupDef.ELEM.BYCOLOR

    Returns:
        Tuple of (lines, nested_class_names) for generating union type alias
    """
    lines = []
    nested_class_names = []

    # Group by secondary prefix (e.g., ELEM, NODE, CSYS, etc.)
    groups = defaultdict(list)
    ungrouped = []

    for const in constants:
        stripped = strip_prefix(const.name, prefix)
        # Split on first underscore to get secondary group
        parts = stripped.split('_', 1)
        if len(parts) == 2:
            group_name = parts[0]
            member_name = parts[1]
            groups[group_name].append((member_name, const))
        else:
            ungrouped.append((stripped, const))

    # Generate nested IntEnum classes for each group
    for group_name in sorted(groups.keys()):
        members = groups[group_name]
        lines.append(f"    class {group_name}(IntEnum):")
        nested_class_names.append(group_name)
        for member_name, const in sorted(members, key=lambda x: x[1].value):
            # Clean up member name
            if member_name and member_name[0].isdigit():
                member_name = '_' + member_name
            lines.append(f"        {member_name} = {const.value}")
        lines.append("")

    # Add ungrouped constants at class level (as plain ints)
    if ungrouped:
        lines.append("    # Ungrouped constants")
        for member_name, const in sorted(ungrouped, key=lambda x: x[1].value):
            lines.append(f"    {member_name} = {const.value}")
        lines.append("")

    return lines, nested_class_names


def generate_flat_class(constants: list[ConstantInfo], prefix: str, enum_name: str, class_name: str) -> list[str]:
    """Generate a flat IntEnum class with all constants as members."""
    lines = []

    for const in sorted(constants, key=lambda x: x.value):
        alias = strip_prefix(const.name, prefix)
        if not alias:
            alias = const.name  # Fallback to full name if prefix is the entire name
        lines.append(f"    {alias} = {const.value}")

    return lines


def detect_prefixes(const_list: List[ConstantInfo]) -> Dict[str, List[ConstantInfo]]:
    """Detect common prefixes in constant names for nested subclassing."""
    prefix_groups = defaultdict(list)

    for const in const_list:
        # Try to find common prefix pattern (e.g., APIWARN_, CTRLDEF_, FCL_, FPF_)
        parts = const.name.split('_', 1)
        if len(parts) > 1:
            prefix = parts[0] + '_'
            prefix_groups[prefix].append(const)
        else:
            # No underscore - put in root
            prefix_groups[''].append(const)

    # Only use nested structure if multiple prefixes detected
    if len(prefix_groups) > 1 and '' not in prefix_groups:
        return dict(prefix_groups)
    else:
        # Single prefix or mixed - flatten
        return {'': const_list}


def generate_tier2_direct(enums: Dict[str, List[ConstantInfo]]) -> List[str]:
    """Generate IntEnum classes with nested subclasses for multi-prefix enums."""
    lines = []
    lines.append('')
    lines.append('# ' + '='*70)
    lines.append('# Tier 2: Auto-Generated Enums (Direct Mapping)')
    lines.append('# ' + '='*70)
    lines.append('# The following enums are generated directly from the .tlb file')
    lines.append('# Enums with multiple prefixes use nested subclasses for organization.')
    lines.append('# Constant names are preserved exactly as they appear in the .tlb.')
    lines.append('')

    tier2_count = 0
    tier2_enums = 0

    for enum_name in sorted(enums.keys()):
        const_list = enums[enum_name]
        if not const_list:
            continue

        tier2_enums += 1
        tier2_count += len(const_list)

        # Detect if this enum has multiple prefixes
        prefix_groups = detect_prefixes(const_list)

        if len(prefix_groups) == 1 and '' in prefix_groups:
            # Simple flat enum
            lines.append(f'class {enum_name}(IntEnum):')
            lines.append(f'    """Constants from {enum_name} enum (auto-generated)."""')
            lines.append('')

            for const in sorted(const_list, key=lambda c: c.value):
                lines.append(f'    {const.name} = {const.value}')

            lines.append('')
            lines.append('')
        else:
            # Nested structure for multiple prefixes
            lines.append(f'class {enum_name}:')
            lines.append(f'    """Constants from {enum_name} enum (auto-generated, nested by prefix)."""')
            lines.append('')

            for prefix in sorted(prefix_groups.keys()):
                # Create nested class for each prefix
                class_name = prefix.rstrip('_') if prefix else 'Other'
                lines.append(f'    class {class_name}(IntEnum):')
                lines.append(f'        """Constants with {prefix} prefix."""')
                lines.append('')

                for const in sorted(prefix_groups[prefix], key=lambda c: c.value):
                    lines.append(f'        {const.name} = {const.value}')

                lines.append('')

            lines.append('')

    return lines, tier2_count, tier2_enums


def generate_constants_file(constants: dict[str, list[ConstantInfo]], output_path: Path):
    """Generate the femap_constants.py file."""

    lines = [
        '"""',
        'femap_constants.py - Type-safe constant aliases for Femap API',
        '',
        'Auto-generated by generate_constants_tlb.py from Femap type library (.tlb)',
        'DO NOT EDIT MANUALLY - regenerate using: python generate_constants_tlb.py',
        '',
        'Uses IntEnum for type safety - each enum is a distinct type that',
        'type checkers can verify (e.g., ReturnCode vs Color are not interchangeable).',
        '',
        'Usage:',
        '    from femap_constants import ReturnCode, Entity, Message, Color',
        '',
        '    if rc == ReturnCode.OK:',
        '        app.feAppMessage(Message.NORMAL, "Success!")',
        '"""',
        '',
        'from enum import IntEnum',
        '',
        '',
    ]

    # Track which enums were processed in Tier 1 (curated aliases)
    tier1_processed = []
    tier1_skipped = []
    tier1_enum_names = set()  # Track which enum names are in ALIAS_CONFIG

    # Tier 1: Curated aliases from ALIAS_CONFIG
    for config_key, config in ALIAS_CONFIG.items():
        # Handle virtual enum syntax: "zColor:FPF_" means filter zColor by FPF_ prefix
        if ':' in config_key:
            enum_name, filter_prefix = config_key.split(':', 1)
        else:
            enum_name = config_key
            filter_prefix = None

        if enum_name not in constants:
            tier1_skipped.append(config_key)
            continue

        # Track that this enum is in ALIAS_CONFIG (for Tier 2 filtering)
        tier1_enum_names.add(enum_name)

        class_name, prefix, use_nested = config
        const_list = constants[enum_name]

        # Filter constants if a filter prefix is specified
        if filter_prefix:
            const_list = [c for c in const_list if c.name.startswith(filter_prefix)]
            if not const_list:
                tier1_skipped.append(config_key)
                continue

        tier1_processed.append((config_key, class_name, len(const_list)))

        # Generate class header
        if use_nested:
            # Nested classes can't be IntEnum, use regular class as container
            lines.append(f'class {class_name}:')
            lines.append(f'    """Constants from {enum_name} enum (nested grouping)."""')
        else:
            lines.append(f'class {class_name}(IntEnum):')
            lines.append(f'    """Constants from {enum_name} enum."""')
        lines.append('')

        # Generate class body
        if use_nested:
            body, nested_names = generate_nested_class(const_list, prefix, enum_name)
            lines.extend(body)
            lines.append('')
            # Generate type alias as union of nested IntEnum classes
            if nested_names:
                union_parts = [f'{class_name}.{n}' for n in nested_names]
                lines.append(f'# Type alias for {class_name} nested IntEnums')
                lines.append(f'{class_name}Type = {" | ".join(union_parts)}')
                lines.append('')
        else:
            body = generate_flat_class(const_list, prefix, enum_name, class_name)
            lines.extend(body)
            lines.append('')

        lines.append('')

    # Tier 2: Auto-generated enums (all enums NOT in ALIAS_CONFIG)
    tier2_enums = {k: v for k, v in constants.items() if k not in tier1_enum_names}

    if tier2_enums:
        tier2_lines, tier2_const_count, tier2_enum_count = generate_tier2_direct(tier2_enums)
        lines.extend(tier2_lines)
    else:
        tier2_const_count = 0
        tier2_enum_count = 0

    # Calculate totals
    tier1_const_count = sum(count for _, _, count in tier1_processed)
    total_enums = len(tier1_processed) + tier2_enum_count
    total_constants = tier1_const_count + tier2_const_count

    # Add summary comment at the end
    lines.append('# ' + '=' * 70)
    lines.append('# Generation Summary')
    lines.append('# ' + '=' * 70)
    lines.append(f'# Tier 1: Curated Aliases ({len(tier1_processed)} enums, {tier1_const_count} constants)')
    for enum_name, class_name, count in tier1_processed:
        lines.append(f'#   {class_name} ({count} constants) from {enum_name}')
    if tier1_skipped:
        lines.append(f'# Tier 1 Skipped: {len(tier1_skipped)} enums (not found in .tlb):')
        for enum_name in tier1_skipped:
            lines.append(f'#   {enum_name}')
    lines.append('#')
    lines.append(f'# Tier 2: Auto-Generated ({tier2_enum_count} enums, {tier2_const_count} constants)')
    lines.append('#   (All remaining enums from .tlb with direct mapping)')
    lines.append('#')
    lines.append(f'# Total: {total_enums} enums, {total_constants} constants')

    # Write file
    output_path.write_text('\n'.join(lines), encoding='utf-8')
    print(f"Generated {output_path}")
    print(f"  Tier 1 (Curated): {len(tier1_processed)} enums, {tier1_const_count} constants")
    print(f"  Tier 2 (Auto-gen): {tier2_enum_count} enums, {tier2_const_count} constants")
    print(f"  Total: {total_enums} enums, {total_constants} constants")
    if tier1_skipped:
        print(f"  Tier 1 Skipped: {len(tier1_skipped)} enums (not in .tlb)")


def print_available_enums(constants: Dict[str, List[ConstantInfo]]):
    """Print all available enums for reference."""
    print("\nAll available enums in .tlb:")
    print("-" * 50)
    for enum_name in sorted(constants.keys()):
        const_list = constants[enum_name]
        # Get common prefix
        if const_list:
            prefixes = set()
            for c in const_list:
                parts = c.name.split('_')
                if len(parts) >= 2:
                    prefixes.add(parts[0] + '_')
            prefix_str = ', '.join(sorted(prefixes)[:3])
            if len(prefixes) > 3:
                prefix_str += '...'
        else:
            prefix_str = ''
        print(f"  {enum_name}: {len(const_list)} constants (prefixes: {prefix_str})")


def main():
    parser = argparse.ArgumentParser(
        description='Generate femap_constants.py from Femap type library'
    )
    parser.add_argument(
        '--tlb',
        default=None,
        help='Path to femap.tlb file (if not specified, will auto-detect or prompt)'
    )
    parser.add_argument(
        '--output',
        default='femap_constants.py',
        help='Output file path'
    )
    parser.add_argument(
        '--list-enums',
        action='store_true',
        help='List all available enums and exit'
    )
    args = parser.parse_args()

    # Resolve the .tlb path using multiple strategies
    tlb_path = get_tlb_path(args.tlb)
    if not tlb_path:
        print("ERROR: No type library selected")
        return 1

    script_dir = Path(__file__).parent
    output_path = script_dir / args.output

    print("Parsing .tlb constants...")
    constants = parse_constants_from_tlb(tlb_path)
    print(f"Found {sum(len(v) for v in constants.values())} constants in {len(constants)} enums")

    if args.list_enums:
        print_available_enums(constants)
        return 0

    print("\nGenerating femap_constants.py...")
    generate_constants_file(constants, output_path)

    return 0


if __name__ == "__main__":
    exit(main())
