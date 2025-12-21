"""
Authoritative Stub Generator from Femap Type Library (.tlb)
============================================================
Generates Pyfemap.pyi with accurate type information directly from the
COM type library, preserving interface types (IMatl, INode) and enum
types (zColor, zReturnCode) that are lost in the makepy-generated Pyfemap.py.

Usage:
    python generate_stubs_tlb.py [--tlb PATH] [--output PATH]

Path Resolution:
    If --tlb is not specified, will use this order:
    1. Environment variable FEMAP_TLB_PATH
    2. Auto-detect in common Femap installation paths
    3. Prompt with file dialog
"""

import argparse
import pythoncom
from typing import Dict, List, Tuple, Set, Any, Optional
from femap_path_utils import get_tlb_path

# VT type codes to Python type names
VT_NAMES = {
    0: 'None',      # VT_EMPTY
    2: 'int',       # VT_I2
    3: 'int',       # VT_I4
    4: 'float',     # VT_R4
    5: 'float',     # VT_R8
    6: 'float',     # VT_CY (currency)
    7: 'float',     # VT_DATE
    8: 'str',       # VT_BSTR
    9: 'Any',       # VT_DISPATCH
    10: 'int',      # VT_ERROR
    11: 'bool',     # VT_BOOL
    12: 'Any',      # VT_VARIANT
    13: 'Any',      # VT_UNKNOWN
    16: 'int',      # VT_I1
    17: 'int',      # VT_UI1
    18: 'int',      # VT_UI2
    19: 'int',      # VT_UI4
    20: 'int',      # VT_I8
    21: 'int',      # VT_UI8
    22: 'int',      # VT_INT
    23: 'int',      # VT_UINT
    24: 'None',     # VT_VOID
    25: 'int',      # VT_HRESULT
}

# Type kind constants
TKIND_ENUM = 0
TKIND_RECORD = 1
TKIND_MODULE = 2
TKIND_INTERFACE = 3
TKIND_DISPATCH = 4
TKIND_COCLASS = 5
TKIND_ALIAS = 6
TKIND_UNION = 7

# INVOKEKIND constants
INVOKE_FUNC = 1
INVOKE_PROPERTYGET = 2
INVOKE_PROPERTYPUT = 4
INVOKE_PROPERTYPUTREF = 8

# Import the alias configuration from generate_constants_tlb.py
# This ensures both generators use the same mapping
from generate_constants_tlb import ALIAS_CONFIG

# Build the enum-to-alias mapping from ALIAS_CONFIG
# ALIAS_CONFIG format: enum_name -> (AliasClassName, prefix_to_strip, use_nested_grouping)
# We also handle virtual entries like 'zColor:FPF_' -> 'BrushPattern'
ENUM_ALIAS_MAP = {}
ENUM_UNION_MAP = {}  # For enums that have virtual subsets (e.g., zColor -> Color | BrushPattern | PenLineStyle)

for enum_key, (alias_name, prefix, nested) in ALIAS_CONFIG.items():
    if ':' in enum_key:
        # Virtual entry like 'zColor:FPF_' -> 'BrushPattern'
        base_enum = enum_key.split(':')[0]
        if base_enum not in ENUM_UNION_MAP:
            ENUM_UNION_MAP[base_enum] = []
        ENUM_UNION_MAP[base_enum].append(alias_name)
    else:
        # Primary mapping
        ENUM_ALIAS_MAP[enum_key] = alias_name

# Add the primary alias to union map entries
for base_enum, subsets in ENUM_UNION_MAP.items():
    if base_enum in ENUM_ALIAS_MAP:
        # Insert primary alias at the front
        subsets.insert(0, ENUM_ALIAS_MAP[base_enum])


def resolve_type(tinfo, typedesc) -> str:
    """
    Resolve a type descriptor to a Python type string.

    Type descriptor formats from COM:
    - Simple int: 3 = VT_I4 = int
    - Tuple (vt, flags, default): (22, 0, None) = VT_INT
    - VT_USERDEFINED: (29, href) = enum/interface reference
    - VT_PTR: (26, inner) = pointer to another type
    - VT_SAFEARRAY: (27, inner) = array of type
    - Nested in elemdesc: ((29, 256), 0, None) = VT_USERDEFINED with flags
    """
    if typedesc is None:
        return 'Any'

    # Handle simple int
    if isinstance(typedesc, int):
        return VT_NAMES.get(typedesc, 'Any')

    # Handle tuple format
    if isinstance(typedesc, tuple):
        if len(typedesc) == 0:
            return 'Any'

        vt = typedesc[0]

        # Check if this is an elemdesc tuple: (type_desc, flags, default)
        # where type_desc itself could be a tuple like (29, href)
        if isinstance(vt, tuple):
            # The first element is the actual type descriptor
            return resolve_type(tinfo, vt)

        # vt is the type code
        if isinstance(vt, int):
            if vt == 26:  # VT_PTR
                # Pointer - get the inner type
                inner = typedesc[1] if len(typedesc) > 1 else None
                inner_type = resolve_type(tinfo, inner)
                # Don't wrap basic types in pointer notation
                return inner_type

            elif vt == 27:  # VT_SAFEARRAY
                inner = typedesc[1] if len(typedesc) > 1 else None
                if inner is not None:
                    elem_type = resolve_type(tinfo, inner)
                    return f'Tuple[{elem_type}, ...]'
                return 'Tuple[Any, ...]'

            elif vt == 29:  # VT_USERDEFINED
                href = typedesc[1] if len(typedesc) > 1 else None
                if href is not None:
                    try:
                        ref_tinfo = tinfo.GetRefTypeInfo(href)
                        name = ref_tinfo.GetDocumentation(-1)[0]
                        return name
                    except Exception:
                        return 'int'
                return 'int'

            # Simple VT code
            return VT_NAMES.get(vt, 'Any')

    return 'Any'


def get_elemdesc_type(tinfo, elemdesc) -> Tuple[str, bool]:
    """Extract type from an ELEMDESC structure.

    elemdesc is a tuple: (type_desc, flags, default_value)
    type_desc can be:
    - Simple int: 22 = VT_INT
    - Tuple: (29, href) = VT_USERDEFINED
    - Nested: ((26, 12), ...) = VT_PTR to VT_VARIANT

    Returns: (type_string, is_output_param)
    PARAMFLAG_FOUT = 0x2 indicates an output parameter
    """
    if elemdesc is None:
        return ('Any', False)

    is_output = False

    # elemdesc is (type_desc, flags, default) or just type_desc
    if isinstance(elemdesc, tuple) and len(elemdesc) >= 2:
        # Check flags for PARAMFLAG_FOUT (0x2)
        flags = elemdesc[1]
        if isinstance(flags, int) and (flags & 0x2):
            is_output = True

    if isinstance(elemdesc, tuple) and len(elemdesc) >= 1:
        return (resolve_type(tinfo, elemdesc), is_output)

    return ('Any', is_output)


def extract_enum_values(typelib, idx: int) -> Dict[str, int]:
    """Extract enum member names and values."""
    tinfo = typelib.GetTypeInfo(idx)
    attr = tinfo.GetTypeAttr()

    members = {}
    for j in range(attr.cVars):
        vardesc = tinfo.GetVarDesc(j)
        name = tinfo.GetNames(vardesc.memid)[0]
        value = vardesc.value
        members[name] = value

    return members


def extract_interface_info(typelib, idx: int, enums: Set[str]) -> Optional[Dict]:
    """Extract properties and methods from a DISPATCH interface."""
    tinfo = typelib.GetTypeInfo(idx)
    name = typelib.GetDocumentation(idx)[0]
    attr = tinfo.GetTypeAttr()

    # Only process DISPATCH interfaces (typekind == 4)
    if attr.typekind != TKIND_DISPATCH:
        return None

    properties = {}  # name -> (type, has_setter)
    methods = []
    indexed_properties = {}  # name -> {'getter_params': [...], 'setter_params': [...], 'type': str}

    # Get variables (properties defined via cVars) - these are simple properties
    for j in range(attr.cVars):
        try:
            vardesc = tinfo.GetVarDesc(j)
            var_names = tinfo.GetNames(vardesc.memid)
            if var_names:
                var_name = var_names[0]
                var_type = resolve_type(tinfo, vardesc.elemdescVar)
                # Properties from vars typically have both getter and setter
                properties[var_name] = (var_type, True)
        except Exception:
            continue

    # Get functions (includes property getters/setters and methods)
    for j in range(attr.cFuncs):
        try:
            funcdesc = tinfo.GetFuncDesc(j)
        except Exception:
            continue

        try:
            func_names = tinfo.GetNames(funcdesc.memid)
        except Exception:
            continue

        if not func_names:
            continue

        func_name = func_names[0]

        # Handle property getters
        if funcdesc.invkind == INVOKE_PROPERTYGET:
            ret_type = resolve_type(tinfo, funcdesc.rettype)

            # Check if this is an indexed property (getter has parameters)
            if funcdesc.args and len(funcdesc.args) > 0:
                # Indexed property - treat as method pair
                params = []
                for k, arg in enumerate(funcdesc.args):
                    param_name = func_names[k + 1] if k + 1 < len(func_names) else f'arg{k}'
                    param_type, _ = get_elemdesc_type(tinfo, arg)  # Ignore is_output for indexed props
                    params.append((param_name, param_type))

                if func_name not in indexed_properties:
                    indexed_properties[func_name] = {'getter_params': params, 'setter_params': None, 'type': ret_type}
                else:
                    indexed_properties[func_name]['getter_params'] = params
                    indexed_properties[func_name]['type'] = ret_type
            else:
                # Simple property (no parameters)
                if func_name not in properties:
                    properties[func_name] = (ret_type, False)
                else:
                    # Update type but preserve setter flag
                    properties[func_name] = (ret_type, properties[func_name][1])
            continue

        # Handle property setters
        if funcdesc.invkind in (INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF):
            # Check if this is an indexed property setter (more than 1 parameter)
            # For indexed setters, args = [index_params..., value]
            if funcdesc.args and len(funcdesc.args) > 1:
                # Indexed property setter
                params = []
                for k, arg in enumerate(funcdesc.args):
                    param_name = func_names[k + 1] if k + 1 < len(func_names) else f'arg{k}'
                    param_type, _ = get_elemdesc_type(tinfo, arg)  # Ignore is_output for indexed props
                    params.append((param_name, param_type))

                if func_name not in indexed_properties:
                    indexed_properties[func_name] = {'getter_params': None, 'setter_params': params, 'type': 'Any'}
                else:
                    indexed_properties[func_name]['setter_params'] = params
            else:
                # Simple property setter
                if func_name in properties:
                    properties[func_name] = (properties[func_name][0], True)
                else:
                    # Setter without getter - get type from parameter
                    if funcdesc.args and len(funcdesc.args) > 0:
                        param_type, _ = get_elemdesc_type(tinfo, funcdesc.args[-1])
                        properties[func_name] = (param_type, True)
                    else:
                        properties[func_name] = ('Any', True)
            continue

        # Regular method (invkind == 1)
        ret_type = resolve_type(tinfo, funcdesc.rettype)

        # Parameters - get count from args tuple length
        # Track output parameters for return type tuple
        params = []
        output_types = []
        if funcdesc.args:
            for k, arg in enumerate(funcdesc.args):
                param_name = func_names[k + 1] if k + 1 < len(func_names) else f'arg{k}'
                param_type, is_output = get_elemdesc_type(tinfo, arg)
                if is_output:
                    # Output param becomes optional input and contributes to return tuple
                    params.append((param_name, param_type, True))  # (name, type, is_optional)
                    output_types.append(param_type)
                else:
                    params.append((param_name, param_type, False))

        # If there are output params, return type is Tuple[ret_type, ...output_types]
        if output_types:
            all_return_types = [ret_type] + output_types
            final_ret_type = f'Tuple[{", ".join(all_return_types)}]'
        else:
            final_ret_type = ret_type

        methods.append({
            'name': func_name,
            'params': params,
            'return_type': final_ret_type
        })

    # Convert indexed properties to method pairs
    for prop_name, prop_info in indexed_properties.items():
        # Add getter method
        if prop_info['getter_params'] is not None:
            methods.append({
                'name': prop_name,
                'params': prop_info['getter_params'],
                'return_type': prop_info['type']
            })

        # Add setter method (SetXxx)
        if prop_info['setter_params'] is not None:
            methods.append({
                'name': f'Set{prop_name}',
                'params': prop_info['setter_params'],
                'return_type': 'None'
            })

    # Convert properties dict to list
    prop_list = []
    for prop_name, (prop_type, has_setter) in properties.items():
        prop_list.append({
            'name': prop_name,
            'type': prop_type,
            'has_setter': has_setter
        })

    return {
        'name': name,
        'properties': prop_list,
        'methods': methods
    }


def translate_type(type_str: str) -> str:
    """Translate .tlb enum names to friendly alias names from femap_constants.

    For enums with virtual subsets (like zColor which contains Color, BrushPattern,
    PenLineStyle), returns a union type.
    """
    # Handle Tuple types like "Tuple[zReturnCode, int, Any]"
    if type_str.startswith('Tuple['):
        inner = type_str[6:-1]  # Remove "Tuple[" and "]"
        parts = []
        depth = 0
        current = ""
        for char in inner:
            if char == '[':
                depth += 1
            elif char == ']':
                depth -= 1
            elif char == ',' and depth == 0:
                parts.append(current.strip())
                current = ""
                continue
            current += char
        if current.strip():
            parts.append(current.strip())
        translated_parts = [translate_type(p) for p in parts]
        return f'Tuple[{", ".join(translated_parts)}]'

    # Check if this enum has virtual subsets (union type)
    if type_str in ENUM_UNION_MAP:
        # Return union of all subset types: Color | BrushPattern | PenLineStyle
        return ' | '.join(ENUM_UNION_MAP[type_str])

    # Direct translation of enum names
    return ENUM_ALIAS_MAP.get(type_str, type_str)


def generate_stub_file(interfaces: List[Dict], enums: Dict[str, Dict[str, int]],
                       output_path: str) -> None:
    """Generate the .pyi stub file."""
    # Collect which aliases are actually used
    used_aliases = set()

    lines = [
        '# Auto-generated from Femap type library (.tlb)',
        '# DO NOT EDIT - regenerate with generate_stubs_tlb.py',
        '#',
        '# This file provides authoritative type information extracted directly',
        '# from the COM type library, including interface types (IMatl, INode)',
        '# and enum types linked to femap_constants.py aliases.',
        '',
        'from typing import Any, Tuple, Optional, overload',
        'from win32com.client import DispatchBaseClass',
    ]

    # We'll add the femap_constants import after we know which aliases are used
    import_line_index = len(lines)
    lines.append('')  # Placeholder for import
    lines.append('')
    lines.append('')

    # Add enum type aliases - both original z* names AND friendly aliases
    lines.append('# Enum types (original .tlb names as aliases to int)')
    for enum_name in sorted(enums.keys()):
        # Keep original z* names for backward compatibility
        lines.append(f'{enum_name} = int')
    lines.append('')
    lines.append('')

    # Add constants class
    lines.append('class constants:')
    lines.append('    """Femap constants from type library."""')
    for enum_name, members in sorted(enums.items()):
        lines.append(f'    # {enum_name}')
        for member_name, value in sorted(members.items(), key=lambda x: x[1]):
            lines.append(f'    {member_name}: int')
    lines.append('')
    lines.append('')

    # Add interface classes (sorted for consistency)
    def track_and_translate(type_str: str) -> str:
        """Translate type and track which aliases are used."""
        translated = translate_type(type_str)
        # Track all aliases used (including union types like "Color | BrushPattern | PenLineStyle")
        for alias in ENUM_ALIAS_MAP.values():
            if alias in translated:
                used_aliases.add(alias)
        # Also track union type components
        for subsets in ENUM_UNION_MAP.values():
            for alias in subsets:
                if alias in translated:
                    used_aliases.add(alias)
        return translated

    for iface in sorted(interfaces, key=lambda x: x['name']):
        lines.append(f'class {iface["name"]}(DispatchBaseClass):')

        has_content = False

        # Properties
        for prop in sorted(iface['properties'], key=lambda x: x['name']):
            has_content = True
            prop_type = track_and_translate(prop['type'])
            lines.append(f'    @property')
            lines.append(f'    def {prop["name"]}(self) -> {prop_type}: ...')
            if prop['has_setter']:
                lines.append(f'    @{prop["name"]}.setter')
                lines.append(f'    def {prop["name"]}(self, value: {prop_type}) -> None: ...')

        # Methods
        for method in sorted(iface['methods'], key=lambda x: x['name']):
            has_content = True
            params = ['self']
            for param_info in method['params']:
                # Handle both 2-tuple (name, type) and 3-tuple (name, type, is_optional) formats
                if len(param_info) == 3:
                    param_name, param_type, is_optional = param_info
                else:
                    param_name, param_type = param_info
                    is_optional = False

                # Translate the parameter type
                param_type = track_and_translate(param_type)

                # Handle reserved words
                safe_name = param_name
                if param_name in ('type', 'id', 'list', 'set', 'from', 'import', 'class', 'in', 'is', 'not', 'and', 'or'):
                    safe_name = param_name + '_'

                if is_optional:
                    params.append(f'{safe_name}: {param_type} = ...')
                else:
                    params.append(f'{safe_name}: {param_type}')

            param_str = ', '.join(params)
            ret_type = track_and_translate(method['return_type'])
            lines.append(f'    def {method["name"]}({param_str}) -> {ret_type}: ...')

        if not has_content:
            lines.append('    ...')

        lines.append('')

    # Now insert the femap_constants import with used aliases
    if used_aliases:
        sorted_aliases = sorted(used_aliases)
        import_stmt = f'from femap_constants import {", ".join(sorted_aliases)}'
        lines[import_line_index] = import_stmt

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))


def main():
    parser = argparse.ArgumentParser(
        description='Generate Pyfemap.pyi type stubs from Femap type library'
    )
    parser.add_argument(
        '--tlb',
        default=None,
        help='Path to femap.tlb file (if not specified, will auto-detect or prompt)'
    )
    parser.add_argument(
        '--output',
        default='Pyfemap.pyi',
        help='Output .pyi file path'
    )
    args = parser.parse_args()

    # Resolve the .tlb path using multiple strategies
    tlb_path = get_tlb_path(args.tlb)
    if not tlb_path:
        print("ERROR: No type library selected")
        return 1

    print(f"Loading type library: {tlb_path}")
    try:
        typelib = pythoncom.LoadTypeLib(tlb_path)
    except Exception as e:
        print(f"Error loading type library: {e}")
        return 1

    count = typelib.GetTypeInfoCount()
    print(f"Found {count} types in library")

    # First pass: collect all enum names and values
    enums: Dict[str, Dict[str, int]] = {}
    enum_names: Set[str] = set()

    for i in range(count):
        tinfo = typelib.GetTypeInfo(i)
        name = typelib.GetDocumentation(i)[0]
        attr = tinfo.GetTypeAttr()

        if attr.typekind == TKIND_ENUM:
            members = extract_enum_values(typelib, i)
            enums[name] = members
            enum_names.add(name)

    print(f"Found {len(enums)} enums")

    # Second pass: extract interfaces
    interfaces: List[Dict] = []

    for i in range(count):
        iface = extract_interface_info(typelib, i, enum_names)
        if iface:
            interfaces.append(iface)

    print(f"Found {len(interfaces)} dispatch interfaces")

    # Generate stub file
    print(f"Generating {args.output}...")
    generate_stub_file(interfaces, enums, args.output)

    # Summary
    total_props = sum(len(i['properties']) for i in interfaces)
    total_methods = sum(len(i['methods']) for i in interfaces)
    print(f"Generated stubs for:")
    print(f"  - {len(enums)} enums")
    print(f"  - {len(interfaces)} interfaces")
    print(f"  - {total_props} properties")
    print(f"  - {total_methods} methods")
    print("Done!")

    return 0


if __name__ == '__main__':
    exit(main())
