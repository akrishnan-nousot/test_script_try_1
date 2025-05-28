#!/usr/bin/env python3
import zipfile, io, os, re, argparse
import pandas as pd
from xml.etree import ElementTree as ET

def extract_field_mappings(wid_path, output_dir):
    """Extract all fields (data + calculated) from a WebI .wid into Excel/CSV."""
    basename = os.path.splitext(os.path.basename(wid_path))[0]
    print(f"\n=== Processing WID: {basename} ===")

    # 1) Read WID, locate embedded ZIP
    raw = open(wid_path, 'rb').read()
    idx = raw.find(b'PK\x03\x04')
    if idx < 0:
        raise ValueError("Not a valid .wid (no ZIP signature).")
    z = zipfile.ZipFile(io.BytesIO(raw[idx:]))

    # --- Print all files/folders inside the embedded ZIP ---
    print("Files inside the WID ZIP:")
    for zi in z.infolist():
        print(f"  - {zi.filename}")

    mappings = []

    # 2) Classify each DP# as BEx / EMP / Universes
    dp_names = extract_data_provider_names(z)

    # 3) Pull document‐level calc fields & variables
    calc_fields = extract_calculated_fields(z)
    report_vars  = extract_report_variables(z)

    # 4) Parse each Data Provider’s DP_Generic
    dp_folders = [f for f in z.namelist() if f.endswith('/') and '/DATAPROVIDERS/DP' in f]
    for dp_folder in dp_folders:
        dp_id = dp_folder.split('/')[-2]
        dp_nm = dp_names.get(dp_id, dp_id)
        dpg   = dp_folder + "DP_Generic"
        if dpg not in z.namelist():
            continue

        raw_txt = z.read(dpg)
        for enc in ('utf-8','utf-16le','latin-1'):
            try:
                text = raw_txt.decode(enc, errors='ignore')
                break
            except:
                continue

        # extract the full QuerySpec XML
        xml = extract_xml(text)
        field_info  = extract_comprehensive_field_info(xml) if xml else {}
        direct_map  = extract_direct_mappings(text)
        dp_forms    = extract_dp_formulas(text, dp_id)

        process_enhanced_mappings(dp_id, dp_nm, field_info, direct_map, dp_forms, mappings)

    # 5) Append document‐level objects (calculated fields & variables)
    for cf in calc_fields:
        mappings.append({
            'Data Provider ID':'CALC',
            'Data Provider':   'Document Variables',
            'Technical ID':    cf['id'],
            'Display Name':    cf['name'],
            'Field Type':      'Calculated Field',
            'Formula':         cf['formula'],
            'Description':     cf['description'],
            'Sample Value':    '',
            'Databricks Table':'',
            'Databricks Column': clean_column_name(cf['name']),
            'XML ID':          ''
        })
    for rv in report_vars:
        mappings.append({
            'Data Provider ID':'VAR',
            'Data Provider':   'Report Variables',
            'Technical ID':    rv['id'],
            'Display Name':    rv['name'],
            'Field Type':      'Report Variable',
            'Formula':         rv['formula'],
            'Description':     rv['description'],
            'Sample Value':    '',
            'Databricks Table':'',
            'Databricks Column': clean_column_name(rv['name']),
            'XML ID':          ''
        })

    # 6) Build DataFrame & write Excel + CSV
    df = pd.DataFrame(mappings)
    if df.empty:
        print("⚠️  No fields extracted—check the WID structure.")
        return

    cols = [
        'Data Provider ID','Data Provider','Technical ID','Display Name',
        'Field Type','Formula','Description','Sample Value',
        'Databricks Table','Databricks Column','XML ID'
    ]
    for c in cols:
        if c not in df.columns:
            df[c] = ''
    df = df[cols].sort_values(['Data Provider','Display Name'])

    # Write Excel
    excel_path = os.path.join(output_dir, f"{basename}_field_map.xlsx")
    with pd.ExcelWriter(excel_path, engine='openpyxl') as w:
        df.to_excel(w, sheet_name='All Fields', index=False)
        cf = df[df['Field Type'].isin(['Calculated Field','Report Variable'])]
        if not cf.empty:
            cf.to_excel(w, sheet_name='Calculated Fields', index=False)
        df[~df['Field Type'].isin(['Calculated Field','Report Variable'])]\
          .to_excel(w, sheet_name='Data Fields', index=False)

    # Write CSV
    csv_path = os.path.join(output_dir, f"{basename}_field_map.csv")
    df.to_csv(csv_path, index=False)

    print(f"✅ Written:\n  • {excel_path}\n  • {csv_path}\n")


def extract_data_provider_names(z):
    """Classify each DP# by inspecting its embedded QuerySpec XML."""
    dp_names = {}
    for fn in z.namelist():
        if fn.endswith('DP_Generic'):
            m = re.search(r'(DP\d+)', fn)
            if not m:
                continue
            dp = m.group(1)
            raw = z.read(fn)
            # decode text
            for enc in ('utf-8','utf-16le','latin-1'):
                try:
                    txt = raw.decode(enc, errors='ignore')
                    break
                except:
                    continue

            # pull the clean QuerySpec XML
            qxml = extract_xml(txt)
            if qxml:
                lo = qxml.lower()
                if 'com.sap.sl.queryspec' in lo or 'bex' in lo:
                    dp_names[dp] = 'BEx'
                elif re.search(r'\bemp\b', lo):
                    dp_names[dp] = 'EMP'
                elif re.search(r'\buniverses?\b', lo):
                    dp_names[dp] = 'universes'
                else:
                    dp_names[dp] = dp
            else:
                dp_names[dp] = dp
    return dp_names


def extract_xml(text):
    """
    Extract the full <…QuerySpec…>…</…QuerySpec…> block, 
    regardless of namespace prefix.
    """
    if not text:
        return None
    m = re.search(r'(<[^>]*QuerySpec[^>]*>.*?</[^>]*QuerySpec[^>]*>)',
                  text, flags=re.DOTALL|re.IGNORECASE)
    return m.group(1) if m else None


def extract_comprehensive_field_info(xml_text):
    """
    Parse QuerySpec XML for:
      • <object id=… name=… expression=…/>
      • <QueryCalculation alias=… calculation=…/>
    """
    info = {}
    if not xml_text:
        return info

    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError as ex:
        print("XML parse error in QuerySpec:", ex)
        return info

    # 1) All normal objects/fields
    for el in root.findall('.//*'):
        fid = (el.get('id') or el.get('identifier')
               or el.get('uniqueName') or el.get('technicalName'))
        disp = (el.get('name') or el.get('displayName')
                or el.get('caption'))
        expr = (el.get('expression') or el.get('formula')
                or el.get('calculation') or '')
        dt   = el.get('dataType') or el.get('type') or ''

        if fid and disp:
            info[fid] = {
                'display_name':  disp.strip(),
                'expression':    expr.strip(),
                'element_type':  el.tag,
                'data_type':     dt,
                'is_calculated': bool(expr.strip())
            }

    # 2) Explicit QueryCalculation elements (ActionCount, etc.)
    for m in re.finditer(
        r'<QueryCalculation[^>]+alias=["\']([^"\']+)["\'][^>]+calculation=["\']([^"\']+)["\']',
        xml_text, flags=re.DOTALL|re.IGNORECASE
    ):
        alias = m.group(1).strip()
        calc  = m.group(2).strip()
        key   = f"QC::{alias}"
        info[key] = {
            'display_name':  alias,
            'expression':    calc,
            'element_type':  'QueryCalculation',
            'data_type':     '',
            'is_calculated': True
        }

    return info


def extract_direct_mappings(text):
    """
    Capture lines like 'Foo = 123, DP2.DO4' and standalone DPx.DOy references.
    """
    dm = {}
    pat1 = re.compile(r'([^=\n]+?)\s*=\s*([^,;\n]+)[,;\s]+(DP\d+\.DO\w+)', re.IGNORECASE)
    for m in pat1.finditer(text):
        dm[m.group(3)] = {
            'display_name': m.group(1).strip(),
            'sample_value': m.group(2).strip()
        }
    pat2 = re.compile(r'(DP\d+\.DO\w+)', re.IGNORECASE)
    for m in pat2.finditer(text):
        dm.setdefault(m.group(1), {'display_name':'','sample_value':''})
    return dm


def extract_dp_formulas(text, dp_id):
    """Grab any formula=… or expression=… or calculation=… attributes."""
    out = {}
    for pat in (r'formula=["\']([^"\']+)["\']',
                r'expression=["\']([^"\']+)["\']',
                r'calculation=["\']([^"\']+)["\']'):
        for i,m in enumerate(re.finditer(pat, text, re.IGNORECASE)):
            out[f"{dp_id}_FORM_{i}"] = m.group(1).strip()
    return out


def process_enhanced_mappings(dp_id, dp_name, field_info, direct_map, dp_forms, out):
    """Merge DP‐level fields, direct mappings, and DP formulas into our master list."""
    for fid,fi in field_info.items():
        out.append({
            'Data Provider ID': dp_id,
            'Data Provider':    dp_name,
            'Technical ID':     fid,
            'Display Name':     fi['display_name'],
            'Field Type':       'Calculated Field' if fi['is_calculated'] else 'Data Field',
            'Formula':          fi['expression'],
            'Description':      f"{fi['element_type']}, type={fi['data_type']}",
            'Sample Value':     '',
            'Databricks Table':'',
            'Databricks Column':clean_column_name(fi['display_name']),
            'XML ID':           fid
        })
    for tid,mm in direct_map.items():
        if any(r['Technical ID']==tid for r in out):
            continue
        out.append({
            'Data Provider ID': dp_id,
            'Data Provider':    dp_name,
            'Technical ID':     tid,
            'Display Name':     mm['display_name'],
            'Field Type':       'Data Field',
            'Formula':          '',
            'Description':      'Direct mapping',
            'Sample Value':     mm['sample_value'],
            'Databricks Table':'',
            'Databricks Column':clean_column_name(mm['display_name'] or tid),
            'XML ID':''
        })
    for fid,formula in dp_forms.items():
        out.append({
            'Data Provider ID': dp_id,
            'Data Provider':    dp_name,
            'Technical ID':     fid,
            'Display Name':     f"Formula {fid}",
            'Field Type':       'Data Provider Formula',
            'Formula':          formula,
            'Description':      'From DP_Generic',
            'Sample Value':'',
            'Databricks Table':'',
            'Databricks Column':clean_column_name(f"formula_{fid}"),
            'XML ID':''
        })


def extract_calculated_fields(z):
    """Pull any <CalculatedField…> or <QueryCalculation…> from document files."""
    results = []
    patterns = [
      r'<CalculatedField[^>]+name=["\']([^"\']+)["\'][^>]*>([^<]+)',
      r'<QueryCalculation[^>]+alias=["\']([^"\']+)["\'][^>]+calculation=["\']([^"\']+)["\']'
    ]
    for fn in z.namelist():
        if any(x in fn.upper() for x in ('DOCUMENT','REPORT','STRUCTURE')):
            raw = z.read(fn)
            try:
                txt = raw.decode('utf-16le', errors='ignore')
            except:
                txt = raw.decode('latin-1', errors='ignore')
            for pat in patterns:
                for m in re.finditer(pat, txt, re.IGNORECASE|re.DOTALL):
                    results.append({
                        'id':        f"CF_{len(results)}",
                        'name':      m.group(1).strip(),
                        'formula':   m.group(2).strip(),
                        'description':f"From {fn}"
                    })
    return results


def extract_report_variables(z):
    """Grab any old‐style `Variable X = formula;` definitions."""
    results = []
    for fn in z.namelist():
        if any(x in fn.upper() for x in ('DOCUMENT','VARIABLE','FORMULA')):
            raw = z.read(fn)
            try:
                txt = raw.decode('utf-16le', errors='ignore')
            except:
                    txt = raw.decode('latin-1', errors='ignore')
            for m in re.finditer(r'Variable\s+([^=\s]+)\s*=\s*([^;]+);', txt, re.IGNORECASE):
                    results.append({
                    'id':         f"RV_{len(results)}",
                    'name':       m.group(1).strip(),
                    'formula':    m.group(2).strip(),
                    'description':f"From {fn}"
                })
    return results


def clean_column_name(name):
    """Normalize to lowercase_underscore, prefix if starts non‐alpha."""
    s = re.sub(r'[^A-Za-z0-9\s]', '', name or '')
    s = re.sub(r'\s+', '_', s).strip('_').lower()
    if s and not s[0].isalpha():
        s = 'f_' + s
    return s


def main():
    p = argparse.ArgumentParser(
        description="Extract WebI .wid fields & calculations into Excel/CSV"
    )
    p.add_argument("input",  help="Path to .wid file or folder containing .wid")
    p.add_argument("output", help="Directory to write outputs")
    args = p.parse_args()
    os.makedirs(args.output, exist_ok=True)

    if os.path.isdir(args.input):
        for f in os.listdir(args.input):
            if f.lower().endswith('.wid'):
                extract_field_mappings(os.path.join(args.input, f), args.output)
    else:
        extract_field_mappings(args.input, args.output)


if __name__ == "__main__":
    main()