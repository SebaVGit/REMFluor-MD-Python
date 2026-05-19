"""
inp_parser.py — parse input.inp and auxiliary .txt files.

Ported from Source_Py/restore_from_example_folder.py (parse_input_inp,
parse_additional_info, parse_retardation_pfas_names).
No xlwings / openpyxl dependency.
"""
import os


# ---------------------------------------------------------------------------
# input.inp parser
# ---------------------------------------------------------------------------

def _is_data_line(line: str) -> bool:
    """True if all comma-separated parts convert to float."""
    if not line:
        return False
    for part in line.split(','):
        try:
            float(part.strip())
        except ValueError:
            return False
    return True


def _next_data(lines, idx):
    for i in range(idx, len(lines)):
        if _is_data_line(lines[i]):
            return i, [x.strip() for x in lines[i].split(',')]
    return None, None


def parse_input_inp(path: str) -> dict:
    """Parse input.inp and return a dict of all model parameters."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"input.inp not found: {path}")

    with open(path) as f:
        lines = [ln.strip() for ln in f
                 if ln.strip() and not ln.strip().startswith('***')]

    data = {}
    idx = 0

    # flags: ncomp, ipre, iwall
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 3:
        data['ncomp'] = int(float(p[0]))
        data['ipre']  = int(float(p[1]))
        data['iwall'] = int(float(p[2]))
        idx += 1

    # source zone concentrations (11 rows)
    src = []
    for _ in range(11):
        idx, p = _next_data(lines, idx)
        if idx is None or len(p) < 5:
            break
        src.append({'time': float(p[0]), 'c1': float(p[1]),
                    'c2': float(p[2]),   'c3': float(p[3]), 'c4': float(p[4])})
        idx += 1
    data['source_zone_data'] = src

    # source zone parameters: lysource, lzsourcemax, lzsourcemin, vd
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 4:
        data['lysource']    = int(float(p[0]))
        data['lzsourcemax'] = int(float(p[1]))
        data['lzsourcemin'] = int(float(p[2]))
        data['vd']          = float(p[3])
        idx += 1

    # high-K transport: porf, tortf, bulkden, foc, alphax, alphay, alphaz
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 7:
        data['porf']   = float(p[0]); data['tortf']  = float(p[1])
        data['bulkden']= float(p[2]); data['foc']    = float(p[3])
        data['alphax'] = float(p[4]); data['alphay'] = float(p[5])
        data['alphaz'] = float(p[6])
        idx += 1

    # organic carbon partition: ock1-4
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 4:
        data['ock1'] = float(p[0]); data['ock2'] = float(p[1])
        data['ock3'] = float(p[2]); data['ock4'] = float(p[3])
        idx += 1

    # zone distances: x1, x2
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 2:
        data['x1'] = float(p[0]); data['x2'] = float(p[1])
        idx += 1

    # plume periods: tplume1, tplume2
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 2:
        data['tplume1'] = float(p[0]); data['tplume2'] = float(p[1])
        idx += 1

    # Freundlich exponent: fra1-4
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 4:
        data['fra1'] = float(p[0]); data['fra2'] = float(p[1])
        data['fra3'] = float(p[2]); data['fra4'] = float(p[3])
        idx += 1

    # Freundlich Kf: fcackf1-4
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 4:
        data['fcackf1'] = float(p[0]); data['fcackf2'] = float(p[1])
        data['fcackf3'] = float(p[2]); data['fcackf4'] = float(p[3])
        idx += 1

    # high-K yield: yieldf2, yieldf3, yieldf4
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 3:
        data['yieldf2'] = float(p[0]); data['yieldf3'] = float(p[1])
        data['yieldf4'] = float(p[2])
        idx += 1

    # high-K decay: decayf1-4
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 4:
        data['decayf1'] = float(p[0]); data['decayf2'] = float(p[1])
        data['decayf3'] = float(p[2]); data['decayf4'] = float(p[3])
        idx += 1

    # low-K transport: mdflag, porm, tortm, focm, diff, volfrac, difflen
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 7:
        data['mdflag']  = int(float(p[0])); data['porm']    = float(p[1])
        data['tortm']   = float(p[2]);      data['focm']    = float(p[3])
        data['diff']    = float(p[4]);      data['volfrac'] = float(p[5])
        data['difflen'] = float(p[6])
        idx += 1

    # low-K yield: yieldm2, yieldm3, yieldm4
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 3:
        data['yieldm2'] = float(p[0]); data['yieldm3'] = float(p[1])
        data['yieldm4'] = float(p[2])
        idx += 1

    # low-K decay: decaym1-4
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 4:
        data['decaym1'] = float(p[0]); data['decaym2'] = float(p[1])
        data['decaym3'] = float(p[2]); data['decaym4'] = float(p[3])
        idx += 1

    # x-grid: xmax, nx1, nxpsb, nx2
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 4:
        data['xmax']  = float(p[0]); data['nx1']   = int(float(p[1]))
        data['nxpsb'] = int(float(p[2])); data['nx2'] = int(float(p[3]))
        idx += 1

    # y-grid: dy, ny
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 2:
        data['dy'] = float(p[0]); data['ny'] = int(float(p[1]))
        idx += 1

    # z-grid: dz, nz
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 2:
        data['dz'] = float(p[0]); data['nz'] = int(float(p[1]))
        idx += 1

    # timesteps: dt, nt, npt
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 3:
        data['dt'] = float(p[0]); data['nt'] = int(float(p[1]))
        data['npt'] = int(float(p[2]))
        idx += 1

    # numerical: tol, maxit
    idx, p = _next_data(lines, idx)
    if idx is not None and len(p) >= 2:
        try:
            data['tol'] = float(p[0])
        except Exception:
            data['tol'] = p[0]
        data['maxit'] = int(float(p[1]))
        idx += 1

    # observation wells
    wells = []
    while idx is not None:
        idx, p = _next_data(lines, idx)
        if idx is None or len(p) < 5:
            break
        wells.append({'nwell': int(float(p[0])), 'xwell': float(p[1]),
                      'ywell': float(p[2]), 'zwelltop': float(p[3]),
                      'zwellbot': float(p[4])})
        idx += 1
    data['wells'] = wells

    return data


# ---------------------------------------------------------------------------
# store_info_additional_input.txt parser
# ---------------------------------------------------------------------------

def parse_additional_info(path: str) -> dict:
    """Parse store_info_additional_input.txt."""
    result = {}
    if not os.path.exists(path):
        return result

    try:
        with open(path) as f:
            lines = [ln.strip() for ln in f
                     if ln.strip() and not ln.strip().startswith('=')]

        in_names = False
        in_concs = False

        for line in lines:
            if "Monitoring Well Names" in line:
                in_names = True; in_concs = False
                result.setdefault('monitoring_well_names', [])
                continue
            if "Monitoring Well Concentrations" in line:
                in_concs = True; in_names = False
                result.setdefault('monitoring_well_concentrations', [])
                continue

            if ',' in line and ':' in line:
                key, _, val = line.partition(':')
                key = key.strip().rstrip(',')
                val = val.strip().lstrip(',')

                def _flt(v):
                    try: return float(v)
                    except: return v if v else None

                def _int(v):
                    try: return int(float(v))
                    except: return v if v else None

                if in_names and "Well" in key:
                    try:
                        idx = int(key.replace("Well", "").strip()) - 1
                        lst = result['monitoring_well_names']
                        while len(lst) <= idx:
                            lst.append("")
                        lst[idx] = val
                    except ValueError:
                        pass
                    continue

                if in_concs and "Well" in key:
                    try:
                        idx = int(key.replace("Well", "").strip()) - 1
                        lst = result['monitoring_well_concentrations']
                        while len(lst) <= idx:
                            lst.append(None)
                        lst[idx] = _flt(val)
                    except ValueError:
                        pass
                    continue

                if "Site Location and ID" in key:
                    result['site_location_id'] = val
                elif "Date" in key and "Treatment" not in key:
                    result['date'] = val
                elif "Thickness" in key:
                    result['thickness'] = _flt(val)
                elif "Source Width" in key:
                    # v102: explicit source-width override — exact, no
                    # lossy recovery from lysource*dy*2.
                    result['source_width'] = _flt(val)
                elif "Model X Size" in key:
                    result['model_x_size'] = _flt(val)
                elif "Model Y Size" in key:
                    result['model_y_size'] = _flt(val)
                elif "Model Z Size" in key:
                    result['model_z_size'] = _flt(val)
                elif "Bulk Darcy Velocity" in key:
                    # v102: explicit vd in user's unit at save time.
                    result['vd_user_unit'] = _flt(val)
                elif "Effective Porosity" in key:
                    result['porf'] = _flt(val)
                elif "Start Year" in key and "Treatment" not in key:
                    result['start_year'] = _flt(val)
                elif "End Year" in key and "Treatment" not in key:
                    result['end_year'] = _flt(val)
                elif "Source Treatment Start Year" in key:
                    result['source_treatment_start_year'] = _flt(val)
                elif "Source Treatment End Year" in key:
                    result['source_treatment_end_year'] = _flt(val)
                elif "Source Concentration Reduction" in key:
                    result['source_concentration_reduction'] = _flt(val)
                elif "Sample Year" in key:
                    result['sample_year'] = _flt(val)
                elif "Unit Flag" in key:
                    result['unit_flag'] = _int(val)
                elif "Dispersivity Flag" in key:
                    result['dispersivity_flag'] = _int(val)
                elif "PSB Loading" in key:
                    result['psb_loading'] = _flt(val)
    except Exception as e:
        print(f"Warning: could not parse additional info: {e}")

    return result


# ---------------------------------------------------------------------------
# retardation_inputs.txt PFAS name parser
# ---------------------------------------------------------------------------

def parse_retardation_pfas_names(path: str) -> dict:
    """Return dict of cell_addr -> pfas_name from retardation_inputs.txt."""
    result = {}
    if not os.path.exists(path):
        return result
    try:
        with open(pat