from _16_tc_specification_import_v4_core import *

def new_file_management(session):
    pdm = getattr(session, "PdmSession", None)
    if pdm is None:
        raise RuntimeError(
            "NXOpen.Session.PdmSession is unavailable. Run J16 in Teamcenter managed mode."
        )
    method = getattr(pdm, "NewFileManagement", None)
    if method is None:
        raise RuntimeError("PdmSession.NewFileManagement is unavailable.")
    return pdm, method()


def invoke_export_files(fm, proposal, relation_type, export_root):
    if os.path.isdir(export_root):
        shutil.rmtree(export_root)
    os.makedirs(export_root)
    args = (
        [proposal["part_number"]], [proposal["revision"]], [proposal["dataset_name"]],
        [proposal["dataset_type"]], [relation_type], [export_root], [proposal["export_tool"]],
    )
    method = getattr(fm, "ExportFiles", None)
    if method is None:
        raise RuntimeError("PDM FileManagement.ExportFiles is unavailable.")
    output_directories = []
    try:
        raw = method(*args)
    except TypeError:
        raw = method(*(args + (output_directories,)))
    codes, returned_paths = parse_codes_and_strings(raw)
    returned_paths.extend(string_list(output_directories))
    return (codes[0] if codes else None), returned_paths


def invoke_import_files(fm, proposal):
    method = getattr(fm, "ImportFiles", None)
    if method is None:
        raise RuntimeError("PDM FileManagement.ImportFiles is unavailable.")
    raw = method(
        [proposal["part_number"]], [proposal["revision"]], [proposal["dataset_name"]],
        [proposal["dataset_type"]], [proposal["relation_type"]], [proposal["stage_dir"]],
    )
    codes, _ = parse_codes_and_strings(raw)
    return codes[0] if codes else None


def all_prt_files(root, returned_paths):
    candidates = []
    for path in returned_paths:
        absolute = os.path.abspath(path)
        if os.path.isfile(absolute) and absolute.lower().endswith(".prt"):
            candidates.append(absolute)
        elif os.path.isdir(absolute):
            for folder, _, files in os.walk(absolute):
                candidates.extend(
                    os.path.join(folder, name) for name in files if name.lower().endswith(".prt")
                )
    if os.path.isdir(root):
        for folder, _, files in os.walk(root):
            candidates.extend(
                os.path.join(folder, name) for name in files if name.lower().endswith(".prt")
            )
    unique = {os.path.normcase(os.path.abspath(path)): os.path.abspath(path) for path in candidates}
    return sorted(unique.values(), key=lambda value: value.lower())


def select_exported_drawing(root, returned_paths, proposal):
    files = all_prt_files(root, returned_paths)
    matches = [
        path for path in files
        if valid_native(path, proposal["part_number"], proposal["revision"], proposal["drawing_index"])
    ]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise RuntimeError("Multiple exported .prt files matched the target: {0}".format(
            " | ".join(matches)
        ))
    if len(files) == 1:
        return files[0]
    if not files:
        raise RuntimeError("Export returned no native .prt file for the target dataset.")
    raise RuntimeError("Export returned multiple unmatched .prt files: {0}".format(
        " | ".join(files)
    ))


def export_exact_dataset(fm, proposal, relation_type, root, pdi_field, file_field):
    code, returned_paths = invoke_export_files(fm, proposal, relation_type, root)
    proposal["report"][pdi_field] = "" if code is None else str(code)
    if code != 0:
        raise RuntimeError("PDM ExportFiles failed for relation '{0}' with PDI code {1}.".format(
            relation_type, "<missing>" if code is None else code
        ))
    exported = select_exported_drawing(root, returned_paths, proposal)
    proposal["report"][file_field] = exported
    return exported, sha256(exported)


def resolve_relation_and_baseline(fm, proposal, root):
    attempts = []
    for index, relation in enumerate(configured_relation_candidates(), 1):
        candidate_root = os.path.join(
            root, "RELATION_{0}_{1}".format(index, re.sub(r"[^A-Za-z0-9_.-]", "_", relation))
        )
        try:
            exported, digest = export_exact_dataset(
                fm, proposal, relation, candidate_root,
                "BASELINE_EXPORT_PDI_CODE", "BASELINE_EXPORT_FILE",
            )
            return relation, exported, digest
        except Exception as exc:
            attempts.append("{0}: {1}".format(relation, error_text(exc)))
    raise RuntimeError(
        "Could not export the existing UGPART specification dataset. Attempts: {0}".format(
            " || ".join(attempts)
        )
    )


def object_tag(value):
    try:
        return int(value.Tag)
    except Exception:
        try:
            return value.Tag
        except Exception:
            return None


def same_nx_object(left, right):
    if left is right:
        return True
    left_tag, right_tag = object_tag(left), object_tag(right)
    return left_tag is not None and right_tag is not None and left_tag == right_tag


def checkedout_arrays(pdm):
    method = getattr(pdm, "GetCheckedoutStatusOfAllObjectsInSession", None)
    if method is None:
        raise RuntimeError("PdmSession.GetCheckedoutStatusOfAllObjectsInSession is unavailable.")
    checked_output = []
    unchecked_output = []
    try:
        raw = method()
    except TypeError:
        raw = method(checked_output, unchecked_output)
    if isinstance(raw, (tuple, list)) and len(raw) >= 2:
        return list(raw[0] or []), list(raw[1] or [])
    if checked_output or unchecked_output:
        return list(checked_output), list(unchecked_output)
    raise RuntimeError("Unexpected checkout-status return: {0}".format(type(raw).__name__))


def close_opened_part(part, log):
    if part is None:
        return
    try:
        whole_tree = resolve_member(
            NXOpen.BasePart.CloseWholeTree,
            ("False_", "False", "CloseWholeTreeFalse"),
            "BasePart.CloseWholeTree false value",
        )
        close_modified = resolve_member(
            NXOpen.BasePart.CloseModified,
            ("UseResponses", "UseLatest", "CloseModified"),
            "BasePart.CloseModified safe value",
        )
        part.Close(whole_tree, close_modified, None)
    except Exception as exc:
        log.write("  WARNING: could not close checkout-probe part: {0}".format(error_text(exc)))


def check_target_checkout(session, pdm, identifier, log):
    try:
        existing = session.Parts.FindObject(identifier)
    except Exception:
        existing = None
    part = existing
    load_status = None
    opened_here = False
    try:
        if part is None:
            part, load_status = session.Parts.OpenBase(identifier)
            opened_here = True
        dispose(load_status)
        load_status = None
        checked, unchecked = checkedout_arrays(pdm)
        if any(same_nx_object(part, item) for item in checked):
            return "CHECKED_OUT"
        if any(same_nx_object(part, item) for item in unchecked):
            return "CLEAR"
        return "UNKNOWN"
    finally:
        dispose(load_status)
        if opened_here:
            close_opened_part(part, log)

