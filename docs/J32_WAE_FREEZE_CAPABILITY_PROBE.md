# J32 WAE Freeze Capability Probe

`from_git/journals/32_probe_wae_freeze_capability.py` is a strictly read-only
NX X 2506 diagnostic. It determines which runtime APIs might support a real
Teamcenter-enforced WAE freeze without assuming that check-in is a lock.

## Target rule

- With no NX preselection, J32 inspects only the active work part.
- With exactly one Assembly Navigator component selected, J32 inspects only
  that component's loaded prototype.
- More than one selection, a non-component selection, a suppressed component,
  or an unmanaged/unloaded target is blocked.

## Safety boundary

J32 reads runtime metadata and the known read-only checkout-status query. It
does not invoke any discovered candidate API. It never checks out, checks in,
saves, changes an attribute, applies a release status, starts a workflow, or
creates a Teamcenter revision.

The JSON includes public runtime member names and filtered candidates matching
lock, release, status, workflow, access, checkout/check-in, lifecycle, and
related terms. V2 reads each candidate member without invoking it and records
its `__doc__`, `__text_signature__`, annotations, overload representation,
`repr`, and `inspect.signature` result. Where the Python binding permits it,
read-only .NET reflection also records candidate parameter and return types.

## NX acceptance run

1. Deploy the complete `from_git` directory to the NX X 2506 machine.
2. Open one disposable Teamcenter-managed CAD part, with no navigator selection.
3. Run J32 and require `PROBE_COMPLETE`.
4. Repeat with exactly one loaded Assembly Navigator component selected.
5. Confirm the target checkout state and `WAE_VERSION` are unchanged.
6. Commit the resulting `WAE_FREEZE_CAPABILITY_<timestamp>.json` under
   `from_git/templates/LOGS/` for analysis.

The report inventories capability only. It does not prove that a candidate API
is authorized by the Teamcenter X tenant or suitable for WAE lifecycle policy.
