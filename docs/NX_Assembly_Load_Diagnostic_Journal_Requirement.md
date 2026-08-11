# NX Assembly Load Diagnostic Journal Requirement

## 1. Objective

Develop an NXOpen journal that diagnoses the root cause of:

> "IM0541: An operation was attempted on an invalid or unsuitable OM object"

during STEP export in Siemens NX.

The journal shall identify the exact assembly component, occurrence, or referenced object that prevents successful STEP translation.

Primary target environment:
- NX X 2506
- Teamcenter X managed assemblies
- Local NX assemblies

---

# 2. Problem Statement

During STEP export, NX may fail with:

```
An operation was attempted on an invalid or unsuitable OM object
```

The translator log usually reports symptoms such as:

- Component left unloaded
- Failed to find file using current search options
- Prototype unavailable

The diagnostic tool shall identify:
- Exact problematic component
- Assembly hierarchy location
- Failure reason
- Recommended corrective action

---

# 3. Functional Requirements

## FR-001: Detect Current Assembly

The journal shall detect the current NX work part.

If no assembly exists:
- Report that the part is not an assembly
- Exit gracefully

---

## FR-002: Recursive Component Scan

Traverse the complete assembly structure:

```
Top Assembly
 |
 +-- Sub Assembly
 |     |
 |     +-- Component
 |
 +-- Component
```

All levels shall be checked.

---

## FR-003: Component Health Check

For every component, check:

- Component name
- Part number
- File path
- Prototype availability
- Load status
- Reference set status

Supported results:

```
OK
MISSING_FILE
PROTOTYPE_UNAVAILABLE
UNLOADED
INVALID_OBJECT
ERROR
```

---

## FR-004: Missing Component Detection

Detect components where:

- Assembly reference exists
- Physical part file cannot be located

Example output:

```
Component:
017679_A.prt

Status:
MISSING_FILE

Reason:
File not found using current search options
```

---

## FR-005: Invalid OM Object Detection

Capture NXOpen exceptions related to:

```
Invalid or unsuitable OM object
```

Report:

- Component name
- Parent assembly
- Failed operation
- Exception message

---

## FR-006: Teamcenter Information (If Available)

For managed mode assemblies, retrieve:

- Item ID
- Revision
- Dataset name
- Status

Example:

```
Component:
264MN033047A01_A.prt

Item:
264MN033047A01

Revision:
A

Status:
PROTOTYPE_UNAVAILABLE
```

---

## FR-007: Diagnostic Report Generation

Generate:

```
NX_Assembly_Load_Diagnostic_Report.txt
```

The report shall contain:

- Assembly name
- Date/time
- Total components scanned
- Failed components
- Failure reason

Example:

```
================================================
NX Assembly Load Diagnostic Report
================================================

Assembly:
264MN021888A01_A.prt

Component:
017679_A.prt

Level:
3

Status:
MISSING_FILE

Reason:
File not found

------------------------------------------------

Component:
264MN033047A01_A.prt

Status:
INVALID_OBJECT

Exception:
An operation was attempted on an invalid
or unsuitable OM object
================================================
```

---

# 4. User Experience Requirements

The journal shall display progress in NX Listing Window:

Example:

```
NX Assembly Diagnostic Started...

Scanning assembly:
264MN021888A01_A.prt

Components found:
154

Errors found:
3

Report generated successfully.
```

The journal shall continue scanning even if individual components fail.

---

# 5. Coding Requirements

Use:

- NXOpen VB.NET
- NX 2506 compatible API
- Modular structure
- Clear comments
- No hard-coded customer paths

The code shall be suitable for:
- Mechanical engineers
- NX administrators
- Teamcenter administrators

---

# 6. Acceptance Criteria

## Test Case 1: Missing File

Given:
- Assembly contains missing component

Expected:
```
Status:
MISSING_FILE
```

The exact component name must be reported.

---

## Test Case 2: Unloaded Prototype

Given:
- Component exists but prototype cannot load

Expected:

```
Status:
PROTOTYPE_UNAVAILABLE
```

---

## Test Case 3: STEP Export Failure

Given:
- STEP export fails due to OM object issue

Expected:

Report identifies:

- Exact component
- Assembly path
- Exception message

---

# 7. Deliverables

Repository:

```
nx-journal
```

Folder:

```
/Assembly/Diagnostic/
```

Files:

```
NX_Assembly_Load_Diagnostic.vb
README.md
Example_Report.txt
```

---

# 8. Future Enhancements

Allow future extension for:

- STEP export pre-check
- Automatic missing component search
- Teamcenter checkout validation
- Duplicate component detection
- WAVE/interpart link checking
- Lightweight representation detection
