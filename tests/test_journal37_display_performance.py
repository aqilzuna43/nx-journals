import contextlib
import csv
import importlib.util
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
JOURNAL = ROOT / "from_git" / "journals" / "37_diagnose_display_performance.py"


def build_nxopen():
    nxopen = types.ModuleType("NXOpen")
    nxopen.NXObject = types.SimpleNamespace(
        AttributeType=types.SimpleNamespace(String="String")
    )

    class Vector3d:
        def __init__(self, x, y, z):
            self.X, self.Y, self.Z = x, y, z

    nxopen.Vector3d = Vector3d
    nxopen.Session = types.SimpleNamespace()
    return nxopen


def load_journal():
    """Return (journal_module, fake_nxopen_module)."""
    nxopen = build_nxopen()
    prior = sys.modules.get("NXOpen")
    sys.modules["NXOpen"] = nxopen
    try:
        spec = importlib.util.spec_from_file_location("journal37", JOURNAL)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module, nxopen
    finally:
        if prior is None:
            sys.modules.pop("NXOpen", None)
        else:
            sys.modules["NXOpen"] = prior


@contextlib.contextmanager
def env_overrides(**values):
    keys = (
        "NX_J37_MODE",
        "NX_J37_SCOPE",
        "NX_J37_VISIBLE_SCAN",
        "NX_J37_ROTATIONS",
        "NX_J37_MAX_BODIES",
        "NX_JOURNALS_IO_DIR",
    )
    saved = {key: os.environ.pop(key, None) for key in keys}
    try:
        for key, value in values.items():
            if value is not None:
                os.environ[key] = value
        yield
    finally:
        for key, value in saved.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value


# ---------------------------------------------------------------------------
# Fakes
# ---------------------------------------------------------------------------


class FakeListingWindow:
    def __init__(self):
        self.lines = []

    def Open(self):
        return None

    def WriteFullline(self, line):
        self.lines.append(str(line))


class FakeCollection:
    def __init__(self, items):
        self._items = list(items)

    def ToArray(self):
        return list(self._items)

    def __iter__(self):
        return iter(self._items)


class FakeBody:
    def __init__(
        self,
        name,
        faces=10,
        edges=20,
        facets=0,
        vertices=0,
        density=7.85e-6,
        layer=1,
        solid=True,
        sheet=False,
        convergent=False,
        blanked=False,
        face_error=None,
    ):
        self.Name = name
        self._faces = list(range(faces))
        self._edges = list(range(edges))
        self._facets = facets
        self._vertices = vertices
        self.Density = density
        self.Layer = layer
        self.IsSolidBody = solid
        self.IsSheetBody = sheet
        self.IsConvergentBody = convergent
        self.IsBlanked = blanked
        self._face_error = face_error

    def GetFaces(self):
        if self._face_error is not None:
            raise RuntimeError(self._face_error)
        return list(self._faces)

    def GetEdges(self):
        return list(self._edges)

    def GetNumberOfFacets(self):
        return self._facets

    def GetNumberOfVertices(self):
        return self._vertices


class FakeLayerManager:
    def __init__(self, states=None, default="Hidden"):
        self._states = dict(states or {})
        self._default = default
        self.state_calls = []

    def GetState(self, layer):
        self.state_calls.append(layer)
        return self._states.get(layer, self._default)


class FakePreferences:
    def __init__(self, values=None):
        for key, value in (values or {}).items():
            setattr(self, key, value)


class FakeView:
    def __init__(
        self,
        name="TFR-ISO",
        visible_objects=None,
        expose_rotate=True,
        expose_regenerate=True,
        rotate_error=None,
    ):
        self.Name = name
        self.RenderingStyle = "SHADED"
        self._visible = list(visible_objects or [])
        self.Matrix = ("MATRIX",)
        self.Origin = ("ORIGIN",)
        self.Scale = 1.0
        self.rotate_calls = []
        self.regenerate_calls = 0
        self.update_display_calls = 0
        self.fit_calls = 0
        self.restore_calls = []
        self._expose_rotate = expose_rotate
        self._expose_regenerate = expose_regenerate
        self._rotate_error = rotate_error

    def AskVisibleObjects(self):
        return list(self._visible)

    def Regenerate(self):
        self.regenerate_calls += 1

    def UpdateDisplay(self):
        self.update_display_calls += 1

    def Rotate(self, origin, vector, angle):
        if self._rotate_error is not None:
            raise RuntimeError(self._rotate_error)
        self.rotate_calls.append((tuple(origin), (vector.X, vector.Y, vector.Z), angle))

    def Fit(self):
        self.fit_calls += 1

    def SetRotationTranslationScale(self, matrix, origin, scale):
        self.restore_calls.append((matrix, origin, scale))


class FakeViewCollection:
    def __init__(self, views):
        self._views = list(views)

    def GetActiveViews(self):
        return list(self._views)

    def ToArray(self):
        return list(self._views)


class FakePart:
    def __init__(
        self,
        name,
        tag=None,
        attributes=None,
        fully_loaded=True,
        load_state="FULLY_LOADED",
        bodies=None,
        datums=None,
        curves=0,
        points=0,
        lines=0,
        csys=0,
        features=0,
        expressions=0,
        true_shading=0,
        true_studio=0,
        point_clouds=0,
        decals=0,
        cameras=0,
        dynamic_sections=0,
        images=0,
        sheets=0,
        layer_states=None,
        prefs=None,
        has_minimal_children=False,
        minimal_children=None,
        assembly_root=None,
    ):
        self.Name = name
        self.Tag = tag if tag is not None else name
        self.Leaf = name
        self.FullPath = name
        self._attributes = dict(attributes or {})
        self.IsFullyLoaded = fully_loaded
        self.PartLoadState = load_state
        self._bodies = list(bodies or [])
        self.Bodies = FakeCollection(self._bodies)
        self.Datums = FakeCollection(
            [types.SimpleNamespace(Layer=1) for _ in range(datums or 0)]
        )
        self.CoordinateSystems = FakeCollection(
            [types.SimpleNamespace(Layer=1) for _ in range(csys or 0)]
        )
        self.Curves = FakeCollection(
            [types.SimpleNamespace(Layer=1) for _ in range(curves or 0)]
        )
        self.Lines = FakeCollection(
            [types.SimpleNamespace(Layer=1) for _ in range(lines or 0)]
        )
        self.Points = FakeCollection(
            [types.SimpleNamespace(Layer=1) for _ in range(points or 0)]
        )
        self.Features = FakeCollection([object() for _ in range(features or 0)])
        self.Expressions = FakeCollection(
            [object() for _ in range(expressions or 0)]
        )
        self.SHEDObjs = FakeCollection(
            [object() for _ in range(true_shading or 0)]
        )
        self.TrueStudioObjs = FakeCollection(
            [object() for _ in range(true_studio or 0)]
        )
        self.PointClouds = FakeCollection(
            [object() for _ in range(point_clouds or 0)]
        )
        self.Decals = FakeCollection([object() for _ in range(decals or 0)])
        self.Cameras = FakeCollection([object() for _ in range(cameras or 0)])
        self.DynamicSections = FakeCollection(
            [object() for _ in range(dynamic_sections or 0)]
        )
        self.Images = FakeCollection([object() for _ in range(images or 0)])
        self.DrawingSheets = FakeCollection(
            [object() for _ in range(sheets or 0)]
        )
        self.Layers = FakeLayerManager(layer_states)
        self.Preferences = FakePreferences(prefs)
        self.Views = FakeViewCollection([])
        self._has_minimal_children = has_minimal_children
        self._minimal_children = list(minimal_children or [])
        self.ComponentAssembly = (
            types.SimpleNamespace(RootComponent=assembly_root)
            if assembly_root is not None
            else None
        )
        self.load_calls = 0

    def GetStringAttribute(self, name):
        if name in self._attributes:
            return self._attributes[name]
        raise RuntimeError("attribute not found: {0}".format(name))

    def GetUserAttribute(self, name, attribute_type, index):
        raise RuntimeError("attribute not found: {0}".format(name))

    def HasAnyMinimallyLoadedChildren(self):
        return self._has_minimal_children

    def GetMinimallyLoadedParts(self, container):
        container.extend(self._minimal_children)

    def LoadThisPartFully(self):  # must never be called by J37
        self.load_calls += 1
        raise AssertionError("J37 must not load parts")


class FakeComponent:
    def __init__(
        self,
        name,
        prototype=None,
        children=None,
        suppressed=False,
        blanked=False,
        layer=1,
        reference_set="MODEL",
        entire_part_refset="Entire Part",
        representation="Lightweight",
        arrangement=None,
        attributes=None,
        prototype_error=None,
    ):
        self.Name = name
        self.DisplayName = name
        self._prototype = prototype
        self._prototype_error = prototype_error
        self._children = list(children or [])
        self.IsSuppressed = suppressed
        self.IsBlanked = blanked
        self.Layer = layer
        self.ReferenceSet = reference_set
        self.EntirePartRefsetName = entire_part_refset
        self.RepresentationMode = representation
        self.UsedArrangement = arrangement
        self._attributes = dict(attributes or {})

    @property
    def Prototype(self):
        if self._prototype_error is not None:
            raise RuntimeError(self._prototype_error)
        return self._prototype

    @Prototype.setter
    def Prototype(self, value):
        self._prototype = value

    def GetChildren(self):
        return list(self._children)

    def GetStringAttribute(self, title):
        if title in self._attributes:
            return self._attributes[title]
        raise RuntimeError("component attribute not found: {0}".format(title))


class FakePartsCollection:
    def __init__(self, session):
        self._session = session
        self.Display = None
        self.Work = None

    def __iter__(self):
        return iter(self._session.loaded_parts)


class FakeSession:
    def __init__(
        self,
        work_part=None,
        loaded_parts=None,
        preferences=None,
        views=None,
    ):
        self.ListingWindow = FakeListingWindow()
        self.Parts = FakePartsCollection(self)
        self.Preferences = FakePreferences(preferences)
        self.loaded_parts = list(loaded_parts or [])
        self.IsManagedMode = False
        if work_part is not None:
            self.Parts.Work = work_part
            self.Parts.Display = work_part
            if work_part not in self.loaded_parts:
                self.loaded_parts.append(work_part)
        if views is not None and work_part is not None:
            work_part.Views = FakeViewCollection(views)


def make_assembly(**overrides):
    """root (L0) -> sub (L1) -> leaf (L2), plus a shared prototype."""
    root = FakePart(
        "ROOT_ASSY",
        tag="root",
        attributes={"DB_PART_NO": "264MN000001A01", "DB_PART_REV": "A"},
    )
    sub = FakePart(
        "SUB_ASSY",
        tag="sub",
        attributes={"DB_PART_NO": "264MN000002A01", "DB_PART_REV": "A"},
    )
    leaf = FakePart(
        "LEAF",
        tag="leaf",
        attributes={
            "DB_PART_NO": "264MN000003A01",
            "DB_PART_REV": "A",
            "Material": "AL-6061",
        },
    )
    shared = FakePart(
        "SHARED",
        tag="shared",
        attributes={"DB_PART_NO": "264MN000004A01", "DB_PART_REV": "A"},
    )
    sub_occurrence = FakeComponent(
        "SUB_ASSY",
        prototype=sub,
        children=[
            FakeComponent("LEAF", prototype=leaf),
            FakeComponent("SHARED", prototype=shared),
        ],
    )
    root_occurrence = FakeComponent(
        "ROOT_ASSY",
        prototype=root,
        children=[
            sub_occurrence,
            FakeComponent("SHARED", prototype=shared),
        ],
    )
    root.ComponentAssembly = types.SimpleNamespace(RootComponent=root_occurrence)
    for key, value in overrides.items():
        setattr(root, key, value)
    return {
        "root": root,
        "sub": sub,
        "leaf": leaf,
        "shared": shared,
        "root_occurrence": root_occurrence,
        "sub_occurrence": sub_occurrence,
    }


SESSION_PREFS = {
    "LoadComponentOnFacetedViewUpdate": False,
    "LoadComponentOnFacetedViewSelection": False,
    "SmartlightweightViewsLoadComponentOnDemand": False,
    "WorkPartDisplayAsEntirePart": False,
    "DisplayUpdateReport": False,
    "RenderSolidsUsingStoredFacets": True,
    "ShowFacetEdges": False,
}

PART_PREFS = {
    "PerformanceVisualization": FakePreferences({"SaveAdvancedDisplayFacets": False}),
    "ShadeVisualization": FakePreferences(
        {
            "RenderSolidsUsingStoredFacets": True,
            "ShowFacetEdges": False,
            "ShadingTolerance": "STANDARD",
            "CustomFaceTolerance": 0.0,
            "CustomEdgeTolerance": 0.0,
            "CustomAngleTolerance": 0.0,
        }
    ),
}


class Journal37DisplayPerformanceTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.journal, cls.nxopen = load_journal()

    def probe_geometry(self, part, **kwargs):
        return self.journal.census_geometry(part, **kwargs)

    # --- modes ---------------------------------------------------------

    def test_mode_defaults(self):
        with env_overrides():
            self.assertEqual("PROBE", self.journal.resolve_mode())
            self.assertEqual("ALL", self.journal.resolve_scope_filter())
            self.assertTrue(self.journal.resolve_visible_scan())
            self.assertEqual(2, self.journal.resolve_rotation_count())
            self.assertEqual(
                self.journal.MAX_BODIES_PER_PART,
                self.journal.resolve_max_bodies(),
            )

    def test_mode_env_overrides(self):
        with env_overrides(
            NX_J37_MODE="timed",
            NX_J37_SCOPE="bom",
            NX_J37_VISIBLE_SCAN="no",
            NX_J37_ROTATIONS="5",
            NX_J37_MAX_BODIES="50",
        ):
            self.assertEqual("TIMED", self.journal.resolve_mode())
            self.assertEqual("BOM", self.journal.resolve_scope_filter())
            self.assertFalse(self.journal.resolve_visible_scan())
            self.assertEqual(5, self.journal.resolve_rotation_count())
            self.assertEqual(50, self.journal.resolve_max_bodies())

    def test_invalid_modes_fail_closed(self):
        for key, value, resolver in (
            ("NX_J37_MODE", "FAST", "resolve_mode"),
            ("NX_J37_SCOPE", "EVERYTHING", "resolve_scope_filter"),
            ("NX_J37_ROTATIONS", "many", "resolve_rotation_count"),
            ("NX_J37_MAX_BODIES", "lots", "resolve_max_bodies"),
        ):
            with env_overrides(**{key: value}):
                with self.assertRaisesRegex(ValueError, key):
                    getattr(self.journal, resolver)()
        with env_overrides(NX_J37_ROTATIONS="-1"):
            with self.assertRaises(ValueError):
                self.journal.resolve_rotation_count()

    # --- availability-aware probes -------------------------------------

    def test_probe_value_ok_unavailable_and_error(self):
        journal = self.journal
        owner = types.SimpleNamespace(Present=5)

        self.assertEqual(("OK", 5, ""), journal.probe_value(owner, "Present"))
        status, value, error = journal.probe_value(owner, "Missing")
        self.assertEqual("UNAVAILABLE", status)
        self.assertIsNone(value)
        self.assertIn("not exposed", error)
        self.assertEqual(
            ("UNAVAILABLE", None, "owner is None"),
            journal.probe_value(None, "Anything"),
        )

        class Raising:
            @property
            def Broken(self):
                raise RuntimeError("boom")

        status, value, error = journal.probe_value(Raising(), "Broken")
        self.assertEqual("ERROR", status)
        self.assertIn("boom", error)

        status, _value, error = journal.probe_value(
            types.SimpleNamespace(Method=lambda: 1), "Method"
        )
        self.assertEqual("UNAVAILABLE", status)
        self.assertIn("method", error)

    def test_call_items_handles_methods_and_properties(self):
        journal = self.journal
        body = FakeBody("B1", faces=7)
        status, items, _error = journal.call_items(body, "GetFaces", "Faces")
        self.assertEqual("OK", status)
        self.assertEqual(7, len(items))

        part = types.SimpleNamespace(Bodies=FakeCollection([1, 2, 3]))
        status, items, _error = journal.call_items(part, "Bodies")
        self.assertEqual("OK", status)
        self.assertEqual(3, len(items))

        failing = types.SimpleNamespace(
            GetFaces=lambda: (_ for _ in ()).throw(RuntimeError("no faces"))
        )
        status, items, error = journal.call_items(failing, "GetFaces")
        self.assertEqual("ERROR", status)
        self.assertIn("no faces", error)
        self.assertEqual([], items)

    def test_call_int_handles_methods_and_properties(self):
        journal = self.journal
        body = FakeBody("B1", facets=123, vertices=45)
        self.assertEqual(
            ("OK", 123, ""), journal.call_int(body, "GetNumberOfFacets")
        )
        self.assertEqual(
            ("OK", 45, ""), journal.call_int(body, "GetNumberOfVertices")
        )
        status, value, _error = journal.call_int(body, "NoSuchThing", "Facets")
        self.assertEqual("UNAVAILABLE", status)

    def test_enum_text(self):
        journal = self.journal
        self.assertEqual("Exact", journal.enum_text(types.SimpleNamespace(name="Exact")))
        self.assertEqual("Lightweight", journal.enum_text("Lightweight"))
        self.assertEqual("", journal.enum_text(None))

    # --- traversal ------------------------------------------------------

    def test_traversal_levels_dedup_and_suppression(self):
        tree = make_assembly()
        suppressed = FakeComponent("SUPPRESSED", prototype=None, suppressed=True)
        tree["root_occurrence"]._children.append(suppressed)

        occurrences, targets, diagnostics, total = self.journal.collect_occurrences(
            tree["root"], "ALL"
        )
        by_name = {t["part"].Name: t for t in targets}
        self.assertEqual(
            {"ROOT_ASSY", "SUB_ASSY", "LEAF", "SHARED"}, set(by_name)
        )
        self.assertEqual(0, by_name["ROOT_ASSY"]["level"])
        self.assertEqual(1, by_name["SUB_ASSY"]["level"])
        self.assertEqual(2, by_name["LEAF"]["level"])
        self.assertEqual(1, by_name["SHARED"]["level"])
        self.assertEqual(2, by_name["SHARED"]["deepest_level"])
        self.assertEqual(2, by_name["SHARED"]["occurrence_count"])
        self.assertEqual(5, total)
        self.assertEqual([], diagnostics)

        suppressed_row = [
            row for row in occurrences if row["IS_SUPPRESSED"] == "YES"
        ]
        self.assertEqual(1, len(suppressed_row))
        self.assertEqual("NO", suppressed_row[0]["COUNTS_FOR_DISPLAY"])
        self.assertEqual("", suppressed_row[0]["PROTOTYPE_NAME"])

    def test_traversal_bom_scope_excludes_keyword_and_reference(self):
        tree = make_assembly()
        tree["root_occurrence"]._children.extend(
            [
                FakeComponent("CSYS_1", prototype=tree["leaf"]),
                FakeComponent(
                    "REF",
                    prototype=tree["leaf"],
                    attributes={"REFERENCE_COMPONENT": "YES"},
                ),
            ]
        )
        occurrences, targets, _diagnostics, _total = (
            self.journal.collect_occurrences(tree["root"], "BOM")
        )
        by_name = {t["part"].Name: t for t in targets}
        self.assertEqual(1, by_name["LEAF"]["occurrence_count"])
        bom_rows = [row for row in occurrences if row["BOM_VISIBLE"] == "NO"]
        self.assertEqual(2, len(bom_rows))
        for row in bom_rows:
            self.assertEqual("NO", row["COUNTS_FOR_DISPLAY"])

    def test_traversal_records_entire_part_reference_set(self):
        tree = make_assembly()
        tree["root_occurrence"]._children[0].ReferenceSet = "Entire Part"
        occurrences, _targets, _diagnostics, _total = (
            self.journal.collect_occurrences(tree["root"], "ALL")
        )
        row = [
            item for item in occurrences if item["COMPONENT_NAME"] == "SUB_ASSY"
        ][0]
        self.assertEqual("YES", row["IS_ENTIRE_PART_REFSET"])
        self.assertEqual("Entire Part", row["REFERENCE_SET"])

    def test_traversal_records_unreadable_prototype(self):
        root = FakePart("ROOT", tag="root", attributes={"DB_PART_NO": "N1"})
        root.ComponentAssembly = types.SimpleNamespace(
            RootComponent=FakeComponent(
                "ROOT",
                prototype=root,
                children=[
                    FakeComponent(
                        "BAD", prototype_error="invalid om object [im0541]"
                    )
                ],
            )
        )
        _occurrences, _targets, diagnostics, _total = (
            self.journal.collect_occurrences(root, "ALL")
        )
        codes = [item["code"] for item in diagnostics]
        self.assertIn("PROTOTYPE_UNAVAILABLE", codes)

    # --- geometry census -----------------------------------------------

    def test_geometry_census_counts_and_density(self):
        part = FakePart(
            "PART",
            tag="p",
            bodies=[
                FakeBody("B1", faces=10, edges=20, facets=0, density=7.85e-6),
                FakeBody(
                    "B2",
                    faces=5,
                    edges=8,
                    facets=100,
                    vertices=50,
                    density=0.0,
                    sheet=True,
                    solid=False,
                    layer=12,
                ),
                FakeBody(
                    "B3", faces=1, edges=1, density=2.7e-6, convergent=True,
                    solid=False,
                ),
            ],
        )
        geometry = self.probe_geometry(part)
        self.assertEqual("OK", geometry["status"])
        self.assertEqual(3, geometry["body_count"])
        self.assertEqual(1, geometry["solid_body_count"])
        self.assertEqual(1, geometry["sheet_body_count"])
        self.assertEqual(1, geometry["convergent_body_count"])
        self.assertEqual(16, geometry["face_count"])
        self.assertEqual(29, geometry["edge_count"])
        self.assertEqual(100, geometry["facet_count"])
        self.assertEqual(50, geometry["vertex_count"])
        self.assertEqual(1, geometry["density_zero_count"])
        self.assertEqual(0.0, geometry["density_min"])
        self.assertEqual(7.85e-6, geometry["density_max"])
        self.assertEqual("NO", geometry["truncated"])

    def test_geometry_census_marks_partial_load_unavailable(self):
        part = FakePart(
            "PART",
            tag="p",
            fully_loaded=False,
            load_state="PartiallyLoaded",
            bodies=[FakeBody("B1")],
        )
        geometry = self.probe_geometry(part)
        self.assertEqual("UNAVAILABLE", geometry["status"])
        self.assertIn("not fully loaded", geometry["error"])
        self.assertEqual("", geometry["body_count"])

    def test_geometry_census_truncates_body_enumeration(self):
        part = FakePart(
            "PART",
            tag="p",
            bodies=[FakeBody("B{0}".format(index), faces=1) for index in range(10)],
        )
        geometry = self.probe_geometry(part, max_bodies=4)
        self.assertEqual(10, geometry["body_count"])
        self.assertEqual(4, geometry["bodies_enumerated"])
        self.assertEqual("YES", geometry["truncated"])
        self.assertEqual(4, geometry["face_count"])

    def test_geometry_census_survives_one_broken_body(self):
        part = FakePart(
            "PART",
            tag="p",
            bodies=[
                FakeBody("BAD", face_error="no geometry"),
                FakeBody("GOOD", faces=3),
            ],
        )
        geometry = self.probe_geometry(part)
        self.assertEqual("OK", geometry["status"])
        self.assertEqual(2, geometry["body_count"])
        self.assertEqual(3, geometry["face_count"])

    def test_geometry_census_counts_blanked_and_heavy_objects(self):
        part = FakePart(
            "PART",
            tag="p",
            bodies=[FakeBody("B1", blanked=True)],
            datums=12,
            curves=3,
            points=4,
            true_shading=2,
            point_clouds=1,
            sheets=2,
            features=25,
            expressions=40,
        )
        geometry = self.probe_geometry(part)
        self.assertEqual(1, geometry["blanked_body_count"])
        self.assertEqual(12, geometry["datum_count"])
        self.assertEqual(3, geometry["curve_count"])
        self.assertEqual(4, geometry["point_count"])
        self.assertEqual(2, geometry["true_shading_count"])
        self.assertEqual(1, geometry["point_cloud_count"])
        self.assertEqual(2, geometry["drawing_sheet_count"])
        self.assertEqual(25, geometry["feature_count"])
        self.assertEqual(40, geometry["expression_count"])

    # --- layers ---------------------------------------------------------

    def test_layer_census_reports_states_and_object_counts(self):
        part = FakePart(
            "PART",
            tag="p",
            bodies=[FakeBody("B1", layer=1), FakeBody("B2", layer=12)],
            datums=2,
            layer_states={1: "Visible", 12: "Hidden", 200: "Visible"},
        )
        rows, status, error, visible_layers = self.journal.census_layers(part)
        self.assertEqual("OK", status)
        self.assertEqual("", error)
        self.assertEqual(2, visible_layers)
        by_layer = {row["LAYER"]: row for row in rows}
        self.assertEqual(1, by_layer[1]["BODY_COUNT"])
        self.assertEqual("NO", by_layer[1]["IS_HIDDEN"])
        self.assertEqual(1, by_layer[12]["BODY_COUNT"])
        self.assertEqual("YES", by_layer[12]["IS_HIDDEN"])
        self.assertEqual(2, by_layer[1]["DATUM_COUNT"])
        self.assertEqual(3, by_layer[1]["TOTAL_COUNT"])

    def test_layer_census_without_layer_manager(self):
        class NoLayers:
            pass

        rows, status, _error, visible_layers = self.journal.census_layers(
            NoLayers()
        )
        self.assertEqual([], rows)
        self.assertEqual("UNAVAILABLE", status)
        self.assertEqual(0, visible_layers)

    # --- visible objects -------------------------------------------------

    def test_visible_object_census(self):
        objects = [
            FakeBody("B1"),
            FakeBody("B2"),
        ] + [FakeComponent("C{0}".format(index)) for index in range(3)]
        view = FakeView(visible_objects=objects)
        status, rows, total, seconds, truncated, error = (
            self.journal.census_visible_objects(view)
        )
        self.assertEqual("OK", status)
        self.assertEqual(5, total)
        self.assertEqual("NO", truncated)
        self.assertEqual("", error)
        self.assertGreaterEqual(float(seconds), 0.0)
        by_type = {row["TYPE_NAME"]: row for row in rows}
        self.assertEqual(2, by_type["FakeBody"]["TYPE_COUNT"])
        self.assertEqual(3, by_type["FakeComponent"]["TYPE_COUNT"])
        self.assertEqual(40.0, by_type["FakeBody"]["TYPE_PERCENT"])

    def test_visible_object_census_truncates(self):
        view = FakeView(visible_objects=[FakeBody("B{0}".format(i)) for i in range(10)])
        status, _rows, total, _seconds, truncated, _error = (
            self.journal.census_visible_objects(view, max_objects=3)
        )
        self.assertEqual("OK", status)
        self.assertEqual(3, total)
        self.assertEqual("YES", truncated)

    def test_visible_object_census_reports_missing_api(self):
        view = types.SimpleNamespace(Name="X")
        status, rows, total, _seconds, _truncated, error = (
            self.journal.census_visible_objects(view)
        )
        self.assertEqual("UNAVAILABLE", status)
        self.assertEqual([], rows)
        self.assertEqual("", total)
        self.assertIn("AskVisibleObjects", error)

    def test_active_view_uses_active_views(self):
        view = FakeView(name="TFR-ISO")
        tree = make_assembly()
        session = FakeSession(work_part=tree["root"], views=[view])
        status, found, error = self.journal.active_view(session)
        self.assertEqual("OK", status)
        self.assertIs(view, found)
        self.assertEqual("", error)

    # --- preferences -----------------------------------------------------

    def test_preference_census_reports_status_per_probe(self):
        tree = make_assembly()
        session = FakeSession(
            work_part=tree["root"], preferences=SESSION_PREFS
        )
        tree["root"].Preferences = FakePreferences(PART_PREFS)
        rows = self.journal.census_preferences(session, [tree["root"]])
        by_key = {}
        for row in rows:
            by_key.setdefault(row["KEY"], []).append(row)
        self.assertTrue(by_key["LoadComponentOnFacetedViewUpdate"])
        session_rows = [
            row
            for row in by_key["LoadComponentOnFacetedViewUpdate"]
            if row["OWNER"] == "session"
        ]
        self.assertTrue(session_rows)
        self.assertTrue(
            all(row["STATUS"] == "OK" for row in session_rows)
        )
        # part-level probes resolve through part.Preferences
        ok_part_rows = [
            row
            for row in by_key["SaveAdvancedDisplayFacets"]
            if row["SCOPE"] == "PART" and row["STATUS"] == "OK"
        ]
        self.assertEqual(1, len(ok_part_rows))
        self.assertEqual("False", ok_part_rows[0]["VALUE"])

    def test_preference_census_marks_missing_containers_unavailable(self):
        tree = make_assembly()
        session = FakeSession(work_part=tree["root"], preferences={})
        rows = self.journal.census_preferences(session, [tree["root"]])
        self.assertTrue(rows)
        for row in rows:
            if row["OWNER"].startswith("session."):
                self.assertEqual("UNAVAILABLE", row["STATUS"])

    def test_pref_helpers(self):
        journal = self.journal
        rows = [
            {
                "SCOPE": "SESSION",
                "OWNER": "session",
                "KEY": "WorkPartDisplayAsEntirePart",
                "STATUS": "OK",
                "VALUE": "True",
                "ERROR": "",
            },
            {
                "SCOPE": "SESSION",
                "OWNER": "session",
                "KEY": "ShowFacetEdges",
                "STATUS": "UNAVAILABLE",
                "VALUE": "",
                "ERROR": "nope",
            },
        ]
        self.assertTrue(
            journal.pref_is_true(
                journal.pref_lookup(rows, "SESSION", "WorkPartDisplay")
            )
        )
        self.assertIsNone(
            journal.pref_lookup(rows, "SESSION", "ShowFacetEdges")
        )
        self.assertFalse(
            journal.pref_is_true(
                journal.pref_lookup(rows, "SESSION", "NotThere")
            )
        )

    # --- suspects --------------------------------------------------------

    def build_suspect_case(
        self,
        bodies=None,
        leaf_attributes=None,
        entire_part=False,
        representation="Exact",
        fully_loaded=True,
        load_state="FULLY_LOADED",
        minimal_children=None,
        visible_total="",
        prefs=None,
        layer_states=None,
        max_bodies=3000,
    ):
        journal = self.journal
        tree = make_assembly()
        leaf = tree["leaf"]
        leaf.Bodies = FakeCollection(list(bodies or []))
        if leaf_attributes:
            leaf._attributes.update(leaf_attributes)
        leaf.IsFullyLoaded = fully_loaded
        leaf.PartLoadState = load_state
        leaf._minimal_children = list(minimal_children or [])
        leaf._has_minimal_children = bool(minimal_children)
        leaf.Layers = FakeLayerManager(layer_states)
        leaf.Preferences = FakePreferences(PART_PREFS)

        for occurrence in tree["sub_occurrence"]._children:
            if occurrence.Name == "LEAF":
                occurrence.RepresentationMode = representation
                if entire_part:
                    occurrence.ReferenceSet = "Entire Part"
                else:
                    occurrence.ReferenceSet = "MODEL"

        session = FakeSession(
            work_part=tree["root"], preferences=prefs or SESSION_PREFS
        )
        ledger = journal.EvidenceLedger("TEST", "TS", "PROBE", "ALL")
        occurrences, targets, diagnostics, _total = journal.collect_occurrences(
            tree["root"], "ALL"
        )
        geometry_by_key = {}
        ledger.add("PartLoadState", "leaf", "OK", journal.load_state_text(leaf))
        ledger.add("HasAnyMinimallyLoadedChildren", "leaf", "OK", 0)
        ledger.add("Bodies", "leaf", "OK", "")
        ledger.add("GetFaces", "leaf", "OK", "")
        ledger.add("GetNumberOfFacets", "leaf", "OK", "")
        ledger.add("Density", "leaf", "OK", "")
        ledger.add("IsConvergentBody", "leaf", "OK", "")
        ledger.add("IsBlanked", "leaf", "OK", "")
        ledger.add("GetNumberOfVertices", "leaf", "OK", "")
        ledger.add("PartLoadState", "root", "OK", "FULLY_LOADED")
        ledger.add("ReferenceSet", "all", "OK", "")
        ledger.add("EntirePartRefsetName", "all", "OK", "")
        ledger.add("RepresentationMode", "all", "OK", "")
        ledger.add("IsSuppressed", "all", "OK", "")
        ledger.add("Layers", "leaf", "OK", "")
        ledger.add("GetState", "leaf", "OK", "")
        ledger.add("Datums", "leaf", "OK", "")
        ledger.add("Curves", "leaf", "OK", "")
        ledger.add("Lines", "leaf", "OK", "")
        ledger.add("Points", "leaf", "OK", "")
        ledger.add("SHEDObjs", "leaf", "OK", "")
        ledger.add("TrueStudioObjs", "leaf", "OK", "")
        ledger.add("PointClouds", "leaf", "OK", "")
        ledger.add("Decals", "leaf", "OK", "")
        ledger.add("AskVisibleObjects", "view", "OK", visible_total)
        ledger.add("Component.GetChildren", "root", "OK", "")
        for target in targets:
            geometry = journal.census_geometry(target["part"], max_bodies=max_bodies)
            geometry_by_key[target["key"]] = geometry

        layer_info_by_key = {}
        for target in targets:
            rows, _status, _error, visible_layers = journal.census_layers(
                target["part"]
            )
            populated = sum(1 for row in rows if int(row["TOTAL_COUNT"]) > 0)
            hidden_populated = sum(
                1
                for row in rows
                if int(row["TOTAL_COUNT"]) > 0 and row["IS_HIDDEN"] == "YES"
            )
            layer_info_by_key[target["key"]] = {
                "visible_layers": visible_layers,
                "populated_layers": populated,
                "hidden_populated_layers": hidden_populated,
            }

        minimal_by_key = {}
        for target in targets:
            minimal_by_key[target["key"]] = journal.minimally_loaded_children(
                target["part"]
            )

        pref_rows = journal.census_preferences(session, [t["part"] for t in targets])
        for row in pref_rows:
            ledger.add(
                "PREF_{0}".format(row["KEY"]),
                row["OWNER"],
                row["STATUS"],
                row["VALUE"],
                row["ERROR"],
            )

        suspects = journal.build_suspects(
            ledger,
            targets,
            occurrences,
            geometry_by_key,
            layer_info_by_key,
            minimal_by_key,
            visible_total,
            pref_rows,
            diagnostics,
        )
        return suspects, ledger

    def codes(self, suspects):
        return [suspect["CODE"] for suspect in suspects]

    def test_suspect_entire_part_reference_set(self):
        suspects, _ledger = self.build_suspect_case(entire_part=True)
        self.assertIn("ENTIRE_PART_REFSET", self.codes(suspects))

    def test_suspect_exact_representation(self):
        suspects, _ledger = self.build_suspect_case(representation="Exact")
        self.assertIn("EXACT_REPRESENTATION", self.codes(suspects))
        self.assertNotIn("NON_LIGHTWEIGHT_REPRESENTATION", self.codes(suspects))

    def test_suspect_not_fully_loaded(self):
        suspects, _ledger = self.build_suspect_case(
            fully_loaded=False, load_state="PartiallyLoaded"
        )
        self.assertIn("NOT_FULLY_LOADED", self.codes(suspects))

    def test_suspect_minimally_loaded_children(self):
        other = FakePart("OTHER", tag="other", attributes={"DB_PART_NO": "X1"})
        suspects, _ledger = self.build_suspect_case(minimal_children=[other])
        self.assertIn("MINIMALLY_LOADED_CHILDREN", self.codes(suspects))

    def test_suspect_zero_density(self):
        suspects, _ledger = self.build_suspect_case(
            bodies=[FakeBody("B1", faces=10, density=0.0)]
        )
        self.assertIn("DENSITY_ZERO_OR_MISSING", self.codes(suspects))

    def test_suspect_high_face_geometry(self):
        suspects, _ledger = self.build_suspect_case(
            bodies=[FakeBody("BIG", faces=9000, density=7.85e-6)]
        )
        self.assertIn("HIGH_FACE_GEOMETRY", self.codes(suspects))
        high = [s for s in suspects if s["CODE"] == "HIGH_FACE_GEOMETRY"][0]
        self.assertEqual(9000, high["MEASURED_VALUE"])

    def test_suspect_heavy_display_objects(self):
        tree_bodies = [FakeBody("B1", faces=10)]
        suspects, _ledger = self.build_suspect_case(bodies=tree_bodies)
        self.assertNotIn("HEAVY_DISPLAY_OBJECTS", self.codes(suspects))

        journal = self.journal
        tree = make_assembly()
        tree["leaf"].SHEDObjs = FakeCollection([object(), object()])
        suspects = self.journal.collect_occurrences(tree["root"], "ALL")
        # direct geometry probe path
        geometry = journal.census_geometry(tree["leaf"])
        self.assertEqual(2, geometry["true_shading_count"])

    def test_suspect_visible_object_count_and_layers(self):
        suspects, _ledger = self.build_suspect_case(
            visible_total=str(self.journal.THRESHOLD_VISIBLE_OBJECTS + 1),
            layer_states={layer: "Visible" for layer in range(1, 40)},
        )
        codes = self.codes(suspects)
        self.assertIn("VISIBLE_OBJECT_COUNT_HIGH", codes)
        self.assertIn("MANY_VISIBLE_LAYERS", codes)
        hidden = [
            s
            for s in suspects
            if s["IDENTITY"].startswith("264MN000003A01")
        ]
        self.assertTrue(hidden)

    def test_suspect_preference_faceted_view_update(self):
        prefs = dict(SESSION_PREFS)
        prefs["LoadComponentOnFacetedViewUpdate"] = True
        suspects, _ledger = self.build_suspect_case(prefs=prefs)
        codes = self.codes(suspects)
        self.assertIn("PREF_LOADCOMPONENTONFACETEDVIEWUPDATE_TRUE", codes)
        suspect = [
            s
            for s in suspects
            if s["CODE"] == "PREF_LOADCOMPONENTONFACETEDVIEWUPDATE_TRUE"
        ][0]
        self.assertEqual("HIGH", suspect["SEVERITY"])
        self.assertEqual("True", suspect["MEASURED_VALUE"])

    def test_every_suspect_cites_recorded_evidence(self):
        prefs = dict(SESSION_PREFS)
        prefs["LoadComponentOnFacetedViewUpdate"] = True
        prefs["WorkPartDisplayAsEntirePart"] = True
        other = FakePart("OTHER", tag="other", attributes={"DB_PART_NO": "X1"})
        suspects, ledger = self.build_suspect_case(
            entire_part=True,
            bodies=[FakeBody("BIG", faces=9000, density=0.0)],
            minimal_children=[other],
            visible_total=str(self.journal.THRESHOLD_VISIBLE_OBJECTS + 1),
            prefs=prefs,
        )
        self.assertTrue(suspects)
        for suspect in suspects:
            self.assertTrue(
                suspect["EVIDENCE_IDS"].strip(),
                "{0} has no evidence ids".format(suspect["CODE"]),
            )
            for fact_id in suspect["EVIDENCE_IDS"].split():
                self.assertTrue(
                    any(fact["id"] == fact_id for fact in ledger.facts),
                    "{0} cites unknown fact {1}".format(
                        suspect["CODE"], fact_id
                    ),
                )

    def test_suspects_are_ranked_by_severity(self):
        prefs = dict(SESSION_PREFS)
        prefs["LoadComponentOnFacetedViewUpdate"] = True
        suspects, _ledger = self.build_suspect_case(
            entire_part=True,
            bodies=[FakeBody("BIG", faces=9000, density=7.85e-6)],
            prefs=prefs,
        )
        severities = [s["SEVERITY"] for s in suspects]
        order = {"HIGH": 0, "MEDIUM": 1, "INFO": 2}
        self.assertEqual(
            sorted(severities, key=lambda s: order[s]), severities
        )

    # --- timing ----------------------------------------------------------

    def test_run_timing_rotates_then_restores(self):
        view = FakeView(name="TFR-ISO")
        rows, total, notes = self.journal.run_timing(
            FakeSession(work_part=make_assembly()["root"]), view, 2
        )
        self.assertEqual([], notes)
        self.assertEqual(1, view.regenerate_calls)
        self.assertEqual(1, view.update_display_calls)
        self.assertEqual(1, view.fit_calls)
        self.assertEqual(4, len(view.rotate_calls))
        angles = [call[2] for call in view.rotate_calls]
        self.assertEqual(
            [
                self.journal.ROTATION_ANGLE_DEGREES,
                self.journal.ROTATION_ANGLE_DEGREES,
                -self.journal.ROTATION_ANGLE_DEGREES,
                -self.journal.ROTATION_ANGLE_DEGREES,
            ],
            angles,
        )
        self.assertEqual(1, len(view.restore_calls))
        self.assertEqual(("MATRIX",), view.restore_calls[0][0])
        self.assertEqual(("ORIGIN",), view.restore_calls[0][1])
        self.assertEqual(1.0, view.restore_calls[0][2])
        self.assertGreater(float(total), 0.0)
        steps = [row["STEP"] for row in rows]
        self.assertEqual(sorted(steps), steps)
        for row in rows:
            self.assertNotIn("ERROR", row["NOTE"])

    def test_run_timing_records_unavailable_operations(self):
        view = types.SimpleNamespace(Name="X")
        rows, _total, notes = self.journal.run_timing(
            FakeSession(work_part=make_assembly()["root"]), view, 1
        )
        self.assertEqual([], notes)
        notes_text = " ".join(row["NOTE"] for row in rows)
        self.assertIn("UNAVAILABLE", notes_text)
        self.assertIn("not restored", notes_text)

    def test_run_timing_records_rotation_error(self):
        view = FakeView(rotate_error="rotation failed")
        rows, _total, _notes = self.journal.run_timing(
            FakeSession(work_part=make_assembly()["root"]), view, 1
        )
        error_rows = [row for row in rows if "ERROR" in row["NOTE"]]
        self.assertTrue(error_rows)

    # --- reports ---------------------------------------------------------

    def test_write_csv_ignores_extra_keys_and_writes_bom(self):
        journal = self.journal
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, "out.csv")
            journal.write_csv(
                path,
                ("A", "B"),
                [{"A": 1, "B": 2, "EXTRA": 3}],
            )
            with open(path, "r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle)
                self.assertEqual(["A", "B"], reader.fieldnames)
                rows = list(reader)
            self.assertEqual([{"A": "1", "B": "2"}], rows)

    def test_format_number(self):
        journal = self.journal
        self.assertEqual("7", journal.format_number(7.0))
        self.assertEqual("7.85e-06", journal.format_number(7.85e-6))
        self.assertEqual("", journal.format_number(""))
        self.assertEqual("abc", journal.format_number("abc"))

    # --- end to end ------------------------------------------------------

    def build_session(self, mode="PROBE", visible_objects=None, **journal_kwargs):
        tree = make_assembly()
        tree["leaf"].Bodies = FakeCollection(
            [FakeBody("B1", faces=12000, density=0.0)]
        )
        tree["leaf"].Layers = FakeLayerManager({1: "Visible"})
        tree["root"].Layers = FakeLayerManager({1: "Visible"})
        tree["root"].Preferences = FakePreferences(PART_PREFS)
        view = FakeView(visible_objects=visible_objects or [])
        session = FakeSession(
            work_part=tree["root"],
            loaded_parts=[tree["sub"], tree["leaf"], tree["shared"]],
            preferences=SESSION_PREFS,
            views=[view],
        )
        self.nxopen.Session.GetSession = staticmethod(lambda: session)
        return tree, view, session

    def run_main(self, folder, **env):
        values = {"NX_JOURNALS_IO_DIR": folder, "NX_J37_MODE": "PROBE"}
        values.update(env)
        with env_overrides(**values):
            self.journal.main()

    def read_run(self, folder):
        run_root = os.path.join(folder, self.journal.OUTPUT_ROOT_FOLDER)
        runs = os.listdir(run_root)
        self.assertEqual(1, len(runs))
        run = os.path.join(run_root, runs[0])
        reports = os.path.join(run, "REPORTS")
        logs = os.path.join(run, "LOGS")
        return run, reports, logs

    def test_main_probe_writes_reports_and_is_read_only(self):
        tree, view, _session = self.build_session()
        try:
            with tempfile.TemporaryDirectory() as folder:
                self.run_main(folder)
                run, reports, logs = self.read_run(folder)

                names = sorted(os.listdir(reports))
                for expected in (
                    "J37_TARGETS_",
                    "J37_OCCURRENCES_",
                    "J37_LAYERS_",
                    "J37_VISIBLE_",
                    "J37_PREFS_",
                    "J37_SUSPECTS_",
                    "J37_TIMING_",
                    "J37_EVIDENCE_",
                ):
                    self.assertTrue(
                        any(name.startswith(expected) for name in names),
                        "{0} missing from {1}".format(expected, names),
                    )

                report_path = [
                    os.path.join(reports, name)
                    for name in names
                    if name.startswith("J37_TARGETS_")
                ][0]
                with open(
                    report_path, "r", encoding="utf-8-sig", newline=""
                ) as handle:
                    rows = list(csv.DictReader(handle))
                self.assertEqual(
                    {"ROOT_ASSY", "SUB_ASSY", "LEAF", "SHARED"},
                    {row["PART_NAME"] for row in rows},
                )
                leaf_row = [row for row in rows if row["PART_NAME"] == "LEAF"][0]
                self.assertEqual("12000", leaf_row["FACE_COUNT"])
                self.assertEqual("1", leaf_row["DENSITY_ZERO_COUNT"])
                self.assertEqual("AL-6061", leaf_row["MATERIAL"])
                self.assertEqual("FULLY_LOADED", leaf_row["LOAD_STATE"])
                self.assertEqual("1", leaf_row["VISIBLE_LAYER_COUNT"])

                suspect_path = [
                    os.path.join(reports, name)
                    for name in names
                    if name.startswith("J37_SUSPECTS_")
                ][0]
                with open(
                    suspect_path, "r", encoding="utf-8-sig", newline=""
                ) as handle:
                    suspect_rows = list(csv.DictReader(handle))
                codes = {row["CODE"] for row in suspect_rows}
                self.assertIn("HIGH_FACE_GEOMETRY", codes)
                self.assertIn("DENSITY_ZERO_OR_MISSING", codes)
                for index, row in enumerate(suspect_rows, start=1):
                    self.assertEqual(str(index), row["RANK"])

                evidence_path = [
                    os.path.join(reports, name)
                    for name in names
                    if name.startswith("J37_EVIDENCE_")
                ][0]
                with open(evidence_path, "r", encoding="utf-8") as handle:
                    payload = json.load(handle)
                self.assertEqual(
                    self.journal.JOURNAL_BUILD_ID, payload["build"]
                )
                self.assertEqual("PROBE", payload["mode"])
                self.assertGreater(payload["fact_count"], 0)
                self.assertEqual(4, payload["prototype_count"])
                self.assertEqual("", payload["timing_total_seconds"])

                log_name = os.listdir(logs)[0]
                with open(
                    os.path.join(logs, log_name), "r", encoding="utf-8"
                ) as handle:
                    log_text = handle.read()
                self.assertIn(
                    "J37-NX2506-DISPLAY-PERF-TRIAGE-V1", log_text
                )
                self.assertIn("Read-only", log_text)

                occurrence_path = [
                    os.path.join(reports, name)
                    for name in names
                    if name.startswith("J37_OCCURRENCES_")
                ][0]
                with open(
                    occurrence_path, "r", encoding="utf-8-sig", newline=""
                ) as handle:
                    occurrence_rows = list(csv.DictReader(handle))
                self.assertEqual(4, len(occurrence_rows))
                reference_sets = {
                    row["REFERENCE_SET"] for row in occurrence_rows
                }
                self.assertEqual({"MODEL"}, reference_sets)
                self.assertEqual(
                    {"YES"},
                    {row["COUNTS_FOR_DISPLAY"] for row in occurrence_rows},
                )

                # read-only: no loads, no view motion, no timing file rows
                for part in (tree["root"], tree["sub"], tree["leaf"], tree["shared"]):
                    self.assertEqual(0, part.load_calls)
                self.assertEqual([], view.rotate_calls)
                self.assertEqual(0, view.fit_calls)
                self.assertEqual(0, view.regenerate_calls)
        finally:
            self.nxopen.Session.__dict__.pop("GetSession", None)

    def test_main_timed_measures_and_restores_view(self):
        visible = [FakeBody("B{0}".format(i)) for i in range(25)]
        tree, view, _session = self.build_session(
            mode="TIMED", visible_objects=visible
        )
        try:
            with tempfile.TemporaryDirectory() as folder:
                self.run_main(folder, NX_J37_MODE="TIMED", NX_J37_ROTATIONS="1")
                _run, reports, _logs = self.read_run(folder)
                timing_path = [
                    os.path.join(reports, name)
                    for name in os.listdir(reports)
                    if name.startswith("J37_TIMING_")
                ][0]
                with open(
                    timing_path, "r", encoding="utf-8-sig", newline=""
                ) as handle:
                    rows = list(csv.DictReader(handle))
                operations = {row["OPERATION"] for row in rows}
                self.assertIn("View.Rotate (+15.0 deg)", operations)
                self.assertIn("View.Rotate (-15.0 deg)", operations)
                self.assertIn("Restore view state", operations)
                self.assertEqual(2, len(view.rotate_calls))
                self.assertEqual(1, len(view.restore_calls))
                self.assertEqual(1, view.regenerate_calls)

                visible_path = [
                    os.path.join(reports, name)
                    for name in os.listdir(reports)
                    if name.startswith("J37_VISIBLE_")
                ][0]
                with open(
                    visible_path, "r", encoding="utf-8-sig", newline=""
                ) as handle:
                    visible_rows = list(csv.DictReader(handle))
                self.assertTrue(visible_rows)
                self.assertEqual(
                    "25", visible_rows[0]["TOTAL_VISIBLE_OBJECTS"]
                )
                self.assertEqual("SHADED", visible_rows[0]["VIEW_RENDERING_STYLE"])
        finally:
            self.nxopen.Session.__dict__.pop("GetSession", None)

    def test_main_reports_missing_work_part(self):
        session = FakeSession()
        self.nxopen.Session.GetSession = staticmethod(lambda: session)
        try:
            with tempfile.TemporaryDirectory() as folder:
                self.run_main(folder)
                _run, _reports, logs = self.read_run(folder)
                log_name = os.listdir(logs)[0]
                with open(
                    os.path.join(logs, log_name), "r", encoding="utf-8"
                ) as handle:
                    log_text = handle.read()
                self.assertIn("No work part is loaded", log_text)
        finally:
            self.nxopen.Session.__dict__.pop("GetSession", None)

    def test_main_visible_scan_can_be_disabled(self):
        tree, _view, _session = self.build_session(
            visible_objects=[FakeBody("B1")]
        )
        try:
            with tempfile.TemporaryDirectory() as folder:
                self.run_main(folder, NX_J37_VISIBLE_SCAN="NO")
                _run, reports, _logs = self.read_run(folder)
                visible_path = [
                    os.path.join(reports, name)
                    for name in os.listdir(reports)
                    if name.startswith("J37_VISIBLE_")
                ][0]
                with open(
                    visible_path, "r", encoding="utf-8-sig", newline=""
                ) as handle:
                    self.assertEqual([], list(csv.DictReader(handle)))
        finally:
            self.nxopen.Session.__dict__.pop("GetSession", None)


if __name__ == "__main__":
    unittest.main()
