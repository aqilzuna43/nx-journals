import NXOpen
import csv
import os
import datetime

# --- CONFIGURATION ---
# The first ten fields follow docs/FZ-PowerSystem_v1_22Jun.csv. The remaining
# fields are the extended columns requested in commit 25105de.
# 2026-08-13: the nine NX-excluded columns (Commodity Code, Country of Origin,
# Export Control Number, Traceability/Serial Numbered Part, Hazardous, Shelf
# Life, Temperature Sensitive, Serviceable Item, Commodity Type) are no longer
# exported — ITEMS is manually maintained and J4 enrichment owns them.
FZ_COLUMNS = [
    "Level",
    "Item Number",
    "Part Description",
    "Item Rev",
    "Lifecycle",
    "Qty",
    "UOM",
    "Mfr. Name",
    "Mfr. Part Number",
    "Reference Notes",
    "WAE_VERSION",
    "NX_MATERIAL",
    "NX_FINISH",
    "NX_MASS",
    "NX_MassPropRollupMass",
    "NX_MassPropRollupArea_m2",
    "COMPONENT_CLASS",
    "Dimensions",
]

# CSV column -> exact internal NX/Teamcenter title -> NX attribute type.
FZ_ATTRIBUTE_SPECS = [
    ("Part Description", "DB_PART_NAME", "String"),
    ("Item Rev", "DB_PART_REV", "String"),
    ("Lifecycle", "ItemRev_REL_STATUS", "String"),
    ("UOM", "Unit_Of_Measure", "String"),
    ("Mfr. Name", "MFG", "String"),
    ("Mfr. Part Number", "MPN", "String"),
    ("Reference Notes", "Stocking_Type", "String"),
    ("WAE_VERSION", "WAE_VERSION", "String"),
    ("NX_MATERIAL", "NX_MATERIAL", "String"),
    ("NX_FINISH", "NX_FINISH", "String"),
    ("NX_MASS", "NX_Mass", "Number"),
    ("NX_MassPropRollupMass", "NX_MassPropRollupMass", "Number"),
    ("NX_MassPropRollupArea", "NX_MassPropRollupArea", "Number"),
    ("COMPONENT_CLASS", "COMPONENT_CLASS", "String"),
    ("Dimensions", "Dimensions", "String"),
]

# The attribute used as the primary identifier (Source of Truth).
SOURCE_OF_TRUTH_ATTR = "DB_PART_NO"
DEFAULT_LIFECYCLE = "DRAFT"

# List of keywords in part names to automatically exclude from the BOM
# (e.g., coordinate systems, datums, skeletons). Case-insensitive.
IGNORE_KEYWORDS = ["CSYS", "COORDINATE", "DATUM", "REFERENCE", "SKELETON"]
# ---------------------

def get_safe_attribute(nx_object, attr_name, attr_type="String"):
    """Read a typed NX attribute, returning None when it is unavailable."""
    if attr_type not in ("String", "Number"):
        raise ValueError("Unsupported NX attribute type: {0}".format(attr_type))

    try:
        if attr_type == "String":
            return nx_object.GetStringAttribute(attr_name)
        return nx_object.GetRealAttribute(attr_name)
    except Exception:
        return None


def get_component_attribute(component, attr_name, attr_type="String"):
    """Read only the 3D master prototype represented by a BOM component."""
    prototype = getattr(component, "Prototype", None)
    target = prototype if prototype is not None else component
    return get_safe_attribute(target, attr_name, attr_type)


def fz_attribute_values(component):
    """Project exact NX attributes into the FZ template column names."""
    values = {}
    for column, attr_name, attr_type in FZ_ATTRIBUTE_SPECS:
        values[column] = get_component_attribute(component, attr_name, attr_type)
    return values


def walk_assembly_tree(component, level, csv_writer, quantity=1):
    # Extract metadata safely
    part_name = component.DisplayName

    # Extract Source of Truth (DB_PART_NO)
    db_part_no = get_component_attribute(component, SOURCE_OF_TRUTH_ATTR)
    # Fallback to DisplayName if the attribute is missing/blank
    if not db_part_no:
        db_part_no = component.DisplayName

    values = fz_attribute_values(component)
    row = {
        "Level": level,
        "Item Number": db_part_no,
        "Qty": quantity,
    }
    for column, _attr_name, _attr_type in FZ_ATTRIBUTE_SPECS:
        value = values[column]
        row[column] = "" if value is None else value

    if not row["Part Description"]:
        row["Part Description"] = part_name
    if not row["Lifecycle"]:
        row["Lifecycle"] = DEFAULT_LIFECYCLE

    # NX stores NX_MassPropRollupArea in square millimetres; present it in
    # square metres so large-system values stay readable.
    raw_area = row.get("NX_MassPropRollupArea")
    if isinstance(raw_area, (int, float)):
        row["NX_MassPropRollupArea_m2"] = round(
            raw_area / 1000000.0, 4
        )
    else:
        row["NX_MassPropRollupArea_m2"] = ""

    csv_writer.writerow([row[column] for column in FZ_COLUMNS])
    
    # Get children and run recursively
    try:
        children = component.GetChildren()
        
        # SMART QUANTITY LOGIC: Group children by their true NX ID (Source of Truth)
        grouped_children = {}
        
        for child in children:
            # Skip suppressed components so they don't appear in the BOM
            if child.IsSuppressed:
                continue
            
            # 1. Skip if the part name contains any of the ignore keywords
            child_name = (child.Name or "").upper()
            child_display_name = (child.DisplayName or "").upper()
            
            should_ignore = False
            for keyword in IGNORE_KEYWORDS:
                if keyword in child_name or keyword in child_display_name:
                    should_ignore = True
                    break
                    
            # 2. Skip if it's marked as a "Reference-Only" component in NX Properties
            is_ref = get_safe_attribute(child, "REFERENCE_COMPONENT")
            is_plist_ignore = get_safe_attribute(child, "PLIST_IGNORE_MEMBER")
            
            # NX natively uses an empty string ("") to mark these, but we keep "YES" just in case of manual overrides
            if is_ref in ["", "YES", "1", "True", "true", "yes"] or is_plist_ignore in ["", "YES", "1", "True", "true", "yes"]:
                should_ignore = True
                
            if should_ignore:
                continue
            
            # Get the child's Source of Truth attribute for accurate grouping
            child_db_part_no = get_component_attribute(
                child, SOURCE_OF_TRUTH_ATTR
            )
            
            # The Source of Truth for grouping (fallback to DisplayName if missing)
            nx_id = child_db_part_no if child_db_part_no else child.DisplayName
            
            if nx_id not in grouped_children:
                # First time seeing this part in this subassembly
                grouped_children[nx_id] = {
                    'instance': child,
                    'count': 1
                }
            else:
                # We already saw this part, just increase the quantity count
                grouped_children[nx_id]['count'] += 1
                
        # Now recurse into each UNIQUE group we found
        for nx_id, data in grouped_children.items():
            walk_assembly_tree(data['instance'], level + 1, csv_writer, quantity=data['count'])
            
    except Exception as e:
        print(f"Warning: Could not get children for {part_name}. Error: {e}")

def main():
    try:
        session = NXOpen.Session.GetSession()
        work_part = session.Parts.Work
        
        # Check if a part is actually open
        if work_part is None:
            print("ERROR: No part is currently open in NX.")
            return
            
        # Check if the open part is an assembly
        root_component = work_part.ComponentAssembly.RootComponent
        if root_component is None:
            print("ERROR: The active part is not an assembly.")
            return

        # Automatically set the output path to the user's Desktop
        desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        
        # Create a unique filename with a timestamp to avoid overwriting
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        csv_filename = f"NX_MultiLevel_BOM_QTY_{timestamp}.csv"
        full_csv_path = os.path.join(desktop_path, csv_filename)

        print(f"Starting BOM extraction for: {work_part.Leaf}")
        
        # Open the CSV file to write the data (added utf-8 encoding for special characters)
        with open(full_csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            
            writer.writerow(FZ_COLUMNS)
            
            # Start walking the tree at Level 0, Quantity is 1 for the very top assembly
            walk_assembly_tree(root_component, 0, writer, quantity=1)
            
        # Notify success via the system console
        print(f"SUCCESS: BOM successfully exported to: {full_csv_path}")

    except Exception as e:
        print(f"ERROR: Failed to run script. Details: {str(e)}")

if __name__ == '__main__':
    main()
