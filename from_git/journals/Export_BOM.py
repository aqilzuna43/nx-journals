import NXOpen
import csv
import os
import datetime

# --- CONFIGURATION ---
# List the exact names of the attributes you want to extract here.
# (e.g., "Material", "Weight", "DB_PART_NO", "DB_PART_NAME")
ATTRIBUTES_TO_EXTRACT = [
    "DB_PART_DESC", 
    "DB_PART_NAME", 
    "DB_PART_REV"
]
# ---------------------

def get_safe_attribute(nx_object, attr_name):
    """Helper to try and read an attribute, returns None if not found."""
    try:
        return nx_object.GetStringAttribute(attr_name)
    except:
        return None

def walk_assembly_tree(component, level, csv_writer):
    # Create a visual indent for the CSV file based on the assembly level
    indent = "    " * level
    
    # Extract metadata safely
    part_name = component.DisplayName
    component_name = component.Name
    
    # Extract custom attributes
    custom_attr_values = []
    for attr in ATTRIBUTES_TO_EXTRACT:
        # Try getting attribute from the component instance first
        val = get_safe_attribute(component, attr)
        
        # If not on the component, try getting it from the actual part file (Prototype)
        if val is None and component.Prototype is not None:
            val = get_safe_attribute(component.Prototype, attr)
            
        custom_attr_values.append(val if val is not None else "")
    
    # Write the row to the CSV file
    row_data = [level, f"{indent}{part_name}", component_name] + custom_attr_values
    csv_writer.writerow(row_data)
    
    # Get children and run recursively
    try:
        children = component.GetChildren()
        
        for child in children:
            # Skip suppressed components so they don't appear in the BOM
            if child.IsSuppressed:
                continue
            
            walk_assembly_tree(child, level + 1, csv_writer)
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
        csv_filename = f"NX_MultiLevel_BOM_{timestamp}.csv"
        full_csv_path = os.path.join(desktop_path, csv_filename)

        print(f"Starting BOM extraction for: {work_part.Leaf}")
        
        # Open the CSV file to write the data (added utf-8 encoding for special characters)
        with open(full_csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            
            # Write the header row
            header_row = ['BOM Level', 'Indented Part Name', 'Component Name'] + ATTRIBUTES_TO_EXTRACT
            writer.writerow(header_row)
            
            # Start walking the tree at Level 0
            walk_assembly_tree(root_component, 0, writer)
            
        # Notify success via the system console
        print(f"SUCCESS: BOM successfully exported to: {full_csv_path}")

    except Exception as e:
        print(f"ERROR: Failed to run script. Details: {str(e)}")

if __name__ == '__main__':
    main()