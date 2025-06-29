#!/usr/bin/env python3
"""
MS Project to Excel Converter
Converts .mpp files to .xlsx format with task hierarchy and resources
"""
import jpype
import jpype.imports
import pandas as pd
import sys
import os
from pathlib import Path
import argparse


def setup_jvm():
    """Setup JVM with MPXJ library"""
    if not jpype.isJVMStarted():
        import mpxj
        mpxj_path = Path(mpxj.__file__).parent / "lib"
        jar_files = list(mpxj_path.glob("*.jar"))
        
        if not jar_files:
            raise RuntimeError("No JAR files found in mpxj package")
        
        jpype.startJVM(classpath=[str(jar) for jar in jar_files])


def format_predecessors(predecessors):
    """Format predecessor relationships in a readable way"""
    if not predecessors:
        return ''
    
    formatted = []
    for relation in predecessors:
        # Get predecessor task
        pred_task = relation.getPredecessorTask()
        task_id = pred_task.getID() if pred_task else 'Unknown'
        
        # Get relationship type (FS, SS, SF, FF)
        rel_type = str(relation.getType()) if relation.getType() else 'FS'
        
        # Get lag
        lag = relation.getLag()
        lag_str = ''
        if lag and lag.getDuration() != 0:
            lag_value = lag.getDuration()
            lag_units = str(lag.getUnits()) if lag.getUnits() else 'd'
            
            # Format based on sign
            if lag_value > 0:
                lag_str = f"+{int(lag_value)}{lag_units}"
            else:
                lag_str = f"{int(lag_value)}{lag_units}"
        
        # Create formatted string (e.g., "3FS", "5SS+2d")
        formatted.append(f"{task_id}{rel_type}{lag_str}")
    
    return '; '.join(formatted)


def read_ms_project(file_path):
    """Read MS Project file and extract task information"""
    setup_jvm()
    
    # Import Java classes
    org = jpype.JPackage("org")
    UniversalProjectReader = org.mpxj.reader.UniversalProjectReader
    
    # Read the project
    reader = UniversalProjectReader()
    project = reader.read(file_path)
    
    tasks_data = []
    
    # Process tasks
    for task in project.getTasks():
        if task is None:
            continue
            
        task_data = {
            'ID': task.getID() if task.getID() else 0,
            'WBS': str(task.getWBS()) if task.getWBS() else '',
            'Name': str(task.getName()) if task.getName() else '',
            'Duration': str(task.getDuration()) if task.getDuration() else '',
            'Start': str(task.getStart()) if task.getStart() else '',
            'Finish': str(task.getFinish()) if task.getFinish() else '',
            'Percent Complete': (task.getPercentageComplete().doubleValue()
                                 if task.getPercentageComplete() else 0),
            'Predecessors': format_predecessors(task.getPredecessors()),
            'Resource Names': (str(task.getResourceNames())
                               if task.getResourceNames() else ''),
            'Cost': (task.getCost().doubleValue()
                     if task.getCost() else 0),
            'Work': str(task.getWork()) if task.getWork() else '',
            'Critical': bool(task.getCritical()) if task.getCritical() else False,
            'Milestone': bool(task.getMilestone()) if task.getMilestone() else False,
            'Summary': bool(task.getSummary()) if task.getSummary() else False,
            'Notes': str(task.getNotes()) if task.getNotes() else '',
            'Outline Level': (task.getOutlineLevel().intValue()
                              if task.getOutlineLevel() else 0)
        }
        
        tasks_data.append(task_data)
    
    return tasks_data, project


def export_to_xlsx(tasks_df, project, output_file):
    """Export project data to XLSX file with formatting"""
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write tasks
        tasks_df.to_excel(writer, sheet_name='Tasks', index=False)
        
        # Auto-adjust column widths
        worksheet = writer.sheets['Tasks']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    # Add resources
    resources_data = []
    for resource in project.getResources():
        if resource and resource.getName():
            resources_data.append({
                'ID': resource.getID() if resource.getID() else 0,
                'Name': str(resource.getName()),
                'Type': str(resource.getType()) if resource.getType() else '',
                'Cost': (resource.getCost().doubleValue()
                         if resource.getCost() else 0),
                'Standard Rate': (str(resource.getStandardRate())
                                  if resource.getStandardRate() else '')
            })
    
    if resources_data:
        resources_df = pd.DataFrame(resources_data)
        with pd.ExcelWriter(output_file, mode='a',
                            engine='openpyxl', if_sheet_exists='new') as writer:
            resources_df.to_excel(writer, sheet_name='Resources', index=False)


def main():
    parser = argparse.ArgumentParser(
        description='Convert MS Project (.mpp) files to Excel (.xlsx) format'
    )
    parser.add_argument('input_file', help='Path to the .mpp file')
    parser.add_argument('-o', '--output', help='Output .xlsx file path',
                        default=None)
    parser.add_argument('-v', '--verbose', action='store_true',
                        help='Show detailed project information')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input_file):
        print(f"Error: File not found - {args.input_file}")
        sys.exit(1)
    
    # Generate output filename if not provided
    if args.output is None:
        output_file = os.path.splitext(args.input_file)[0] + '.xlsx'
    else:
        output_file = args.output
    
    print(f"Converting: {os.path.basename(args.input_file)}")
    
    try:
        # Read project
        tasks_data, project = read_ms_project(args.input_file)
        
        if not tasks_data:
            print("No tasks found in project file!")
            sys.exit(1)
        
        # Convert to DataFrame
        tasks_df = pd.DataFrame(tasks_data)
        
        # Show summary if verbose
        if args.verbose:
            print(f"\nTotal Tasks: {len(tasks_df)}")
            print(f"Milestones: {tasks_df['Milestone'].sum()}")
            print(f"Critical Tasks: {tasks_df['Critical'].sum()}")
            
        # Export to Excel
        export_to_xlsx(tasks_df, project, output_file)
        print(f"✓ Exported to: {output_file}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)
    finally:
        if jpype.isJVMStarted():
            jpype.shutdownJVM()


if __name__ == "__main__":
    main()