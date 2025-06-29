#!/usr/bin/env python3
import jpype
import jpype.imports
import pandas as pd
import sys
import os
from pathlib import Path


def setup_jvm():
    """Setup JVM with MPXJ"""
    if not jpype.isJVMStarted():
        # Use the mpxj module's JAR files
        import mpxj
        mpxj_path = Path(mpxj.__file__).parent / "lib"
        
        # Find all JAR files in the lib directory
        jar_files = list(mpxj_path.glob("*.jar"))
        
        if not jar_files:
            raise RuntimeError("No JAR files found in mpxj package lib directory")
        
        print(f"Found {len(jar_files)} JAR files in mpxj package")
        
        # Start JVM with all JAR files
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
    
    # Import Java classes using JPackage
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
            'WBS': str(task.getWBS()) if task.getWBS() else '',
            'Outline Level': (task.getOutlineLevel().intValue()
                              if task.getOutlineLevel() else 0)
        }
        
        tasks_data.append(task_data)
    
    return tasks_data, project


def visualize_project_summary(tasks_df):
    """Print project summary information"""
    print("\n=== PROJECT SUMMARY ===")
    print(f"Total Tasks: {len(tasks_df)}")
    print(f"Milestones: {tasks_df['Milestone'].sum()}")
    print(f"Summary Tasks: {tasks_df['Summary'].sum()}")
    print(f"Critical Tasks: {tasks_df['Critical'].sum()}")
    print(f"Average Completion: {tasks_df['Percent Complete'].mean():.1f}%")
    
    print("\n=== TASK HIERARCHY ===")
    # Show tasks with outline levels for hierarchy
    hierarchy_view = tasks_df[['WBS', 'Name', 'Duration', 'Start', 'Finish',
                                'Percent Complete', 'Outline Level']].copy()
    
    # Add indentation based on outline level
    hierarchy_view['Name'] = hierarchy_view.apply(
        lambda x: '  ' * x['Outline Level'] + x['Name'], axis=1
    )
    
    print(hierarchy_view[['WBS', 'Name', 'Duration',
                          'Percent Complete']].head(30).to_string())


def export_to_xlsx(tasks_df, project, output_file):
    """Export project data to XLSX file with formatting"""
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write tasks to Excel
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
        
        # Style header row
        for row in worksheet.iter_rows(min_row=1, max_row=1):
            for cell in row:
                cell.style = 'Headline 3'
    
    # Add resources sheet
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
                                  if resource.getStandardRate() else ''),
                'Max Units': (resource.getMaxUnits().doubleValue()
                              if resource.getMaxUnits() else 100)
            })
    
    if resources_data:
        resources_df = pd.DataFrame(resources_data)
        with pd.ExcelWriter(output_file, mode='a',
                            engine='openpyxl', if_sheet_exists='new') as writer:
            resources_df.to_excel(writer, sheet_name='Resources', index=False)
    
    print(f"\nProject exported successfully to: {output_file}")


def main():
    mpp_file = ('/Users/fernandosimich/Desktop/Workspacegit/MS-Project/'
                'Programa FCC II 2025 Rev 5 19-6 FALTA AFINAR PEM Y DETALLES.mpp')
    output_file = ('/Users/fernandosimich/Desktop/Workspacegit/MS-Project/'
                   'proyecto_exportado.xlsx')
    
    if not os.path.exists(mpp_file):
        print(f"Error: File not found - {mpp_file}")
        sys.exit(1)
    
    print(f"Reading MS Project file: {os.path.basename(mpp_file)}")
    
    try:
        # Read project data
        tasks_data, project = read_ms_project(mpp_file)
        
        if not tasks_data:
            print("No tasks found in the project file!")
            sys.exit(1)
        
        # Convert to DataFrame
        tasks_df = pd.DataFrame(tasks_data)
        
        # Visualize summary
        visualize_project_summary(tasks_df)
        
        # Export to Excel
        export_to_xlsx(tasks_df, project, output_file)
        
        # Show project properties
        print("\n=== PROJECT PROPERTIES ===")
        props = project.getProjectProperties()
        if props.getProjectTitle():
            print(f"Project Title: {props.getProjectTitle()}")
        if props.getManager():
            print(f"Project Manager: {props.getManager()}")
        if props.getStartDate():
            print(f"Start Date: {props.getStartDate()}")
        if props.getFinishDate():
            print(f"Finish Date: {props.getFinishDate()}")
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        if jpype.isJVMStarted():
            jpype.shutdownJVM()


if __name__ == "__main__":
    main()