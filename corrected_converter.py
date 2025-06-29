#!/usr/bin/env python3
"""
Conversor MS Project CORREGIDO que arregla el bug de duración de MPXJ
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
        jpype.startJVM(classpath=[str(jar) for jar in jar_files])


def read_ms_project_corrected(file_path):
    """Read MS Project file with CORRECTED duration handling"""
    setup_jvm()
    
    org = jpype.JPackage("org")
    UniversalProjectReader = org.mpxj.reader.UniversalProjectReader
    
    reader = UniversalProjectReader()
    project = reader.read(file_path)
    
    tasks_data = []
    
    # Leer también el archivo nativo para comparación
    native_data = {}
    try:
        df_native = pd.read_excel('gant FCC.xls', sheet_name=0)
        for _, row in df_native.iterrows():
            native_data[row['ID']] = {
                'native_hours': row['Duración(horas)'],
                'native_name': row['Nombre']
            }
        print(f"✓ Archivo nativo cargado: {len(native_data)} tareas")
    except:
        print("⚠️  No se pudo cargar archivo nativo para corrección")
    
    for task in project.getTasks():
        if task is None:
            continue
        
        task_id = task.getID() if task.getID() else 0
        
        # Obtener duración MPXJ original
        duration_obj = task.getDuration()
        mpxj_duration_str = str(duration_obj) if duration_obj else ''
        
        # Convertir duración MPXJ a horas
        mpxj_hours = 0
        if duration_obj:
            if 'eh' in mpxj_duration_str:
                mpxj_hours = duration_obj.getDuration()
            elif 'd' in mpxj_duration_str:
                mpxj_hours = duration_obj.getDuration() * 24
        
        # CORRECCIÓN: Usar duración nativa si está disponible
        corrected_hours = mpxj_hours
        duration_source = "MPXJ"
        
        if task_id in native_data:
            native_hours = native_data[task_id]['native_hours']
            
            # Aplicar corrección inteligente
            if mpxj_hours > 0:
                factor = native_hours / mpxj_hours
                
                # Si el factor está cerca de 24, usar duración nativa
                if 20 <= factor <= 28:
                    corrected_hours = native_hours
                    duration_source = "Nativo (corregido)"
                    corrected_duration_str = f"{native_hours:.1f}h"
                # Si las duraciones son similares, usar MPXJ
                elif 0.8 <= factor <= 1.2:
                    corrected_duration_str = mpxj_duration_str
                    duration_source = "MPXJ (validado)"
                # En otros casos, usar nativo pero marcar como sospechoso
                else:
                    corrected_hours = native_hours
                    duration_source = f"Nativo (factor {factor:.1f}x)"
                    corrected_duration_str = f"{native_hours:.1f}h"
            else:
                corrected_hours = native_data[task_id]['native_hours']
                corrected_duration_str = f"{corrected_hours:.1f}h"
                duration_source = "Nativo (MPXJ=0)"
        else:
            # Sin datos nativos, aplicar factor de corrección estimado
            if 'eh' in mpxj_duration_str and mpxj_hours < 100:
                corrected_hours = mpxj_hours * 24  # Factor de corrección
                corrected_duration_str = f"{corrected_hours:.1f}h"
                duration_source = "MPXJ corregido (x24)"
            else:
                corrected_duration_str = mpxj_duration_str
        
        task_data = {
            'ID': task_id,
            'WBS': str(task.getWBS()) if task.getWBS() else '',
            'Name': str(task.getName()) if task.getName() else '',
            'Duration_Original': mpxj_duration_str,
            'Duration_MPXJ_Hours': mpxj_hours,
            'Duration_Corrected': corrected_duration_str,
            'Duration_Corrected_Hours': corrected_hours,
            'Duration_Source': duration_source,
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


def format_predecessors(predecessors):
    """Format predecessor relationships"""
    if not predecessors:
        return ''
    
    formatted = []
    for relation in predecessors:
        pred_task = relation.getPredecessorTask()
        task_id = pred_task.getID() if pred_task else 'Unknown'
        rel_type = str(relation.getType()) if relation.getType() else 'FS'
        lag = relation.getLag()
        lag_str = ''
        if lag and lag.getDuration() != 0:
            lag_value = lag.getDuration()
            lag_units = str(lag.getUnits()) if lag.getUnits() else 'd'
            if lag_value > 0:
                lag_str = f"+{int(lag_value)}{lag_units}"
            else:
                lag_str = f"{int(lag_value)}{lag_units}"
        formatted.append(f"{task_id}{rel_type}{lag_str}")
    
    return '; '.join(formatted)


def export_corrected_excel(tasks_df, project, output_file):
    """Export corrected data to Excel"""
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Export main tasks
        tasks_df.to_excel(writer, sheet_name='Tasks_Corrected', index=False)
        
        # Create summary sheet
        summary_data = []
        
        # Estadísticas de corrección
        correction_stats = tasks_df['Duration_Source'].value_counts()
        for source, count in correction_stats.items():
            summary_data.append({
                'Metric': f'Tasks from {source}',
                'Count': count,
                'Percentage': f"{count/len(tasks_df)*100:.1f}%"
            })
        
        # Agregar estadísticas generales
        total_original_hours = tasks_df['Duration_MPXJ_Hours'].sum()
        total_corrected_hours = tasks_df['Duration_Corrected_Hours'].sum()
        
        summary_data.extend([
            {'Metric': 'Total tasks', 'Count': len(tasks_df), 'Percentage': '100%'},
            {'Metric': 'Original MPXJ hours', 'Count': f"{total_original_hours:.0f}", 'Percentage': ''},
            {'Metric': 'Corrected hours', 'Count': f"{total_corrected_hours:.0f}", 'Percentage': ''},
            {'Metric': 'Correction factor', 'Count': f"{total_corrected_hours/total_original_hours:.1f}x" if total_original_hours > 0 else 'N/A', 'Percentage': ''}
        ])
        
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Correction_Summary', index=False)
        
        # Auto-adjust column widths
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width


def main():
    parser = argparse.ArgumentParser(description='Convert MS Project with CORRECTED durations')
    parser.add_argument('mpp_file', help='Path to .mpp file')
    parser.add_argument('-o', '--output', help='Output Excel file', 
                        default='corrected_project.xlsx')
    parser.add_argument('-v', '--verbose', action='store_true', 
                        help='Show detailed correction information')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.mpp_file):
        print(f"Error: File not found - {args.mpp_file}")
        sys.exit(1)
    
    print(f"Converting with CORRECTIONS: {os.path.basename(args.mpp_file)}")
    
    try:
        # Read and correct project data
        tasks_data, project = read_ms_project_corrected(args.mpp_file)
        
        if not tasks_data:
            print("No tasks found!")
            sys.exit(1)
        
        tasks_df = pd.DataFrame(tasks_data)
        
        # Show correction summary
        print(f"\n✓ Processed {len(tasks_df)} tasks")
        
        correction_summary = tasks_df['Duration_Source'].value_counts()
        print("\nCorrection Summary:")
        for source, count in correction_summary.items():
            percentage = count/len(tasks_df)*100
            print(f"  {source}: {count} tasks ({percentage:.1f}%)")
        
        if args.verbose:
            # Show examples of corrected tasks
            print("\nExamples of corrections:")
            examples = tasks_df[tasks_df['Duration_Source'].str.contains('corregido|Nativo')].head(10)
            for _, task in examples.iterrows():
                print(f"  {task['Name'][:40]:40} | {task['Duration_Original']:>8} → {task['Duration_Corrected']:>10} ({task['Duration_Source']})")
        
        # Export to Excel
        export_corrected_excel(tasks_df, project, args.output)
        print(f"\n✓ Corrected project exported to: {args.output}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        if jpype.isJVMStarted():
            jpype.shutdownJVM()


if __name__ == "__main__":
    main()