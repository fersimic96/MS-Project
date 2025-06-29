#!/usr/bin/env python3
"""
Gantt Chart Visualizer for MS Project Excel exports
Reads Excel files exported from MS Project and creates interactive Gantt charts
"""
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import argparse
import sys


def parse_duration(duration_str):
    """Convert MS Project duration format to days"""
    if not duration_str or pd.isna(duration_str):
        return 0
    
    duration_str = str(duration_str).lower()
    
    # Extract number and unit
    import re
    match = re.match(r'([\d.]+)\s*([a-z]+)', duration_str)
    if not match:
        return 0
    
    value = float(match.group(1))
    unit = match.group(2)
    
    # Convert to days
    if unit.startswith('d'):  # days
        return value
    elif unit.startswith('h') or unit.startswith('e'):  # hours (eh = elapsed hours)
        return value / 8  # 8 hours per day
    elif unit.startswith('w'):  # weeks
        return value * 5  # 5 days per week
    elif unit.startswith('m'):  # months
        return value * 20  # 20 days per month (approx)
    else:
        return value


def prepare_gantt_data(excel_file, sheet_name='Tasks'):
    """Read Excel file and prepare data for Gantt chart"""
    
    # Detectar si es archivo del conversor corregido
    try:
        xl_file = pd.ExcelFile(excel_file)
        if 'Tasks_Corrected' in xl_file.sheet_names:
            sheet_name = 'Tasks_Corrected'
            print(f"✓ Detected corrected converter file, using sheet: {sheet_name}")
    except:
        pass
    
    # Read Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    
    # Adaptar columnas según el tipo de archivo
    if 'Duration_Corrected' in df.columns:
        # Archivo del conversor corregido
        df['Duration'] = df['Duration_Corrected']
        duration_col = 'Duration_Corrected_Hours'
        if duration_col in df.columns:
            df['Duration_Hours'] = df[duration_col]
        print("✓ Using corrected duration data")
    else:
        # Archivo estándar
        duration_col = 'Duration'
        print("✓ Using standard duration data")
    
    # Filter out rows without dates
    df = df[df['Start'].notna() & df['Finish'].notna()].copy()
    
    # Parse dates - handle different formats
    for col in ['Start', 'Finish']:
        # Try to parse as string first
        df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Add completion status for coloring
    df['Status'] = df['Percent Complete'].apply(
        lambda x: 'Complete' if x >= 100 else
                  'In Progress' if x > 0 else
                  'Not Started'
    )
    
    # Create hover text adaptado
    duration_info = df['Duration'] if 'Duration' in df.columns else 'N/A'
    
    df['HoverText'] = df.apply(
        lambda row: (
            f"<b>{row['Name']}</b><br>"
            f"WBS: {row.get('WBS', 'N/A')}<br>"
            f"Duration: {row.get('Duration', 'N/A')}<br>"
            f"Progress: {row['Percent Complete']:.0f}%<br>"
            f"Resources: {row.get('Resource Names', 'N/A')}<br>"
            f"Predecessors: {row.get('Predecessors', 'N/A')}<br>"
            + (f"Source: {row['Duration_Source']}<br>" if 'Duration_Source' in df.columns else "")
        ), axis=1
    )
    
    # Add hierarchy levels for better visualization
    df['DisplayName'] = df.apply(
        lambda row: '  ' * int(row.get('Outline Level', 0)) + row['Name'], 
        axis=1
    )
    
    return df


def create_gantt_chart(df, title='Project Gantt Chart'):
    """Create interactive Gantt chart using Plotly"""
    # Create figure
    fig = px.timeline(
        df,
        x_start='Start',
        x_end='Finish',
        y='DisplayName',
        color='Status',
        hover_name='HoverText',
        title=title,
        color_discrete_map={
            'Complete': '#28a745',
            'In Progress': '#ffc107',
            'Not Started': '#6c757d'
        }
    )
    
    # Update layout
    fig.update_layout(
        height=max(600, len(df) * 20),  # Dynamic height
        xaxis_title='Timeline',
        yaxis_title='Tasks',
        yaxis_autorange='reversed',  # Show tasks from top to bottom
        showlegend=True,
        hovermode='closest',
        plot_bgcolor='white'
    )
    
    # Add grid lines
    fig.update_xaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor='lightgray',
        dtick='M1',  # Monthly grid
        tickformat='%b %Y'
    )
    
    fig.update_yaxes(
        showgrid=True,
        gridwidth=1,
        gridcolor='lightgray'
    )
    
    # Add critical path highlighting
    critical_tasks = df[df['Critical'] == True]
    if not critical_tasks.empty:
        for _, task in critical_tasks.iterrows():
            fig.add_shape(
                type='rect',
                x0=task['Start'],
                x1=task['Finish'],
                y0=task['DisplayName'],
                y1=task['DisplayName'],
                line=dict(color='red', width=3),
                fillcolor='rgba(255,0,0,0.1)'
            )
    
    # Add milestone markers
    milestones = df[df['Milestone'] == True]
    if not milestones.empty:
        fig.add_trace(go.Scatter(
            x=milestones['Start'],
            y=milestones['DisplayName'],
            mode='markers',
            marker=dict(
                symbol='diamond',
                size=12,
                color='purple'
            ),
            name='Milestones',
            showlegend=True,
            hovertext=milestones['Name']
        ))
    
    return fig


def create_resource_chart(excel_file):
    """Create resource allocation chart"""
    try:
        # Intentar leer hoja de recursos
        xl_file = pd.ExcelFile(excel_file)
        resource_sheet = None
        
        # Buscar hoja de recursos
        for sheet in xl_file.sheet_names:
            if 'Resource' in sheet:
                resource_sheet = sheet
                break
        
        if resource_sheet:
            resources_df = pd.read_excel(excel_file, sheet_name=resource_sheet)
            print(f"✓ Found resource sheet: {resource_sheet}")
        else:
            print("⚠️ No resource sheet found")
            return None
        
        # Verificar que hay datos
        if resources_df.empty:
            return None
        
        # Create bar chart
        fig = px.bar(
            resources_df,
            x='Name',
            y='Cost',
            title='Resource Costs',
            text='Cost',
            color='Type' if 'Type' in resources_df.columns else None
        )
        
        fig.update_traces(texttemplate='%{text:.2f}', textposition='outside')
        fig.update_layout(
            xaxis_title='Resources',
            yaxis_title='Cost',
            showlegend=True
        )
        
        return fig
    except Exception as e:
        print(f"⚠️ Could not create resource chart: {e}")
        return None


def main():
    parser = argparse.ArgumentParser(
        description='Create interactive Gantt chart from MS Project Excel export'
    )
    parser.add_argument('excel_file', help='Path to Excel file')
    parser.add_argument('-o', '--output', help='Output HTML file',
                        default='gantt_chart.html')
    parser.add_argument('-t', '--title', help='Chart title',
                        default='Project Gantt Chart')
    parser.add_argument('--no-browser', action='store_true',
                        help='Do not open browser after generation')
    parser.add_argument('--resources', action='store_true',
                        help='Include resource chart')
    parser.add_argument('--sheet-name', help='Name of Excel sheet to read',
                        default='Tasks')
    
    args = parser.parse_args()
    
    try:
        print(f"Reading Excel file: {args.excel_file}")
        
        # Prepare data
        df = prepare_gantt_data(args.excel_file, args.sheet_name)
        
        print(f"Found {len(df)} tasks with valid dates")
        
        # Create Gantt chart
        gantt_fig = create_gantt_chart(df, args.title)
        
        if args.resources:
            # Create resource chart
            resource_fig = create_resource_chart(args.excel_file)
            
            if resource_fig:
                # Create subplot with both charts
                from plotly.subplots import make_subplots
                
                fig = make_subplots(
                    rows=2, cols=1,
                    row_heights=[0.7, 0.3],
                    subplot_titles=(args.title, 'Resource Allocation'),
                    specs=[[{"type": "timeline"}], [{"type": "bar"}]]
                )
                
                # Add traces
                for trace in gantt_fig.data:
                    fig.add_trace(trace, row=1, col=1)
                for trace in resource_fig.data:
                    fig.add_trace(trace, row=2, col=1)
                
                fig.update_layout(height=1000, showlegend=True)
                fig.write_html(args.output)
            else:
                gantt_fig.write_html(args.output)
        else:
            gantt_fig.write_html(args.output)
        
        print(f"✓ Gantt chart saved to: {args.output}")
        
        # Open in browser
        if not args.no_browser:
            import webbrowser
            webbrowser.open(args.output)
            
    except Exception as e:
        print(f"Error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()