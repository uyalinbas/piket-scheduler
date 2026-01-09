"""
Excel import/export utilities for schedule data.
"""
import pandas as pd
from datetime import date, timedelta
from typing import List, Optional
from pathlib import Path
import io

from models import (
    Employee, FixedAssignment, ScheduleConfig, ScheduleResult,
    WEEKDAY_NAMES, WEEKDAYS, WEEKEND_DAYS
)


def export_schedule_to_excel(result: ScheduleResult, config: ScheduleConfig, path: str) -> None:
    """
    Export schedule to Excel with multiple sheets:
    - Calendar view
    - Employee statistics
    - Daily assignments
    """
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        # Sheet 1: Daily assignments
        if result.assignments:
            daily_data = []
            for d in sorted(result.assignments.keys()):
                emp = result.assignments[d]
                daily_data.append({
                    'Date': d.strftime('%Y-%m-%d'),
                    'Day': WEEKDAY_NAMES[d.weekday()],
                    'Week': d.isocalendar()[1],
                    'Employee': emp,
                    'Type': 'Weekday' if d.weekday() in WEEKDAYS else 'Weekend'
                })
            df_daily = pd.DataFrame(daily_data)
            df_daily.to_excel(writer, sheet_name='Daily Schedule', index=False)
        
        # Sheet 2: Calendar view (weeks as rows, days as columns)
        if result.assignments:
            # Group by ISO week
            week_data = {}
            for d, emp in result.assignments.items():
                iso = d.isocalendar()
                week_key = (iso[0], iso[1])
                if week_key not in week_data:
                    week_data[week_key] = {day: '' for day in WEEKDAY_NAMES}
                week_data[week_key][WEEKDAY_NAMES[d.weekday()]] = emp
            
            calendar_rows = []
            for (year, week), days in sorted(week_data.items()):
                row = {'Year': year, 'Week': week}
                row.update(days)
                calendar_rows.append(row)
            
            df_cal = pd.DataFrame(calendar_rows)
            cols = ['Year', 'Week'] + WEEKDAY_NAMES
            df_cal = df_cal[cols]
            df_cal.to_excel(writer, sheet_name='Calendar View', index=False)
        
        # Sheet 3: Employee statistics
        if result.employee_stats:
            stats_data = []
            for name, s in result.employee_stats.items():
                stats_data.append({
                    'Employee': name,
                    'Total Duties': s.total_duties,
                    'Weekdays (Mon-Fri)': s.weekday_duties,
                    'Weekends (Sat+Sun)': s.weekend_duties,
                    'Saturdays': s.saturday_count,
                    'Sundays': s.sunday_count,
                    'Fridays': s.friday_count,
                    'Fixed Weekdays': s.fixed_weekdays,
                    'Fixed Weekends': s.fixed_weekends,
                    'Variable Weekdays': s.variable_weekdays,
                    'Variable Weekends': s.variable_weekends,
                })
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='Statistics', index=False)
        
        # Sheet 4: Summary info
        summary_data = [{
            'Metric': 'Tolerance Used',
            'Value': result.tolerance_used
        }, {
            'Metric': 'Solve Time (s)',
            'Value': f'{result.solve_time_seconds:.2f}'
        }, {
            'Metric': 'Total Weekdays',
            'Value': result.total_weekdays
        }, {
            'Metric': 'Total Weekends',
            'Value': result.total_weekends
        }, {
            'Metric': 'Remaining Weekdays Pool',
            'Value': result.remaining_weekdays
        }, {
            'Metric': 'Remaining Weekends Pool',
            'Value': result.remaining_weekends
        }]
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name='Summary', index=False)


def create_template_excel(path: str) -> None:
    """
    Create a template Excel file for importing configuration.
    """
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        # Sheet 1: Employees
        employees_df = pd.DataFrame({
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Forbidden Days': ['', 'Monday,Tuesday', 'Saturday'],
            'Is Extra Weekend': [False, False, True],
        })
        employees_df.to_excel(writer, sheet_name='Employees', index=False)
        
        # Sheet 2: Vacations
        vacations_df = pd.DataFrame({
            'Employee': ['Alice', 'Bob'],
            'Start Date': ['2026-01-15', '2026-02-01'],
            'End Date': ['2026-01-20', '2026-02-05'],
        })
        vacations_df.to_excel(writer, sheet_name='Vacations', index=False)
        
        # Sheet 3: Fixed Assignments
        fixed_df = pd.DataFrame({
            'Day': ['Monday', 'Friday'],
            'Employee': ['Alice', 'Bob'],
        })
        fixed_df.to_excel(writer, sheet_name='Fixed Assignments', index=False)
        
        # Sheet 4: Instructions
        instructions_df = pd.DataFrame({
            'Instructions': [
                'Employees sheet: List all employees with their forbidden days (comma-separated)',
                'Forbidden days: Monday, Tuesday, Wednesday, Thursday, Friday, Saturday, Sunday',
                'Vacations sheet: Define vacation periods with start and end dates (YYYY-MM-DD)',
                'Fixed Assignments: Recurring weekly assignments by day of week',
                'Only one employee can be marked as Is Extra Weekend = True',
            ]
        })
        instructions_df.to_excel(writer, sheet_name='Instructions', index=False)


def import_config_from_excel(path: str) -> tuple:
    """
    Import configuration from Excel template.
    Returns (employees, fixed_assignments, error_message).
    """
    try:
        # Read employees
        df_emp = pd.read_excel(path, sheet_name='Employees')
        employees = []
        
        day_name_to_num = {name: i for i, name in enumerate(WEEKDAY_NAMES)}
        
        for _, row in df_emp.iterrows():
            name = str(row['Name']).strip()
            if not name or name == 'nan':
                continue
            
            # Parse forbidden days
            forbidden = set()
            forbidden_str = str(row.get('Forbidden Days', '')).strip()
            if forbidden_str and forbidden_str != 'nan':
                for day_name in forbidden_str.split(','):
                    day_name = day_name.strip()
                    if day_name in day_name_to_num:
                        forbidden.add(day_name_to_num[day_name])
            
            is_extra = bool(row.get('Is Extra Weekend', False))
            
            emp = Employee(
                name=name,
                forbidden_weekdays=forbidden,
                is_extra_weekend=is_extra
            )
            employees.append(emp)
        
        # Read vacations
        try:
            df_vac = pd.read_excel(path, sheet_name='Vacations')
            for _, row in df_vac.iterrows():
                emp_name = str(row['Employee']).strip()
                
                if emp_name == 'nan' or pd.isna(row.get('Start Date')) or pd.isna(row.get('End Date')):
                    continue
                
                # Parse dates - handle multiple formats
                def parse_date(val):
                    if hasattr(val, 'date'):  # pandas Timestamp
                        return val.date()
                    if hasattr(val, 'strftime'):  # datetime
                        return val if isinstance(val, date) else val.date()
                    s = str(val).strip()[:10]
                    # Try YYYY-MM-DD
                    try:
                        return date.fromisoformat(s)
                    except:
                        pass
                    # Try DD/MM/YYYY
                    try:
                        parts = s.split('/')
                        if len(parts) == 3:
                            return date(int(parts[2]), int(parts[1]), int(parts[0]))
                    except:
                        pass
                    return None
                
                start_date = parse_date(row['Start Date'])
                end_date = parse_date(row['End Date'])
                
                if start_date and end_date:
                    for emp in employees:
                        if emp.name == emp_name:
                            emp.vacation_ranges.append((start_date, end_date))
                            break
        except Exception:
            pass  # Vacations sheet optional
        
        # Read fixed assignments
        fixed_assignments = []
        try:
            df_fixed = pd.read_excel(path, sheet_name='Fixed Assignments')
            for _, row in df_fixed.iterrows():
                day_name = str(row['Day']).strip()
                emp_name = str(row['Employee']).strip()
                
                if day_name == 'nan' or emp_name == 'nan':
                    continue
                
                if day_name in day_name_to_num:
                    fixed_assignments.append(FixedAssignment(
                        day_of_week=day_name_to_num[day_name],
                        employee_name=emp_name
                    ))
        except Exception:
            pass  # Fixed assignments sheet optional
        
        return employees, fixed_assignments, None
        
    except Exception as e:
        return [], [], f"Error reading Excel file: {str(e)}"


def export_schedule_to_bytes(result: ScheduleResult, config: ScheduleConfig) -> bytes:
    """
    Export schedule to Excel and return as bytes (for Streamlit download).
    """
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Same logic as export_schedule_to_excel
        if result.assignments:
            daily_data = []
            for d in sorted(result.assignments.keys()):
                emp = result.assignments[d]
                daily_data.append({
                    'Date': d.strftime('%Y-%m-%d'),
                    'Day': WEEKDAY_NAMES[d.weekday()],
                    'Week': d.isocalendar()[1],
                    'Employee': emp,
                    'Type': 'Weekday' if d.weekday() in WEEKDAYS else 'Weekend'
                })
            df_daily = pd.DataFrame(daily_data)
            df_daily.to_excel(writer, sheet_name='Daily Schedule', index=False)
        
        if result.assignments:
            week_data = {}
            for d, emp in result.assignments.items():
                iso = d.isocalendar()
                week_key = (iso[0], iso[1])
                if week_key not in week_data:
                    week_data[week_key] = {day: '' for day in WEEKDAY_NAMES}
                week_data[week_key][WEEKDAY_NAMES[d.weekday()]] = emp
            
            calendar_rows = []
            for (year, week), days in sorted(week_data.items()):
                row = {'Year': year, 'Week': week}
                row.update(days)
                calendar_rows.append(row)
            
            df_cal = pd.DataFrame(calendar_rows)
            cols = ['Year', 'Week'] + WEEKDAY_NAMES
            df_cal = df_cal[cols]
            df_cal.to_excel(writer, sheet_name='Calendar View', index=False)
        
        if result.employee_stats:
            stats_data = []
            for name, s in result.employee_stats.items():
                stats_data.append({
                    'Employee': name,
                    'Total Duties': s.total_duties,
                    'Weekdays (Mon-Fri)': s.weekday_duties,
                    'Weekends (Sat+Sun)': s.weekend_duties,
                    'Saturdays': s.saturday_count,
                    'Sundays': s.sunday_count,
                    'Variable Weekdays': s.variable_weekdays,
                    'Variable Weekends': s.variable_weekends,
                })
            df_stats = pd.DataFrame(stats_data)
            df_stats.to_excel(writer, sheet_name='Statistics', index=False)
    
    return buffer.getvalue()
