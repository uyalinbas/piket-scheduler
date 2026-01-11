"""
Roster Scheduling System - Streamlit Application
A constraint-based piket scheduler using OR-Tools CP-SAT.

Made by Utku Yalinbas
"""
import streamlit as st
import pandas as pd
from datetime import date, timedelta
from typing import List, Dict, Optional
import json

from models import (
    Employee, FixedAssignment, ScheduleConfig, ScheduleResult,
    WEEKDAY_NAMES, WEEKDAYS, WEEKEND_DAYS, get_date_range_from_weeks
)
from solver import solve_schedule, compute_pool_stats
from validation import validate_config, compute_theoretical_bounds
from excel_utils import export_schedule_to_bytes

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="Piket Scheduler",
    page_icon="🚢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== CUSTOM CSS ====================
st.markdown("""
<style>
    /* Light theme with clean look */
    .stApp {
        background: linear-gradient(135deg, #e8f4f8 0%, #f0f4f8 50%, #e8eef8 100%);
    }
    
    /* Sidebar styling */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
    }
    
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] h1, h2, h3 {
        color: white !important;
    }
    
    /* Keep input text dark for readability */
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] select {
        color: #333 !important;
        background: white !important;
    }
    
    /* Main content cards */
    .main-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        margin: 1rem 0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    
    /* Section headers */
    .section-header {
        font-size: 1.5rem;
        font-weight: 600;
        color: #1a1a2e;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    .section-desc {
        color: #666;
        font-size: 0.9rem;
        margin-bottom: 1rem;
    }
    
    /* Quick start buttons */
    .quick-btn {
        background: linear-gradient(90deg, #667eea, #764ba2);
        color: white;
        border: none;
        padding: 0.75rem 1.5rem;
        border-radius: 8px;
        font-weight: 500;
    }
    
    /* Success/info boxes */
    .success-box {
        background: #d4edda;
        border-left: 4px solid #28a745;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
    }
    
    .info-box {
        background: #d1ecf1;
        border-left: 4px solid #17a2b8;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
    }
    
    .warning-box {
        background: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 1rem;
        border-radius: 8px;
        margin: 0.5rem 0;
    }
    
    /* Data table styling */
    .dataframe {
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Computed values in sidebar */
    .computed-value {
        background: rgba(255,255,255,0.15);
        padding: 0.5rem 1rem;
        border-radius: 8px;
        margin: 0.25rem 0;
    }
</style>
""", unsafe_allow_html=True)

# ==================== DEMO DATA ====================
def get_demo_employees() -> List[Employee]:
    """Return demo employee data."""
    return [
        Employee(name="Alice", forbidden_weekdays=set(), is_extra_weekend=True),
        Employee(name="Bob", forbidden_weekdays=set()),
        Employee(name="Charlie", forbidden_weekdays={0}),  # No Monday
        Employee(name="Diana", forbidden_weekdays={4}),    # No Friday
        Employee(name="Eve", forbidden_weekdays=set()),
        Employee(name="Frank", forbidden_weekdays={5, 6}), # No weekends
        Employee(name="Grace", forbidden_weekdays=set()),
        Employee(name="Henry", forbidden_weekdays=set()),
    ]

def get_demo_fixed_assignments() -> List[FixedAssignment]:
    """Return demo fixed assignments."""
    return [
        FixedAssignment(day_of_week=0, employee_name="Alice"),  # Monday -> Alice
    ]

# ==================== SESSION STATE ====================
def init_session_state():
    if 'employees' not in st.session_state:
        st.session_state.employees = []  # Start empty, no demo data
    if 'fixed_assignments' not in st.session_state:
        st.session_state.fixed_assignments = []  # Start empty
    if 'schedule_result' not in st.session_state:
        st.session_state.schedule_result = None

init_session_state()

# ==================== SIDEBAR ====================
with st.sidebar:
    # NextLogic Logo
    try:
        st.image("nextlogic_logo.png", width=150)
    except:
        pass  # Logo not found, continue without it
    
    st.markdown("## 🚢 Schedule Settings")
    
    # Year
    current_year = date.today().year
    year = st.selectbox("Year", options=list(range(current_year - 1, current_year + 3)), index=1)
    
    # Week range
    col1, col2 = st.columns(2)
    with col1:
        start_week = st.number_input("Start Week", min_value=1, max_value=53, value=1)
    with col2:
        end_week = st.number_input("End Week", min_value=1, max_value=53, value=27)
    
    # Compute date range from weeks
    try:
        start_date, end_date = get_date_range_from_weeks(year, start_week, end_week)
    except:
        start_date = date(year, 1, 6)
        end_date = date(year, 7, 5)
    
    st.divider()
    
    # Computed Values
    st.markdown("## 📊 Computed Values")
    
    config = ScheduleConfig(
        start_date=start_date,
        end_date=end_date,
        employees=st.session_state.employees,
        fixed_assignments=st.session_state.fixed_assignments,
    )
    
    H = config.get_num_weeks()
    N = len(config.employees)
    weekday_dates = config.get_weekday_dates()
    weekend_dates = config.get_weekend_dates()
    
    st.markdown(f"**H (weeks):** {H}")
    st.markdown(f"**N (employees):** {N}")
    
    # Spacing window
    L = max(4, min(10, N - 1)) if N > 1 else 4
    st.markdown(f"**L (spacing window):** {L} weeks")
    
    st.markdown(f"**Total weekday slots:** {len(weekday_dates)}")
    
    if N > 0:
        target_wd = len(weekday_dates) / N
        st.markdown(f"**Target weekdays/employee:** {target_wd:.1f}")
    
    st.divider()
    
    # Branding
    st.markdown("---")
    st.markdown("*Made by Utku Yalinbas*")


# ==================== MAIN CONTENT ====================
st.markdown("*A constraint-based piket scheduler using OR-Tools CP-SAT*")

# ==================== QUICK START ====================
st.markdown('<div class="section-header">🚀 Quick Start</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Clear All", use_container_width=True):
        st.session_state.employees = []
        st.session_state.fixed_assignments = []
        st.session_state.schedule_result = None
        st.session_state.last_uploaded_file = None
        st.rerun()

with col2:
    # Download template button
    from excel_utils import import_config_from_excel
    import io
    
    # Create empty template for users to fill
    template_buffer = io.BytesIO()
    with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
        # Employees sheet - empty with headers and example
        emp_df = pd.DataFrame({
            'Name': ['ExampleEmployee'],
            'Forbidden Days': [''],
            'Is Extra Weekend': [False],
        })
        emp_df.to_excel(writer, sheet_name='Employees', index=False)
        
        # Vacations sheet - empty but with headers
        vac_df = pd.DataFrame({
            'Employee': [],
            'Start Date': [],
            'End Date': [],
        })
        vac_df.to_excel(writer, sheet_name='Vacations', index=False)
        
        # Fixed Assignments sheet - empty with headers
        fixed_df = pd.DataFrame({
            'Day': [],
            'Employee': [],
        })
        fixed_df.to_excel(writer, sheet_name='Fixed Assignments', index=False)
    
    st.download_button(
        "📥 Download Template",
        data=template_buffer.getvalue(),
        file_name="employee_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# Excel Import
st.markdown("**Import from Excel:**")
uploaded_file = st.file_uploader("Upload employee data (XLSX)", type=['xlsx'], label_visibility="collapsed", key="excel_upload")

# Track if we already processed this file
if 'last_uploaded_file' not in st.session_state:
    st.session_state.last_uploaded_file = None

if uploaded_file is not None and uploaded_file.name != st.session_state.last_uploaded_file:
    employees, fixed_assignments, error = import_config_from_excel(uploaded_file)
    if error:
        st.error(error)
    else:
        st.session_state.employees = employees
        st.session_state.fixed_assignments = fixed_assignments
        st.session_state.schedule_result = None
        st.session_state.last_uploaded_file = uploaded_file.name
        st.success(f"✅ Imported {len(employees)} employees, {len(fixed_assignments)} fixed assignments")
        st.rerun()

st.divider()

# ==================== EMPLOYEES ====================
st.markdown('<div class="section-header">👥 Employees</div>', unsafe_allow_html=True)

col_add, col_count = st.columns([1, 3])
with col_add:
    if st.button("➕ Add Employee", use_container_width=True):
        new_name = f"User {len(st.session_state.employees) + 1}"
        st.session_state.employees.append(Employee(name=new_name))
        st.rerun()

with col_count:
    st.markdown(f"*Total: {len(st.session_state.employees)} employees*")

# Display employees
for i, emp in enumerate(st.session_state.employees):
    with st.expander(f"👤 {emp.name}", expanded=False):
        col1, col2 = st.columns([4, 1])
        
        with col1:
            # Name
            new_name = st.text_input("Name", value=emp.name, key=f"emp_name_{i}")
            if new_name != emp.name:
                # Update fixed assignments too
                for fa in st.session_state.fixed_assignments:
                    if fa.employee_name == emp.name:
                        fa.employee_name = new_name
                st.session_state.employees[i].name = new_name
            
            # Extra weekend employee
            is_extra = st.checkbox(
                "Extra weekend employee ⓘ",
                value=emp.is_extra_weekend,
                key=f"emp_extra_{i}",
                help="This employee works exactly H weekend days"
            )
            if is_extra != emp.is_extra_weekend:
                # Only one can be extra
                if is_extra:
                    for j, e in enumerate(st.session_state.employees):
                        if j != i:
                            e.is_extra_weekend = False
                st.session_state.employees[i].is_extra_weekend = is_extra
            
            # Forbidden Days
            st.markdown("**Forbidden Days** (checked = cannot work)")
            fcols = st.columns(7)
            for d, day_name in enumerate(WEEKDAY_NAMES):
                with fcols[d]:
                    is_forbidden = st.checkbox(
                        day_name[:3],
                        value=d in emp.forbidden_weekdays,
                        key=f"emp_forbidden_{i}_{d}"
                    )
                    if is_forbidden and d not in emp.forbidden_weekdays:
                        st.session_state.employees[i].forbidden_weekdays.add(d)
                    elif not is_forbidden and d in emp.forbidden_weekdays:
                        st.session_state.employees[i].forbidden_weekdays.discard(d)
            
            # Vacation Periods
            st.markdown("**Vacation Periods** (date ranges when employee is unavailable)")
            
            # Display existing vacations
            for v_idx, (vac_start, vac_end) in enumerate(emp.vacation_ranges):
                vcol1, vcol2, vcol3 = st.columns([2, 2, 1])
                with vcol1:
                    st.text(f"{vac_start.strftime('%Y-%m-%d')}")
                with vcol2:
                    st.text(f"→ {vac_end.strftime('%Y-%m-%d')}")
                with vcol3:
                    if st.button("❌", key=f"vac_del_{i}_{v_idx}"):
                        st.session_state.employees[i].vacation_ranges.pop(v_idx)
                        st.rerun()
            
            # Add new vacation
            st.markdown("*Add new vacation:*")
            vcol1, vcol2, vcol3 = st.columns([2, 2, 1])
            with vcol1:
                st.caption("Start Date")
                vac_start_input = st.date_input(
                    "Start",
                    value=start_date,
                    key=f"vac_start_{i}",
                    label_visibility="collapsed"
                )
            with vcol2:
                st.caption("End Date")
                vac_end_input = st.date_input(
                    "End",
                    value=start_date + timedelta(days=6),
                    key=f"vac_end_{i}",
                    label_visibility="collapsed"
                )
            with vcol3:
                st.caption(" ")  # Spacer
                if st.button("➕ Add", key=f"vac_add_{i}"):
                    if vac_start_input <= vac_end_input:
                        st.session_state.employees[i].vacation_ranges.append((vac_start_input, vac_end_input))
                        st.rerun()
        
        with col2:
            if st.button("🗑️ Remove", key=f"emp_del_{i}"):
                # Remove from fixed assignments
                st.session_state.fixed_assignments = [
                    fa for fa in st.session_state.fixed_assignments
                    if fa.employee_name != emp.name
                ]
                st.session_state.employees.pop(i)
                st.rerun()

st.divider()

# ==================== SCHEDULING OPTIONS ====================
st.markdown('<div class="section-header">⚙️ Scheduling Options</div>', unsafe_allow_html=True)

link_fri_sat = st.checkbox(
    "🔗 Link Friday → Saturday (same employee) ⓘ",
    value=True,
    help="When enabled, the same employee works both Friday and Saturday each week"
)

# ==================== FAIRNESS CONSTRAINTS ====================
st.markdown('<div class="section-header">⚖️ Fairness Constraints</div>', unsafe_allow_html=True)
st.markdown('<div class="section-desc">Control how evenly work is distributed (fix-aware).</div>', unsafe_allow_html=True)

enforce_hard_fairness = st.checkbox(
    "✅ Enforce hard fairness (guarantee nearly equal totals) ⓘ",
    value=True,
    help="Makes fairness a hard constraint with the specified tolerance"
)

col1, col2 = st.columns(2)
with col1:
    weekday_tolerance = st.number_input("Weekday tolerance (±)", min_value=1, max_value=10, value=1)
with col2:
    weekend_tolerance = st.number_input("Weekend tolerance (±)", min_value=1, max_value=10, value=1)

auto_relax = st.checkbox(
    "Auto-relax fairness if infeasible ⓘ",
    value=True,
    help="Automatically increase tolerance if no solution is found"
)

col1, col2 = st.columns([3, 1])
with col2:
    max_tolerance = st.number_input("Max tolerance", min_value=1, max_value=10, value=4)

st.divider()

# ==================== SMOOTHING ====================
st.markdown('<div class="section-header">📉 Smoothing / Anti-Consecutive</div>', unsafe_allow_html=True)
st.markdown('<div class="section-desc">Reduce consecutive day assignments</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    penalize_consecutive = st.checkbox("✅ Penalize consecutive days ⓘ", value=True)
with col2:
    penalize_consecutive_wd = st.checkbox("✅ Penalize consecutive weekdays (Mon-Thu) ⓘ", value=True)

col1, col2 = st.columns(2)
with col1:
    consec_weight = st.number_input("Consecutive day weight", min_value=1, max_value=1000, value=50)
with col2:
    consec_wd_weight = st.number_input("Consecutive weekday weight", min_value=1, max_value=1000, value=20)

st.divider()

# ==================== FIXED DAY ASSIGNMENTS ====================
st.markdown('<div class="section-header">📌 Fixed Day Assignments</div>', unsafe_allow_html=True)
st.markdown('<div class="section-desc">Assign specific employees to work on specific days every week</div>', unsafe_allow_html=True)

# Build current fixed map
fixed_map = {fa.day_of_week: fa.employee_name for fa in st.session_state.fixed_assignments}
emp_names = ["(None)"] + [e.name for e in st.session_state.employees]

fixed_cols = st.columns(7)
for dow in range(7):
    with fixed_cols[dow]:
        st.markdown(f"**{WEEKDAY_NAMES[dow]}**")
        current = fixed_map.get(dow, "(None)")
        selected = st.selectbox(
            f"fixed_{dow}",
            options=emp_names,
            index=emp_names.index(current) if current in emp_names else 0,
            key=f"fixed_dow_{dow}",
            label_visibility="collapsed"
        )
        
        # Update fixed assignments
        if selected != "(None)" and (dow not in fixed_map or fixed_map[dow] != selected):
            # Remove old
            st.session_state.fixed_assignments = [fa for fa in st.session_state.fixed_assignments if fa.day_of_week != dow]
            # Add new
            st.session_state.fixed_assignments.append(FixedAssignment(day_of_week=dow, employee_name=selected))
        elif selected == "(None)" and dow in fixed_map:
            # Remove
            st.session_state.fixed_assignments = [fa for fa in st.session_state.fixed_assignments if fa.day_of_week != dow]

st.divider()

# ==================== SOLVE BUTTON ====================
if st.button("🔧 Solve Schedule", type="primary", use_container_width=True):
    config = ScheduleConfig(
        start_date=start_date,
        end_date=end_date,
        employees=st.session_state.employees,
        fixed_assignments=st.session_state.fixed_assignments,
        link_friday_saturday=link_fri_sat
    )
    
    is_valid, errors = validate_config(config)
    
    if not is_valid:
        for err in errors:
            st.error(err)
    else:
        with st.spinner("Solving... This may take a moment."):
            result = solve_schedule(
                config,
                max_tolerance=max_tolerance if auto_relax else weekday_tolerance,
                time_limit_seconds=60
            )
            st.session_state.schedule_result = result

# ==================== RESULTS ====================
if st.session_state.schedule_result:
    result = st.session_state.schedule_result
    
    if result.success:
        st.markdown('<div class="success-box">✅ Solution found! Status: OPTIMAL</div>', unsafe_allow_html=True)
        
        st.markdown(f'<div class="info-box">📊 Fairness tolerance used: ±{result.tolerance_used} (WD), ±{result.tolerance_used} (WE)</div>', unsafe_allow_html=True)
        
        # Count consecutive pairs
        sorted_dates = sorted(result.assignments.keys())
        consec_pairs = 0
        for i in range(len(sorted_dates) - 1):
            d1, d2 = sorted_dates[i], sorted_dates[i + 1]
            if (d2 - d1).days == 1 and result.assignments[d1] == result.assignments[d2]:
                consec_pairs += 1
        
        st.markdown(f'<div class="info-box">ℹ️ Consecutive day pairs: {consec_pairs}</div>', unsafe_allow_html=True)
        
        # Show released vacations if any
        if hasattr(result, 'released_vacations') and result.released_vacations:
            released_text = ", ".join([f"{name} ({start} to {end})" for name, start, end in result.released_vacations])
            st.markdown(f'<div class="warning-box">⚠️ Released vacations for fairness: {released_text}</div>', unsafe_allow_html=True)
        
        st.divider()
        
        # Results table
        st.markdown('<div class="section-header">📊 Results</div>', unsafe_allow_html=True)
        
        results_data = []
        for d in sorted(result.assignments.keys()):
            emp = result.assignments[d]
            results_data.append({
                'Week': d.isocalendar()[1],
                'Date': d.strftime('%Y-%m-%d'),
                'Day': WEEKDAY_NAMES[d.weekday()],
                'Employee': emp
            })
        
        df_results = pd.DataFrame(results_data)
        st.dataframe(df_results, use_container_width=True, hide_index=True, height=400)
        
        # Export
        col1, col2, col3 = st.columns([1, 1, 2])
        with col1:
            excel_bytes = export_schedule_to_bytes(result, config)
            st.download_button(
                "📥 Download Excel",
                data=excel_bytes,
                file_name=f"schedule_{start_date}_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        st.divider()
        
        # Per-Employee Statistics
        st.markdown('<div class="section-header">📈 Per-Employee Statistics</div>', unsafe_allow_html=True)
        
        if result.employee_stats:
            # Get bounds for targets
            bounds = compute_theoretical_bounds(config)
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Weekday Slots", bounds['total_weekdays'])
            with col2:
                st.metric("WD Target (var)", f"{bounds['remaining_weekdays'] / len(result.employee_stats):.1f}" if result.employee_stats else "N/A")
            with col3:
                # WD Band - exclude extra WE employees
                pool_wd = [s.variable_weekdays for n, s in result.employee_stats.items()
                           if not config.get_employee_by_name(n).is_extra_weekend]
                if pool_wd:
                    st.metric("WD Band", f"{min(pool_wd)}-{max(pool_wd)}")
                else:
                    st.metric("WD Band", "N/A")
            with col4:
                # WE Band - all employees (extra WE's Var WE = pool share)
                all_we = [s.variable_weekends for n, s in result.employee_stats.items()]
                if all_we:
                    st.metric("WE Band", f"{min(all_we)}-{max(all_we)}")
                else:
                    st.metric("WE Band", "N/A")
            
            # Stats table
            stats_data = []
            for name, s in result.employee_stats.items():
                emp = config.get_employee_by_name(name)
                stats_data.append({
                    'Employee': name,
                    'Extra WE': '✅' if emp and emp.is_extra_weekend else '',
                    'Total': s.total_duties,
                    'Weekdays': s.weekday_duties,
                    'Var WD': s.variable_weekdays,
                    'Weekends': s.weekend_duties,
                    'Var WE': s.variable_weekends,
                    'Sat': s.saturday_count,
                    'Sun': s.sunday_count,
                })
            
            df_stats = pd.DataFrame(stats_data)
            st.dataframe(df_stats, use_container_width=True, hide_index=True)
    
    else:
        st.error(f"❌ {result.error_message}")
