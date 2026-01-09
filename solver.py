"""
OR-Tools CP-SAT constraint solver for roster scheduling.
Implements the pool-based fair distribution paradigm.
"""
from ortools.sat.python import cp_model
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple
import time

from models import (
    Employee, FixedAssignment, ScheduleConfig, ScheduleResult,
    EmployeeStats, WEEKDAYS, WEEKEND_DAYS, SATURDAY, SUNDAY, FRIDAY,
    WEEKDAY_NAMES
)


def compute_pool_stats(config: ScheduleConfig) -> Dict:
    """
    Compute pool statistics for pre-solve analysis.
    Returns detailed breakdown of pools and availability.
    """
    all_dates = config.get_all_dates()
    weekday_dates = config.get_weekday_dates()
    weekend_dates = config.get_weekend_dates()
    H = config.get_num_weeks()
    N = len(config.employees)
    
    # Build fixed assignment lookup: day_of_week -> employee_name
    fixed_dow = {}
    for fa in config.fixed_assignments:
        fixed_dow[fa.day_of_week] = fa.employee_name
    
    # Count fixed assignments per employee - QUOTA BASED (always H per assignment)
    # Vacation doesn't reduce quota, employee compensates from pool
    fixed_weekdays_per_emp = {e.name: 0 for e in config.employees}
    fixed_weekends_per_emp = {e.name: 0 for e in config.employees}
    
    for dow, emp_name in fixed_dow.items():
        if emp_name in fixed_weekdays_per_emp:
            if dow in WEEKDAYS:
                fixed_weekdays_per_emp[emp_name] = H  # Full quota
            if dow in WEEKEND_DAYS:
                fixed_weekends_per_emp[emp_name] = H  # Full quota
    
    total_fixed_weekdays = sum(fixed_weekdays_per_emp.values())
    total_fixed_weekends = sum(fixed_weekends_per_emp.values())
    
    remaining_weekdays = len(weekday_dates) - total_fixed_weekdays
    remaining_weekends = len(weekend_dates) - total_fixed_weekends
    
    # Account for extra weekend employee
    extra_emp = config.get_extra_weekend_employee()
    extra_weekend_quota = H if extra_emp else 0
    remaining_weekends_effective = remaining_weekends - extra_weekend_quota
    
    return {
        'H': H,
        'N': N,
        'total_weekdays': len(weekday_dates),
        'total_weekends': len(weekend_dates),
        'total_fixed_weekdays': total_fixed_weekdays,
        'total_fixed_weekends': total_fixed_weekends,
        'remaining_weekdays': remaining_weekdays,
        'remaining_weekends': remaining_weekends,
        'extra_weekend_quota': extra_weekend_quota,
        'remaining_weekends_effective': remaining_weekends_effective,
        'fixed_weekdays_per_emp': fixed_weekdays_per_emp,
        'fixed_weekends_per_emp': fixed_weekends_per_emp,
        'fixed_dow': fixed_dow,
    }


def solve_schedule(config: ScheduleConfig, max_tolerance: int = 5, time_limit_seconds: int = 60) -> ScheduleResult:
    """
    Solve the roster scheduling problem using CP-SAT.
    Implements auto-relax: tries tolerance 1, 2, ... up to max_tolerance.
    """
    start_time = time.time()
    
    # Basic validation
    if not config.employees:
        return ScheduleResult(
            success=False,
            error_message="No employees defined."
        )
    
    all_dates = config.get_all_dates()
    if not all_dates:
        return ScheduleResult(
            success=False,
            error_message="No dates in the specified range."
        )
    
    # Compute pool stats
    stats = compute_pool_stats(config)
    N = stats['N']
    H = stats['H']
    
    # Check basic feasibility
    if stats['remaining_weekdays'] < 0:
        return ScheduleResult(
            success=False,
            error_message=f"Fixed weekday assignments exceed available weekdays. "
                          f"Fixed: {stats['total_fixed_weekdays']}, Available: {stats['total_weekdays']}"
        )
    
    if stats['remaining_weekends_effective'] < 0:
        return ScheduleResult(
            success=False,
            error_message=f"Weekend pool exhausted. "
                          f"Fixed: {stats['total_fixed_weekends']}, Extra quota: {stats['extra_weekend_quota']}, "
                          f"Available: {stats['total_weekends']}"
        )
    
    # Try solving with increasing tolerance
    for tolerance in range(1, max_tolerance + 1):
        result = _solve_with_tolerance(config, stats, tolerance, time_limit_seconds)
        if result.success:
            result.solve_time_seconds = time.time() - start_time
            return result
    
    # All tolerances failed - try releasing vacations (longest first)
    # Collect all vacations with their lengths
    all_vacations = []
    for e in config.employees:
        for v_idx, (vac_start, vac_end) in enumerate(e.vacation_ranges):
            length = (vac_end - vac_start).days + 1
            all_vacations.append((e.name, v_idx, length, vac_start, vac_end))
    
    # Sort by length descending (release longest first)
    all_vacations.sort(key=lambda x: x[2], reverse=True)
    
    released_vacations = []
    
    for emp_name, v_idx, length, vac_start, vac_end in all_vacations:
        # Find and release this vacation
        for e in config.employees:
            if e.name == emp_name and len(e.vacation_ranges) > v_idx:
                released_vacations.append((emp_name, vac_start, vac_end))
                # Create a modified config without this vacation
                e.vacation_ranges = [v for i, v in enumerate(e.vacation_ranges) if i != v_idx]
                break
        
        # Recompute stats and try again
        stats = compute_pool_stats(config)
        for tolerance in range(1, max_tolerance + 1):
            result = _solve_with_tolerance(config, stats, tolerance, time_limit_seconds)
            if result.success:
                result.solve_time_seconds = time.time() - start_time
                if released_vacations:
                    released_str = ", ".join([f"{n}: {s.strftime('%m/%d')}-{e.strftime('%m/%d')}" 
                                             for n, s, e in released_vacations])
                    result.error_message = f"Released vacations: {released_str}"
                return result
    
    # Still failed after releasing all vacations
    return ScheduleResult(
        success=False,
        error_message=f"Could not find feasible solution with tolerance up to {max_tolerance}. "
                      f"Check employee availability, fixed assignments, and forbidden days.",
        solve_time_seconds=time.time() - start_time,
        total_weekdays=stats['total_weekdays'],
        total_weekends=stats['total_weekends'],
        remaining_weekdays=stats['remaining_weekdays'],
        remaining_weekends=stats['remaining_weekends_effective']
    )


def _solve_with_tolerance(config: ScheduleConfig, stats: Dict, tolerance: int, time_limit: int) -> ScheduleResult:
    """Attempt to solve with a specific tolerance level."""
    
    model = cp_model.CpModel()
    
    all_dates = config.get_all_dates()
    weekday_dates = config.get_weekday_dates()
    weekend_dates = config.get_weekend_dates()
    saturday_dates = config.get_saturday_dates()
    sunday_dates = config.get_sunday_dates()
    friday_dates = config.get_friday_dates()
    
    employees = config.employees
    N = len(employees)
    H = stats['H']
    fixed_dow = stats['fixed_dow']
    
    # ==================== VARIABLES ====================
    
    # assign[e, d] = 1 if employee e works on date d
    assign = {}
    for e in employees:
        for d in all_dates:
            assign[e.name, d] = model.NewBoolVar(f'assign_{e.name}_{d}')
    
    # allowed[e, dow] = 1 if employee e can work weekday dow (for pattern consistency)
    allowed = {}
    for e in employees:
        for dow in WEEKDAYS:
            allowed[e.name, dow] = model.NewBoolVar(f'allowed_{e.name}_{dow}')
    
    # ==================== HARD CONSTRAINTS ====================
    
    # 1. Exactly one employee per day
    for d in all_dates:
        model.AddExactlyOne(assign[e.name, d] for e in employees)
    
    # 2. Fixed assignments (recurring by day-of-week)
    # BUT skip if the fixed employee is on vacation
    for d in all_dates:
        dow = d.weekday()
        if dow in fixed_dow:
            emp_name = fixed_dow[dow]
            emp = config.get_employee_by_name(emp_name)
            if emp and emp.is_available(d):
                # Employee is available - enforce the fixed assignment
                model.Add(assign[emp_name, d] == 1)
            # If not available (vacation), the day is free for others
    
    # 3. Forbidden days and vacations - employee cannot work these days
    for e in employees:
        for d in all_dates:
            if not e.is_available(d):
                model.Add(assign[e.name, d] == 0)
    
    # Identify vacation replacement days (fixed DOW but fixed employee is on vacation)
    vacation_replacement_dates = set()
    for d in all_dates:
        dow = d.weekday()
        if dow in fixed_dow:
            fixed_emp = config.get_employee_by_name(fixed_dow[dow])
            if fixed_emp and not fixed_emp.is_available(d):
                vacation_replacement_dates.add(d)
    
    # 4. Pattern consistency: max 2 distinct weekdays per employee
    # Only applies to Mon-Thu (0-3), Friday (4) is exempt because of Fri-Sat link
    # EXCLUDE vacation replacement days from pattern consistency
    MON_THU = [0, 1, 2, 3]  # Monday=0, Tuesday=1, Wednesday=2, Thursday=3
    for e in employees:
        model.Add(sum(allowed[e.name, dow] for dow in MON_THU) <= 2)
        
        # If assigned to a Mon-Thu weekday, must have that day allowed
        # BUT skip vacation replacement days - they don't count towards pattern
        for d in weekday_dates:
            dow = d.weekday()
            if dow in MON_THU and d not in vacation_replacement_dates:
                model.AddImplication(assign[e.name, d], allowed[e.name, dow])
    
    # 5. Extra weekend employee constraint
    # The extra employee gets their H quota PLUS participates in remaining pool.
    # This constraint will be combined with fairness below.
    # (We remove the "exactly H" constraint - fairness will handle it)
    
    # 6. Friday -> Saturday link (optional)
    if config.link_friday_saturday:
        for fri in friday_dates:
            sat = fri + timedelta(days=1)
            if sat in all_dates:
                for e in employees:
                    # If e works Friday, e must work Saturday
                    model.AddImplication(assign[e.name, fri], assign[e.name, sat])
                    model.AddImplication(assign[e.name, sat], assign[e.name, fri])
    
    # 7. Sat/Sun balance: implemented below in fairness section with priority logic
    
    # ==================== FAIRNESS CONSTRAINTS ====================
    
    # Calculate pool statistics
    remaining_wd = stats['remaining_weekdays']
    remaining_we = stats['remaining_weekends_effective']
    total_remaining = remaining_wd + remaining_we
    
    # Per-employee targets
    wd_floor = remaining_wd // N if N > 0 else 0
    wd_ceil = (remaining_wd + N - 1) // N if N > 0 else 0
    we_floor = remaining_we // N if N > 0 else 0
    we_ceil = (remaining_we + N - 1) // N if N > 0 else 0
    total_floor = total_remaining // N if N > 0 else 0
    total_ceil = (total_remaining + N - 1) // N if N > 0 else 0
    
    # Classify employees
    restricted_sat_only = []
    restricted_sun_only = []
    unrestricted = []
    
    for e in employees:
        can_work_saturday = SATURDAY not in e.forbidden_weekdays
        can_work_sunday = SUNDAY not in e.forbidden_weekdays
        
        if can_work_saturday and not can_work_sunday:
            restricted_sat_only.append(e)
        elif can_work_sunday and not can_work_saturday:
            restricted_sun_only.append(e)
        elif can_work_saturday and can_work_sunday:
            unrestricted.append(e)
    
    # Create variables and constraints for all employees
    wd_vars = {}
    we_vars = {}
    total_vars = {}
    
    for e in employees:
        wd_total = sum(assign[e.name, d] for d in weekday_dates)
        we_total = sum(assign[e.name, d] for d in weekend_dates)
        sat_count = sum(assign[e.name, d] for d in saturday_dates)
        sun_count = sum(assign[e.name, d] for d in sunday_dates)
        
        # QUOTA-BASED: Fixed quota is always H (weeks), not actual worked days
        # If employee has fixed weekday assignment, their quota is H
        # If they miss days due to vacation, they compensate from pool
        has_fixed_weekday = any(
            fixed_dow.get(dow) == e.name 
            for dow in WEEKDAYS
        )
        has_fixed_weekend = any(
            fixed_dow.get(dow) == e.name 
            for dow in WEEKEND_DAYS
        )
        fixed_wd_quota = H if has_fixed_weekday else 0
        fixed_we_quota = H if has_fixed_weekend else 0
        extra_we_quota = H if e.is_extra_weekend else 0
        
        # Variable counts: total worked - full quota (not actual fixed days worked)
        wd_var = model.NewIntVar(0, len(weekday_dates), f'wd_var_{e.name}')
        we_var = model.NewIntVar(0, len(weekend_dates), f'we_var_{e.name}')
        total_var = model.NewIntVar(0, len(all_dates), f'total_var_{e.name}')
        
        model.Add(wd_var == wd_total - fixed_wd_quota)
        model.Add(we_var == we_total - fixed_we_quota - extra_we_quota)
        model.Add(total_var == wd_var + we_var)
        
        wd_vars[e.name] = wd_var
        we_vars[e.name] = we_var
        total_vars[e.name] = total_var
        
        # BOUNDS: Ensure fair distribution with some flexibility
        # Weekday bounds
        model.Add(wd_var >= max(0, wd_floor - 1))
        model.Add(wd_var <= wd_ceil + tolerance + 1)
        
        # Weekend bounds - CRITICAL: everyone must get at least we_floor - 1 weekends
        model.Add(we_var >= max(0, we_floor - 1))
        model.Add(we_var <= we_ceil + tolerance + 1)
        
        # Sat/Sun restrictions based on employee type
        if e in restricted_sat_only:
            model.Add(sun_count == 0)
            # Restricted employee gets exactly we_ceil (max) weekends from Saturdays
            model.Add(we_var == we_ceil)
        elif e in restricted_sun_only:
            model.Add(sat_count == 0)
            # Restricted employee gets exactly we_ceil (max) weekends from Sundays
            model.Add(we_var == we_ceil)
        else:
            # Unrestricted: balance Sat/Sun (soft - allow tolerance)
            model.Add(sat_count - sun_count <= 1 + tolerance)
            model.Add(sun_count - sat_count <= 1 + tolerance)
    
    # SPREAD CONSTRAINTS (using tolerance)
    
    # 1. TOTAL spread
    total_var_list = list(total_vars.values())
    total_max_var = model.NewIntVar(0, len(all_dates), 'total_max')
    total_min_var = model.NewIntVar(0, len(all_dates), 'total_min')
    model.AddMaxEquality(total_max_var, total_var_list)
    model.AddMinEquality(total_min_var, total_var_list)
    model.Add(total_max_var - total_min_var <= tolerance)
    
    # 2. Weekday spread
    wd_var_list = list(wd_vars.values())
    wd_max_var = model.NewIntVar(0, len(weekday_dates), 'wd_max')
    wd_min_var = model.NewIntVar(0, len(weekday_dates), 'wd_min')
    model.AddMaxEquality(wd_max_var, wd_var_list)
    model.AddMinEquality(wd_min_var, wd_var_list)
    model.Add(wd_max_var - wd_min_var <= tolerance)
    
    # 3. Weekend spread
    we_var_list = list(we_vars.values())
    we_max_var = model.NewIntVar(0, len(weekend_dates), 'we_max')
    we_min_var = model.NewIntVar(0, len(weekend_dates), 'we_min')
    model.AddMaxEquality(we_max_var, we_var_list)
    model.AddMinEquality(we_min_var, we_var_list)
    model.Add(we_max_var - we_min_var <= tolerance)
    
    # ==================== SOFT OBJECTIVES ====================
    
    objectives = []
    
    # PRIORITY 1: Minimize the actual spread (max - min) - HIGH WEIGHT
    # This encourages solver to find tightest possible distribution
    wd_spread = model.NewIntVar(0, len(weekday_dates), 'wd_spread')
    model.Add(wd_spread == wd_max_var - wd_min_var)
    objectives.append(wd_spread * 1000)  # Very high weight
    
    we_spread = model.NewIntVar(0, len(weekend_dates), 'we_spread')
    model.Add(we_spread == we_max_var - we_min_var)
    objectives.append(we_spread * 1000)  # Very high weight
    
    total_spread = model.NewIntVar(0, len(all_dates), 'total_spread')
    model.Add(total_spread == total_max_var - total_min_var)
    objectives.append(total_spread * 1000)  # Very high weight
    
    # 1. Minimize fairness deviation (weekday)
    for e in employees:
        wd_total = sum(assign[e.name, d] for d in weekday_dates)
        fixed_wd = stats['fixed_weekdays_per_emp'].get(e.name, 0)
        
        # |N * wd_var - remaining_wd| is hard; use linear proxy
        dev = model.NewIntVar(0, len(weekday_dates) * N, f'wd_dev_{e.name}')
        wd_scaled = model.NewIntVar(-len(weekday_dates) * N, len(weekday_dates) * N, f'wd_scaled_{e.name}')
        model.Add(wd_scaled == N * (wd_total - fixed_wd) - remaining_wd)
        model.AddAbsEquality(dev, wd_scaled)
        objectives.append(dev * 10)  # Weight
    
    # 2. Minimize fairness deviation (weekend)
    for e in employees:
        we_total = sum(assign[e.name, d] for d in weekend_dates)
        fixed_we = stats['fixed_weekends_per_emp'].get(e.name, 0)
        extra_taken = H if e.is_extra_weekend else 0
        
        dev = model.NewIntVar(0, len(weekend_dates) * N, f'we_dev_{e.name}')
        we_scaled = model.NewIntVar(-len(weekend_dates) * N, len(weekend_dates) * N, f'we_scaled_{e.name}')
        model.Add(we_scaled == N * (we_total - fixed_we - extra_taken) - remaining_we)
        model.AddAbsEquality(dev, we_scaled)
        objectives.append(dev * 10)
    
    # 3. Balance Saturday vs Sunday per employee (soft)
    for e in employees:
        sat_count = sum(assign[e.name, d] for d in saturday_dates)
        sun_count = sum(assign[e.name, d] for d in sunday_dates)
        
        sat_sun_diff = model.NewIntVar(-len(weekend_dates), len(weekend_dates), f'sat_sun_diff_{e.name}')
        model.Add(sat_sun_diff == sat_count - sun_count)
        
        sat_sun_abs = model.NewIntVar(0, len(weekend_dates), f'sat_sun_abs_{e.name}')
        model.AddAbsEquality(sat_sun_abs, sat_sun_diff)
        objectives.append(sat_sun_abs * 5)
    
    # 4. Penalize consecutive days
    sorted_dates = sorted(all_dates)
    for e in employees:
        for i in range(len(sorted_dates) - 1):
            d1, d2 = sorted_dates[i], sorted_dates[i + 1]
            if (d2 - d1).days == 1:
                consec = model.NewBoolVar(f'consec_{e.name}_{d1}')
                model.AddBoolAnd([assign[e.name, d1], assign[e.name, d2]]).OnlyEnforceIf(consec)
                model.AddBoolOr([assign[e.name, d1].Not(), assign[e.name, d2].Not()]).OnlyEnforceIf(consec.Not())
                objectives.append(consec * 100)
    
    # 5. Weekend spacing for non-extra employees
    # For each L-week window, prefer at most 1 weekend duty
    L = max(4, min(10, N - 1)) if N > 1 else 4
    for e in employees:
        if e.is_extra_weekend:
            continue
        # Group weekends by ISO week
        week_weekends = {}
        for d in weekend_dates:
            iso_week = d.isocalendar()[:2]
            if iso_week not in week_weekends:
                week_weekends[iso_week] = []
            week_weekends[iso_week].append(d)
        
        sorted_weeks = sorted(week_weekends.keys())
        for i in range(len(sorted_weeks)):
            window_weeks = sorted_weeks[i:i + L]
            if len(window_weeks) < 2:
                continue
            window_dates = []
            for w in window_weeks:
                window_dates.extend(week_weekends[w])
            
            # Count weekends in window
            weekend_count = sum(assign[e.name, d] for d in window_dates)
            violation = model.NewIntVar(0, len(window_dates), f'spacing_viol_{e.name}_{i}')
            model.Add(violation >= weekend_count - 1)
            objectives.append(violation * 50)
    
    # Set objective
    model.Minimize(sum(objectives))
    
    # ==================== SOLVE ====================
    
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit
    solver.parameters.num_search_workers = 4
    
    status = solver.Solve(model)
    
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        # Extract solution
        assignments = {}
        for d in all_dates:
            for e in employees:
                if solver.Value(assign[e.name, d]) == 1:
                    assignments[d] = e.name
                    break
        
        # Compute employee stats
        employee_stats = {}
        for e in employees:
            s = EmployeeStats(name=e.name)
            for d, emp_name in assignments.items():
                if emp_name == e.name:
                    s.total_duties += 1
                    dow = d.weekday()
                    if dow in WEEKDAYS:
                        s.weekday_duties += 1
                    if dow in WEEKEND_DAYS:
                        s.weekend_duties += 1
                    if dow == SATURDAY:
                        s.saturday_count += 1
                    if dow == SUNDAY:
                        s.sunday_count += 1
                    if dow == FRIDAY:
                        s.friday_count += 1
            
            s.fixed_weekdays = stats['fixed_weekdays_per_emp'].get(e.name, 0)
            s.fixed_weekends = stats['fixed_weekends_per_emp'].get(e.name, 0)
            s.variable_weekdays = s.weekday_duties - s.fixed_weekdays
            s.variable_weekends = s.weekend_duties - s.fixed_weekends
            if e.is_extra_weekend:
                s.variable_weekends -= H
            
            employee_stats[e.name] = s
        
        return ScheduleResult(
            success=True,
            assignments=assignments,
            employee_stats=employee_stats,
            tolerance_used=tolerance,
            total_weekdays=stats['total_weekdays'],
            total_weekends=stats['total_weekends'],
            remaining_weekdays=stats['remaining_weekdays'],
            remaining_weekends=stats['remaining_weekends_effective']
        )
    
    return ScheduleResult(
        success=False,
        error_message=f"No solution found with tolerance {tolerance}.",
        tolerance_used=tolerance
    )
