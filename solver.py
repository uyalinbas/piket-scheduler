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
    friday_dates = config.get_friday_dates()
    H = config.get_num_weeks()
    N = len(config.employees)
    
    # When Friday-Saturday link is active, Fridays are NOT part of the variable weekday pool
    # They are automatically distributed with Saturdays
    effective_weekday_dates = weekday_dates
    if config.link_friday_saturday:
        effective_weekday_dates = [d for d in weekday_dates if d.weekday() != FRIDAY]
    
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
            # When link is active, Friday is NOT a fixed weekday (it's tied to Saturday)
            if dow in WEEKDAYS and not (dow == FRIDAY and config.link_friday_saturday):
                fixed_weekdays_per_emp[emp_name] = H  # Full quota
            if dow in WEEKEND_DAYS:
                fixed_weekends_per_emp[emp_name] = H  # Full quota
    
    total_fixed_weekdays = sum(fixed_weekdays_per_emp.values())
    total_fixed_weekends = sum(fixed_weekends_per_emp.values())
    
    remaining_weekdays = len(effective_weekday_dates) - total_fixed_weekdays
    remaining_weekends = len(weekend_dates) - total_fixed_weekends
    
    # Account for extra weekend employee
    extra_emp = config.get_extra_weekend_employee()
    extra_weekend_quota = H if extra_emp else 0
    remaining_weekends_effective = remaining_weekends - extra_weekend_quota
    
    return {
        'H': H,
        'N': N,
        'total_weekdays': len(effective_weekday_dates),  # Excludes Fridays when link active
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
        'link_friday_saturday': config.link_friday_saturday,
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
    
    # SIMPLE STRATEGY: Try tolerances from 1 to max
    # Give more time for complex scenarios (extra weekend + restrictions)
    time_per_tolerance = max(30, time_limit_seconds // max_tolerance)
    for tolerance in range(1, max_tolerance + 1):
        result = _solve_with_tolerance(config, stats, tolerance, time_per_tolerance)
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
    # EXCLUDE vacation replacement days AND fixed days from pattern consistency
    MON_THU = [0, 1, 2, 3]  # Monday=0, Tuesday=1, Wednesday=2, Thursday=3
    
    # Identify fixed Mon-Thu days per employee (these don't count towards 2-day limit)
    emp_fixed_dow = {e.name: set() for e in employees}
    for dow, emp_name in fixed_dow.items():
        if dow in MON_THU and emp_name in emp_fixed_dow:
            emp_fixed_dow[emp_name].add(dow)
    
    for e in employees:
        # Pattern limit: 2 different Mon-Thu days from POOL only (excluding fixed days)
        pool_days = [dow for dow in MON_THU if dow not in emp_fixed_dow[e.name]]
        if pool_days:
            model.Add(sum(allowed[e.name, dow] for dow in pool_days) <= 2)
        
        # If assigned to a Mon-Thu weekday, must have that day allowed
        # BUT skip vacation replacement days AND fixed days - they don't count towards pattern
        for d in weekday_dates:
            dow = d.weekday()
            if dow in MON_THU and d not in vacation_replacement_dates:
                # Skip if this is a fixed day for this employee
                if dow in emp_fixed_dow[e.name]:
                    continue
                model.AddImplication(assign[e.name, d], allowed[e.name, dow])
    
    # 5. Extra weekend employee constraint
    # The extra employee must work at least H weekends (their quota)
    for e in employees:
        if e.is_extra_weekend:
            we_total = sum(assign[e.name, d] for d in weekend_dates)
            model.Add(we_total >= H)  # Must work at least H weekends
    
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
    
    # NOTE: Extra WE weekday fairness is handled differently:
    # - When link is active: Sat = Fri (linked), so more Saturdays = more Fridays = more weekdays
    # - Weekend spread constraint (includes extra WE) ensures equal WE pool share
    # - Extra WE pool WD: try to match others (SOFT), but cannot exceed (HARD)
    
    # Extra WE weekday constraint:
    # - Pool WD = wd_total - fixed_wd - linked_fri (Mon-Thu only when link active)
    # - Pool WD <= wd_ceil (cannot exceed others) - HARD
    # - Pool WD should match wd_floor (try to be equal) - SOFT (maximize)
    extra_we_wd_deviation = []  # Track deviation from target for soft objective
    
    for e in employees:
        if e.is_extra_weekend:
            wd_total = sum(assign[e.name, d] for d in weekday_dates)
            fixed_wd = stats['fixed_weekdays_per_emp'].get(e.name, 0)
            
            # Pool WD = weekdays - fixed (includes Fridays for everyone)
            # This ensures Ritesh's total weekdays match others
            pool_wd = wd_total - fixed_wd
            
            # Calculate wd_floor including Fridays
            # Total weekday pool = len(weekday_dates) - total fixed
            total_fixed_wd = sum(stats['fixed_weekdays_per_emp'].values())
            wd_pool_with_fri = len(weekday_dates) - total_fixed_wd
            wd_floor_with_fri = wd_pool_with_fri // N if N > 0 else 0
            
            pool_wd_var = model.NewIntVar(0, len(weekday_dates), f'pool_wd_{e.name}')
            model.Add(pool_wd_var == pool_wd)
            
            # HARD: Absolute cap - cannot exceed floor + 1 (ensures feasibility while limiting)
            model.Add(pool_wd <= wd_floor_with_fri + 1)
            
            # SOFT: Try to match wd_floor_with_fri exactly
            # Deviation = how far from target (negative = below, positive = above)
            deviation = model.NewIntVar(-len(weekday_dates), len(weekday_dates), f'wd_dev_{e.name}')
            model.Add(deviation == pool_wd_var - wd_floor_with_fri)
            
            # Penalize going ABOVE floor EXTREMELY heavily (should not exceed others)
            above_floor = model.NewIntVar(0, len(weekday_dates), f'above_floor_{e.name}')
            model.AddMaxEquality(above_floor, [deviation, model.NewConstant(0)])
            extra_we_wd_deviation.append(above_floor * 10000)  # EXTREME penalty for exceeding
            
            # Also penalize going below (try to maximize up to floor)
            below_floor = model.NewIntVar(0, len(weekday_dates), f'below_floor_{e.name}')
            neg_dev = model.NewIntVar(-len(weekday_dates), 0, f'neg_dev_{e.name}')
            model.Add(neg_dev == -deviation)
            model.AddMaxEquality(below_floor, [neg_dev, model.NewConstant(0)])
            extra_we_wd_deviation.append(below_floor * 10)  # Low penalty for being below
    
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
    actual_total_counts = {}  # For total spread: raw wd_total + we_total
    sat_counts = {}
    sun_counts = {}
    
    for e in employees:
        wd_total = sum(assign[e.name, d] for d in weekday_dates)
        we_total = sum(assign[e.name, d] for d in weekend_dates)
        sat_count = sum(assign[e.name, d] for d in saturday_dates)
        sun_count = sum(assign[e.name, d] for d in sunday_dates)
        
        # Store actual total for total spread constraint (raw counts, no adjustments)
        actual_total_counts[e.name] = wd_total + we_total
        
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
        
        # When Friday→Saturday link is active, Fridays are not "free" weekdays
        # They are tied to Saturday assignments, so we exclude them from wd_var
        # This applies to ALL employees, not just extra weekend
        friday_total = sum(assign[e.name, d] for d in friday_dates)
        linked_friday_adjustment = friday_total if config.link_friday_saturday else 0
        
        fixed_we_quota = H if has_fixed_weekend else 0
        # Extra WE quota: H extra weekend assignments on top of normal pool share
        extra_we_quota = H if e.is_extra_weekend else 0
        
        # Variable counts: total worked - full quota (not actual fixed days worked)
        # When link is active, subtract Friday count since they're tied to Saturdays
        wd_var = model.NewIntVar(-len(weekday_dates), len(weekday_dates), f'wd_var_{e.name}')
        we_var = model.NewIntVar(-len(weekend_dates), len(weekend_dates), f'we_var_{e.name}')
        total_var = model.NewIntVar(-len(all_dates), len(all_dates), f'total_var_{e.name}')
        
        model.Add(wd_var == wd_total - fixed_wd_quota - linked_friday_adjustment)
        model.Add(we_var == we_total - fixed_we_quota - extra_we_quota)
        model.Add(total_var == wd_var + we_var)
        
        wd_vars[e.name] = wd_var
        we_vars[e.name] = we_var
        total_vars[e.name] = total_var
        sat_counts[e.name] = sat_count
        sun_counts[e.name] = sun_count
        
        # BOUNDS: Ensure fair distribution with some flexibility
        # Weekday bounds
        # For extra WE with link active: Mon-Thu can be 0 if Fridays fill quota
        wd_min = 0 if (e.is_extra_weekend and config.link_friday_saturday) else max(0, wd_floor - 1)
        model.Add(wd_var >= wd_min)
        model.Add(wd_var <= wd_ceil + tolerance + 1)
        
        # Weekend bounds - CRITICAL: everyone must get at least we_floor - 1 weekends
        model.Add(we_var >= max(0, we_floor - 1))
        model.Add(we_var <= we_ceil + tolerance + 1)
        
        # Sat/Sun restrictions based on employee type
        if e in restricted_sat_only:
            model.Add(sun_count == 0)
        elif e in restricted_sun_only:
            model.Add(sat_count == 0)
        elif e.is_extra_weekend:
            # Extra WE (Ritesh): No Sat/Sun balance constraint
            # This allows more flexibility to minimize Saturdays (and linked Fridays)
            pass
        # No per-person Sat/Sun balance for unrestricted - use spread instead
    
    # SATURDAY SPREAD (excluding restricted_sat_only AND extra_weekend employees)
    # Extra WE employees (Ritesh) have higher counts, exclude from spread
    pool_sat_employees = [e for e in employees 
                          if e not in restricted_sat_only and not e.is_extra_weekend]
    non_sat_only_sat_counts = [sat_counts[e.name] for e in pool_sat_employees]
    if non_sat_only_sat_counts:
        sat_max = model.NewIntVar(0, len(saturday_dates), 'pool_sat_max')
        sat_min = model.NewIntVar(0, len(saturday_dates), 'pool_sat_min')
        model.AddMaxEquality(sat_max, non_sat_only_sat_counts)
        model.AddMinEquality(sat_min, non_sat_only_sat_counts)
        model.Add(sat_max - sat_min <= 1)  # HARD: spread 1
    
    # SUNDAY SPREAD (excluding restricted_sun_only, restricted_sat_only, AND extra_weekend)
    pool_sun_employees = [e for e in employees 
                          if e not in restricted_sun_only 
                          and e not in restricted_sat_only 
                          and not e.is_extra_weekend]
    non_sun_only_sun_counts = [sun_counts[e.name] for e in pool_sun_employees]
    if non_sun_only_sun_counts:
        sun_max = model.NewIntVar(0, len(sunday_dates), 'pool_sun_max')
        sun_min = model.NewIntVar(0, len(sunday_dates), 'pool_sun_min')
        model.AddMaxEquality(sun_max, non_sun_only_sun_counts)
        model.AddMinEquality(sun_min, non_sun_only_sun_counts)
        model.Add(sun_max - sun_min <= 1)  # HARD: spread 1
    
    # SPREAD CONSTRAINTS - ALL HARD RULES: max spread 1
    # Exclude extra WE employees (their we_var=0 vs others we_var≈we_floor makes spread impossible)
    pool_employees = [e for e in employees if not e.is_extra_weekend]
    
    # 1. TOTAL spread via ANTI-CORRELATION (excluding extra WE)
    # Instead of a hard max-min constraint, we use anti-correlation:
    # If an employee gets ceil in WD pool, they must get floor in WE pool.
    # This mathematically guarantees Total spread <= 1 when combined with
    # individual WD/WE spread constraints.
    
    # Calculate actual pool statistics for anti-correlation
    # For employees without extra WE: wd_actual = wd_total (raw weekday count)
    # we_actual = we_total (raw weekend count)
    
    # Pool parameters (using actual assignment counts, not vars)
    N_pool = len(pool_employees)
    
    if N_pool > 0:
        # Calculate what the actual weekday pool looks like for pool employees
        # When link is active and extra WE exists, their Fridays are committed to Saturdays
        # So we subtract extra WE's Friday quota (wd_ceil) from the pool
        extra_we_friday_quota = wd_ceil if (config.link_friday_saturday and config.get_extra_weekend_employee()) else 0
        actual_wd_pool = len(weekday_dates) - stats['total_fixed_weekdays'] - extra_we_friday_quota
        actual_we_pool = len(weekend_dates) - stats['total_fixed_weekends'] - stats['extra_weekend_quota']
        
        actual_wd_floor = actual_wd_pool // N_pool
        actual_wd_ceil = (actual_wd_pool + N_pool - 1) // N_pool
        actual_we_floor = actual_we_pool // N_pool
        actual_we_ceil = (actual_we_pool + N_pool - 1) // N_pool
        
        wd_has_remainder = (actual_wd_pool % N_pool) > 0
        we_has_remainder = (actual_we_pool % N_pool) > 0
        
        # Anti-correlation: if both pools have remainders, prevent getting ceil in both
        if wd_has_remainder and we_has_remainder:
            for e in pool_employees:
                # Get actual counts (not adjusted for linked Fridays etc)
                actual_wd = sum(assign[e.name, d] for d in weekday_dates)
                fixed_wd_quota = stats['fixed_weekdays_per_emp'].get(e.name, 0)
                actual_wd_var = actual_wd - fixed_wd_quota  # Actual variable weekdays
                
                actual_we = sum(assign[e.name, d] for d in weekend_dates)
                fixed_we_quota = stats['fixed_weekends_per_emp'].get(e.name, 0)
                actual_we_var = actual_we - fixed_we_quota  # Actual variable weekends
                
                # Boolean: is this employee at WD ceil?
                is_wd_high = model.NewBoolVar(f'is_wd_high_{e.name}')
                model.Add(actual_wd_var > actual_wd_floor).OnlyEnforceIf(is_wd_high)
                model.Add(actual_wd_var <= actual_wd_floor).OnlyEnforceIf(is_wd_high.Not())
                
                # If WD is high (ceil), then WE must be low (floor)
                model.Add(actual_we_var <= actual_we_floor).OnlyEnforceIf(is_wd_high)
    
    # 2. Weekday spread - HARD (pool employees only)
    # Extra WE is handled by total pool share constraint
    actual_wd_totals = []
    for e in pool_employees:  # Exclude extra WE
        actual_wd = sum(assign[e.name, d] for d in weekday_dates)
        fixed_wd = stats['fixed_weekdays_per_emp'].get(e.name, 0)
        actual_wd_var = model.NewIntVar(0, len(weekday_dates), f'actual_wd_var_{e.name}')
        model.Add(actual_wd_var == actual_wd - fixed_wd)
        actual_wd_totals.append(actual_wd_var)
    
    if actual_wd_totals:
        wd_max_var = model.NewIntVar(0, len(weekday_dates), 'wd_max')
        wd_min_var = model.NewIntVar(0, len(weekday_dates), 'wd_min')
        model.AddMaxEquality(wd_max_var, actual_wd_totals)
        model.AddMinEquality(wd_min_var, actual_wd_totals)
        model.Add(wd_max_var - wd_min_var <= 1)  # HARD: max spread 1
    
    # 3. Weekend spread - HARD (ALL employees, VarWE = we - fixed - H for extra WE)
    actual_we_totals = []
    for e in employees:  # ALL employees including extra WE
        actual_we = sum(assign[e.name, d] for d in weekend_dates)
        fixed_we = stats['fixed_weekends_per_emp'].get(e.name, 0)
        extra_we = H if e.is_extra_weekend else 0  # Subtract H for extra WE
        actual_we_var = model.NewIntVar(0, len(weekend_dates), f'actual_we_var_{e.name}')
        model.Add(actual_we_var == actual_we - fixed_we - extra_we)
        actual_we_totals.append(actual_we_var)
    
    if actual_we_totals:
        we_max_var = model.NewIntVar(0, len(weekend_dates), 'we_max')
        we_min_var = model.NewIntVar(0, len(weekend_dates), 'we_min')
        model.AddMaxEquality(we_max_var, actual_we_totals)
        model.AddMinEquality(we_min_var, actual_we_totals)
        model.Add(we_max_var - we_min_var <= 1)  # HARD: max spread 1
    
    # NOTE: Total fairness is ensured by WD spread (pool) + WE spread (all with H subtracted)
    # WD + WE spread together provide balanced distribution
    
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
    
    # PRIORITY 2: Minimize extra WE weekday deviation (try to match others' pool WD)
    for dev in extra_we_wd_deviation:
        objectives.append(dev * 500)  # High weight - maximize extra WE pool WD
    
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
            
            # When link is active, linked Fridays are tied to Saturdays
            # So subtract Friday count from variable weekdays for ALL employees
            if config.link_friday_saturday:
                s.variable_weekdays -= s.friday_count
            
            # Extra WE adjustments
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
