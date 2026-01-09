"""
Validation and sanity checks for schedule configuration.
Provides pre-solve feasibility analysis with clear error messages.
"""
from typing import List, Tuple, Optional
from models import ScheduleConfig, WEEKDAYS, WEEKEND_DAYS, WEEKDAY_NAMES


def validate_config(config: ScheduleConfig) -> Tuple[bool, List[str]]:
    """
    Validate schedule configuration before solving.
    Returns (is_valid, list_of_error_messages).
    """
    errors = []
    
    # 1. Check employees exist
    if not config.employees:
        errors.append("No employees defined. Add at least one employee.")
        return False, errors
    
    N = len(config.employees)
    
    # 2. Check date range
    all_dates = config.get_all_dates()
    if not all_dates:
        errors.append("Invalid date range. Start date must be before or equal to end date.")
        return False, errors
    
    # 3. Check for duplicate employee names
    names = [e.name for e in config.employees]
    if len(names) != len(set(names)):
        errors.append("Duplicate employee names found. Each employee must have a unique name.")
        return False, errors
    
    # 4. Check at most one extra weekend employee
    extra_count = sum(1 for e in config.employees if e.is_extra_weekend)
    if extra_count > 1:
        errors.append(f"At most one employee can be marked as extra weekend. Found {extra_count}.")
        return False, errors
    
    # 5. Check fixed assignments reference valid employees
    valid_names = set(names)
    for fa in config.fixed_assignments:
        if fa.employee_name not in valid_names:
            errors.append(f"Fixed assignment for {WEEKDAY_NAMES[fa.day_of_week]} references unknown employee '{fa.employee_name}'.")
    
    # 6. Check for conflicting fixed assignments (multiple employees for same day)
    fixed_days = {}
    for fa in config.fixed_assignments:
        if fa.day_of_week in fixed_days:
            errors.append(
                f"Conflicting fixed assignments: both '{fixed_days[fa.day_of_week]}' and '{fa.employee_name}' "
                f"assigned to {WEEKDAY_NAMES[fa.day_of_week]}."
            )
        fixed_days[fa.day_of_week] = fa.employee_name
    
    if errors:
        return False, errors
    
    # 7. Pool feasibility checks
    weekday_dates = config.get_weekday_dates()
    weekend_dates = config.get_weekend_dates()
    H = config.get_num_weeks()
    
    # Count fixed days
    total_fixed_weekdays = sum(1 for d in weekday_dates if d.weekday() in fixed_days)
    total_fixed_weekends = sum(1 for d in weekend_dates if d.weekday() in fixed_days)
    
    remaining_weekdays = len(weekday_dates) - total_fixed_weekdays
    remaining_weekends = len(weekend_dates) - total_fixed_weekends
    
    # Extra weekend quota
    extra_emp = config.get_extra_weekend_employee()
    extra_quota = H if extra_emp else 0
    remaining_weekends_effective = remaining_weekends - extra_quota
    
    if remaining_weekdays < 0:
        errors.append(
            f"WEEKDAY POOL EXHAUSTED: Fixed assignments ({total_fixed_weekdays}) "
            f"exceed available weekdays ({len(weekday_dates)})."
        )
    
    if remaining_weekends_effective < 0:
        errors.append(
            f"WEEKEND POOL EXHAUSTED: Fixed ({total_fixed_weekends}) + Extra quota ({extra_quota}) "
            f"exceed available weekend days ({len(weekend_dates)})."
        )
    
    # 8. Check availability per pool
    # Count available days per employee for each pool
    weekday_availability = {e.name: 0 for e in config.employees}
    weekend_availability = {e.name: 0 for e in config.employees}
    
    for e in config.employees:
        for d in weekday_dates:
            dow = d.weekday()
            # Skip fixed days assigned to others
            if dow in fixed_days and fixed_days[dow] != e.name:
                continue
            if e.is_available(d):
                weekday_availability[e.name] += 1
        
        for d in weekend_dates:
            dow = d.weekday()
            if dow in fixed_days and fixed_days[dow] != e.name:
                continue
            if e.is_available(d):
                weekend_availability[e.name] += 1
    
    # Check if any employee has zero availability in a pool where they're needed
    for e in config.employees:
        if remaining_weekdays > 0 and weekday_availability[e.name] == 0:
            # Check if they have fixed weekdays
            has_fixed_wd = any(
                fa.employee_name == e.name and fa.day_of_week in WEEKDAYS
                for fa in config.fixed_assignments
            )
            if not has_fixed_wd:
                errors.append(
                    f"Employee '{e.name}' has no available weekdays due to forbidden days/vacations. "
                    f"This may cause fairness issues."
                )
    
    # Extra weekend employee must have enough weekend availability
    if extra_emp:
        if weekend_availability[extra_emp.name] < extra_quota:
            errors.append(
                f"Extra weekend employee '{extra_emp.name}' needs {extra_quota} weekend days "
                f"but only {weekend_availability[extra_emp.name]} are available (due to forbidden days/vacations)."
            )
    
    # 9. Friday-Saturday link feasibility
    if config.link_friday_saturday:
        friday_dates = config.get_friday_dates()
        for fri in friday_dates:
            from datetime import timedelta
            sat = fri + timedelta(days=1)
            if sat not in all_dates:
                continue
            
            # Find who can work both days
            can_work_both = []
            for e in config.employees:
                fri_dow = fri.weekday()
                sat_dow = sat.weekday()
                
                fri_ok = e.is_available(fri) and (fri_dow not in fixed_days or fixed_days[fri_dow] == e.name)
                sat_ok = e.is_available(sat) and (sat_dow not in fixed_days or fixed_days[sat_dow] == e.name)
                
                if fri_ok and sat_ok:
                    can_work_both.append(e.name)
            
            if not can_work_both:
                errors.append(
                    f"Friday-Saturday link active but no employee can work both {fri} and {sat}."
                )
    
    return len(errors) == 0, errors


def compute_theoretical_bounds(config: ScheduleConfig) -> dict:
    """
    Compute theoretical min/max assignments per employee for each pool.
    Useful for debugging and UI display.
    """
    N = len(config.employees)
    if N == 0:
        return {}
    
    weekday_dates = config.get_weekday_dates()
    weekend_dates = config.get_weekend_dates()
    H = config.get_num_weeks()
    
    # Fixed day map
    fixed_dow = {}
    for fa in config.fixed_assignments:
        fixed_dow[fa.day_of_week] = fa.employee_name
    
    total_fixed_weekdays = sum(1 for d in weekday_dates if d.weekday() in fixed_dow)
    total_fixed_weekends = sum(1 for d in weekend_dates if d.weekday() in fixed_dow)
    
    remaining_weekdays = len(weekday_dates) - total_fixed_weekdays
    remaining_weekends = len(weekend_dates) - total_fixed_weekends
    
    extra_emp = config.get_extra_weekend_employee()
    extra_quota = H if extra_emp else 0
    remaining_weekends_effective = remaining_weekends - extra_quota
    
    return {
        'total_weekdays': len(weekday_dates),
        'total_weekends': len(weekend_dates),
        'fixed_weekdays': total_fixed_weekdays,
        'fixed_weekends': total_fixed_weekends,
        'remaining_weekdays': remaining_weekdays,
        'remaining_weekends_effective': remaining_weekends_effective,
        'weekday_floor': remaining_weekdays // N,
        'weekday_ceil': (remaining_weekdays + N - 1) // N if remaining_weekdays > 0 else 0,
        'weekend_floor': remaining_weekends_effective // N,
        'weekend_ceil': (remaining_weekends_effective + N - 1) // N if remaining_weekends_effective > 0 else 0,
        'extra_weekend_quota': extra_quota,
        'num_employees': N,
        'num_weeks': H,
    }
