"""
Data models for the roster scheduling system.
"""
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Dict, List, Optional, Set, Tuple
import calendar


# Day of week constants (Monday=0, Sunday=6)
MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY = range(7)
WEEKDAY_NAMES = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
WEEKDAYS = {MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY}  # Mon-Fri
WEEKEND_DAYS = {SATURDAY, SUNDAY}  # Sat+Sun only


@dataclass
class Employee:
    """Represents an employee with their constraints."""
    name: str
    forbidden_weekdays: Set[int] = field(default_factory=set)  # 0-6 (Mon-Sun)
    vacation_ranges: List[Tuple[date, date]] = field(default_factory=list)  # (start, end) inclusive
    is_extra_weekend: bool = False
    
    def is_available(self, d: date) -> bool:
        """Check if employee is available on a specific date."""
        dow = d.weekday()
        # Check forbidden weekday
        if dow in self.forbidden_weekdays:
            return False
        # Check vacations
        for vac_start, vac_end in self.vacation_ranges:
            if vac_start <= d <= vac_end:
                return False
        return True
    
    def to_dict(self) -> dict:
        """Convert to dictionary for serialization."""
        return {
            'name': self.name,
            'forbidden_weekdays': list(self.forbidden_weekdays),
            'vacation_ranges': [(s.isoformat(), e.isoformat()) for s, e in self.vacation_ranges],
            'is_extra_weekend': self.is_extra_weekend
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Employee':
        """Create from dictionary."""
        return cls(
            name=data['name'],
            forbidden_weekdays=set(data.get('forbidden_weekdays', [])),
            vacation_ranges=[
                (date.fromisoformat(s), date.fromisoformat(e))
                for s, e in data.get('vacation_ranges', [])
            ],
            is_extra_weekend=data.get('is_extra_weekend', False)
        )


@dataclass
class FixedAssignment:
    """A recurring fixed assignment: every week, this day goes to this employee."""
    day_of_week: int  # 0-6 (Mon-Sun)
    employee_name: str
    
    def to_dict(self) -> dict:
        return {'day_of_week': self.day_of_week, 'employee_name': self.employee_name}
    
    @classmethod
    def from_dict(cls, data: dict) -> 'FixedAssignment':
        return cls(day_of_week=data['day_of_week'], employee_name=data['employee_name'])


@dataclass
class ScheduleConfig:
    """Configuration for schedule generation."""
    start_date: date
    end_date: date
    employees: List[Employee]
    fixed_assignments: List[FixedAssignment] = field(default_factory=list)
    link_friday_saturday: bool = False
    
    def get_all_dates(self) -> List[date]:
        """Get all dates in the range."""
        dates = []
        d = self.start_date
        while d <= self.end_date:
            dates.append(d)
            d += timedelta(days=1)
        return dates
    
    def get_num_weeks(self) -> int:
        """Get number of ISO weeks covered (H)."""
        weeks = set()
        for d in self.get_all_dates():
            weeks.add(d.isocalendar()[:2])  # (year, week)
        return len(weeks)
    
    def get_weekday_dates(self) -> List[date]:
        """Get all Mon-Fri dates."""
        return [d for d in self.get_all_dates() if d.weekday() in WEEKDAYS]
    
    def get_weekend_dates(self) -> List[date]:
        """Get all Sat+Sun dates."""
        return [d for d in self.get_all_dates() if d.weekday() in WEEKEND_DAYS]
    
    def get_saturday_dates(self) -> List[date]:
        """Get all Saturday dates."""
        return [d for d in self.get_all_dates() if d.weekday() == SATURDAY]
    
    def get_sunday_dates(self) -> List[date]:
        """Get all Sunday dates."""
        return [d for d in self.get_all_dates() if d.weekday() == SUNDAY]
    
    def get_friday_dates(self) -> List[date]:
        """Get all Friday dates."""
        return [d for d in self.get_all_dates() if d.weekday() == FRIDAY]
    
    def get_employee_by_name(self, name: str) -> Optional[Employee]:
        """Find employee by name."""
        for e in self.employees:
            if e.name == name:
                return e
        return None
    
    def get_extra_weekend_employee(self) -> Optional[Employee]:
        """Get the extra weekend employee if any."""
        for e in self.employees:
            if e.is_extra_weekend:
                return e
        return None


@dataclass
class EmployeeStats:
    """Statistics for a single employee."""
    name: str
    total_duties: int = 0
    weekday_duties: int = 0  # Mon-Fri
    weekend_duties: int = 0  # Sat+Sun
    saturday_count: int = 0
    sunday_count: int = 0
    friday_count: int = 0
    fixed_weekdays: int = 0
    fixed_weekends: int = 0
    variable_weekdays: int = 0
    variable_weekends: int = 0


@dataclass
class ScheduleResult:
    """Result of schedule generation."""
    success: bool
    assignments: Dict[date, str] = field(default_factory=dict)  # date -> employee_name
    employee_stats: Dict[str, EmployeeStats] = field(default_factory=dict)
    tolerance_used: int = 1
    solve_time_seconds: float = 0.0
    error_message: str = ""
    
    # Pool info for debugging
    total_weekdays: int = 0
    total_weekends: int = 0
    remaining_weekdays: int = 0
    remaining_weekends: int = 0
    
    def get_fairness_spread(self, pool: str = 'weekday') -> Tuple[int, int, int]:
        """Get (min, max, spread) for variable assignments in a pool."""
        if not self.employee_stats:
            return (0, 0, 0)
        
        if pool == 'weekday':
            values = [s.variable_weekdays for s in self.employee_stats.values()]
        else:  # weekend
            values = [s.variable_weekends for s in self.employee_stats.values()]
        
        if not values:
            return (0, 0, 0)
        return (min(values), max(values), max(values) - min(values))


def get_iso_week_dates(year: int, week: int) -> Tuple[date, date]:
    """Get the Monday and Sunday of a given ISO week."""
    # Find Jan 4 of the year (always in week 1)
    jan4 = date(year, 1, 4)
    # Find the Monday of week 1
    week1_monday = jan4 - timedelta(days=jan4.weekday())
    # Calculate the Monday of the target week
    target_monday = week1_monday + timedelta(weeks=week - 1)
    target_sunday = target_monday + timedelta(days=6)
    return target_monday, target_sunday


def get_date_range_from_weeks(year: int, start_week: int, end_week: int) -> Tuple[date, date]:
    """Get date range from ISO week range."""
    start_date, _ = get_iso_week_dates(year, start_week)
    _, end_date = get_iso_week_dates(year, end_week)
    return start_date, end_date
