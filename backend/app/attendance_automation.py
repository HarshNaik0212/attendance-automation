"""
Attendance Automation
======================
Reads the MAIN sheet of an attendance export and computes, per employee:
    - No. of Days Present
    - No. of Days Absent / LWP
    - No. of Days Sandwich Rule (Nopay)

Business rules implemented (confirmed with user):
  1. Only the day-columns actually present in the sheet are used as the
     "period" (works for any payroll cycle, not just a calendar month).
  2. Effective working window per employee:
       - Active: [max(DOJ, period_start), period_end]. Always.
       - Resigned:
           * ActualLWD blank, or falling in a month OTHER than the
             period's ending month (still serving notice past this
             cycle) -> unchanged, same as Active:
             [max(DOJ, period_start), period_end].
           * ActualLWD falls in the period's ENDING calendar month
             (e.g. August in a 26-Jul-25-Aug cycle) -> NARROW window:
             [max(DOJ, 1st-of-that-month), ActualLWD].
       - Exited:
           * DOJ == ActualLWD (joined and exited the exact same day) ->
             0 payable days for the whole period, no matter what.
           * ActualLWD falls in the period's ENDING calendar month ->
             same NARROW window as Resigned above:
             [max(DOJ, 1st-of-that-month), ActualLWD].
           * Otherwise (ActualLWD in an earlier month, e.g. already left
             during the period's starting-month tail) -> unchanged:
             [max(DOJ, period_start), min(ActualLWD, period_end)].
     When the NARROW window's end (ActualLWD) falls beyond this file's
     last tracked date (period_end), the untracked tail days (period_end+1
     through ActualLWD) are assumed PRESENT (the employee is still
     formally on payroll until ActualLWD and nothing suggests otherwise) -
     Present/Absent/Sandwich are all computed consistently over this same
     window.

     IMPORTANT: the window described above (working_window /
     is_narrow_month_window) is used ONLY for "No. of Days Present
     (Payable Days only)" and "Total No. of Payed Days". "No. of Days
     Absent/LWP" and "No. of Days Sandwich (Nopay)" use a SEPARATE, WIDE
     window (wide_window()) that always starts at max(DOJ, period_start)
     - i.e. it can include the 26-31 July tail even for Resigned/Exited
     employees whose LWD falls in August - and ends at ActualLWD (for
     Resigned/Exited) or period_end (for Active). This means for an
     employee who both joined in the 26-31 July tail AND
     resigned/exited in August, "Present(Payable Days only)" and
     "Absent/LWP + Sandwich" are computed over two DIFFERENT windows and
     will not necessarily sum to the same total day count - this is
     intentional, per confirmed requirements.
  3. Day status classification (within the working window). Matching is
     CASE-INSENSITIVE (the source data is inconsistent, e.g. "Paternity
     leave" vs "Paternity Leave") - see normalize_status().
       Paid   : Present, Weekly Off, Holiday, Casual Leave, Sick Leave,
                Privilege Leave, Maternity Leave, Paternity Leave,
                Menstrual Leave, Miss Punch Out, Half Day
       Unpaid : Absent, Leave without Pay
     Any non-blank status matching neither list is flagged as a warning
     at the end of the run instead of being silently dropped.
  4. Sandwich rule (per employee, based on their own weekly-off weekday(s)
     detected from their row):
       - Sat+Sun off employees : Friday & the following Monday both
         Absent/LWP  -> one sandwich INSTANCE, worth 4 nopay days
         (Fri+Sat+Sun+Mon). Saturday & Sunday (originally Weekly Off) flip
         to unpaid.
       - Sun-only off employees: Saturday & the following Monday both
         Absent/LWP  -> one sandwich INSTANCE, worth 3 nopay days
         (Sat+Sun+Mon). Sunday flips to unpaid.
       The Sandwich column reports "<instance count> (<total nopay days>
       days)", e.g. "2 (8 days)" if it triggered twice for a Sat+Sun-off
       employee, or "0" if it never triggered.
       Only the newly-flipped Weekly-Off day(s) (Sat+Sun, or just Sun) are
       removed from the Present count; Friday/Monday were already unpaid
       and stay counted once in the Absent/LWP column (not duplicated).
  5. Total No. of Payed Days (new joiner bonus):
       - Exited employees: always Present (no bonus, regardless of window).
       - Resigned employees using the NARROW window (point 2 above):
         always Present (no bonus - Present already reflects exactly the
         days owed).
       - Resigned employees NOT using the narrow window, and Active
         employees, with Present == 0: no bonus, Total = 0.
       - Resigned (non-narrow) / Active employees with Present > 0 and DOJ
         in the period's ending calendar month (e.g. Aug in a
         26-Jul-25-Aug cycle): bonus = days_in_that_month - 25 (fills the
         real gap between period-end and the actual calendar month-end,
         which this cycle's data doesn't cover).
       - Everyone else (DOJ in the period's starting month, e.g. the
         26-31 July tail, or an already-established employee): no bonus.
         Their Present already has zero gaps for the whole window, so
         nothing is owed on top.
  6. "No. of Days Present(Payable Days only)" column - normally identical
     to the Present value used everywhere else, EXCEPT for employees whose
     DOJ falls in the period's STARTING calendar month (e.g. the 26-31
     July tail): for THIS column only, they're recalculated as if they
     were an ending-month (August) joiner - window [1-Aug, period_end]
     with the sandwich rule applied within that sub-window, plus the same
     monthly bonus (+6 for a 31-day ending month, etc.) applied
     UNCONDITIONALLY (regardless of Active/Resigned/Exited status, unlike
     point 5's Total Paid Days bonus which is Active-only). Absent/LWP,
     Sandwich, and Total Paid Days are unaffected by this and keep using
     the normal wide window.
"""

import calendar
import datetime
import openpyxl
from openpyxl.utils import get_column_letter

PAID_STATUSES = {
    "Present", "Weekly Off", "Holiday", "Casual Leave", "Sick Leave",
    "Privilege Leave", "Maternity Leave", "Paternity Leave",
    "Menstrual Leave", "Miss Punch Out", "Half Day",
}
UNPAID_STATUSES = {"Absent", "Leave without Pay"}
WEEKLY_OFF_STATUS = "Weekly Off"

# Case-insensitive lookup sets. The source data isn't perfectly consistent
# about capitalization (e.g. "Paternity leave" vs "Paternity Leave"), so all
# status comparisons go through normalize_status() below rather than
# comparing raw strings directly.
_PAID_STATUSES_NORM = {s.lower() for s in PAID_STATUSES}
_UNPAID_STATUSES_NORM = {s.lower() for s in UNPAID_STATUSES}
_WEEKLY_OFF_NORM = WEEKLY_OFF_STATUS.lower()


def normalize_status(status):
    """Lowercase + strip a raw status cell value for reliable comparison.
    Returns None for blank/None cells."""
    if status is None:
        return None
    return str(status).strip().lower()


def is_paid_status(status):
    return normalize_status(status) in _PAID_STATUSES_NORM


def is_unpaid_status(status):
    return normalize_status(status) in _UNPAID_STATUSES_NORM


def is_weekly_off(status):
    return normalize_status(status) == _WEEKLY_OFF_NORM


SATURDAY, SUNDAY, MONDAY, FRIDAY = 5, 6, 0, 4


class Employee:
    """Wraps a single employee's row: identity fields + date->status map."""

    def __init__(self, code, name, status, doj, actual_lwd, extra_fields, day_status):
        self.code = code
        self.name = name
        self.status = status  # Active / Resigned / Exited
        self.doj = doj.date() if isinstance(doj, datetime.datetime) else doj
        self.actual_lwd = actual_lwd.date() if isinstance(actual_lwd, datetime.datetime) else actual_lwd
        self.extra_fields = extra_fields  # Grade, Designation, ... kept for passthrough
        self.day_status = day_status  # dict: date -> status string (or None)

    def weekly_off_weekdays(self):
        """Detect this employee's weekly-off weekday(s) from their own row."""
        offs = {d.weekday() for d, s in self.day_status.items() if is_weekly_off(s)}
        if offs == {SATURDAY, SUNDAY}:
            return frozenset({SATURDAY, SUNDAY})
        if offs == {SUNDAY}:
            return frozenset({SUNDAY})
        # Fallback: default to the common Sat+Sun pattern if we can't tell
        # (e.g. employee absent the whole window, no Weekly Off visible).
        return frozenset({SATURDAY, SUNDAY})


class AttendanceCalculator:
    """Applies the working-window, tenure, and sandwich rules to compute
    the three output metrics for one employee over the given period dates."""

    def __init__(self, period_dates):
        self.period_dates = period_dates  # sorted list of date objects
        self.period_start = period_dates[0]
        self.period_end = period_dates[-1]

    def working_window(self, emp: Employee):
        """Return (start, end, is_narrow_month_window) inclusive, or None if
        the employee gets 0 payable days for the whole period (short-tenure
        Exited rule).

        is_narrow_month_window=True marks the new special case: ActualLWD
        falls in the period's ending calendar month (e.g. August in a
        26-Jul-25-Aug cycle). For that case the window becomes
        [max(DOJ, 1st-of-that-month), ActualLWD] for BOTH Resigned and
        Exited employees, and all three attendance columns are recomputed
        over it - even extending past this file's last tracked date
        (period_end) if ActualLWD falls after it (those untracked tail
        days are assumed Present, since the employee is still formally on
        payroll until ActualLWD and nothing suggests otherwise).
        """
        month_start = datetime.date(self.period_end.year, self.period_end.month, 1)
        lwd_in_ending_month = (
            emp.actual_lwd is not None
            and emp.actual_lwd.year == self.period_end.year
            and emp.actual_lwd.month == self.period_end.month
        )

        if emp.status == "Exited":
            if emp.doj and emp.actual_lwd and emp.doj == emp.actual_lwd:
                return None  # joined & exited on the exact same day -> no pay at all
            if lwd_in_ending_month:
                start = max(emp.doj, month_start) if emp.doj else month_start
                end = emp.actual_lwd
                if start > end:
                    return None
                return start, end, True
            # Unchanged original behavior: LWD not in the ending month.
            start = max(emp.doj, self.period_start) if emp.doj else self.period_start
            end = min(emp.actual_lwd, self.period_end) if emp.actual_lwd else self.period_end
            if start > end:
                return None
            return start, end, False

        elif emp.status == "Resigned":
            if lwd_in_ending_month:
                start = max(emp.doj, month_start) if emp.doj else month_start
                end = emp.actual_lwd
                if start > end:
                    return None
                return start, end, True
            # Blank / future / other-month ActualLWD -> unchanged, full
            # period same as Active.
            start = max(emp.doj, self.period_start) if emp.doj else self.period_start
            end = self.period_end
            if start > end:
                return None
            return start, end, False

        else:
            # Active -> always full period, only truncated by DOJ
            start = max(emp.doj, self.period_start) if emp.doj else self.period_start
            end = self.period_end
            if start > end:
                return None
            return start, end, False

    def wide_window(self, emp: Employee):
        """Return (start, end) inclusive, or None for zero-payable, used
        ONLY for the 'No. of Days Absent/LWP' and 'No. of Days Sandwich
        (Nopay)' columns. Unlike working_window() (which switches to a
        narrow 1st-of-month start for Resigned/Exited employees whose
        ActualLWD falls in the ending month), this window ALWAYS anchors
        its start at max(DOJ, period_start) - i.e. it can include the
        26-31 July tail - and its end is ActualLWD for Resigned/Exited
        (regardless of which month that LWD falls in) or period_end for
        Active / Resigned-with-no-LWD-yet. Untracked dates beyond
        period_end (if LWD extends past it) simply have no data and are
        not tallied here (they're assumed Present elsewhere, so they
        can't be Absent/LWP or sandwiched anyway)."""
        if emp.doj and emp.actual_lwd and emp.status == "Exited" and emp.doj == emp.actual_lwd:
            return None  # zero-tenure -> zero across every column, including this one

        start = max(emp.doj, self.period_start) if emp.doj else self.period_start
        if emp.status in ("Resigned", "Exited"):
            end = emp.actual_lwd if emp.actual_lwd else self.period_end
        else:
            end = self.period_end
        if start > end:
            return None
        return start, end

    def compute_absent_and_sandwich(self, emp: Employee):
        """Returns {"absent_lwp": N, "sandwich": {...}} computed over
        wide_window() - see that method's docstring for why this window
        differs from the one used for Present(Payable Days only)."""
        window = self.wide_window(emp)
        if window is None:
            return {"absent_lwp": 0, "sandwich": {"instances": 0, "nopay_days": 0, "flipped_days": 0, "label": "0"}}
        start, end = window
        real_end = min(end, self.period_end)
        window_dates = [d for d in self.period_dates if start <= d <= real_end]
        tally = self._tally_window(emp, window_dates)
        return {"absent_lwp": tally["absent_lwp"], "sandwich": tally["sandwich"]}

    def _tally_window(self, emp: Employee, window_dates):
        """Shared core: given a list of tracked dates (already scoped to
        whatever window the caller wants), applies the sandwich rule and
        tallies present/absent-lwp over exactly those dates. Returns a
        dict with present, absent_lwp, sandwich info, and unrecognized
        statuses. Does NOT handle untracked-tail-days-as-Present (that's
        the caller's job, since it only makes sense for the main window)."""
        present = 0
        absent_lwp = 0
        sandwiched_dates = set()
        instance_count = 0

        off_weekdays = emp.weekly_off_weekdays()
        days_per_instance = 4 if off_weekdays == frozenset({SATURDAY, SUNDAY}) else 3

        # ---- Step 1: detect sandwich instances (only over these dates) ----
        for d in window_dates:
            status_d = emp.day_status.get(d)
            if not is_unpaid_status(status_d):
                continue

            if off_weekdays == frozenset({SATURDAY, SUNDAY}) and d.weekday() == FRIDAY:
                monday = d + datetime.timedelta(days=3)
                if monday in emp.day_status and is_unpaid_status(emp.day_status.get(monday)):
                    instance_count += 1
                    saturday = d + datetime.timedelta(days=1)
                    sunday = d + datetime.timedelta(days=2)
                    for extra in (saturday, sunday):
                        if extra in window_dates and is_weekly_off(emp.day_status.get(extra)):
                            sandwiched_dates.add(extra)

            elif off_weekdays == frozenset({SUNDAY}) and d.weekday() == SATURDAY:
                monday = d + datetime.timedelta(days=2)
                if monday in emp.day_status and is_unpaid_status(emp.day_status.get(monday)):
                    instance_count += 1
                    sunday = d + datetime.timedelta(days=1)
                    if sunday in window_dates and is_weekly_off(emp.day_status.get(sunday)):
                        sandwiched_dates.add(sunday)

        # ---- Step 2: tally present / absent-lwp, skipping sandwiched days from Present ----
        unrecognized = set()
        for d in window_dates:
            status_d = emp.day_status.get(d)
            if d in sandwiched_dates:
                continue  # neither present nor counted again in absent_lwp
            if is_paid_status(status_d):
                present += 1
            elif is_unpaid_status(status_d):
                absent_lwp += 1
            elif status_d is not None:
                # A non-blank status that matches neither list - flag it
                # rather than silently dropping the day (this is exactly
                # the class of bug that under-counted "Paternity leave").
                unrecognized.add(status_d)
            # status_d is None (blank) within window -> shouldn't normally
            # happen, but if it does, we don't count it either way.

        nopay_days = instance_count * days_per_instance
        sandwich_label = f"{instance_count} ({nopay_days} days)" if instance_count else "0"

        return {
            "present": present,
            "absent_lwp": absent_lwp,
            "unrecognized_statuses": unrecognized,
            "sandwich": {
                "instances": instance_count,
                "nopay_days": nopay_days,
                "flipped_days": len(sandwiched_dates),
                "label": sandwich_label,
            },
        }

    def compute(self, emp: Employee):
        """Returns dict with present, absent_lwp counts, a sandwich dict
        {"instances": N, "nopay_days": M, "label": "N (M days)"}, and
        narrow_window (bool, whether the new month-start-to-LWD window
        applied - needed downstream to decide bonus eligibility)."""
        window = self.working_window(emp)
        if window is None:
            return {"present": 0, "absent_lwp": 0, "narrow_window": False,
                     "sandwich": {"instances": 0, "nopay_days": 0, "flipped_days": 0, "label": "0"}}

        start, end, is_narrow = window

        # This file only has actual day-by-day data up to period_end. If the
        # window's end (ActualLWD) falls beyond that, the extra tail days
        # have no data - assume Present for each of them (per confirmed
        # rule), rather than trying to look them up.
        tracked_end = self.period_end
        real_end = min(end, tracked_end)
        window_dates = [d for d in self.period_dates if start <= d <= real_end]
        untracked_present_days = (end - tracked_end).days if end > tracked_end else 0

        result = self._tally_window(emp, window_dates)
        result["present"] += untracked_present_days
        result["narrow_window"] = is_narrow
        return result

    def present_payable_days_for_starting_month_joiner(self, emp: Employee):
        """Special case: an employee whose DOJ falls in the period's
        STARTING calendar month (e.g. the 26-31 July tail of a
        26-Jul-25-Aug cycle). For the "No. of Days Present(Payable Days
        only)" column ONLY, such an employee is treated as if they were an
        ending-month (August) joiner: present is computed over
        [1st-of-ending-month, period_end] (sandwich rule applied within
        that sub-window), plus the same monthly bonus used for real
        ending-month joiners - applied unconditionally here (regardless of
        Active/Resigned/Exited status, unlike the Total Paid Days bonus
        which is Active-only)."""
        month_start = datetime.date(self.period_end.year, self.period_end.month, 1)
        sub_window_dates = [d for d in self.period_dates if month_start <= d <= self.period_end]
        tally = self._tally_window(emp, sub_window_dates)

        days_in_month = calendar.monthrange(self.period_end.year, self.period_end.month)[1]
        bonus = days_in_month - 25

        return tally["present"] + bonus

    def paid_days_in_month(self, emp: Employee):
        """'No. of Paid Days in a Month' column.

        - Active: if DOJ falls in the ending calendar month (e.g. August),
          prorated from DOJ: (month-end - DOJ) + 1 (e.g. DOJ=5-Aug ->
          31-5+1=27). Otherwise (DOJ in an earlier month, an established
          employee), the full calendar length of the month (e.g. 31).
        - Resigned, ActualLWD blank/future/other-month: same DOJ-based
          rule as Active above (prorated if DOJ is in the ending month,
          else full month length).
        - Resigned, ActualLWD in the ending month: inclusive day count
          from max(DOJ, 1st-of-month) to min(ActualLWD, month-end).
        - Exited, DOJ == ActualLWD (zero-tenure): 0 (matches the existing
          zero-tenure override on the other 4 columns).
        - Exited, ActualLWD in the ending month: same inclusive-day-count
          formula as Resigned above.
        - Exited, ActualLWD NOT in the ending month (already left before
          this cycle even started, e.g. both DOJ and LWD in the prior
          month): 0 - they have no payable days in this month at all.
        """
        year, month = self.period_end.year, self.period_end.month
        month_start = datetime.date(year, month, 1)
        days_in_month = calendar.monthrange(year, month)[1]
        month_end = datetime.date(year, month, days_in_month)

        lwd_in_ending_month = (
            emp.actual_lwd is not None
            and emp.actual_lwd.year == year
            and emp.actual_lwd.month == month
        )
        doj_in_ending_month = (
            emp.doj is not None
            and emp.doj.year == year
            and emp.doj.month == month
        )

        def full_month_or_prorated_by_doj():
            if doj_in_ending_month:
                return (month_end - emp.doj).days + 1
            return days_in_month

        if emp.status == "Active":
            return full_month_or_prorated_by_doj()

        if emp.status == "Resigned":
            if not lwd_in_ending_month:
                return full_month_or_prorated_by_doj()
            start = max(emp.doj, month_start) if emp.doj else month_start
            end = min(month_end, emp.actual_lwd)
            if start > end:
                return 0
            return (end - start).days + 1

        if emp.status == "Exited":
            if emp.doj and emp.actual_lwd and emp.doj == emp.actual_lwd:
                return 0  # zero-tenure override, consistent with the other 4 columns
            if not lwd_in_ending_month:
                return 0  # already left before this cycle even started
            start = max(emp.doj, month_start) if emp.doj else month_start
            end = min(month_end, emp.actual_lwd)
            if start > end:
                return 0
            return (end - start).days + 1

        return days_in_month


class AttendanceWorkbookProcessor:
    """Reads the source workbook, computes metrics for every employee, and
    writes a new workbook with 3 extra columns appended."""

    FIXED_COLS = ["Employee Code", "Employee Name", "EmploymentStatus", "DOJ",
                  "ActualLWD", "Grade", "Designation"]

    def __init__(self, input_path, sheet_name="MAIN", header_row=2):
        self.input_path = input_path
        self.sheet_name = sheet_name
        self.header_row = header_row
        self.wb = openpyxl.load_workbook(input_path, data_only=True)
        self.ws = self.wb[sheet_name]

    def _parse_period_dates(self):
        """Read date-column headers, stopping as soon as a non-date column
        is hit (e.g. previously-appended output columns like "No. of Days
        Present"). This makes the processor safe to re-run on a file that
        was already processed - it will detect where the real date columns
        end and overwrite the old summary columns in place instead of
        misreading them as dates."""
        header_vals = [c.value for c in self.ws[self.header_row]]
        date_cols = header_vals[len(self.FIXED_COLS):]
        dates = []
        for v in date_cols:
            if isinstance(v, datetime.datetime):
                dates.append(v.date())
                continue
            if isinstance(v, str):
                try:
                    dates.append(datetime.datetime.strptime(v, "%Y-%m-%d").date())
                    continue
                except ValueError:
                    break  # hit a non-date column (e.g. old output columns) -> stop
            break  # blank or unrecognized cell -> stop
        return dates

    def _load_employees(self, dates):
        employees = []
        n_fixed = len(self.FIXED_COLS)
        for row in self.ws.iter_rows(min_row=self.header_row + 1, values_only=True):
            code, name, status, doj, actual_lwd, grade, designation = row[:n_fixed]
            if code is None and name is None:
                continue
            day_vals = row[n_fixed:]
            day_status = {}
            for d, v in zip(dates, day_vals):
                if d is not None:
                    day_status[d] = v
            employees.append(Employee(
                code=code, name=name, status=status, doj=doj, actual_lwd=actual_lwd,
                extra_fields={"Grade": grade, "Designation": designation},
                day_status=day_status,
            ))
        return employees

    @staticmethod
    def _is_starting_month_tail_joiner(emp: "Employee", metrics: dict, calc: "AttendanceCalculator") -> bool:
        """True for an employee whose DOJ falls in the period's STARTING
        calendar month (e.g. the 26-31 July tail of a 26-Jul-25-Aug
        cycle), and who is NOT already on the narrow Resigned/Exited
        LWD-based window (that window already starts at max(DOJ,1-Aug),
        so there's nothing to override for them)."""
        doj = emp.doj
        return bool(
            doj and not metrics.get("narrow_window")
            and doj >= calc.period_start
            and doj.year == calc.period_start.year
            and doj.month == calc.period_start.month
        )

    @staticmethod
    def _total_payed_days(emp: "Employee", metrics: dict, calc: "AttendanceCalculator") -> int:
        """Total No. of Payed Days.

        - Employees whose DOJ falls in the period's STARTING calendar
          month (e.g. 26-31 July tail), and who aren't already on the
          narrow Resigned/Exited LWD-window: treated EXACTLY like the
          "No. of Days Present(Payable Days only)" column - recalculated
          as if they joined the ending month (August), using
          [1-Aug, period_end] + the monthly bonus, applied unconditionally
          regardless of Active/Resigned/Exited status. Total will equal
          that same Present(Payable Days only) value.
        - Exited employees whose ActualLWD falls in the period's ENDING
          calendar month (e.g. August): Total = "No. of Paid Days in a
          Month" - "No. of Days Absent/LWP" (the two columns as already
          computed - this can go negative, by design, since those two
          columns use different windows and that's accepted).
        - Exited employees NOT covered by either point above (e.g.
          ActualLWD in an earlier month): always just their Present count
          (never gets the new-joiner bonus).
        - Resigned employees whose ActualLWD triggered the narrow
          month-start-to-LWD window: always just their Present count (no
          bonus - Present already reflects exactly the days owed).
        - Resigned employees NOT in that narrow window (blank/future/
          other-month ActualLWD): treated exactly like Active - eligible
          for the same bonus rule below.
        - Active employees (and Resigned employees per the point above)
          with 0 Present days: no bonus, Total = Present (0).
        - Active/Resigned(non-narrow) employees with Present > 0:
            * DOJ in the period's ENDING calendar month (e.g. Aug in a
              26-Jul-to-25-Aug cycle) -> bonus = days_in_that_month - 25
              (covers the days after the 25th that this cycle's data
              doesn't include).
            * DOJ elsewhere (an older employee) -> no bonus, Total = Present.
        """
        if AttendanceWorkbookProcessor._is_starting_month_tail_joiner(emp, metrics, calc):
            return calc.present_payable_days_for_starting_month_joiner(emp)

        if emp.status == "Exited" and metrics.get("narrow_window"):
            return metrics["paid_days_in_month"] - metrics["absent_lwp"]

        if emp.status == "Exited":
            return metrics["present"]

        if emp.status == "Resigned" and metrics.get("narrow_window"):
            return metrics["present"]

        # Active, or Resigned with the unchanged full-period window.
        if metrics["present"] <= 0 or not emp.doj:
            return metrics["present"]

        doj = emp.doj
        period_end = calc.period_end

        if doj.year == period_end.year and doj.month == period_end.month:
            days_in_month = calendar.monthrange(doj.year, doj.month)[1]
            bonus = days_in_month - 25
        else:
            # An older employee (DOJ well before this period) -> Present
            # already has no gaps, no bonus.
            bonus = 0

        return metrics["present"] + bonus

    @staticmethod
    def _present_payable_days(emp: "Employee", metrics: dict, calc: "AttendanceCalculator") -> int:
        """'No. of Days Present(Payable Days only)' - normally identical to
        metrics["present"], EXCEPT for employees whose DOJ falls in the
        period's STARTING calendar month (e.g. 26-31 July): for THIS
        column, they're recalculated as if they joined the ending month
        (August), using [1-Aug, period_end] + the same monthly bonus -
        applied unconditionally, regardless of employment status.
        Absent/LWP and Sandwich continue to use the normal wide window
        (unaffected). Total No. of Payed Days now mirrors this same
        override for this subgroup - see _total_payed_days above.
        Employees already on the narrow Resigned/Exited LWD-based window
        are excluded (their window already starts at max(DOJ, 1-Aug), so
        there's nothing to override)."""
        if AttendanceWorkbookProcessor._is_starting_month_tail_joiner(emp, metrics, calc):
            return calc.present_payable_days_for_starting_month_joiner(emp)
        return metrics["present"]

    def process(self, output_path):
        dates = self._parse_period_dates()
        dates = [d for d in dates if d is not None]
        employees = self._load_employees(dates)
        calc = AttendanceCalculator(dates)

        results = []
        all_unrecognized = set()
        for emp in employees:
            metrics = calc.compute(emp)

            # Absent/LWP and Sandwich use their OWN wide window (always
            # anchored at max(DOJ, period_start), i.e. can include the
            # 26-31 July tail) - independent of the narrow window used for
            # Present(Payable Days only) and Total No. of Payed Days.
            # Computed BEFORE total_payed_days/paid_days_in_month, since
            # the new Exited-in-ending-month Total formula depends on both
            # of these already being final.
            wide = calc.compute_absent_and_sandwich(emp)
            metrics["absent_lwp"] = wide["absent_lwp"]
            metrics["sandwich"] = wide["sandwich"]
            # "Total No. of Absent" = raw Absent/LWP + the sandwich-flipped
            # days folded in (e.g. Sat+Sun for a Sat-Sun-off employee, or
            # just Sun for a Sun-off employee) - NOT the full 4/3-day
            # block, since Friday/Monday are already inside absent_lwp.
            metrics["total_no_of_absent"] = wide["absent_lwp"] + wide["sandwich"]["flipped_days"]

            metrics["paid_days_in_month"] = calc.paid_days_in_month(emp)
            metrics["total_payed_days"] = self._total_payed_days(emp, metrics, calc)
            metrics["present_payable_days"] = self._present_payable_days(emp, metrics, calc)

            all_unrecognized |= metrics.get("unrecognized_statuses", set())
            results.append((emp, metrics))

        if all_unrecognized:
            print("WARNING: unrecognized day-status values found (excluded from "
                  "Present and Absent/LWP counts - check spelling/casing against "
                  "PAID_STATUSES / UNPAID_STATUSES):")
            for s in sorted(all_unrecognized):
                print(f"   - {s!r}")

        self._write_output(output_path, dates, results)
        return results

    def _write_output(self, output_path, dates, results):
        out_wb = openpyxl.load_workbook(self.input_path)  # preserve formatting, formulas untouched
        out_ws = out_wb[self.sheet_name]

        n_fixed = len(self.FIXED_COLS)
        n_date_cols = len(dates)
        new_col_start = n_fixed + n_date_cols + 1  # 1-indexed first new column

        headers = [
            "No. of Paid Days in a Month",
            "No. of Days Present(Payable Days only)",
            "Total No. of Absent",
            "No. of Days Sandwich (Nopay)",
            "Final No. of Payed Days",
        ]
        for i, h in enumerate(headers):
            out_ws.cell(row=self.header_row, column=new_col_start + i, value=h)

        for idx, (emp, metrics) in enumerate(results):
            excel_row = self.header_row + 1 + idx
            out_ws.cell(row=excel_row, column=new_col_start, value=metrics["paid_days_in_month"])
            out_ws.cell(row=excel_row, column=new_col_start + 1, value=metrics["present_payable_days"])
            out_ws.cell(row=excel_row, column=new_col_start + 2, value=metrics["total_no_of_absent"])
            out_ws.cell(row=excel_row, column=new_col_start + 3, value=metrics["sandwich"]["label"])
            out_ws.cell(row=excel_row, column=new_col_start + 4, value=metrics["total_payed_days"])

        out_wb.save(output_path)


if __name__ == "__main__":
    import os
    import sys

    # Usage: python attendance_automation.py [path/to/input.xlsx]
    # Defaults to "ATTENDANCE_LOGIC_Harsh.xlsx" in the current folder if no
    # argument is given. Output is always saved next to the input file as
    # "<input-name>_output<ext>" (e.g. "MyFile.xlsx" -> "MyFile_output.xlsx").
    input_path = sys.argv[1] if len(sys.argv) > 1 else "ATTENDANCE_LOGIC_Harsh_reviewed.xlsx"
    base, ext = os.path.splitext(input_path)
    output_path = f"{base}_output{ext}"

    processor = AttendanceWorkbookProcessor(input_path)
    results = processor.process(output_path)
    print(f"Processed {len(results)} employees.")
    print(f"Output saved to: {output_path}")
