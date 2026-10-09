from datetime import datetime, time, timedelta
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.page import PageMargins
import os
import shutil
import subprocess
import tempfile
from copy import deepcopy
import json
import re
import threading
from openpyxl.utils import get_column_letter

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import (
    WD_TABLE_ALIGNMENT,
    WD_CELL_VERTICAL_ALIGNMENT,
)

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import (
    Inches,
    Pt,
    RGBColor,
)

from pathlib import Path

from django.contrib.auth import get_user_model
from django.contrib.auth.models import Group
from django.contrib.auth import login

from django.db import models
from django.contrib import messages
from django.contrib.auth.decorators import login_required, user_passes_test
from django.db.models import Q
from django.shortcuts import get_object_or_404, redirect, render
from django.http import (
    HttpResponse,
    JsonResponse,
    FileResponse,
)

from django.utils import timezone
from django.urls import reverse

from docx import Document
from docx.enum.section import WD_ORIENT


# Optional OCR dependencies are imported only when an image/scanned PDF is read.

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, letter, legal, landscape, portrait
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)

from .forms import (
    InstructorForm,
    SubjectForm,
    RoomForm,
    ScheduleForm,
    MonitoringRecordForm,
    LateMonitoringRecordForm,
    MissedMonitoringForm,
    ReportTemplateForm,
    FirstAdminSetupForm,
    AdminAccountForm,
)

from .models import (
    Instructor,
    Subject,
    Room,
    Schedule,
    MonitoringRecord,
    ReportTemplate,
    SystemAdministrator,
)


# ==========================================
# DASHBOARD
# ==========================================


from django.shortcuts import redirect
from django.contrib.auth import get_user_model

from .models import SystemAdministrator


def home_redirect(request):
    """
    Send visitors to setup only while there is no active system administrator.
    Once an administrator exists, anonymous visitors go to login and logged-in
    users go to the dashboard.
    """
    User = get_user_model()

    active_main_admin_exists = SystemAdministrator.objects.filter(
        is_main_admin=True,
        user__is_active=True,
    ).exists()

    active_admin_group_exists = User.objects.filter(
        groups__name="System Administrator",
        is_active=True,
    ).exists()

    if not active_main_admin_exists and not active_admin_group_exists:
        return redirect("system_setup")

    if request.user.is_authenticated:
        return redirect("dashboard")

    return redirect("login")


@login_required
def dashboard(request):

    context = {
        "instructor_count": Instructor.objects.count(),
        "subject_count": Subject.objects.count(),
        "room_count": Room.objects.count(),
        "schedule_count": Schedule.objects.count(),
        "monitoring_count": MonitoringRecord.objects.count(),

        "recent_records": MonitoringRecord.objects.select_related(
            "schedule__instructor",
            "schedule__subject",
            "schedule__room"
        ).order_by("-record_datetime")[:10],
    }

    return render(
        request,
        "dashboard.html",
        context
    )

# =========================================================
# SYSTEM ADMINISTRATOR SETUP
# =========================================================



def system_setup(request):
    """Allow the Dean to create the first Main Administrator account."""

    admin_group, _ = Group.objects.get_or_create(
        name="System Administrator"
    )

    # Check whether the Dean's Main Administrator account
    # has already been registered.
    main_admin_exists = SystemAdministrator.objects.filter(
        is_main_admin=True,
        user__is_active=True,
    ).exists()

    if main_admin_exists:
        return render(
            request,
            "registration/setup_locked.html",
            status=403,
        )

    form = FirstAdminSetupForm(request.POST or None)

    if request.method == "POST" and form.is_valid():
        user = form.save()

        admin_group.user_set.add(user)

        SystemAdministrator.objects.update_or_create(
            user=user,
            defaults={"is_main_admin": True},
        )

        login(request, user)

        messages.success(
            request,
            "Main Administrator account created successfully.",
        )

        return redirect("dashboard")

    return render(
        request,
        "registration/setup.html",
        {"form": form},
    )

# =========================================================
# ADMIN ACCOUNT MANAGEMENT
# =========================================================


def is_system_admin(user):
    if not user.is_authenticated:
        return False

    # Technical Django superuser
    if user.is_superuser:
        return True

    # Only the Dean/Main Administrator can manage admin accounts
    return SystemAdministrator.objects.filter(
        user=user,
        is_main_admin=True,
    ).exists()

@login_required
@user_passes_test(is_system_admin)
def admin_accounts(request):
    User = get_user_model()

    admin_group, _ = Group.objects.get_or_create(
        name="System Administrator"
    )

    admins = admin_group.user_set.all().order_by("username")

    return render(
        request,
        "registration/admin_accounts.html",
        {
            "admins": admins,
        }
    )


@login_required
@user_passes_test(is_system_admin)
def admin_account_create(request):
    User = get_user_model()
    admin_group, _ = Group.objects.get_or_create(name="System Administrator")
    form = AdminAccountForm(request.POST or None)

    if request.method == "POST" and form.is_valid():
        user = form.save()
        admin_group.user_set.add(user)
        SystemAdministrator.objects.update_or_create(
            user=user,
            defaults={"is_main_admin": False},
        )
        messages.success(request, f"Administrator account '{user.username}' created successfully.")
        return redirect("admin_accounts")

    return render(
        request,
        "registration/admin_account_form.html",
        {"form": form},
    )

@login_required
@user_passes_test(lambda user: user.is_superuser)
def admin_account_password(request, user_id):

    User = get_user_model()

    admin_group, created = Group.objects.get_or_create(
        name="System Administrator"
    )

    user = get_object_or_404(
        User,
        id=user_id
    )

    # Main Admin cannot change their own password here.
    if user.id == request.user.id:

        messages.error(
            request,
            "You cannot change the Main Administrator password from this page."
        )

        return redirect("admin_accounts")

    # Only System Administrators can be managed here.
    if not user.groups.filter(
        name="System Administrator"
    ).exists():

        messages.error(
            request,
            "That account is not a System Administrator account."
        )

        return redirect("admin_accounts")

    if request.method == "POST":

        password = request.POST.get(
            "password",
            ""
        )

        confirm_password = request.POST.get(
            "confirm_password",
            ""
        )

        if not password or not confirm_password:

            messages.error(
                request,
                "Please complete both password fields."
            )

        elif password != confirm_password:

            messages.error(
                request,
                "Passwords do not match."
            )

        elif len(password) < 8:

            messages.error(
                request,
                "Password must be at least 8 characters."
            )

        else:

            user.set_password(password)

            user.save()

            messages.success(
                request,
                f"Password for administrator '{user.username}' was changed successfully."
            )

            return redirect("admin_accounts")

    return render(
        request,
        "registration/admin_account_password.html",
        {
            "admin": user,
        }
    )
@login_required
@user_passes_test(lambda user: user.is_superuser)
def admin_account_deactivate(request, user_id):

    User = get_user_model()

    user = get_object_or_404(
        User,
        id=user_id
    )

    # Main Admin cannot deactivate themselves.
    if user.id == request.user.id:

        messages.error(
            request,
            "You cannot deactivate the Main Administrator account."
        )

        return redirect("admin_accounts")

    # Only System Administrators can be managed.
    if not user.groups.filter(
        name="System Administrator"
    ).exists():

        messages.error(
            request,
            "That account is not a System Administrator account."
        )

        return redirect("admin_accounts")

    if request.method == "POST":

        user.is_active = False

        user.save()

        messages.success(
            request,
            f"Administrator account '{user.username}' has been deactivated."
        )

    return redirect("admin_accounts")

@login_required
@user_passes_test(lambda user: user.is_superuser)
def admin_account_delete(request, user_id):

    User = get_user_model()

    user = get_object_or_404(
        User,
        id=user_id
    )

    # Main Admin cannot delete themselves.
    if user.id == request.user.id:

        messages.error(
            request,
            "You cannot delete the Main Administrator account."
        )

        return redirect("admin_accounts")

    # Only System Administrators can be deleted.
    if not user.groups.filter(
        name="System Administrator"
    ).exists():

        messages.error(
            request,
            "That account is not a System Administrator account."
        )

        return redirect("admin_accounts")

    if request.method == "POST":

        username = user.username

        user.delete()

        messages.success(
            request,
            f"Administrator account '{username}' was deleted successfully."
        )

    return redirect("admin_accounts")

# ==========================================
# INSTRUCTORS
# ==========================================

@login_required
def instructor_list(request):

    search = request.GET.get("search", "").strip()
    schedule_date_text = request.GET.get("schedule_date", "").strip()

    current_datetime = timezone.localtime()
    today = current_datetime.date()
    current_time = current_datetime.time()

    if schedule_date_text:
        try:
            selected_date = datetime.strptime(
                schedule_date_text, "%Y-%m-%d"
            ).date()
        except ValueError:
            selected_date = today
            schedule_date_text = today.strftime("%Y-%m-%d")
    else:
        selected_date = today
        schedule_date_text = today.strftime("%Y-%m-%d")

    day_code_map = {
        0: "M", 1: "T", 2: "W", 3: "TH",
        4: "F", 5: "S", 6: "SU",
    }

    selected_day = day_code_map[selected_date.weekday()]
    is_today = selected_date == today

    instructors = Instructor.objects.all().order_by("name")
    if search:
        instructors = instructors.filter(name__icontains=search)

    all_instructors = Instructor.objects.all().order_by("name")

    schedules = list(
        Schedule.objects
        .select_related("instructor", "subject", "room")
        .filter(day=selected_day)
        .order_by("instructor__name", "start_time")
    )

    if search:
        schedules = [
            schedule for schedule in schedules
            if search.lower() in schedule.instructor.name.lower()
        ]

    # A monitoring record is tied to a specific schedule and date.
    # Therefore, monitoring an 8:00-10:00 class at 9:00 AM removes
    # only that schedule from the active list. A later class for the
    # same instructor can appear again later in the day.
    day_start = timezone.make_aware(datetime.combine(selected_date, time.min))
    day_end = timezone.make_aware(datetime.combine(selected_date + timedelta(days=1), time.min))

    monitored_schedule_ids = set(
        MonitoringRecord.objects.filter(
            record_datetime__gte=day_start,
            record_datetime__lt=day_end,
        ).values_list("schedule_id", flat=True)
    )

    active_schedules = []
    upcoming_schedules = []

    for schedule in schedules:
        if schedule.id in monitored_schedule_ids:
            continue

        if is_today:
            if schedule.start_time <= current_time <= schedule.end_time:
                active_schedules.append(schedule)
            elif schedule.start_time > current_time:
                upcoming_schedules.append(schedule)
        else:
            # For a manually selected date, show that day's schedules
            # without using today's clock to hide them.
            upcoming_schedules.append(schedule)

    # Active classes are alphabetized by instructor, then time.
    active_schedules.sort(key=lambda x: (x.instructor.name.lower(), x.start_time))

    # Upcoming classes are also alphabetized by instructor, then time.
    upcoming_schedules.sort(key=lambda x: (x.instructor.name.lower(), x.start_time))

    return render(
        request,
        "instructors/list.html",
        {
            "instructors": instructors,
            "all_instructors": all_instructors,
            "search": search,
            "schedule_date": schedule_date_text,
            "selected_date": selected_date,
            "selected_day_name": selected_date.strftime("%A"),
            "is_today": is_today,
            "current_datetime": current_datetime,
            "current_time": current_time,
            "active_schedules": active_schedules,
            "upcoming_schedules": upcoming_schedules,
        },
    )

@login_required
def instructor_bulk_delete(request):

    if request.method == "POST":

        instructor_ids = request.POST.getlist(
            "selected_instructors"
        )

        if instructor_ids:

            Instructor.objects.filter(
                id__in=instructor_ids
            ).delete()

            messages.success(
                request,
                "Selected instructors were deleted successfully."
            )

        else:

            messages.warning(
                request,
                "Please select at least one instructor to delete."
            )

    return redirect("instructor_list")

@login_required
def instructor_create(request):

    if request.method == "POST":

        form = InstructorForm(request.POST)

        if form.is_valid():

            form.save()

            messages.success(
                request,
                "Instructor added successfully!"
            )

            return redirect("instructor_list")

    else:

        form = InstructorForm()

    return render(
        request,
        "instructors/form.html",
        {
            "form": form,
            "title": "Add Instructor",
        }
    )


@login_required
def instructor_update(request, pk):

    instructor = get_object_or_404(
        Instructor,
        pk=pk
    )

    if request.method == "POST":

        form = InstructorForm(
            request.POST,
            instance=instructor
        )

        if form.is_valid():

            form.save()

            messages.success(
                request,
                "Instructor updated successfully!"
            )

            return redirect("instructor_list")

    else:

        form = InstructorForm(
            instance=instructor
        )

    return render(
        request,
        "instructors/form.html",
        {
            "form": form,
            "title": "Edit Instructor",
        }
    )


@login_required
def instructor_delete(request, pk):

    instructor = get_object_or_404(
        Instructor,
        pk=pk
    )

    if request.method == "POST":

        instructor.delete()

        messages.success(
            request,
            "Instructor deleted successfully!"
        )

        return redirect("instructor_list")

    return render(
        request,
        "confirm_delete.html",
        {
            "object": instructor,
            "back_url": "instructor_list",
        }
    )

# ==========================================
# AUTOMATICALLY CREATE MISSED RECORDS
# ==========================================

def create_missed_records_for_today():

    current_datetime = timezone.localtime()
    today = current_datetime.date()
    current_time = current_datetime.time()

    day_code_map = {
        0: "M", 1: "T", 2: "W", 3: "TH",
        4: "F", 5: "S", 6: "SU",
    }

    today_code = day_code_map[today.weekday()]

    schedules = Schedule.objects.filter(day=today_code)

    day_start = timezone.make_aware(datetime.combine(today, time.min))
    day_end = timezone.make_aware(datetime.combine(today + timedelta(days=1), time.min))

    for schedule in schedules:
        if current_time <= schedule.end_time:
            continue

        exists = MonitoringRecord.objects.filter(
            schedule=schedule,
            record_datetime__gte=day_start,
            record_datetime__lt=day_end,
        ).exists()

        if not exists:
            MonitoringRecord.objects.create(
                schedule=schedule,
                record_datetime=timezone.now(),
                remarks="Missed",
                missed_reason="Class was not monitored.",
            )





# ==========================================
# INSTRUCTOR SCHEDULE
# ==========================================

@login_required
def instructor_schedule(request, instructor_id):

    instructor = get_object_or_404(
        Instructor,
        pk=instructor_id
    )

    # ==========================================
    # GET ALL SCHEDULES FOR THIS INSTRUCTOR
    # ==========================================

    schedules = list(
        Schedule.objects
        .filter(
            instructor=instructor
        )
        .select_related(
            "subject",
            "room"
        )
    )

    # ==========================================
    # WEEKLY DAY ORDER
    # ==========================================

    day_order = {
        "M": 0,
        "T": 1,
        "W": 2,
        "TH": 3,
        "F": 4,
        "S": 5,
    }

    day_names = {
        "M": "Monday",
        "T": "Tuesday",
        "W": "Wednesday",
        "TH": "Thursday",
        "F": "Friday",
        "S": "Saturday",
    }

    # ==========================================
    # SORT ALL SCHEDULES
    # ==========================================

    schedules.sort(
        key=lambda schedule: (
            day_order.get(schedule.day, 99),
            schedule.start_time
        )
    )

    # ==========================================
    # BUILD COMPLETE WEEK
    # ==========================================

    weekly_schedule = []

    for day_code in ["M", "T", "W", "TH", "F", "S"]:

        day_schedules = [
            schedule
            for schedule in schedules
            if schedule.day == day_code
        ]

        day_schedules.sort(
            key=lambda schedule: schedule.start_time
        )

        weekly_schedule.append({
            "code": day_code,
            "name": day_names[day_code],
            "schedules": day_schedules,
        })

    # ==========================================
    # RENDER
    # ==========================================

    return render(
        request,
        "instructors/schedules/schedule.html",
        {
            "instructor": instructor,
            "weekly_schedule": weekly_schedule,
        }
    )

# ==========================================
# SUBJECTS
# ==========================================

@login_required
def subject_list(request):

    search = request.GET.get(
        "search",
        ""
    ).strip()

    subjects = Subject.objects.all().order_by(
        "subject_code"
    )

    if search:

        subjects = subjects.filter(
            Q(
                subject_code__icontains=search
            )
            |
            Q(
                subject_name__icontains=search
            )
        )

    all_subjects = Subject.objects.all().order_by(
        "subject_code"
    )

    return render(
        request,
        "instructors/subjects/list.html",
        {
            "subjects": subjects,
            "all_subjects": all_subjects,
            "search": search,
        }
    )


@login_required
def subject_create(request):

    if request.method == "POST":

        subjects_text = request.POST.get(
            "subjects",
            ""
        ).strip()

        if not subjects_text:

            messages.error(
                request,
                "Please enter at least one subject."
            )

            return redirect(
                "subject_create"
            )

        # ==========================================
        # SPLIT SUBJECTS BY LINES AND COMMAS
        # ==========================================

        parts = []

        for line in subjects_text.splitlines():

            parts.extend(
                line.split(",")
            )

        subject_entries = []

        # ==========================================
        # READ EACH SUBJECT
        # ==========================================

        for part in parts:

            part = part.strip()

            if not part:
                continue

            # --------------------------------------
            # FORMAT:
            # GE 5 - Ethics
            # --------------------------------------

            if " - " in part:

                code, name = part.split(
                    " - ",
                    1
                )

                code = code.strip()
                name = name.strip()

            # --------------------------------------
            # FORMAT:
            # GE 5-Ethics
            # --------------------------------------

            elif "-" in part:

                code, name = part.split(
                    "-",
                    1
                )

                code = code.strip()
                name = name.strip()

            # --------------------------------------
            # FORMAT:
            # GE 5
            # --------------------------------------

            else:

                code = part.strip()
                name = ""

            # --------------------------------------
            # CODE IS REQUIRED
            # --------------------------------------

            if not code:
                continue

            subject_entries.append(
                (
                    code,
                    name
                )
            )

        # ==========================================
        # REMOVE DUPLICATE CODES
        # ==========================================

        unique_entries = []

        seen_codes = set()

        for code, name in subject_entries:

            code_key = code.lower().strip()

            if code_key in seen_codes:
                continue

            seen_codes.add(code_key)

            unique_entries.append(
                (
                    code,
                    name
                )
            )

        # ==========================================
        # NO VALID SUBJECTS
        # ==========================================

        if not unique_entries:

            messages.error(
                request,
                "No valid subjects were found."
            )

            return redirect(
                "subject_create"
            )

        created_count = 0
        skipped_count = 0

        # ==========================================
        # CREATE SUBJECTS
        # ==========================================

        for code, name in unique_entries:

            exists = Subject.objects.filter(
                subject_code__iexact=code
            ).exists()

            if exists:

                skipped_count += 1

                continue

            Subject.objects.create(
                subject_code=code,
                subject_name=name
            )

            created_count += 1

        # ==========================================
        # SUCCESS MESSAGE
        # ==========================================

        if created_count > 0:

            message = (
                f"{created_count} subject"
                f"{'s' if created_count != 1 else ''} "
                f"created successfully!"
            )

            if skipped_count > 0:

                message += (
                    f" {skipped_count} existing subject"
                    f"{'s were' if skipped_count != 1 else ' was'} "
                    f"skipped."
                )

            messages.success(
                request,
                message
            )

        elif skipped_count > 0:

            messages.warning(
                request,
                "All entered subject codes already exist. "
                "No duplicate subjects were created."
            )

        return redirect(
            "subject_list"
        )

    return render(
        request,
        "instructors/subjects/form.html",
        {
            "title": "Add Subject",
            "is_multiple": True,
        }
    )


@login_required
def subject_update(request, pk):

    subject = get_object_or_404(
        Subject,
        pk=pk
    )

    if request.method == "POST":

        form = SubjectForm(
            request.POST,
            instance=subject
        )

        if form.is_valid():

            code = form.cleaned_data[
                "subject_code"
            ]

            name = form.cleaned_data[
                "subject_name"
            ]

            duplicate = Subject.objects.filter(
                subject_code__iexact=code
            ).exclude(
                pk=subject.pk
            ).exists()

            if duplicate:

                messages.error(
                    request,
                    "A subject with this subject code already exists."
                )

            else:

                form.save()

                messages.success(
                    request,
                    "Subject updated successfully!"
                )

                return redirect(
                    "subject_list"
                )

    else:

        form = SubjectForm(
            instance=subject
        )

    return render(
        request,
        "instructors/subjects/form.html",
        {
            "form": form,
            "title": "Edit Subject",
            "is_multiple": False,
            "subject": subject,
        }
    )


@login_required
def subject_delete(request, pk):

    subject = get_object_or_404(
        Subject,
        pk=pk
    )

    if request.method == "POST":

        subject.delete()

        messages.success(
            request,
            "Subject deleted successfully!"
        )

        return redirect(
            "subject_list"
        )

    return render(
        request,
        "instructors/subjects/confirm_delete.html",
        {
            "object": subject,
            "back_url": "subject_list",
        }
    )


@login_required
def subject_bulk_delete(request):

    if request.method == "POST":

        subject_ids = request.POST.getlist(
            "subject_ids"
        )

        if subject_ids:

            subjects = Subject.objects.filter(
                id__in=subject_ids
            )

            deleted_count, _ = subjects.delete()

            if deleted_count > 0:

                messages.success(
                    request,
                    f"{deleted_count} subject"
                    f"{'s' if deleted_count != 1 else ''} "
                    f"deleted successfully!"
                )

            else:

                messages.warning(
                    request,
                    "No subjects were deleted."
                )

        else:

            messages.warning(
                request,
                "No subjects were selected."
            )

    return redirect(
        "subject_list"
    )


# ==========================================
# ROOMS
# ==========================================

@login_required
def room_list(request):

    search = request.GET.get(
        "search",
        ""
    ).strip()

    rooms = Room.objects.all().order_by(
        "room_name"
    )

    if search:

        rooms = rooms.filter(
            room_name__icontains=search
        )

    all_rooms = Room.objects.all().order_by(
        "room_name"
    )

    return render(
        request,
        "instructors/rooms/list.html",
        {
            "rooms": rooms,
            "search": search,
            "all_rooms": all_rooms,
        }
    )


@login_required
def room_create(request):

    if request.method == "POST":

        room_names_text = request.POST.get(
            "room_names",
            ""
        ).strip()

        if not room_names_text:

            messages.error(
                request,
                "Please enter at least one room."
            )

            return redirect(
                "room_create"
            )

        import re

        room_names = []

        parts = []

        for line in room_names_text.splitlines():

            parts.extend(
                line.split(",")
            )

        for part in parts:

            part = part.strip()

            if not part:
                continue

            range_match = re.match(
                r"^(.+?)\s+(\d+)\s*-\s*(\d+)$",
                part
            )

            if range_match:

                prefix = (
                    range_match
                    .group(1)
                    .strip()
                    .upper()
                )

                start = int(
                    range_match.group(2)
                )

                end = int(
                    range_match.group(3)
                )

                if start > end:

                    start, end = end, start

                digit_count = max(
                    2,
                    len(str(end))
                )

                for number in range(
                    start,
                    end + 1
                ):

                    room_name = (
                        f"{prefix} "
                        f"{number:0{digit_count}d}"
                    )

                    room_names.append(
                        room_name
                    )

            else:

                room_names.append(
                    part
                )

        room_names = list(
            dict.fromkeys(
                room_names
            )
        )

        created_count = 0
        skipped_count = 0

        for room_name in room_names:

            if Room.objects.filter(
                room_name__iexact=room_name
            ).exists():

                skipped_count += 1

                continue

            Room.objects.create(
                room_name=room_name
            )

            created_count += 1

        if created_count > 0:

            message = (
                f"{created_count} room"
                f"{'s' if created_count != 1 else ''} "
                f"created successfully!"
            )

            if skipped_count > 0:

                message += (
                    f" {skipped_count} existing room"
                    f"{'s were' if skipped_count != 1 else ' was'} "
                    f"skipped."
                )

            messages.success(
                request,
                message
            )

        elif skipped_count > 0:

            messages.warning(
                request,
                "All entered rooms already exist. "
                "No duplicate rooms were created."
            )

        return redirect(
            "room_list"
        )

    return render(
        request,
        "instructors/rooms/form.html",
        {
            "title": "Add Room",
            "is_multiple": True,
        }
    )


@login_required
def room_update(request, pk):

    room = get_object_or_404(
        Room,
        pk=pk
    )

    if request.method == "POST":

        form = RoomForm(
            request.POST,
            instance=room
        )

        if form.is_valid():

            form.save()

            messages.success(
                request,
                "Room updated successfully!"
            )

            return redirect(
                "room_list"
            )

    else:

        form = RoomForm(
            instance=room
        )

    return render(
        request,
        "instructors/rooms/form.html",
        {
            "form": form,
            "title": "Edit Room",
            "is_multiple": False,
        }
    )


@login_required
def room_delete(request, pk):

    room = get_object_or_404(
        Room,
        pk=pk
    )

    if request.method == "POST":

        room.delete()

        messages.success(
            request,
            "Room deleted successfully!"
        )

        return redirect(
            "room_list"
        )

    return render(
        request,
        "confirm_delete.html",
        {
            "object": room,
            "back_url": "room_list",
        }
    )


@login_required
def room_bulk_delete(request):

    if request.method == "POST":

        selected_rooms = request.POST.getlist(
            "selected_rooms"
        )

        if selected_rooms:

            deleted_count, _ = Room.objects.filter(
                id__in=selected_rooms
            ).delete()

            messages.success(
                request,
                f"{deleted_count} room(s) deleted successfully!"
            )

        else:

            messages.warning(
                request,
                "No rooms were selected."
            )

    return redirect(
        "room_list"
    )


# ==========================================
# SCHEDULES
# ==========================================

@login_required
def schedule_list(request):
    search = request.GET.get("search", "").strip()

    instructors = (
        Instructor.objects
        .filter(schedules__isnull=False)
        .distinct()
        .order_by("name")
    )

    if search:
        instructors = instructors.filter(
            name__icontains=search
        )

    return render(
        request,
        "instructors/schedules/list.html",
        {
            "instructors": instructors,
            "search": search,
        }
    )


@login_required
def schedule_create(request):

    if request.method == "POST":

        form = ScheduleForm(
            request.POST
        )

        if form.is_valid():

            instructor = form.cleaned_data[
                "instructor"
            ]

            subject = form.cleaned_data[
                "subject"
            ]

            room = form.cleaned_data[
                "room"
            ]

            day = form.cleaned_data[
                "day"
            ]

            start_time = form.cleaned_data[
                "start_time"
            ]

            end_time = form.cleaned_data[
                "end_time"
            ]

            # ==========================================
            # PREVENT DUPLICATE SCHEDULES
            # ==========================================
            # A schedule is considered a duplicate when
            # the same instructor, subject, day, start time,
            # and end time already exist.
            #
            # ROOM IS NOT USED for duplicate checking.
            # ==========================================

            existing_schedule = Schedule.objects.filter(
                instructor=instructor,
                subject=subject,
                day=day,
                start_time=start_time,
                end_time=end_time,
            ).exists()

            if existing_schedule:

                messages.error(
                    request,
                    "Schedule already exists. "
                    "This schedule was not added."
                )

            else:

                form.save()

                messages.success(
                    request,
                    "Schedule added successfully!"
                )

                return redirect(
                    "schedule_list"
                )

    else:

        form = ScheduleForm()

    return render(
        request,
        "instructors/schedules/form.html",
        {
            "form": form,
            "title": "Add Schedule",
        }
    )


@login_required
def schedule_update(request, pk):

    schedule = get_object_or_404(
        Schedule,
        pk=pk
    )

    if request.method == "POST":

        form = ScheduleForm(
            request.POST,
            instance=schedule
        )

        if form.is_valid():

            form.save()

            messages.success(
                request,
                "Schedule updated successfully!"
            )

            return redirect(
                "schedule_list"
            )

    else:

        form = ScheduleForm(
            instance=schedule
        )

    return render(
        request,
        "instructors/schedules/form.html",
        {
            "form": form,
            "title": "Edit Schedule",
        }
    )


@login_required
def schedule_delete(request, pk):

    schedule = get_object_or_404(
        Schedule,
        pk=pk
    )

    if request.method == "POST":

        schedule.delete()

        messages.success(
            request,
            "Schedule deleted successfully!"
        )

        return redirect(
            "schedule_list"
        )

    return render(
        request,
        "confirm_delete.html",
        {
            "object": schedule,
            "back_url": "schedule_list",
        }
    )


@login_required
def schedule_bulk_delete(request):

    if request.method == "POST":

        schedule_ids = request.POST.getlist(
            "schedule_ids"
        )

        if schedule_ids:

            deleted_count, _ = Schedule.objects.filter(
                id__in=schedule_ids
            ).delete()

            if deleted_count > 0:

                messages.success(
                    request,
                    f"{deleted_count} schedule(s) deleted successfully!"
                )

        else:

            messages.warning(
                request,
                "No schedules were selected."
            )

    return redirect(
        "schedule_list"
    )


# ==========================================
# MONITORING RECORDS
# ==========================================

# ==========================================
# MONITORING RECORDS
# ==========================================

# ==========================================
# DAILY CLASS MONITORING
# ==========================================

@login_required
def monitoring_list(request):

    # ==========================================
    # SELECTED DATE
    # ==========================================

    date_text = request.GET.get(
        "date",
        ""
    ).strip()

    # ==========================================
    # CURRENT PHILIPPINE DATE AND TIME
    # ==========================================

    current_datetime = timezone.localtime()

    today = current_datetime.date()

    current_time = current_datetime.time()

    # ==========================================
    # DETERMINE SELECTED DATE
    # ==========================================

    if date_text:

        try:

            selected_date = datetime.strptime(
                date_text,
                "%Y-%m-%d"
            ).date()

        except ValueError:

            selected_date = today

    else:

        selected_date = today

    date_text = selected_date.strftime(
        "%Y-%m-%d"
    )

    # ==========================================
    # DAY CODE
    # ==========================================

    day_code_map = {

        0: "M",

        1: "T",

        2: "W",

        3: "TH",

        4: "F",

        5: "S",

    }

    selected_day = day_code_map.get(
        selected_date.weekday()
    )

    # Sunday has no schedule code
    # because Schedule.DAYS currently
    # contains Monday-Saturday only.

    is_today = (
        selected_date == today
    )

    # ==========================================
    # GET SCHEDULES FOR SELECTED DAY
    # ==========================================

    if selected_day:

        schedules = list(
            Schedule.objects
            .select_related(
                "instructor",
                "subject",
                "room"
            )
            .filter(
                day=selected_day
            )
            .order_by(
                "instructor__name",
                "start_time"
            )
        )

    else:

        schedules = []

    # ==========================================
    # SELECTED DATE TIME RANGE
    # ==========================================

    day_start = timezone.make_aware(
        datetime.combine(
            selected_date,
            time.min
        )
    )

    day_end = timezone.make_aware(
        datetime.combine(
            selected_date + timedelta(days=1),
            time.min
        )
    )

    # ==========================================
    # GET SCHEDULES ALREADY MONITORED
    # ON THE SELECTED DATE
    # ==========================================

    monitored_schedule_ids = set(
        MonitoringRecord.objects
        .filter(
            record_datetime__gte=day_start,
            record_datetime__lt=day_end
        )
        .values_list(
            "schedule_id",
            flat=True
        )
    )

    # ==========================================
    # ACTIVE / UPCOMING CLASSES
    # ==========================================

    active_schedules = []

    upcoming_schedules = []

    for schedule in schedules:

        # --------------------------------------
        # Already monitored?
        # --------------------------------------

        if schedule.id in monitored_schedule_ids:

            continue

        # --------------------------------------
        # TODAY
        # --------------------------------------

        if is_today:

            # Currently teaching
            if (
                schedule.start_time
                <= current_time
                <= schedule.end_time
            ):

                active_schedules.append(
                    schedule
                )

            # Upcoming
            elif (
                schedule.start_time
                > current_time
            ):

                upcoming_schedules.append(
                    schedule
                )

        # --------------------------------------
        # FUTURE / OTHER SELECTED DATE
        # --------------------------------------

        else:

            upcoming_schedules.append(
                schedule
            )

    # ==========================================
    # ALPHABETICAL ORDER
    # ==========================================

    active_schedules.sort(
        key=lambda schedule: (
            schedule.instructor.name.lower(),
            schedule.start_time
        )
    )

    upcoming_schedules.sort(
        key=lambda schedule: (
            schedule.instructor.name.lower(),
            schedule.start_time
        )
    )

    # ==========================================
    # MONITORING RECORDS FOR SELECTED DATE
    # ==========================================

    monitoring_records = (
        MonitoringRecord.objects
        .select_related(
            "schedule__instructor",
            "schedule__subject",
            "schedule__room"
        )
        .filter(
            record_datetime__gte=day_start,
            record_datetime__lt=day_end
        )
        .order_by(
            "schedule__instructor__name",
            "record_datetime"
        )
    )

    # ==========================================
    # CONTEXT
    # ==========================================

    context = {

        "selected_date":
            selected_date,

        "schedule_date":
            date_text,

        "selected_day":
            selected_day,

        "selected_day_name":
            selected_date.strftime("%A"),

        "today":
            today,

        "is_today":
            is_today,

        "current_datetime":
            current_datetime,

        "current_time":
            current_time,

        "active_schedules":
            active_schedules,

        "upcoming_schedules":
            upcoming_schedules,

        "monitoring_records":
            monitoring_records,
    }

    # ==========================================
    # RENDER DAILY MONITORING PAGE
    # ==========================================

    return render(
        request,
        "instructors/monitoring/list.html",
        context
    )

# ==========================================
# NORMAL MONITORING
# ==========================================

@login_required
def monitoring_create(request, schedule_id):

    schedule = get_object_or_404(
        Schedule.objects.select_related("instructor", "subject", "room"),
        pk=schedule_id,
    )

    today = timezone.localdate()
    day_start = timezone.make_aware(datetime.combine(today, time.min))
    day_end = timezone.make_aware(datetime.combine(today + timedelta(days=1), time.min))

    if MonitoringRecord.objects.filter(
        schedule=schedule,
        record_datetime__gte=day_start,
        record_datetime__lt=day_end,
    ).exists():
        messages.warning(request, "This class has already been monitored today.")
        return redirect("instructor_list")

    if request.method == "POST":
        form = MonitoringRecordForm(request.POST)
        if form.is_valid():
            record = form.save(commit=False)
            record.schedule = schedule
            record.record_datetime = timezone.now()
            record.remarks = form.cleaned_data.get("remarks") or "Monitored"
            record.other_remarks = (
                form.cleaned_data.get("other_remarks", "")
                if record.remarks == "Not Monitored"
                else ""
            )
            # Signature is intentionally blank for physical signing.
            record.signature = ""
            record.save()

            messages.success(request, "Monitoring record saved successfully!")
            return redirect("instructor_list")
    else:
        form = MonitoringRecordForm(initial={"remarks": "Monitored"})

    return render(
        request,
        "instructors/monitoring/form.html",
        {
            "form": form,
            "schedule": schedule,
            "title": "Monitor Class",
            "submit_text": "Save Monitoring Record",
            "cancel_url": reverse("instructor_list"),
        },
    )


# ==========================================
# LATE MONITORING
# ==========================================

@login_required
def monitor_late(request, schedule_id):

    schedule = get_object_or_404(
        Schedule.objects.select_related("instructor", "subject", "room"),
        pk=schedule_id,
    )

    if request.method == "POST":
        form = LateMonitoringRecordForm(request.POST)
        if form.is_valid():
            record = form.save(commit=False)
            record.schedule = schedule
            record.remarks = "Late"
            record.signature = ""
            record.save()

            messages.success(request, "Late monitoring record saved successfully!")
            return redirect("instructor_list")
    else:
        form = LateMonitoringRecordForm(
            initial={
                "monitoring_date": timezone.localdate(),
                "monitoring_time": timezone.localtime().time(),
                "remarks": "Late",
            }
        )

    return render(
        request,
        "instructors/monitoring/late_form.html",
        {"form": form, "schedule": schedule},
    )


# ==========================================
# MISSED MONITORING
# ==========================================

@login_required
def monitoring_missed(request, schedule_id):

    schedule = get_object_or_404(
        Schedule.objects.select_related("instructor", "subject", "room"),
        pk=schedule_id,
    )

    today = timezone.localdate()
    day_start = timezone.make_aware(datetime.combine(today, time.min))
    day_end = timezone.make_aware(datetime.combine(today + timedelta(days=1), time.min))

    existing_record = MonitoringRecord.objects.filter(
        schedule=schedule,
        record_datetime__gte=day_start,
        record_datetime__lt=day_end,
    ).order_by("-record_datetime").first()

    if existing_record:
        return redirect("monitoring_update", pk=existing_record.pk)

    if request.method == "POST":
        form = MissedMonitoringForm(request.POST)
        if form.is_valid():
            record = form.save(commit=False)
            record.schedule = schedule
            record.remarks = "Missed"
            record.record_datetime = timezone.now()
            record.signature = ""
            record.save()

            messages.success(request, "Class marked as Missed.")
            return redirect("instructor_list")
    else:
        form = MissedMonitoringForm()

    return render(
        request,
        "instructors/monitoring/missed_form.html",
        {"form": form, "schedule": schedule},
    )


# ==========================================
# UPDATE MONITORING
# ==========================================

@login_required
def monitoring_update(request, pk):

    record = get_object_or_404(
        MonitoringRecord.objects.select_related(
            "schedule__instructor",
            "schedule__subject",
            "schedule__room",
        ),
        pk=pk,
    )

    old_remarks = record.remarks

    if request.method == "POST":
        form = MonitoringRecordForm(request.POST, instance=record)

        if form.is_valid():
            updated_record = form.save(commit=False)
            updated_record.signature = ""

            if updated_record.remarks != "Not Monitored":
                updated_record.other_remarks = ""

            if old_remarks == "Missed" and updated_record.remarks in ["Monitored", "Late"]:
                updated_record.record_datetime = timezone.now()

            updated_record.save()

            messages.success(request, "Monitoring record updated successfully!")
            return redirect("monitoring_list")
    else:
        form = MonitoringRecordForm(instance=record)

    return render(
        request,
        "instructors/monitoring/form.html",
        {
            "form": form,
            "title": "Edit Monitoring Record",
            "submit_text": "Update Monitoring Record",
            "schedule": record.schedule,
            "record": record,
            "cancel_url": reverse("monitoring_list"),
        },
    )




# ==========================================
# DELETE MONITORING
# ==========================================

@login_required
def monitoring_delete(request, pk):

    record = get_object_or_404(
        MonitoringRecord,
        pk=pk
    )

    if request.method == "POST":

        record.delete()

        messages.success(
            request,
            "Monitoring record deleted successfully!"
        )

        return redirect(
            "monitoring_list"
        )

    return render(
        request,
        "confirm_delete.html",
        {
            "object": record,
            "back_url": "monitoring_list",
        }
    )


# ==========================================
# BULK DELETE MONITORING
# ==========================================

@login_required
def monitoring_bulk_delete(request):

    if request.method == "POST":

        selected_ids = request.POST.getlist(
            "selected_records"
        )

        if selected_ids:

            records = MonitoringRecord.objects.filter(
                pk__in=selected_ids
            )

            deleted_count = records.count()

            records.delete()

            messages.success(
                request,
                f"{deleted_count} monitoring record(s) deleted successfully!"
            )

        else:

            messages.warning(
                request,
                "No monitoring records were selected."
            )

    return redirect("reports")


# ==========================================
# ==========================================
# REPORTS
# ==========================================


def _request_value(request, key, default=""):
    """Read a report setting from POST first, then GET."""
    if request.method == "POST" and key in request.POST:
        return request.POST.get(key, default)
    return request.GET.get(key, default)


def _report_week_range(request):
    today = timezone.localdate()
    current_week_start = today - timedelta(days=today.weekday())
    selected_week = _request_value(
        request,
        "week",
        current_week_start.strftime("%Y-%m-%d"),
    ).strip()

    try:
        week_start = datetime.strptime(
            selected_week,
            "%Y-%m-%d",
        ).date()
    except ValueError:
        week_start = current_week_start

    week_start -= timedelta(days=week_start.weekday())
    week_end = week_start + timedelta(days=6)
    return week_start, week_end, current_week_start


def _report_classification(request):
    classification = _request_value(
        request,
        "classification",
        "ALL",
    ).strip().upper()

    if classification not in {
        "ALL",
        "PART_TIME",
        "REGULAR",
    }:
        classification = "ALL"

    return classification


def _report_records(request, week_start, week_end):
    start_dt = timezone.make_aware(
        datetime.combine(week_start, time.min)
    )
    end_dt = timezone.make_aware(
        datetime.combine(
            week_end + timedelta(days=1),
            time.min,
        )
    )

    records = (
        MonitoringRecord.objects
        .filter(
            record_datetime__gte=start_dt,
            record_datetime__lt=end_dt,
        )
        .select_related(
            "schedule__instructor",
            "schedule__subject",
            "schedule__room",
        )
        .order_by(
            "schedule__instructor__name",
            "record_datetime",
        )
    )

    classification = _report_classification(request)

    if classification != "ALL":
        records = records.filter(
            schedule__instructor__classification=classification
        )

    instructor_search = _request_value(
        request,
        "instructor",
        "",
    ).strip()

    if instructor_search:
        records = records.filter(
            schedule__instructor__name__icontains=instructor_search
        )

    selected_records = _request_value(
        request,
        "records",
        "",
    ).strip()

    selected_ids = []

    if selected_records:
        for value in selected_records.split(","):
            value = value.strip()
            if value.isdigit():
                selected_ids.append(int(value))

        if selected_ids:
            records = records.filter(pk__in=selected_ids)

    return list(records), instructor_search


def _report_settings(request):
    paper_size = _request_value(
        request,
        "paper_size",
        "A4",
    ).upper()

    if paper_size not in {
        "A4",
        "LETTER",
        "LEGAL",
        "FOLIO",
    }:
        paper_size = "A4"

    orientation = _request_value(
        request,
        "orientation",
        "landscape",
    ).lower()

    if orientation not in {
        "portrait",
        "landscape",
    }:
        orientation = "landscape"

    color_mode = _request_value(
        request,
        "color",
        "color",
    ).lower()

    if color_mode not in {
        "color",
        "bw",
    }:
        color_mode = "color"

    margins = _request_value(
        request,
        "margins",
        "normal",
    ).lower()

    if margins not in {
        "normal",
        "narrow",
    }:
        margins = "normal"

    return {
        "paper_size": paper_size,
        "orientation": orientation,
        "color": color_mode,
        "margins": margins,
    }


def _report_page_size(settings):
    sizes = {
        "A4": A4,
        "LETTER": letter,
        "LEGAL": legal,
        "FOLIO": (8.5 * inch, 13 * inch),
    }

    page_size = sizes.get(
        settings["paper_size"],
        A4,
    )

    return (
        landscape(page_size)
        if settings["orientation"] == "landscape"
        else portrait(page_size)
    )


def _report_margin_inches(settings):
    return (
        0.30
        if settings["margins"] == "narrow"
        else 0.50
    )


def _report_subject(record):
    subject = record.schedule.subject.subject_code

    if record.schedule.subject.subject_name:
        subject += (
            f" - {record.schedule.subject.subject_name}"
        )

    return subject


def _report_actual_room(record):
    return (
        record.transfer_room
        or record.schedule.room.room_name
    )


def _report_remarks(record):
    remarks = (record.remarks or "").strip()
    other = (record.other_remarks or "").strip()
    missed = (record.missed_reason or "").strip()

    if remarks == "Not Monitored" and other:
        return f"Not Monitored - {other}"

    if remarks == "Missed" and missed:
        return f"Missed - {missed}"

    if other:
        return (
            f"{remarks} - {other}"
            if remarks
            else other
        )

    if missed:
        return (
            f"{remarks} - {missed}"
            if remarks
            else missed
        )

    return remarks


def _report_row(record):
    schedule = record.schedule
    monitored_dt = timezone.localtime(
        record.record_datetime
    )

    return [
        schedule.instructor.name,
        _report_subject(record),
        schedule.room.room_name,
        _report_actual_room(record),
        (
            f"{schedule.start_time.strftime('%I:%M %p')}"
            f" - {schedule.end_time.strftime('%I:%M %p')}"
        ),
        monitored_dt.strftime(
            "%b %d, %Y %I:%M %p"
        ),
        schedule.get_day_display(),
        "",
        _report_remarks(record),
    ]


# ==========================================
# WINDOWS PRINTER DETECTION
# ==========================================


def _windows_printers():
    """
    Return printers installed on the Windows computer running Django.

    This deliberately does NOT hard-code printer names. The names come
    directly from the Windows print subsystem through pywin32.
    """
    if os.name != "nt":
        return []

    try:
        import win32print
    except ImportError:
        return []

    printers = []

    flags = (
        win32print.PRINTER_ENUM_LOCAL
        | win32print.PRINTER_ENUM_CONNECTIONS
    )

    try:
        printer_info = win32print.EnumPrinters(
            flags,
            None,
            2,
        )
    except Exception:
        return []

    try:
        default_printer = win32print.GetDefaultPrinter()
    except Exception:
        default_printer = ""

    seen = set()

    for info in printer_info:
        printer_name = str(
            info.get("pPrinterName", "")
        ).strip()

        if not printer_name or printer_name in seen:
            continue

        seen.add(printer_name)

        printers.append({
            "name": printer_name,
            "is_default": (
                printer_name == default_printer
            ),
        })

    printers.sort(
        key=lambda item: (
            not item["is_default"],
            item["name"].lower(),
        )
    )

    return printers

def _get_report_template(template_id=None):
    """
    Return the exact Excel report template selected
    by the user.

    Instructor classification does NOT determine
    the report layout.
    """

    from pathlib import Path

    if not template_id:

        raise ValueError(
            "Please choose a report layout before "
            "generating the report."
        )

    try:

        template = (
            ReportTemplate.objects.get(
                pk=template_id
            )
        )

    except ReportTemplate.DoesNotExist:

        raise ValueError(
            "The selected report layout "
            "could not be found."
        )

    if not template.file:

        raise ValueError(
            "The selected report layout "
            "does not contain a file."
        )

    try:

        template_path = Path(
            template.file.path
        )

    except Exception as exc:

        raise ValueError(
            "The selected report layout file "
            "could not be accessed."
        ) from exc

    if not template_path.exists():

        raise ValueError(
            "The selected report layout file "
            "could not be found."
        )

    if template_path.suffix.lower() != ".xlsx":

        raise ValueError(
            "The selected report layout must "
            "be an Excel .xlsx file."
        )

    return (
        template,
        template_path,
    )

@login_required
def report_printers(request):
    """
    AJAX endpoint used by reports/list.html.

    Returns all printers Windows reports as installed on the SAME
    computer where this Django application is running.
    """
    if request.method != "GET":
        return JsonResponse(
            {"error": "GET request required."},
            status=405,
        )

    if os.name != "nt":
        return JsonResponse({
            "printers": [],
            "default_printer": "",
            "available": False,
            "message": (
                "Windows printer detection is available when "
                "Django is running on Windows."
            ),
        })

    try:
        import win32print  # noqa: F401
    except ImportError:
        return JsonResponse({
            "printers": [],
            "default_printer": "",
            "available": False,
            "message": (
                "pywin32 is not installed. Run: "
                "python -m pip install pywin32"
            ),
        })

    printers = _windows_printers()
    default_printer = next(
        (
            item["name"]
            for item in printers
            if item["is_default"]
        ),
        "",
    )

    return JsonResponse({
        "printers": printers,
        "default_printer": default_printer,
        "available": bool(printers),
        "message": (
            ""
            if printers
            else "No Windows printers were detected."
        ),
    })

def _clean_template_instructor_name(value):

    if not value:
        return ""

    value = str(value).strip()

    # Remove excessive whitespace
    value = " ".join(value.split())

    return value.upper()

def _instructor_name_matches(
    template_name,
    django_name
):

    template_name = _clean_template_instructor_name(
        template_name
    )

    django_name = _clean_template_instructor_name(
        django_name
    )

    if not template_name or not django_name:
        return False

    return (
        template_name.startswith(django_name)
        or django_name.startswith(template_name)
    )

def _build_report_excel_file(
    records,
    week_start,
    week_end,
    template_id,
):
    """
    Build the report from the ACTIVE uploaded Excel template.

    IMPORTANT:
    - The Excel file is the master layout.
    - Python does NOT create report headers/columns.
    - Python does NOT define colors, borders, widths, merged cells,
      page layout, semester, school year, etc.
    - The template controls all of those things.
    - Python only replaces {{PLACEHOLDERS}} with system data.

    Template placeholders:

        {{WEEK_START}}
        {{WEEK_END}}
        {{WEEK}}
        {{INSTRUCTOR}}
        {{SUBJECT}}
        {{TIME}}
        {{DAY}}
        {{ROOM}}
        {{ACTUAL_ROOM}}
        {{MONITORING_TIME}}
        {{SIGNATURE}}
        {{REMARKS}}

    For a repeating monitoring row, put the placeholders on ONE
    template row. The system will copy that row for every record.
    """

    import os
    import re
    import shutil
    import tempfile

    from copy import copy
    from pathlib import Path

    from openpyxl import load_workbook
    from openpyxl.formula.translate import Translator

    # ==========================================================
    # 1. GET USER-SELECTED UPLOADED TEMPLATE
    # ==========================================================

    template, template_path = _get_report_template(
        template_id
    )

    extension = os.path.splitext(
        str(template_path)
    )[1].lower()

    if extension != ".xlsx":
        raise ValueError(
            "The active report template must be an .xlsx Excel file."
        )

    # ==========================================================
    # 2. CREATE TEMPORARY WORKING DIRECTORY
    # ==========================================================

    temp_dir = tempfile.mkdtemp(
        prefix="teacher_monitoring_report_"
    )

    output_xlsx = os.path.join(
        temp_dir,
        "Monitoring_Report.xlsx"
    )

    # ==========================================================
    # 3. COPY THE ORIGINAL TEMPLATE
    # ==========================================================

    shutil.copy2(
        str(template_path),
        output_xlsx
    )

    # ==========================================================
    # 4. OPEN THE COPIED TEMPLATE
    # ==========================================================

    workbook = load_workbook(
        output_xlsx,
        data_only=False
    )

    # ==========================================================
    # 5. DATE VALUES
    # ==========================================================

    week_start_text = week_start.strftime(
        "%B %d, %Y"
    )

    week_end_text = week_end.strftime(
        "%B %d, %Y"
    )

    week_text = (
        f"{week_start_text} - {week_end_text}"
    )

    # ==========================================================
    # 6. CREATE DATA FOR EACH MONITORING RECORD
    # ==========================================================

    def build_record_data(record):

        schedule = getattr(
            record,
            "schedule",
            None
        )

        if not schedule:
            return {}

        instructor = getattr(
            schedule,
            "instructor",
            None
        )

        subject = getattr(
            schedule,
            "subject",
            None
        )

        room = getattr(
            schedule,
            "room",
            None
        )

        # ------------------------------------------------------
        # INSTRUCTOR
        # ------------------------------------------------------

        instructor_name = str(
            getattr(
                instructor,
                "name",
                ""
            ) or ""
        ).strip()

        # ------------------------------------------------------
        # SUBJECT
        # ------------------------------------------------------

        subject_code = str(
            getattr(
                subject,
                "subject_code",
                ""
            ) or ""
        ).strip()

        subject_name = str(
            getattr(
                subject,
                "subject_name",
                ""
            ) or ""
        ).strip()

        if subject_code and subject_name:
            subject_text = (
                f"{subject_code} - {subject_name}"
            )
        else:
            subject_text = (
                subject_code
                or subject_name
            )

        # ------------------------------------------------------
        # SCHEDULED TIME
        # ------------------------------------------------------

        start_time = getattr(
            schedule,
            "start_time",
            None
        )

        end_time = getattr(
            schedule,
            "end_time",
            None
        )

        if start_time and end_time:

            scheduled_time = (
                start_time.strftime("%I:%M %p")
                + " - "
                + end_time.strftime("%I:%M %p")
            )

        else:

            scheduled_time = ""

        # ------------------------------------------------------
        # DAY
        # ------------------------------------------------------

        try:

            day_text = (
                schedule.get_day_display()
            )

        except Exception:

            day_text = str(
                getattr(
                    schedule,
                    "day",
                    ""
                ) or ""
            )

        # ------------------------------------------------------
        # ROOM
        # ------------------------------------------------------

        assigned_room = str(
            getattr(
                room,
                "room_name",
                ""
            ) or ""
        ).strip()

        actual_room = str(
            getattr(
                record,
                "transfer_room",
                ""
            ) or ""
        ).strip()

        if not actual_room:
            actual_room = assigned_room

        # ------------------------------------------------------
        # MONITORING DATE/TIME
        # ------------------------------------------------------

        record_datetime = getattr(
            record,
            "record_datetime",
            None
        )

        if record_datetime:

            monitoring_time = timezone.localtime(
                record_datetime
            ).strftime(
                "%B %d, %Y %I:%M %p"
            )

        else:

            monitoring_time = ""

        # ------------------------------------------------------
        # SIGNATURE
        # ------------------------------------------------------

        signature = str(
            getattr(
                record,
                "signature",
                ""
            ) or ""
        ).strip()

        # ------------------------------------------------------
        # REMARKS
        # ------------------------------------------------------

        remarks = str(
            getattr(
                record,
                "remarks",
                ""
            ) or ""
        ).strip()

        other_remarks = str(
            getattr(
                record,
                "other_remarks",
                ""
            ) or ""
        ).strip()

        missed_reason = str(
            getattr(
                record,
                "missed_reason",
                ""
            ) or ""
        ).strip()

        if remarks == "Not Monitored" and other_remarks:

            remarks_text = (
                f"{remarks} - {other_remarks}"
            )

        elif remarks == "Missed" and missed_reason:

            remarks_text = (
                f"{remarks} - {missed_reason}"
            )

        elif other_remarks:

            remarks_text = (
                f"{remarks} - {other_remarks}"
                if remarks
                else other_remarks
            )

        elif missed_reason:

            remarks_text = (
                f"{remarks} - {missed_reason}"
                if remarks
                else missed_reason
            )

        else:

            remarks_text = remarks

        # ------------------------------------------------------
        # RETURN PLACEHOLDER DATA
        # ------------------------------------------------------

        return {
            "{{WEEK_START}}": week_start_text,
            "{{WEEK_END}}": week_end_text,
            "{{WEEK}}": week_text,

            "{{INSTRUCTOR}}": instructor_name,
            "{{SUBJECT}}": subject_text,
            "{{TIME}}": scheduled_time,
            "{{DAY}}": str(day_text or ""),
            "{{ROOM}}": assigned_room,
            "{{ACTUAL_ROOM}}": actual_room,

            "{{MONITORING_TIME}}": monitoring_time,
            "{{SIGNATURE}}": signature,
            "{{REMARKS}}": remarks_text,
        }

    # ==========================================================
    # 7. BUILD DATA LIST
    # ==========================================================

    record_data = []

    for record in records:

        data = build_record_data(
            record
        )

        if data:
            record_data.append(data)

    # ==========================================================
    # 8. PLACEHOLDER REGEX
    # ==========================================================

    placeholder_pattern = re.compile(
        r"\{\{[A-Z0-9_]+\}\}"
    )

    # ==========================================================
    # 9. REPLACE SIMPLE PLACEHOLDERS
    # ==========================================================

    def replace_placeholders_in_workbook():

        for worksheet in workbook.worksheets:

            for row in worksheet.iter_rows():

                for cell in row:

                    value = cell.value

                    if not isinstance(
                        value,
                        str
                    ):
                        continue

                    matches = (
                        placeholder_pattern.findall(
                            value
                        )
                    )

                    if not matches:
                        continue

                    new_value = value

                    for placeholder in matches:

                        # This is a header/week/global
                        # placeholder, so use the first
                        # record's data if available.
                        replacement = ""

                        if record_data:

                            replacement = (
                                record_data[0]
                                .get(
                                    placeholder,
                                    ""
                                )
                            )

                        else:

                            global_values = {
                                "{{WEEK_START}}":
                                    week_start_text,

                                "{{WEEK_END}}":
                                    week_end_text,

                                "{{WEEK}}":
                                    week_text,
                            }

                            replacement = (
                                global_values.get(
                                    placeholder,
                                    ""
                                )
                            )

                        new_value = new_value.replace(
                            placeholder,
                            str(
                                replacement
                                or ""
                            )
                        )

                    cell.value = new_value

    # ==========================================================
    # 10. FIND THE REPEATING DATA ROW
    # ==========================================================

    detail_row = None

    detail_placeholders = {
        "{{INSTRUCTOR}}",
        "{{SUBJECT}}",
        "{{TIME}}",
        "{{DAY}}",
        "{{ROOM}}",
        "{{ACTUAL_ROOM}}",
        "{{MONITORING_TIME}}",
        "{{SIGNATURE}}",
        "{{REMARKS}}",
    }

    for worksheet in workbook.worksheets:

        for row in worksheet.iter_rows():

            row_placeholders = set()

            for cell in row:

                value = cell.value

                if not isinstance(
                    value,
                    str
                ):
                    continue

                matches = (
                    placeholder_pattern.findall(
                        value
                    )
                )

                row_placeholders.update(
                    matches
                )

            if row_placeholders.intersection(
                detail_placeholders
            ):

                detail_row = (
                    worksheet,
                    row[0].row
                )

                break

        if detail_row:
            break

    # ==========================================================
    # 11. FUNCTION TO COPY A COMPLETE EXCEL ROW
    # ==========================================================

    def copy_row_with_style(
        worksheet,
        source_row,
        target_row
    ):
        """
        Copy the template row exactly enough for normal Excel
        report templates.

        This copies:
        - values
        - formulas
        - font
        - fill
        - border
        - alignment
        - number format
        - protection
        - comments
        - row height
        """

        source_dimension = worksheet.row_dimensions[
            source_row
        ]

        target_dimension = worksheet.row_dimensions[
            target_row
        ]

        if source_dimension.height is not None:

            target_dimension.height = (
                source_dimension.height
            )

        target_dimension.hidden = (
            source_dimension.hidden
        )

        target_dimension.outlineLevel = (
            source_dimension.outlineLevel
        )

        for column in range(
            1,
            worksheet.max_column + 1
        ):

            source_cell = worksheet.cell(
                source_row,
                column
            )

            target_cell = worksheet.cell(
                target_row,
                column
            )

            if source_cell.data_type == "f":

                try:

                    target_cell.value = (
                        Translator(
                            source_cell.value,
                            origin=source_cell.coordinate
                        ).translate_formula(
                            target_cell.coordinate
                        )
                    )

                except Exception:

                    target_cell.value = (
                        source_cell.value
                    )

            else:

                target_cell.value = (
                    source_cell.value
                )

            if source_cell.has_style:

                target_cell._style = copy(
                    source_cell._style
                )

            if source_cell.number_format:

                target_cell.number_format = (
                    source_cell.number_format
                )

            if source_cell.alignment:

                target_cell.alignment = copy(
                    source_cell.alignment
                )

            if source_cell.protection:

                target_cell.protection = copy(
                    source_cell.protection
                )

            if source_cell.comment:

                target_cell.comment = copy(
                    source_cell.comment
                )

    # ==========================================================
    # 12. COPY MERGED-CELL STRUCTURE FOR A ROW
    # ==========================================================

    def copy_row_merges(
        worksheet,
        source_row,
        target_row
    ):

        merged_ranges = list(
            worksheet.merged_cells.ranges
        )

        for merged_range in merged_ranges:

            if (
                merged_range.min_row
                <= source_row
                <= merged_range.max_row
            ):

                row_height = (
                    merged_range.max_row
                    - merged_range.min_row
                )

                if row_height != 0:
                    continue

                new_min_row = target_row

                new_max_row = target_row

                new_range = (
                    f"{get_column_letter(merged_range.min_col)}"
                    f"{new_min_row}:"
                    f"{get_column_letter(merged_range.max_col)}"
                    f"{new_max_row}"
                )

                try:

                    worksheet.merge_cells(
                        new_range
                    )

                except Exception:

                    pass

    # ==========================================================
    # 13. FILL REPEATING DATA ROW
    # ==========================================================

    if detail_row:

        worksheet, template_row = detail_row

        # ------------------------------------------------------
        # MAKE SPACE FOR ADDITIONAL RECORDS
        # ------------------------------------------------------

        number_of_records = len(
            record_data
        )

        if number_of_records > 1:

            worksheet.insert_rows(
                template_row + 1,
                amount=number_of_records - 1
            )

            # Copy the template row's formatting
            # into all newly created rows.
            for index in range(
                1,
                number_of_records
            ):

                target_row = (
                    template_row + index
                )

                copy_row_with_style(
                    worksheet,
                    template_row,
                    target_row
                )

                copy_row_merges(
                    worksheet,
                    template_row,
                    target_row
                )

        # ------------------------------------------------------
        # FILL EVERY RECORD
        # ------------------------------------------------------

        for index, data in enumerate(
            record_data
        ):

            target_row = (
                template_row + index
            )

            for column in range(
                1,
                worksheet.max_column + 1
            ):

                cell = worksheet.cell(
                    target_row,
                    column
                )

                value = cell.value

                if not isinstance(
                    value,
                    str
                ):
                    continue

                matches = (
                    placeholder_pattern.findall(
                        value
                    )
                )

                if not matches:
                    continue

                new_value = value

                for placeholder in matches:

                    replacement = data.get(
                        placeholder,
                        ""
                    )

                    new_value = (
                        new_value.replace(
                            placeholder,
                            str(
                                replacement
                                or ""
                            )
                        )
                    )

                cell.value = new_value

    # ==========================================================
    # 14. IF THERE IS NO DETAIL ROW
    # ==========================================================

    else:

        # The template does not contain a repeating
        # data placeholder.
        #
        # Do NOT invent a table or columns here.
        # This protects the "template is the master"
        # requirement.

        if record_data:

            raise ValueError(
                "The active Excel report template does not "
                "contain a repeating monitoring row. "
                "Add at least one row containing one or more "
                "of these placeholders: "
                "{{INSTRUCTOR}}, {{SUBJECT}}, {{TIME}}, "
                "{{DAY}}, {{ROOM}}, {{ACTUAL_ROOM}}, "
                "{{MONITORING_TIME}}, {{SIGNATURE}}, "
                "{{REMARKS}}."
            )

    # ==========================================================
    # 15. REPLACE GLOBAL PLACEHOLDERS
    # ==========================================================

    # We do this after the repeating row so that the
    # repeating row has already received its own values.

    global_values = {
        "{{WEEK_START}}": week_start_text,
        "{{WEEK_END}}": week_end_text,
        "{{WEEK}}": week_text,
    }

    for worksheet in workbook.worksheets:

        for row in worksheet.iter_rows():

            for cell in row:

                value = cell.value

                if not isinstance(
                    value,
                    str
                ):
                    continue

                new_value = value

                for placeholder, replacement in (
                    global_values.items()
                ):

                    new_value = (
                        new_value.replace(
                            placeholder,
                            replacement
                        )
                    )

                cell.value = new_value

    # ==========================================================
    # 16. REMOVE UNUSED TEMPLATE PLACEHOLDERS
    # ==========================================================

    for worksheet in workbook.worksheets:

        for row in worksheet.iter_rows():

            for cell in row:

                value = cell.value

                if not isinstance(
                    value,
                    str
                ):
                    continue

                if (
                    "{{" in value
                    and "}}" in value
                ):

                    # Only replace placeholders that
                    # are actually known system fields.
                    for placeholder in (
                        detail_placeholders
                    ):

                        if placeholder in value:

                            # If this cell was outside the
                            # repeating row, blank it rather
                            # than inventing data.
                            value = value.replace(
                                placeholder,
                                ""
                            )

                    cell.value = value

    # ==========================================================
    # 17. SAVE GENERATED EXCEL
    # ==========================================================

    workbook.save(
        output_xlsx
    )

    return (
        output_xlsx,
        temp_dir,
        template,
    )

def _build_report_docx_file(records, week_start, week_end):
    """
    Build an editable DOCX from the generated Excel report.

    The uploaded Excel template is used as the master source.
    The generated Excel report is converted into an editable
    Word table while preserving:
    - cell values
    - merged cells
    - column widths
    - row heights
    - alignment
    - fonts
    - background colors
    - borders
    - Excel images/logos
    - page size
    - margins
    """

    import os
    from io import BytesIO

    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    # =========================================================
    # 1. BUILD THE EXCEL REPORT FIRST
    # =========================================================

    excel_path, temp_dir, template = _build_report_excel_file(
        records,
        week_start,
        week_end,
    )

    # =========================================================
    # 2. OPEN GENERATED EXCEL REPORT
    # =========================================================

    workbook = load_workbook(
        excel_path,
        data_only=False,
    )

    worksheet = workbook.active

    # =========================================================
    # 3. DEFINE EXCEL RANGE
    # =========================================================

    min_row = 1
    max_row = worksheet.max_row

    min_col = 1
    max_col = worksheet.max_column

    # =========================================================
    # 4. FIND ACTUAL REPORT WIDTH
    # =========================================================

    header_row = None

    for row in worksheet.iter_rows():

        values = []

        for cell in row:

            value = cell.value

            if value is None:
                values.append("")
            else:
                values.append(
                    str(value).strip().upper()
                )

        if (
            "NAME" in values
            or "SUBJECT" in values
            or "REMARKS" in values
        ):
            header_row = row[0].row
            break

    if header_row is not None:

        meaningful_columns = []

        # Find columns containing actual values
        for column in range(
            1,
            worksheet.max_column + 1,
        ):

            value = worksheet.cell(
                header_row,
                column,
            ).value

            if value is not None:
                meaningful_columns.append(
                    column
                )

        # Include columns covered by merged cells
        for merged_range in worksheet.merged_cells.ranges:

            if (
                merged_range.min_row
                <= header_row
                <= merged_range.max_row
            ):

                meaningful_columns.append(
                    merged_range.max_col
                )

        if meaningful_columns:

            max_col = max(
                meaningful_columns
            )

    # =========================================================
    # 5. CREATE WORD DOCUMENT
    # =========================================================

    document = Document()

    section = document.sections[0]

    # =========================================================
    # 6. PAGE SIZE
    # =========================================================

    orientation = worksheet.page_setup.orientation
    paper_size = worksheet.page_setup.paperSize

    paper_sizes = {
        "1": (8.5, 11),       # Letter
        "2": (8.5, 11),       # Letter
        "3": (11.69, 16.54),  # A3
        "5": (8.27, 11.69),   # A4
        "6": (8.5, 13),       # Legal
        "7": (8.5, 14),       # Legal
    }

    if str(paper_size) in paper_sizes:

        width, height = paper_sizes[
            str(paper_size)
        ]

        if orientation == "landscape":

            width, height = (
                height,
                width,
            )

        section.page_width = Inches(
            width
        )

        section.page_height = Inches(
            height
        )

    # =========================================================
    # 7. PAGE MARGINS
    # =========================================================

    margins = worksheet.page_margins

    if margins.left is not None:

        section.left_margin = Inches(
            margins.left
        )

    if margins.right is not None:

        section.right_margin = Inches(
            margins.right
        )

    if margins.top is not None:

        section.top_margin = Inches(
            margins.top
        )

    if margins.bottom is not None:

        section.bottom_margin = Inches(
            margins.bottom
        )

    # =========================================================
    # 8. COLUMN WIDTHS
    # =========================================================

    column_widths = {}

    for column in range(
        min_col,
        max_col + 1,
    ):

        # IMPORTANT:
        # Use the column number directly.
        # This avoids MergedCell.column_letter errors.

        letter = get_column_letter(
            column
        )

        width = worksheet.column_dimensions[
            letter
        ].width

        if width is None:
            width = 8.43

        # Excel column width -> approximate Word inches
        column_widths[column] = (
            width * 0.095
        )

    # =========================================================
    # 9. CREATE WORD TABLE
    # =========================================================

    table = document.add_table(
        rows=max_row,
        cols=max_col,
    )

    table.autofit = False

    # =========================================================
    # 10. COPY EXCEL CELLS
    # =========================================================

    for row_number in range(
        min_row,
        max_row + 1,
    ):

        excel_row = worksheet.row_dimensions[
            row_number
        ]

        word_row = table.rows[
            row_number - 1
        ]

        # -----------------------------------------------------
        # ROW HEIGHT
        # -----------------------------------------------------

        if excel_row.height:

            word_row.height = Pt(
                excel_row.height
            )

        # -----------------------------------------------------
        # CELLS
        # -----------------------------------------------------

        for column_number in range(
            min_col,
            max_col + 1,
        ):

            excel_cell = worksheet.cell(
                row_number,
                column_number,
            )

            word_cell = word_row.cells[
                column_number - 1
            ]

            # =================================================
            # VALUE
            # =================================================

            value = excel_cell.value

            if value is None:
                value = ""

            word_cell.text = str(
                value
            )

            # =================================================
            # WIDTH
            # =================================================

            word_cell.width = Inches(
                column_widths[
                    column_number
                ]
            )

            # =================================================
            # ALIGNMENT
            # =================================================

            excel_alignment = (
                excel_cell.alignment
            )

            paragraph = (
                word_cell.paragraphs[0]
            )

            horizontal = (
                excel_alignment.horizontal
            )

            if horizontal == "center":

                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                )

            elif horizontal == "right":

                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.RIGHT
                )

            else:

                paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT
                )

            vertical = (
                excel_alignment.vertical
            )

            if vertical == "center":

                word_cell.vertical_alignment = (
                    WD_CELL_VERTICAL_ALIGNMENT.CENTER
                )

            elif vertical == "bottom":

                word_cell.vertical_alignment = (
                    WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
                )

            else:

                word_cell.vertical_alignment = (
                    WD_CELL_VERTICAL_ALIGNMENT.TOP
                )

            # =================================================
            # FONT
            # =================================================

            excel_font = excel_cell.font

            for run in paragraph.runs:

                if excel_font.name:

                    run.font.name = (
                        excel_font.name
                    )

                if excel_font.sz:

                    run.font.size = Pt(
                        excel_font.sz
                    )

                run.bold = bool(
                    excel_font.bold
                )

                run.italic = bool(
                    excel_font.italic
                )

                if excel_font.underline:

                    run.underline = True

            # =================================================
            # CELL BACKGROUND
            # =================================================

            fill = excel_cell.fill

            if (
                fill
                and fill.fill_type
                and fill.fgColor
            ):

                color = fill.fgColor

                # Only process real RGB colors.
                # Theme/indexed colors are skipped.
                if (
                    color.type == "rgb"
                    and color.rgb
                ):

                    rgb_value = str(
                        color.rgb
                    )

                    # Excel normally uses ARGB:
                    # FFFFFFFF
                    #
                    # Word needs RGB:
                    # FFFFFF

                    if len(rgb_value) == 8:

                        rgb_value = (
                            rgb_value[-6:]
                        )

                    elif len(rgb_value) == 6:

                        rgb_value = rgb_value

                    else:

                        rgb_value = None

                    if rgb_value:

                        tc_pr = (
                            word_cell
                            ._tc
                            .get_or_add_tcPr()
                        )

                        shading = OxmlElement(
                            "w:shd"
                        )

                        shading.set(
                            qn("w:fill"),
                            rgb_value,
                        )

                        tc_pr.append(
                            shading
                        )

    # =========================================================
    # 11. MERGED CELLS
    # =========================================================

    for merged_range in (
        worksheet.merged_cells.ranges
    ):

        start_row = (
            merged_range.min_row - 1
        )

        start_col = (
            merged_range.min_col - 1
        )

        end_row = (
            merged_range.max_row - 1
        )

        end_col = (
            merged_range.max_col - 1
        )

        # Make sure the merge is inside
        # the generated Word table.

        if (
            start_row >= 0
            and start_col >= 0
            and end_row < len(table.rows)
            and end_col < len(table.columns)
        ):

            start_cell = table.cell(
                start_row,
                start_col,
            )

            end_cell = table.cell(
                end_row,
                end_col,
            )

            start_cell.merge(
                end_cell
            )

    # =========================================================
    # 12. COPY BORDERS
    # =========================================================

    for row_number in range(
        min_row,
        max_row + 1,
    ):

        for column_number in range(
            min_col,
            max_col + 1,
        ):

            excel_cell = worksheet.cell(
                row_number,
                column_number,
            )

            word_cell = table.cell(
                row_number - 1,
                column_number - 1,
            )

            tc_pr = (
                word_cell
                ._tc
                .get_or_add_tcPr()
            )

            borders = tc_pr.find(
                qn("w:tcBorders")
            )

            if borders is None:

                borders = OxmlElement(
                    "w:tcBorders"
                )

                tc_pr.append(
                    borders
                )

            excel_border = (
                excel_cell.border
            )

            border_map = {
                "top": excel_border.top,
                "left": excel_border.left,
                "bottom": excel_border.bottom,
                "right": excel_border.right,
            }

            for edge_name, edge in (
                border_map.items()
            ):

                if (
                    edge is None
                    or edge.style is None
                ):
                    continue

                border = OxmlElement(
                    f"w:{edge_name}"
                )

                border.set(
                    qn("w:val"),
                    "single",
                )

                border.set(
                    qn("w:sz"),
                    "4",
                )

                border.set(
                    qn("w:space"),
                    "0",
                )

                if (
                    edge.color
                    and edge.color.type == "rgb"
                    and edge.color.rgb
                ):

                    border_color = str(
                        edge.color.rgb
                    )

                    if len(border_color) == 8:

                        border_color = (
                            border_color[-6:]
                        )

                    elif len(border_color) != 6:

                        border_color = None

                    if border_color:

                        border.set(
                            qn("w:color"),
                            border_color,
                        )

                borders.append(
                    border
                )

    # =========================================================
    # 13. COPY EXCEL IMAGES / LOGOS
    # =========================================================

    for image in getattr(
        worksheet,
        "_images",
        [],
    ):

        try:

            image_data = image._data()

            anchor = image.anchor

            row = anchor._from.row
            column = anchor._from.col

            if (
                row < len(table.rows)
                and column < len(table.columns)
            ):

                cell = table.cell(
                    row,
                    column,
                )

                paragraph = (
                    cell.paragraphs[0]
                )

                run = paragraph.add_run()

                run.add_picture(
                    BytesIO(
                        image_data
                    )
                )

        except Exception:

            # If an Excel image cannot be
            # converted, continue generating
            # the rest of the report.
            pass

    # =========================================================
    # 14. SAVE EDITABLE DOCX
    # =========================================================

    output_path = os.path.join(
        temp_dir,
        "Monitoring_Report.docx",
    )

    document.save(
        output_path
    )

    return (
        output_path,
        temp_dir,
        template,
    )


# ==========================================
# REPORT PDF BUILDER
# ==========================================
def _get_template_extension(template):
    """
    Return the uploaded template file extension.
    """
    return os.path.splitext(template.file.name)[1].lower()

def _build_report_file(
    records,
    week_start,
    week_end,
    template_id,
):
    """
    Build the report using the selected
    uploaded Excel template.

    The Excel template is the master layout.
    """

    output_path, temp_dir, template = (
        _build_report_excel_file(
            records,
            week_start,
            week_end,
            template_id,
        )
    )

    return (
        output_path,
        temp_dir,
        template,
        "excel",
    )

def _build_report_pdf_bytes(
    records,
    week_start,
    week_end,
    template_id,
):
    """
    Generate PDF from the generated Excel report.

    There is NO ReportLab report layout here.

    Excel is the master template.
    LibreOffice performs the Excel -> PDF conversion.
    """

    import os
    import subprocess

    output_xlsx, temp_dir, template = (
        _build_report_excel_file(
            records,
            week_start,
            week_end,
            template_id,
        )
    )

    pdf_dir = os.path.join(
        temp_dir,
        "pdf"
    )

    os.makedirs(
        pdf_dir,
        exist_ok=True
    )

    # ==========================================================
    # FIND LIBREOFFICE
    # ==========================================================

    libreoffice_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]

    libreoffice_path = None

    for path in libreoffice_paths:

        if os.path.isfile(path):

            libreoffice_path = path
            break

    if not libreoffice_path:

        raise FileNotFoundError(
            "LibreOffice was not found. "
            "Please install LibreOffice so the Excel "
            "template can be converted to PDF."
        )

    # ==========================================================
    # EXCEL -> PDF
    # ==========================================================

    command = [
        libreoffice_path,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        pdf_dir,
        output_xlsx,
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=120,
        creationflags=getattr(
            subprocess,
            "CREATE_NO_WINDOW",
            0,
        ),
    )

    pdf_path = os.path.join(
        pdf_dir,
        "Monitoring_Report.pdf"
    )

    if not os.path.exists(pdf_path):

        raise RuntimeError(
            "LibreOffice could not create the PDF.\n\n"
            f"Output:\n{result.stdout}\n\n"
            f"Error:\n{result.stderr}"
        )

    with open(
        pdf_path,
        "rb"
    ) as pdf_file:

        pdf_bytes = pdf_file.read()

    return pdf_bytes

# ==========================================
# REPORTS PAGE
# ==========================================
@login_required
def reports(request):
    week_start, week_end, current_week_start = (
        _report_week_range(request)
    )

    all_records = (
        MonitoringRecord.objects
        .select_related(
            "schedule__instructor",
            "schedule__subject",
            "schedule__room",
        )
        .order_by("-record_datetime")
    )

    week_starts = {
        current_week_start
    }

    for record in all_records:

        record_date = timezone.localtime(
            record.record_datetime
        ).date()

        week_starts.add(
            record_date
            - timedelta(
                days=record_date.weekday()
            )
        )

    weeks = []

    for start in sorted(
        week_starts,
        reverse=True,
    ):

        end = start + timedelta(days=6)

        weeks.append({
            "value": start.strftime(
                "%Y-%m-%d"
            ),
            "label": (
                f"{start.strftime('%B %d, %Y')} - "
                f"{end.strftime('%B %d, %Y')}"
            ),
        })

    records, instructor_search = (
        _report_records(
            request,
            week_start,
            week_end,
        )
    )

    selected_classification = (
        _report_classification(request)
    )

    report_templates = (
        ReportTemplate.objects
        .all()
        .order_by("name")
    )

    return render(
        request,
        "instructors/reports/list.html",
        {
            "records": records,

            "weeks": weeks,

            "selected_week":
                week_start.strftime(
                    "%Y-%m-%d"
                ),

            "week_start":
                week_start,

            "week_end":
                week_end,

            "previous_week":
                week_start
                - timedelta(days=7),

            "next_week":
                week_start
                + timedelta(days=7),

            "current_week_start":
                current_week_start,

            "current_week_end":
                current_week_start
                + timedelta(days=6),

            "instructors":
                Instructor.objects
                .all()
                .order_by("name"),

            "instructor_search":
                instructor_search,

            "selected_classification":
                selected_classification,

            "report_settings":
                _report_settings(request),

            "report_templates":
                report_templates,

            "windows_printers":
                _windows_printers(),
        },
    )

# ==========================================
# REPORT TEMPLATES
# ==========================================


@login_required
def report_template_list(request):
    templates = ReportTemplate.objects.all().order_by(
        "-is_active",
        "-uploaded_at",
    )

    return render(
        request,
        "instructors/reports/template.html",
        {"templates": templates},
    )


@login_required
def report_template_create(request):

    if request.method == "POST":

        uploaded_file = request.FILES.get("file")

        # Only allow Excel .xlsx templates
        if not uploaded_file:
            messages.error(
                request,
                "Please select an Excel template."
            )
            return redirect("report_template_create")

        if not uploaded_file.name.lower().endswith(".xlsx"):
            messages.error(
                request,
                "Only Excel .xlsx report templates are allowed."
            )
            return redirect("report_template_create")

        form = ReportTemplateForm(
            request.POST,
            request.FILES,
        )

        if form.is_valid():

            template = form.save(
                commit=False
            )

            # If this template is being activated,
            # deactivate all other templates first.
            if template.is_active:

                ReportTemplate.objects.update(
                    is_active=False
                )

            template.save()

            messages.success(
                request,
                "Excel report template uploaded successfully."
            )

            return redirect(
                "report_template_list"
            )

    else:

        form = ReportTemplateForm()

    return render(
        request,
        "instructors/reports/template_form.html",
        {
            "form": form
        },
    )

@login_required
def report_template_update(request, pk):

    template = get_object_or_404(
        ReportTemplate,
        pk=pk
    )

    if request.method == "POST":

        form = ReportTemplateForm(
            request.POST,
            request.FILES,
            instance=template
        )

        if form.is_valid():

            updated_template = form.save(
                commit=False
            )

            if updated_template.is_active:

                ReportTemplate.objects.exclude(
                    pk=template.pk
                ).update(
                    is_active=False
                )

            updated_template.save()

            messages.success(
                request,
                "Report template updated successfully."
            )

            return redirect(
                "report_template_list"
            )

    else:

        form = ReportTemplateForm(
            instance=template
        )

    return render(
        request,
        "instructors/reports/template_form.html",
        {
            "form": form,
            "template": template,
            "is_edit": True,
        }
    )


@login_required
def report_template_delete(request, pk):

    template = get_object_or_404(
        ReportTemplate,
        pk=pk
    )

    if request.method == "POST":

        template.delete()

        messages.success(
            request,
            "Report template deleted successfully."
        )

    return redirect("report_template_list")


@login_required
def report_template_bulk_delete(request):

    if request.method == "POST":

        template_ids = request.POST.getlist(
            "selected_templates"
        )

        if template_ids:

            ReportTemplate.objects.filter(
                id__in=template_ids
            ).delete()

            messages.success(
                request,
                "Selected report templates were deleted successfully."
            )

        else:

            messages.warning(
                request,
                "Please select at least one report template."
            )

    return redirect("report_template_list")


# ==========================================
# REPORT PRINT PAGE / WINDOWS PRINTING
# ==========================================
@login_required
def report_print(request):
    """
    Generate the selected report using the selected
    Excel report template and send it directly
    to the selected Windows printer.
    """

    # ==========================================================
    # REPORT WEEK
    # ==========================================================

    week_start, week_end, _ = (
        _report_week_range(request)
    )

    # ==========================================================
    # REPORT DATA
    # ==========================================================

    records, instructor_search = (
        _report_records(
            request,
            week_start,
            week_end,
        )
    )

    # ==========================================================
    # SELECTED EXCEL TEMPLATE
    # ==========================================================

    template_id = request.GET.get(
        "template_id",
        "",
    ).strip()

    if not template_id:

        template_id = request.POST.get(
            "template_id",
            "",
        ).strip()

    if not template_id:

        return JsonResponse(
            {
                "success": False,
                "message": (
                    "Please choose a report layout "
                    "before printing."
                ),
            },
            status=400,
        )

    # ==========================================================
    # CHECK WINDOWS
    # ==========================================================

    if os.name != "nt":

        return JsonResponse(
            {
                "success": False,
                "message": (
                    "Direct printing is only available "
                    "when Django is running on Windows."
                ),
            },
            status=400,
        )

    # ==========================================================
    # SELECTED PRINTER
    # ==========================================================

    printer_name = request.GET.get(
        "printer",
        "",
    ).strip()

    if not printer_name:

        printer_name = request.POST.get(
            "printer",
            "",
        ).strip()

    if not printer_name:

        return JsonResponse(
            {
                "success": False,
                "message":
                    "Please select a printer.",
            },
            status=400,
        )

    # ==========================================================
    # VERIFY PRINTER
    # ==========================================================

    printers = _windows_printers()

    valid_printers = {
        item["name"]
        for item in printers
    }

    if printer_name not in valid_printers:

        return JsonResponse(
            {
                "success": False,
                "message": (
                    f"The printer '{printer_name}' "
                    "was not found."
                ),
            },
            status=400,
        )

    # ==========================================================
    # GENERATE PDF FROM SELECTED EXCEL TEMPLATE
    # ==========================================================

    temp_path = None

    try:

        temp_path = (
            _build_report_pdf_bytes(
                records,
                week_start,
                week_end,
                template_id,
            )
        )

        if not temp_path:

            return JsonResponse(
                {
                    "success": False,
                    "message": (
                        "The report PDF could not "
                        "be generated."
                    ),
                },
                status=500,
            )

        # ======================================================
        # FIND LIBREOFFICE
        # ======================================================

        libreoffice_paths = [
            (
                r"C:\Program Files\LibreOffice"
                r"\program\soffice.exe"
            ),
            (
                r"C:\Program Files (x86)"
                r"\LibreOffice\program\soffice.exe"
            ),
        ]

        libreoffice_path = None

        for path in libreoffice_paths:

            if os.path.isfile(path):

                libreoffice_path = path
                break

        # ======================================================
        # WINDOWS PRINTTO FALLBACK
        # ======================================================

        if not libreoffice_path:

            try:

                os.startfile(
                    temp_path,
                    "printto",
                    printer_name,
                )

                def remove_temp_file(path):

                    try:

                        if os.path.exists(path):

                            os.remove(path)

                    except Exception:
                        pass

                threading.Timer(
                    120,
                    remove_temp_file,
                    args=(temp_path,),
                ).start()

                return JsonResponse(
                    {
                        "success": True,
                        "message": (
                            f"The report was sent to "
                            f"'{printer_name}'."
                        ),
                        "printer":
                            printer_name,
                        "method":
                            "Windows PrintTo",
                    }
                )

            except Exception as exc:

                return JsonResponse(
                    {
                        "success": False,
                        "message": (
                            "LibreOffice was not found "
                            "and Windows could not "
                            "print the PDF."
                        ),
                        "detail": str(exc),
                    },
                    status=500,
                )

        # ======================================================
        # LIBREOFFICE PRINT
        # ======================================================

        command = [
            libreoffice_path,
            "--headless",
            "--pt",
            printer_name,
            temp_path,
        ]

        try:

            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=120,
                creationflags=getattr(
                    subprocess,
                    "CREATE_NO_WINDOW",
                    0,
                ),
            )

        except subprocess.TimeoutExpired:

            return JsonResponse(
                {
                    "success": False,
                    "message": (
                        "Printing timed out while "
                        "LibreOffice was processing "
                        "the report."
                    ),
                },
                status=500,
            )

        # ======================================================
        # PRINT RESULT
        # ======================================================

        if result.returncode != 0:

            error_text = (
                result.stderr.strip()
                or result.stdout.strip()
                or "LibreOffice returned an unknown error."
            )

            return JsonResponse(
                {
                    "success": False,
                    "message": (
                        f"The report could not be "
                        f"printed to '{printer_name}'."
                    ),
                    "detail": error_text,
                },
                status=500,
            )

        # ======================================================
        # CLEANUP
        # ======================================================

        def remove_temp_file(path):

            try:

                if os.path.exists(path):

                    os.remove(path)

            except Exception:
                pass

        threading.Timer(
            10,
            remove_temp_file,
            args=(temp_path,),
        ).start()

        # ======================================================
        # SUCCESS
        # ======================================================

        return JsonResponse(
            {
                "success": True,
                "message": (
                    f"The report was sent successfully "
                    f"to '{printer_name}'."
                ),
                "printer": printer_name,
                "method": "LibreOffice",
            }
        )

    except Exception as exc:

        if temp_path:

            try:

                if os.path.exists(temp_path):

                    os.remove(temp_path)

            except Exception:
                pass

        return JsonResponse(
            {
                "success": False,
                "message": (
                    "The report could not be printed."
                ),
                "detail": str(exc),
            },
            status=500,
        )

# ==========================================
# REPORT PDF DOWNLOAD
# ==========================================
@login_required
def report_pdf(request):
    """
    Generate a PDF from the Excel report layout
    selected by the user.
    """

    week_start, week_end, _ = (
        _report_week_range(request)
    )

    records, _ = _report_records(
        request,
        week_start,
        week_end,
    )

    # ==========================================================
    # SELECT TEMPLATE
    # ==========================================================

    template_id = request.GET.get(
        "template_id",
        "",
    ).strip()

    if not template_id:

        template_id = request.POST.get(
            "template_id",
            "",
        ).strip()

    if not template_id:

        return HttpResponse(
            (
                "Please choose a report layout "
                "before generating the PDF."
            ),
            status=400,
        )

    # ==========================================================
    # BUILD EXCEL USING SELECTED TEMPLATE
    # ==========================================================

    try:

        output_xlsx, temp_dir, template = (
            _build_report_excel_file(
                records,
                week_start,
                week_end,
                template_id,
            )
        )

    except Exception as exc:

        return HttpResponse(
            (
                "Report could not be generated: "
                f"{exc}"
            ),
            status=500,
        )

    # ==========================================================
    # PDF DIRECTORY
    # ==========================================================

    pdf_dir = os.path.join(
        temp_dir,
        "pdf",
    )

    os.makedirs(
        pdf_dir,
        exist_ok=True,
    )

    # ==========================================================
    # LIBREOFFICE
    # ==========================================================

    libreoffice_paths = [
        (
            r"C:\Program Files\LibreOffice"
            r"\program\soffice.exe"
        ),
        (
            r"C:\Program Files (x86)"
            r"\LibreOffice\program\soffice.exe"
        ),
    ]

    libreoffice_path = None

    for path in libreoffice_paths:

        if os.path.isfile(path):

            libreoffice_path = path
            break

    if not libreoffice_path:

        shutil.rmtree(
            temp_dir,
            ignore_errors=True,
        )

        return HttpResponse(
            "LibreOffice was not found.",
            status=500,
        )

    # ==========================================================
    # EXCEL → PDF
    # ==========================================================

    command = [
        libreoffice_path,
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        pdf_dir,
        output_xlsx,
    ]

    try:

        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=getattr(
                subprocess,
                "CREATE_NO_WINDOW",
                0,
            ),
        )

    except subprocess.TimeoutExpired:

        shutil.rmtree(
            temp_dir,
            ignore_errors=True,
        )

        return HttpResponse(
            "Excel-to-PDF conversion timed out.",
            status=500,
        )

    # ==========================================================
    # FIND PDF
    # ==========================================================

    pdf_path = os.path.join(
        pdf_dir,
        "Monitoring_Report.pdf",
    )

    if not os.path.exists(pdf_path):

        shutil.rmtree(
            temp_dir,
            ignore_errors=True,
        )

        return HttpResponse(
            (
                "LibreOffice could not create the PDF.\n\n"
                f"Output:\n{result.stdout}\n\n"
                f"Error:\n{result.stderr}"
            ),
            status=500,
        )

    # ==========================================================
    # RETURN PDF
    # ==========================================================

    response = FileResponse(
        open(
            pdf_path,
            "rb",
        ),
        content_type="application/pdf",
    )

    response["Content-Disposition"] = (
        "attachment; "
        f'filename="Monitoring_Report_'
        f'{week_start.strftime("%Y-%m-%d")}.pdf"'
    )

    return response

# ==========================================
# REPORT DOCX
# ==========================================

@login_required
def report_docx(request):

    week_start, week_end, _ = _report_week_range(request)

    records, _ = _report_records(
        request,
        week_start,
        week_end,
    )

    output_path, temp_dir, template = (
        _build_report_docx_file(
            records,
            week_start,
            week_end,
        )
    )

    with open(output_path, "rb") as file:

        response = HttpResponse(
            file.read(),
            content_type=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )

    response["Content-Disposition"] = (
        "attachment; "
        f'filename="Monitoring_Report_'
        f'{week_start.strftime("%Y-%m-%d")}.docx"'
    )

    shutil.rmtree(
        temp_dir,
        ignore_errors=True,
    )

    return response


# ==========================================
# REPORT EXCEL
# ==========================================
@login_required
def report_excel(request):

    week_start, week_end, _ = (
        _report_week_range(request)
    )

    records, _ = _report_records(
        request,
        week_start,
        week_end,
    )

    # ==========================================================
    # SELECT TEMPLATE
    # ==========================================================

    template_id = request.GET.get(
        "template_id",
        "",
    ).strip()

    if not template_id:

        template_id = request.POST.get(
            "template_id",
            "",
        ).strip()

    if not template_id:

        return HttpResponse(
            (
                "Please choose a report layout "
                "before generating the Excel report."
            ),
            status=400,
        )

    # ==========================================================
    # BUILD REPORT
    # ==========================================================

    try:

        output_path, temp_dir, template = (
            _build_report_excel_file(
                records,
                week_start,
                week_end,
                template_id,
            )
        )

    except Exception as exc:

        return HttpResponse(
            (
                "Report could not be generated: "
                f"{exc}"
            ),
            status=500,
        )

    # ==========================================================
    # RETURN EXCEL
    # ==========================================================

    response = FileResponse(
        open(
            output_path,
            "rb",
        ),
        content_type=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )

    response["Content-Disposition"] = (
        "attachment; "
        f'filename="Monitoring_Report_'
        f'{week_start.strftime("%Y-%m-%d")}.xlsx"'
    )

    return response
# UNIVERSAL SCHEDULE IMPORTER
# Supports Excel, XLSM, XLS, CSV, PDF, DOCX, PPTX, TXT and images.
# ============================================================


def _normalize_text(value):
    if value is None:
        return ""
    return " ".join(str(value).replace("\u00a0", " ").strip().split())


def _canonical_text(value):
    return re.sub(r"[^A-Z0-9]+", "", _normalize_text(value).upper())


def _normalize_room_name(value):
    value = _normalize_text(value)
    value = re.sub(r"\s*,\s*", ", ", value)
    value = re.sub(r"\s*/\s*", " / ", value)
    return value or "UNASSIGNED"


def _extract_instructor_and_room(value):
    raw = str(value or "").replace("\u00a0", " ").strip()
    parts = re.split(r"\s{2,}", raw)
    if len(parts) >= 2:
        return _normalize_text(parts[0]), _normalize_room_name(" ".join(parts[1:]))
    return _normalize_text(raw), "UNASSIGNED"


def _parse_subject(value):
    return _normalize_text(value)


_TIME_RE = re.compile(
    r"(?P<start>\d{1,2}(?::\d{2})?\s*(?:A\.?M\.?|P\.?M\.?|AM|PM))"
    r"\s*(?:-|–|—|TO)\s*"
    r"(?P<end>\d{1,2}(?::\d{2})?\s*(?:A\.?M\.?|P\.?M\.?|AM|PM))",
    re.IGNORECASE,
)


def _parse_time_value(value):
    value = _normalize_text(value).upper().replace(".", "")
    for fmt in ("%I:%M %p", "%I %p", "%H:%M"):
        try:
            return datetime.strptime(value, fmt).time()
        except ValueError:
            continue
    return None


def _parse_excel_time_range(value):
    if value is None or value == "":
        return None, None
    if isinstance(value, time):
        return value, None
    if isinstance(value, datetime):
        return value.time(), None
    raw = _normalize_text(value).replace("–", "-").replace("—", "-")
    match = _TIME_RE.search(raw)
    if not match:
        return None, None
    return _parse_time_value(match.group("start")), _parse_time_value(match.group("end"))


def _normalize_day(value):
    raw = _normalize_text(value).upper().replace(".", "")
    raw = raw.replace("–", "-").replace("—", "-").replace("/", "-")
    day_map = {
        "MONDAY": "M", "MON": "M", "M": "M",
        "TUESDAY": "T", "TUE": "T", "TUES": "T", "TU": "T", "T": "T",
        "WEDNESDAY": "W", "WED": "W", "WE": "W", "W": "W",
        "THURSDAY": "TH", "THU": "TH", "THURS": "TH", "THUR": "TH", "TH": "TH", "R": "TH",
        "FRIDAY": "F", "FRI": "F", "F": "F",
        "SATURDAY": "S", "SAT": "S", "S": "S",
    }
    return day_map.get(raw.replace(" ", ""), "")


def _expand_days(value):
    raw = _normalize_text(value).upper().replace(".", "")
    raw = raw.replace("–", "-").replace("—", "-").replace("/", "-").replace(" ", "")
    if not raw:
        return []

    special = {
        "T-TH": ["T", "TH"],
        "TTH": ["T", "TH"],
        "W-F": ["W", "F"],
        "WF": ["W", "F"],
        "M-W-F": ["M", "W", "F"],
        "MWF": ["M", "W", "F"],
        "M-W": ["M", "W"],
        "MW": ["M", "W"],
        "T-TH-S": ["T", "TH", "S"],
        "TTHS": ["T", "TH", "S"],
        "THS": ["TH", "S"],
    }
    if raw in special:
        return special[raw]

    if "-" in raw:
        result = []
        for part in raw.split("-"):
            day = _normalize_day(part)
            if day and day not in result:
                result.append(day)
        return result

    day = _normalize_day(raw)
    if day:
        return [day]

    # Compact forms such as MWF, TTH, TTHS.
    result = []
    remaining = raw
    while remaining:
        if remaining.startswith("TH"):
            code = "TH"
            remaining = remaining[2:]
        else:
            code = remaining[0]
            remaining = remaining[1:]
        if code not in {"M", "T", "W", "TH", "F", "S"} or code in result:
            return []
        result.append(code)
    return result


def _find_days_in_text(value):
    text = _normalize_text(value)
    if not text:
        return []
    upper = text.upper()
    if sum(name in upper for name in ("MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY")) >= 4:
        return []

    patterns = [
        r"\bT-TH-S\b", r"\bM-W-F\b", r"\bT-TH\b", r"\bW-F\b", r"\bM-W\b",
        r"\bTTHS\b", r"\bTTH\b", r"\bMWF\b", r"\bMW\b", r"\bWF\b",
        r"\bMONDAY\b", r"\bTUESDAY\b", r"\bWEDNESDAY\b", r"\bTHURSDAY\b", r"\bFRIDAY\b", r"\bSATURDAY\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, upper)
        if match:
            days = _expand_days(match.group(0))
            if days:
                return days

    # A standalone abbreviated day cell.
    for token in re.findall(r"\b(?:M|T|W|TH|F|S)\b", upper):
        days = _expand_days(token)
        if days:
            return days
    return []


def _extract_name_from_text(text):
    text = str(text or "")
    patterns = [
        r"Name\s*:\s*(.+?)(?=\s+Acad\.?\s*Rank\s*:|\s+Address\s*:|\s+Department\s*:|\n|\r|$)",
        r"NAME\s*[:\-]\s*([^\n\r]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            name = _normalize_text(match.group(1)).strip(".,")
            if name:
                return name
    return ""


def _make_import_row(instructor, subject, day, start_time, end_time, room, row_number=0):
    return {
        "row_number": row_number,
        "instructor": _normalize_text(instructor),
        "subject": _parse_subject(subject),
        "day": day,
        "start_time": start_time.strftime("%H:%M") if isinstance(start_time, time) else _normalize_text(start_time),
        "end_time": end_time.strftime("%H:%M") if isinstance(end_time, time) else _normalize_text(end_time),
        "room": _normalize_room_name(room),
        "time_error": "",
    }

def _parse_uploaded_schedule_excel(uploaded_file):
    """
    Reads the actual Regular Monitoring Excel format.

    Structure of the Excel:

        NAME
        SUBJECT | TIME | DAY

        Instructor Name | Subject | Time | Day
                         Subject | Time
                         Subject | Time
                         Subject | Time | New Day

    The instructor is carried down until a new instructor appears.

    The day is carried down until a new day appears.

    Room information attached to the instructor name is ignored.
    """

    workbook = load_workbook(
        uploaded_file,
        data_only=True
    )

    rows = []

    # ==========================================================
    # PROCESS EVERY SHEET
    # ==========================================================

    for worksheet in workbook.worksheets:

        current_instructor = ""
        current_day = ""

        # ======================================================
        # READ EVERY ROW
        # ======================================================

        for row_number in range(
            1,
            worksheet.max_row + 1
        ):

            # --------------------------------------------------
            # Read first four columns:
            #
            # A = Instructor
            # B = Subject
            # C = Time
            # D = Day
            # --------------------------------------------------

            instructor_value = worksheet.cell(
                row_number,
                1
            ).value

            subject_value = worksheet.cell(
                row_number,
                2
            ).value

            time_value = worksheet.cell(
                row_number,
                3
            ).value

            day_value = worksheet.cell(
                row_number,
                4
            ).value

            instructor_text = _normalize_text(
                instructor_value
            )

            subject_text = _normalize_text(
                subject_value
            )

            day_text = _normalize_text(
                day_value
            )

            # ==================================================
            # IGNORE COMPLETELY EMPTY ROWS
            # ==================================================

            if not any(
                [
                    instructor_text,
                    subject_text,
                    _normalize_text(time_value),
                    day_text,
                ]
            ):
                continue

            # ==================================================
            # NEW INSTRUCTOR
            # ==================================================

            if instructor_text:

                upper_instructor = (
                    instructor_text.upper().strip()
                )

                # ----------------------------------------------
                # Ignore repeated section header
                # ----------------------------------------------

                if upper_instructor == "NAME":
                    current_instructor = ""
                    current_day = ""
                    continue

                # ----------------------------------------------
                # Ignore the title/header area
                # ----------------------------------------------

                if (
                    "OFFICE OF THE DEAN" in upper_instructor
                    or "MONITORING OF CLASSES" in upper_instructor
                    or "PART-TIME INSTRUCTORS" in upper_instructor
                    or "SCHOOL YEAR" in upper_instructor
                ):
                    continue

                # ----------------------------------------------
                # Detect actual instructor.
                #
                # Instructor names in this file are formatted
                # like:
                #
                # ABEQUIBEL, FRANCO T.     EDTECH ROOM
                #
                # We only keep the instructor's name.
                # The room is intentionally ignored.
                # ----------------------------------------------

                possible_name, possible_room = (
                    _extract_instructor_and_room(
                        instructor_text
                    )
                )

                if (
                    possible_name
                    and "," in possible_name
                ):
                    current_instructor = (
                        possible_name.strip()
                    )

            # ==================================================
            # NEW DAY
            # ==================================================

            if day_text:

                normalized_day = _normalize_day(
                    day_text
                )

                if normalized_day:
                    current_day = normalized_day

                else:

                    # Try the more flexible day detector.
                    detected_days = (
                        _find_days_in_text(
                            day_text
                        )
                    )

                    if detected_days:
                        current_day = (
                            detected_days[0]
                        )

            # ==================================================
            # WE CANNOT CREATE A SCHEDULE WITHOUT:
            #
            # Instructor
            # Subject
            # Day
            # Time
            # ==================================================

            if not current_instructor:
                continue

            if not subject_text:
                continue

            if not current_day:
                continue

            # ==================================================
            # IGNORE HEADER ROWS
            # ==================================================

            if subject_text.upper() in {
                "SUBJECT",
                "SUBJECT W/ SECTION",
            }:
                continue

            # ==================================================
            # READ CLASS TIME
            # ==================================================

            start_time, end_time = (
                _parse_excel_time_range(
                    time_value
                )
            )

            if not start_time or not end_time:
                continue

            # ==================================================
            # ROOM
            #
            # IMPORTANT:
            #
            # The Excel contains room information attached to
            # instructor names, such as:
            #
            # ABEQUIBEL, FRANCO T. EDTECH ROOM
            #
            # We DO NOT use that as the assigned room.
            #
            # Leave room blank.
            # ==================================================

            room = ""

            # ==================================================
            # CREATE IMPORT ROW
            # ==================================================

            rows.append(
                _make_import_row(
                    current_instructor,
                    subject_text,
                    current_day,
                    start_time,
                    end_time,
                    room,
                    row_number,
                )
            )

    return rows

def _extract_schedule_rows_from_text(text, default_instructor=""):
    rows = []
    lines = [_normalize_text(line) for line in str(text or "").splitlines() if _normalize_text(line)]
    instructor = default_instructor or _extract_name_from_text(text)

    for index, line in enumerate(lines):
        time_match = _TIME_RE.search(line)
        if not time_match:
            continue

        start_time = _parse_time_value(time_match.group("start"))
        end_time = _parse_time_value(time_match.group("end"))
        if not start_time or not end_time:
            continue

        combined = line
        days = _find_days_in_text(combined)
        if not days:
            for nearby in lines[max(0, index - 2):min(len(lines), index + 3)]:
                candidate_days = _find_days_in_text(nearby)
                if candidate_days:
                    days = candidate_days
                    combined = f"{nearby} {combined}"
                    break
        if not days:
            continue

        day_match = None
        for pattern in [r"T-TH-S|M-W-F|T-TH|W-F|M-W|TTHS|TTH|MWF|MW|WF|MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY"]:
            day_match = re.search(pattern, combined, re.IGNORECASE)
            if day_match:
                break
        if day_match:
            subject = _normalize_text(combined[:day_match.start()]).strip(" |:;-_")
        else:
            subject = _normalize_text(combined[:time_match.start()]).strip(" |:;-_")

        subject = re.sub(r"^(?:SUBJECT(?:\s+W/\s+SECTION)?|DAY|TIME)\s*[:\-]?\s*", "", subject, flags=re.IGNORECASE)
        if not subject or subject.upper() in {"TOTAL", "CONSULTATION PERIOD"}:
            continue

        after_time = _normalize_text(line[time_match.end():]).strip(" |:;-_")
        room = after_time if after_time else "UNASSIGNED"
        # Prefer known room forms when OCR has trailing unrelated text.
        room_matches = re.findall(
            r"(?:CAS|CIT|IT|LAB|RM|ROOM)\s*[A-Z0-9\-]+|SPEECH\s+LAB(?:ORATORY)?|SCIENCE\s+LAB(?:ORATORY)?|COM(?:PUTER)?\s*LAB(?:ORATORY)?|FACULTY\s+ROOM|AVR|LIBRARY",
            line,
            flags=re.IGNORECASE,
        )
        if room_matches:
            room = ", ".join(dict.fromkeys(_normalize_text(x) for x in room_matches))

        for day in days:
            rows.append(_make_import_row(instructor, subject, day, start_time, end_time, room, index + 1))
    return rows


def _parse_uploaded_csv(uploaded_file):
    import csv
    import io
    content = uploaded_file.read()
    text = content.decode("utf-8-sig", errors="ignore")
    rows = []
    for row_number, row in enumerate(csv.reader(io.StringIO(text)), start=1):
        values = [_normalize_text(v) for v in row]
        if not any(values):
            continue
        time_index = next((i for i, v in enumerate(values) if _parse_excel_time_range(v)[0] and _parse_excel_time_range(v)[1]), None)
        day_index = next((i for i, v in enumerate(values) if _find_days_in_text(v)), None)
        if time_index is None or day_index is None:
            continue
        start_time, end_time = _parse_excel_time_range(values[time_index])
        subject_candidates = [v for i, v in enumerate(values[:time_index]) if v and i != day_index]
        if not subject_candidates:
            continue
        instructor = values[0]
        room = values[-1] or "UNASSIGNED"
        for day in _find_days_in_text(values[day_index]):
            rows.append(_make_import_row(instructor, subject_candidates[-1], day, start_time, end_time, room, row_number))
    return rows


def _parse_uploaded_docx(uploaded_file):
    document = Document(uploaded_file)
    full_text = "\n".join(p.text for p in document.paragraphs)
    instructor = _extract_name_from_text(full_text)
    rows = []
    for row_number, row in enumerate((r for table in document.tables for r in table.rows), start=1):
        values = [_normalize_text(cell.text) for cell in row.cells]
        if len(values) < 3:
            continue
        time_index = next((i for i, v in enumerate(values) if _parse_excel_time_range(v)[0] and _parse_excel_time_range(v)[1]), None)
        day_index = next((i for i, v in enumerate(values) if _find_days_in_text(v)), None)
        if time_index is None or day_index is None:
            continue
        start_time, end_time = _parse_excel_time_range(values[time_index])
        candidates = [v for i, v in enumerate(values[:time_index]) if v and i != day_index]
        if not candidates:
            continue
        for day in _find_days_in_text(values[day_index]):
            rows.append(_make_import_row(instructor, candidates[-1], day, start_time, end_time, values[-1], row_number))
    return rows or _extract_schedule_rows_from_text(full_text, instructor)


def _configure_tesseract():
    import pytesseract
    possible_paths = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for path in possible_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return path
    path_from_windows = shutil.which("tesseract")
    if path_from_windows:
        pytesseract.pytesseract.tesseract_cmd = path_from_windows
        return path_from_windows
    return ""


def _ocr_pdf_page(image):
    """OCR a rendered PDF page and return reconstructed visual lines."""
    import pytesseract
    from pytesseract import Output

    data = pytesseract.image_to_data(
        image,
        output_type=Output.DICT,
        config="--psm 6",
    )

    groups = {}
    count = len(data.get("text", []))
    for i in range(count):
        text = _normalize_text(data["text"][i])
        if not text:
            continue
        try:
            confidence = float(data["conf"][i])
        except Exception:
            confidence = 0
        if confidence < 15:
            continue

        key = (
            data.get("block_num", [0] * count)[i],
            data.get("par_num", [0] * count)[i],
            data.get("line_num", [0] * count)[i],
        )
        left = int(data["left"][i])
        top = int(data["top"][i])
        width = int(data["width"][i])
        height = int(data["height"][i])
        groups.setdefault(key, []).append({
            "text": text,
            "left": left,
            "top": top,
            "right": left + width,
            "bottom": top + height,
        })

    lines = []
    for words in groups.values():
        words.sort(key=lambda w: w["left"])
        lines.append({
            "text": _normalize_text(" ".join(w["text"] for w in words)),
            "top": min(w["top"] for w in words),
            "bottom": max(w["bottom"] for w in words),
            "left": min(w["left"] for w in words),
            "right": max(w["right"] for w in words),
        })
    lines.sort(key=lambda x: (x["top"], x["left"]))
    return lines


def _parse_part_time_teaching_load_pdf(uploaded_file):
    """Parse scanned/image-based part-time teaching-load PDFs using OCR."""
    import fitz
    import pytesseract
    from PIL import Image

    tesseract_path = _configure_tesseract()
    if not tesseract_path:
        raise RuntimeError(
            "Tesseract OCR was not found. Expected it at "
            r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        )

    uploaded_file.seek(0)
    pdf_bytes = uploaded_file.read()
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    imported_rows = []

    for page_number, page in enumerate(document, start=1):
        # 3x is enough for most scanned forms while keeping OCR manageable.
        pixmap = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=False)
        image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
        lines = _ocr_pdf_page(image)
        if not lines:
            continue

        page_text = "\n".join(line["text"] for line in lines)
        instructor = _extract_name_from_text(page_text)

        # Locate schedule rows by their time ranges. Teaching-load PDFs normally
        # put Subject, Day, Time and Room on the same visual row.
        for line_index, line in enumerate(lines):
            row_text = line["text"]
            time_match = _TIME_RE.search(row_text)
            if not time_match:
                continue

            start_time = _parse_time_value(time_match.group("start"))
            end_time = _parse_time_value(time_match.group("end"))
            if not start_time or not end_time:
                continue

            # Combine a small vertical neighborhood in case OCR split the row.
            neighborhood = [row_text]
            for j in range(max(0, line_index - 2), min(len(lines), line_index + 3)):
                if j != line_index and abs(lines[j]["top"] - line["top"]) < 180:
                    neighborhood.append(lines[j]["text"])
            combined = _normalize_text(" ".join(neighborhood))

            if re.search(r"\b(?:TOTAL|CONSULTATION\s+PERIOD)\b", combined, re.IGNORECASE):
                continue

            days = _find_days_in_text(combined)
            if not days:
                continue

            # Subject = text before the first day notation on the same/neighborhood line.
            day_match = None
            for pattern in [r"T-TH-S|M-W-F|T-TH|W-F|M-W|TTHS|TTH|MWF|MW|WF|MONDAY|TUESDAY|WEDNESDAY|THURSDAY|FRIDAY|SATURDAY"]:
                day_match = re.search(pattern, combined, re.IGNORECASE)
                if day_match:
                    break

            if day_match:
                subject = _normalize_text(combined[:day_match.start()]).strip(" |:;-_")
            else:
                subject = _normalize_text(combined[:time_match.start()]).strip(" |:;-_")

            # Remove common table labels and instructor header fragments.
            subject = re.sub(r"^(?:SUBJECT(?:\s+W/\s+SECTION)?|DAY|TIME|ROOM)\s*[:\-]?\s*", "", subject, flags=re.IGNORECASE)
            if not subject:
                continue

            # If the text before the day contains an instructor name, remove it.
            if instructor and _canonical_text(instructor) in _canonical_text(subject):
                subject = re.sub(re.escape(instructor), "", subject, flags=re.IGNORECASE).strip(" |:;-_")

            # Require a plausible subject code OR meaningful course text.
            if not re.search(r"\b[A-Z]{1,8}\s*\d{1,4}\b", subject, re.IGNORECASE):
                # OCR may split a code, so allow subject text if it is not a header.
                if subject.upper() in {"SUBJECT", "DAY", "TIME", "ROOM"} or len(subject) < 3:
                    continue

            # Extract all known room tokens from the row. Multiple rooms remain ONE string.
            room_matches = re.findall(
                r"(?:CAS|CIT|IT|LAB|RM|ROOM)\s*[A-Z0-9\-]+|"
                r"SPEECH\s+LAB(?:ORATORY)?|SCIENCE\s+LAB(?:ORATORY)?|"
                r"COM(?:PUTER)?\s*LAB(?:ORATORY)?|FACULTY\s+ROOM|AVR|LIBRARY|"
                r"[A-Z]{2,}\s+\d{1,3}",
                combined,
                flags=re.IGNORECASE,
            )
            cleaned_rooms = []
            for room in room_matches:
                room = _normalize_room_name(room)
                if room.upper() not in {"SUBJECT", "DAY", "TIME", "ROOM"} and room not in cleaned_rooms:
                    cleaned_rooms.append(room)
            room = ", ".join(cleaned_rooms) if cleaned_rooms else "UNASSIGNED"

            # If OCR did not find a room on the combined line, use text after the time.
            if room == "UNASSIGNED":
                after = _normalize_text(row_text[time_match.end():]).strip(" |:;-_")
                if after and not _find_days_in_text(after):
                    room = _normalize_room_name(after)

            for day in days:
                imported_rows.append(_make_import_row(
                    instructor,
                    subject,
                    day,
                    start_time,
                    end_time,
                    room,
                    page_number,
                ))

    return imported_rows


def _parse_uploaded_pdf(uploaded_file):
    """Use OCR first for scanned teaching-load PDFs, then fall back to PDF text/tables."""
    imported_rows = []
    ocr_error = None

    try:
        imported_rows = _parse_part_time_teaching_load_pdf(uploaded_file)
    except Exception as exc:
        ocr_error = str(exc)

    if not imported_rows:
        try:
            import pdfplumber
            uploaded_file.seek(0)
            pdf_bytes = uploaded_file.read()
            text_parts = []
            tables = []
            with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text_parts.append(page.extract_text() or "")
                    try:
                        tables.extend(page.extract_tables() or [])
                    except Exception:
                        pass

            text = "\n".join(text_parts)
            instructor = _extract_name_from_text(text)

            for table in tables:
                for row_number, row in enumerate(table, start=1):
                    values = [_normalize_text(v) for v in (row or [])]
                    time_index = next((i for i, v in enumerate(values) if _parse_excel_time_range(v)[0] and _parse_excel_time_range(v)[1]), None)
                    day_index = next((i for i, v in enumerate(values) if _find_days_in_text(v)), None)
                    if time_index is None or day_index is None:
                        continue
                    start_time, end_time = _parse_excel_time_range(values[time_index])
                    candidates = [v for i, v in enumerate(values[:time_index]) if v and i != day_index]
                    if not candidates:
                        continue
                    for day in _find_days_in_text(values[day_index]):
                        imported_rows.append(_make_import_row(instructor, candidates[-1], day, start_time, end_time, values[-1], row_number))

            if not imported_rows:
                imported_rows = _extract_schedule_rows_from_text(text, instructor)
        except Exception as exc:
            if not ocr_error:
                ocr_error = str(exc)

    cleaned = []
    seen = set()
    for row in imported_rows:
        row = {
            "instructor": _normalize_text(row.get("instructor", "")),
            "subject": _normalize_text(row.get("subject", "")),
            "day": _normalize_day(row.get("day", "")),
            "start_time": _normalize_text(row.get("start_time", "")),
            "end_time": _normalize_text(row.get("end_time", "")),
            "room": _normalize_room_name(row.get("room", "")),
            "row_number": row.get("row_number", 0),
        }
        if not row["day"] or not row["start_time"] or not row["end_time"] or not row["subject"]:
            continue
        key = (
            _canonical_text(row["instructor"]),
            _canonical_text(row["subject"]),
            _canonical_text(row["room"]),
            row["day"],
            row["start_time"],
            row["end_time"],
        )
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(row)

    if not cleaned and ocr_error:
        print("Teaching-load PDF parser:", ocr_error)
    return cleaned


def _parse_uploaded_image(uploaded_file):
    import pytesseract
    from PIL import Image
    if not _configure_tesseract():
        raise RuntimeError("Tesseract OCR was not found.")
    uploaded_file.seek(0)
    image = Image.open(uploaded_file).convert("RGB")
    text = pytesseract.image_to_string(image, config="--psm 6")
    instructor = _extract_name_from_text(text)
    return _extract_schedule_rows_from_text(text, instructor)


def _parse_uploaded_pptx(uploaded_file):
    from pptx import Presentation
    presentation = Presentation(uploaded_file)
    text_parts = []
    for slide in presentation.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text_parts.append(shape.text)
    text = "\n".join(text_parts)
    instructor = _extract_name_from_text(text)
    return _extract_schedule_rows_from_text(text, instructor)


def _parse_uploaded_schedule_file(uploaded_file):
    filename = uploaded_file.name.lower()
    extension = os.path.splitext(filename)[1]

    if extension in {".xlsx", ".xlsm"}:
        return _parse_uploaded_schedule_excel(uploaded_file)

    if extension == ".xls":
        import pandas as pd
        frames = pd.read_excel(uploaded_file, sheet_name=None, header=None)
        rows = []
        for frame in frames.values():
            csv_text = frame.to_csv(index=False, header=False).encode("utf-8")
            buffer = BytesIO(csv_text)
            rows.extend(_parse_uploaded_csv(buffer))
        return rows

    if extension == ".csv":
        return _parse_uploaded_csv(uploaded_file)
    if extension == ".docx":
        return _parse_uploaded_docx(uploaded_file)
    if extension == ".pdf":
        return _parse_uploaded_pdf(uploaded_file)
    if extension in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}:
        return _parse_uploaded_image(uploaded_file)
    if extension == ".pptx":
        return _parse_uploaded_pptx(uploaded_file)
    if extension == ".txt":
        raw = uploaded_file.read().decode("utf-8-sig", errors="ignore")
        return _extract_schedule_rows_from_text(raw, _extract_name_from_text(raw))

    raise ValueError(
        "Unsupported file type. Use XLSX, XLSM, XLS, CSV, PDF, DOCX, PPTX, TXT, PNG, JPG, JPEG, WEBP, BMP, TIF or TIFF."
    )


def _build_existing_import_cache():
    instructor_cache = {_canonical_text(obj.name): obj for obj in Instructor.objects.all()}
    subject_cache = {}
    for obj in Subject.objects.all():
        full_key = _canonical_text(f"{obj.subject_code} {obj.subject_name}")
        subject_cache[full_key] = obj
        subject_cache.setdefault(_canonical_text(obj.subject_code), obj)
    room_cache = {_canonical_text(obj.room_name): obj for obj in Room.objects.all()}
    return instructor_cache, subject_cache, room_cache


@login_required
def _schedule_import(request, classification):
    """
    Common schedule importer.

    classification:
        PART_TIME
        REGULAR
    """

    if classification not in {"PART_TIME", "REGULAR"}:
        messages.error(
            request,
            "Invalid instructor classification."
        )
        return redirect("schedule_list")

    # Store the selected classification in the session
    request.session["schedule_import_classification"] = classification
    request.session.modified = True

    if request.method == "POST":

        uploaded_files = request.FILES.getlist(
            "schedule_files"
        )

        if not uploaded_files:

            messages.error(
                request,
                "Please select at least one schedule file."
            )

            return redirect(
                "schedule_import_part_time"
                if classification == "PART_TIME"
                else "schedule_import_regular"
            )

        imported_rows = []
        file_errors = []

        # ==========================================
        # READ ALL UPLOADED FILES
        # ==========================================

        for uploaded_file in uploaded_files:

            try:

                rows = _parse_uploaded_schedule_file(
                    uploaded_file
                )

                for row in rows:

                    row["source_file"] = (
                        uploaded_file.name
                    )

                imported_rows.extend(rows)

            except Exception as exc:

                file_errors.append(
                    f"{uploaded_file.name}: {exc}"
                )

        # ==========================================
        # REMOVE DUPLICATES
        #
        # Duplicate identity:
        #
        # Instructor
        # Subject
        # Day
        # Start Time
        # End Time
        #
        # ROOM IS NOT PART OF THE DUPLICATE RULE.
        # ==========================================

        unique_rows = []
        seen = set()

        for row in imported_rows:

            key = (
                _canonical_text(
                    row.get("instructor")
                ),

                _canonical_text(
                    row.get("subject")
                ),

                _canonical_text(
                    row.get("day")
                ),

                _canonical_text(
                    row.get("start_time")
                ),

                _canonical_text(
                    row.get("end_time")
                ),
            )

            if key in seen:
                continue

            seen.add(key)

            unique_rows.append(row)

        # ==========================================
        # PREPARE ROW JSON
        # ==========================================

        for row in unique_rows:

            row["row_json"] = json.dumps(
                {
                    key: value
                    for key, value in row.items()
                    if key != "row_json"
                }
            )

        # ==========================================
        # NO VALID ROWS
        # ==========================================

        if not unique_rows:

            detail = (
                " | ".join(file_errors)
                if file_errors
                else "No schedule rows were found."
            )

            messages.error(
                request,
                detail
            )

            return redirect(
                "schedule_import_part_time"
                if classification == "PART_TIME"
                else "schedule_import_regular"
            )

        # ==========================================
        # SAVE IMPORT DATA IN SESSION
        # ==========================================

        request.session[
            "schedule_import_rows"
        ] = unique_rows

        request.session[
            "schedule_import_file_errors"
        ] = file_errors

        request.session[
            "schedule_import_classification"
        ] = classification

        request.session.modified = True

        # ==========================================
        # PREVIEW
        # ==========================================

        return render(
            request,
            "instructors/schedules/import.html",
            {
                "rows": unique_rows,

                "preview": True,

                "file_errors": file_errors,

                "uploaded_count": len(
                    uploaded_files
                ),

                "import_classification": (
                    "Part-Time"
                    if classification == "PART_TIME"
                    else "Regular"
                ),

                "import_classification_code": (
                    classification
                ),
            },
        )

    # ==========================================
    # NORMAL IMPORT PAGE
    # ==========================================

    return render(
        request,
        "instructors/schedules/import.html",
        {
            "rows": [],

            "preview": False,

            "file_errors": [],

            "uploaded_count": 0,

            "import_classification": (
                "Part-Time"
                if classification == "PART_TIME"
                else "Regular"
            ),

            "import_classification_code": (
                classification
            ),
        },
    )


@login_required
def schedule_import_part_time(request):

    return _schedule_import(
        request,
        "PART_TIME"
    )


@login_required
def schedule_import_regular(request):

    return _schedule_import(
        request,
        "REGULAR"
    )


@login_required
def schedule_import(request):
    """
    Keeps the old generic URL working.

    The old generic importer defaults to Part-Time.
    The new Schedules page uses the dedicated
    Part-Time and Regular buttons.
    """

    return _schedule_import(
        request,
        "PART_TIME"
    )


@login_required
def schedule_import_confirm(request):

    if request.method != "POST":

        return redirect(
            "schedule_import"
        )

    rows = request.POST.getlist(
        "row_data"
    )

    if not rows:

        messages.error(
            request,
            "No schedule data was submitted."
        )

        return redirect(
            "schedule_import"
        )

    # ==========================================
    # GET CLASSIFICATION FROM SESSION
    # ==========================================

    classification = request.session.get(
        "schedule_import_classification",
        "PART_TIME"
    )

    if classification not in {
        "PART_TIME",
        "REGULAR"
    }:

        classification = "PART_TIME"

    classification_label = (
        "Part-Time"
        if classification == "PART_TIME"
        else "Regular"
    )

    # ==========================================
    # COUNTERS
    # ==========================================

    created_count = 0
    skipped_count = 0
    error_count = 0

    # ==========================================
    # EXISTING OBJECT CACHE
    # ==========================================

    (
        instructor_cache,
        subject_cache,
        room_cache
    ) = _build_existing_import_cache()

    # ==========================================
    # EXISTING SCHEDULE KEYS
    #
    # ROOM IS NOT INCLUDED.
    #
    # Duplicate identity:
    # Instructor + Subject + Day +
    # Start Time + End Time
    # ==========================================

    existing_schedule_keys = set(
        Schedule.objects.values_list(
            "instructor_id",
            "subject_id",
            "day",
            "start_time",
            "end_time",
        )
    )

    # ==========================================
    # PROCESS EVERY IMPORTED ROW
    # ==========================================

    for row_data in rows:

        try:

            data = json.loads(
                row_data
            )

            # ======================================
            # INSTRUCTOR
            # ======================================

            instructor_name = _normalize_text(
                data.get("instructor")
            )

            # ======================================
            # SUBJECT
            # ======================================

            subject_text = _parse_subject(
                data.get("subject")
            )

            # ======================================
            # ROOM
            #
            # Keep the room supplied by the
            # imported row.
            #
            # Room is NOT used for duplicate
            # detection.
            # ======================================

            room_name = _normalize_room_name(
                data.get("room")
            )

            # ======================================
            # DAY
            # ======================================

            day = _normalize_day(
                data.get("day")
            )

            # ======================================
            # TIME
            # ======================================

            start_text = _normalize_text(
                data.get("start_time")
            )

            end_text = _normalize_text(
                data.get("end_time")
            )

            # ======================================
            # VALIDATE REQUIRED DATA
            # ======================================

            if not all([
                instructor_name,
                subject_text,
                day,
                start_text,
                end_text,
            ]):

                error_count += 1

                continue

            # ======================================
            # CONVERT TIME
            # ======================================

            start_time = datetime.strptime(
                start_text,
                "%H:%M"
            ).time()

            end_time = datetime.strptime(
                end_text,
                "%H:%M"
            ).time()

            # ======================================
            # FIND / CREATE INSTRUCTOR
            # ======================================

            instructor_key = _canonical_text(
                instructor_name
            )

            instructor = instructor_cache.get(
                instructor_key
            )

            if instructor is None:

                instructor = Instructor.objects.create(
                    name=instructor_name,
                    classification=classification,
                )

                instructor_cache[
                    instructor_key
                ] = instructor

            else:

                # ==================================
                # KEEP THE EXISTING INSTRUCTOR
                # BUT MAKE SURE ITS CLASSIFICATION
                # MATCHES THE IMPORT TYPE.
                #
                # This allows an existing instructor
                # to be correctly classified when
                # importing their schedule.
                # ==================================

                if instructor.classification != classification:

                    instructor.classification = (
                        classification
                    )

                    instructor.save(
                        update_fields=[
                            "classification"
                        ]
                    )

            # ======================================
            # FIND / CREATE SUBJECT
            # ======================================

            subject_key = _canonical_text(
                subject_text
            )

            subject = subject_cache.get(
                subject_key
            )

            if subject is None:

                subject = Subject.objects.create(
                    subject_code=subject_text,
                    subject_name=""
                )

                subject_cache[
                    subject_key
                ] = subject

            # ======================================
            # FIND / CREATE ROOM
            # ======================================

            room_key = _canonical_text(
                room_name
            )

            room = room_cache.get(
                room_key
            )

            if room is None:

                room = Room.objects.create(
                    room_name=room_name
                )

                room_cache[
                    room_key
                ] = room

            # ======================================
            # DUPLICATE CHECK
            #
            # IMPORTANT:
            # ROOM IS NOT INCLUDED.
            # ======================================

            schedule_key = (
                instructor.id,
                subject.id,
                day,
                start_time,
                end_time,
            )

            if schedule_key in existing_schedule_keys:

                skipped_count += 1

                continue

            # ======================================
            # CREATE SCHEDULE
            # ======================================

            Schedule.objects.create(
                instructor=instructor,
                subject=subject,
                room=room,
                day=day,
                start_time=start_time,
                end_time=end_time,
            )

            existing_schedule_keys.add(
                schedule_key
            )

            created_count += 1

        except Exception as exc:

            error_count += 1

            print(
                "Schedule import row error:",
                exc
            )

    # ==========================================
    # CLEAR IMPORT SESSION DATA
    # ==========================================

    request.session.pop(
        "schedule_import_rows",
        None
    )

    file_errors = request.session.pop(
        "schedule_import_file_errors",
        []
    )

    request.session.pop(
        "schedule_import_classification",
        None
    )

    # ==========================================
    # SUCCESS MESSAGE
    # ==========================================

    if created_count:

        messages.success(
            request,
            (
                f"{created_count} schedule(s) "
                f"imported successfully as "
                f"{classification_label}."
            )
        )

    # ==========================================
    # DUPLICATES
    # ==========================================

    if skipped_count:

        messages.warning(
            request,
            (
                f"{skipped_count} duplicate "
                f"schedule(s) were skipped."
            )
        )

    # ==========================================
    # ERRORS
    # ==========================================

    if error_count:

        messages.error(
            request,
            (
                f"{error_count} schedule row(s) "
                f"could not be imported."
            )
        )

    # ==========================================
    # FILE ERRORS
    # ==========================================

    for error in file_errors:

        messages.warning(
            request,
            f"File issue: {error}"
        )

    # ==========================================
    # RETURN TO SCHEDULE LIST
    # ==========================================

    return redirect(
        "schedule_list"
    )