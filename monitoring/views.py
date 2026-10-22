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
from django.conf import settings
from django.contrib.auth.forms import PasswordResetForm
from django.shortcuts import render, redirect
from .forms import MonitoringStaffAccountForm
from .models import MonitoringStaff

import logging
logger = logging.getLogger(__name__)

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
    MonitoringStaffAccountForm,
)

from .models import (
    Instructor,
    Subject,
    Room,
    Schedule,
    MonitoringRecord,
    ReportTemplate,
    SystemAdministrator,
    MonitoringStaff,
)

from django.contrib.auth.forms import PasswordResetForm
from django.contrib.auth.views import PasswordResetView
from django.core.exceptions import ValidationError
from django.urls import reverse_lazy

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
    """Allow active administrator accounts only."""
    if not user.is_authenticated or not user.is_active:
        return False

    return SystemAdministrator.objects.filter(
        user=user,
        user__is_active=True,
    ).exists()


def is_admin_or_monitoring_staff(user):
    """Allow registered administrators and Monitoring Staff."""
    if not user.is_authenticated or not user.is_active:
        return False

    return (
        is_system_admin(user)
        or is_monitoring_staff(user)
    )


def is_main_admin(user):
    """Only the designated Main Administrator may manage admin accounts."""
    if not user.is_authenticated or not user.is_active:
        return False

    return SystemAdministrator.objects.filter(
        user=user,
        is_main_admin=True,
        user__is_active=True,
    ).exists()

def is_monitoring_staff(user):
    """Check whether the user is an active Monitoring Staff account."""
    if not user.is_authenticated or not user.is_active:
        return False

    return MonitoringStaff.objects.filter(user=user).exists()


def is_not_monitoring_staff(user):
    """Allow only administrators into administrator-only sections."""
    return user.is_authenticated and not is_monitoring_staff(user)


def can_manage_schedules(user):
    """Administrators and Monitoring Staff can manage schedules."""
    if not user.is_authenticated or not user.is_active:
        return False

    return (
        is_main_admin(user)
        or is_monitoring_staff(user)
        or SystemAdministrator.objects.filter(user=user).exists()
    )


def is_not_monitoring_staff_or_schedule_manager(user):
    """Keep restricted sections available only to administrators."""
    return is_not_monitoring_staff(user)

@login_required
@user_passes_test(is_main_admin)
def admin_accounts(request):
    admin_group, _ = Group.objects.get_or_create(
        name="System Administrator"
    )

    # Get the Main Admin account
    main_admin = SystemAdministrator.objects.filter(
        is_main_admin=True
    ).select_related("user").first()

    # Get regular administrators only
    admins = admin_group.user_set.exclude(
        system_administrator__is_main_admin=True
    ).order_by("username")

    # Get Monitoring Staff accounts separately
    monitoring_staff = MonitoringStaff.objects.select_related(
        "user"
    ).order_by("user__username")

    return render(
        request,
        "registration/admin_accounts.html",
        {
            "main_admin": main_admin,
            "admins": admins,
            "monitoring_staff": monitoring_staff,
            "is_main_admin": is_main_admin(request.user),
        }
    )


@login_required
@user_passes_test(is_main_admin)
def admin_account_create(request):
    admin_group, _ = Group.objects.get_or_create(
        name="System Administrator"
    )

    form = AdminAccountForm(request.POST or None)

    if request.method == "POST" and form.is_valid():
        user = form.save()
        admin_group.user_set.add(user)

        SystemAdministrator.objects.update_or_create(
            user=user,
            defaults={"is_main_admin": False},
        )

        messages.success(
            request,
            f"Administrator account '{user.username}' created successfully.",
        )
        return redirect("admin_accounts")

    return render(
        request,
        "registration/admin_account_form.html",
        {"form": form},
    )



@login_required
@user_passes_test(is_main_admin)
def monitoring_staff_create(request):
    if request.method == "POST":
        form = MonitoringStaffAccountForm(request.POST)

        if form.is_valid():
            user = form.save()

            MonitoringStaff.objects.update_or_create(
                user=user,
            )

            messages.success(
                request,
                f"Monitoring Staff account '{user.username}' created successfully.",
            )
            return redirect("admin_accounts")
    else:
        form = MonitoringStaffAccountForm()

    return render(
        request,
        "registration/monitoring_staff_form.html",
        {"form": form},
    )

@login_required
@user_passes_test(is_main_admin)
def admin_account_password(request, user_id):
    User = get_user_model()

    user = get_object_or_404(User, pk=user_id)

    # Only the Main Administrator can change passwords
    # for administrator accounts.
    if not SystemAdministrator.objects.filter(
        user=request.user,
        is_main_admin=True,
        user__is_active=True,
    ).exists():
        messages.error(
            request,
            "Only the Main Administrator can manage administrator passwords."
        )
        return redirect("admin_accounts")

    # Do not allow changing another Main Administrator account.
    # The Main Administrator may change their OWN password.
    target_is_main_admin = SystemAdministrator.objects.filter(
        user=user,
        is_main_admin=True,
    ).exists()

    if target_is_main_admin and user.pk != request.user.pk:
        messages.error(
            request,
            "You cannot change another Main Administrator's password here."
        )
        return redirect("admin_accounts")

    # Only System Administrator accounts can be managed here.
    if not (
            user.groups.filter(
                name="System Administrator"
            ).exists()
            or MonitoringStaff.objects.filter(
        user=user
    ).exists()
    ):
        messages.error(
            request,
            "That account is not a System Administrator account."
        )
        return redirect("admin_accounts")

    if request.method == "POST":
        password = request.POST.get("password", "")
        confirm_password = request.POST.get("confirm_password", "")

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

            # Keep the Main Administrator logged in after
            # changing their own password.
            if user.pk == request.user.pk:
                from django.contrib.auth import update_session_auth_hash
                update_session_auth_hash(request, user)

            messages.success(
                request,
                f"Password for administrator '{user.username}' was changed successfully."
            )
            return redirect("admin_accounts")

    return render(
        request,
        "registration/admin_account_password.html",
        {"admin": user},
    )

@login_required
@user_passes_test(is_main_admin)
def admin_account_deactivate(request, user_id):
    User = get_user_model()

    user = get_object_or_404(
        User,
        pk=user_id,
    )

    target_is_main_admin = SystemAdministrator.objects.filter(
        user=user,
        is_main_admin=True,
    ).exists()

    if user.pk == request.user.pk or target_is_main_admin:
        messages.error(
            request,
            "The Main Administrator account cannot be deactivated here.",
        )
        return redirect("admin_accounts")

    if not (
            user.groups.filter(
                name="System Administrator"
            ).exists()
            or MonitoringStaff.objects.filter(
        user=user
    ).exists()
    ):
        messages.error(
            request,
            "That account is not a System Administrator account.",
        )
        return redirect("admin_accounts")

    if request.method != "POST":
        return redirect("admin_accounts")

    user.is_active = False
    user.save(update_fields=["is_active"])

    messages.success(
        request,
        f"Administrator account '{user.username}' has been deactivated.",
    )

    return redirect("admin_accounts")


@login_required
@user_passes_test(is_main_admin)
def admin_account_delete(request, user_id):
    User = get_user_model()

    user = get_object_or_404(
        User,
        pk=user_id,
    )

    target_is_main_admin = SystemAdministrator.objects.filter(
        user=user,
        is_main_admin=True,
    ).exists()

    if user.pk == request.user.pk or target_is_main_admin:
        messages.error(
            request,
            "The Main Administrator account cannot be deleted here.",
        )
        return redirect("admin_accounts")

    if not (
            user.groups.filter(
                name="System Administrator"
            ).exists()
            or MonitoringStaff.objects.filter(
        user=user
    ).exists()
    ):
        messages.error(
            request,
            "That account is not a System Administrator account.",
        )
        return redirect("admin_accounts")

    if request.method != "POST":
        return redirect("admin_accounts")

    username = user.username
    user.delete()

    messages.success(
        request,
        f"Administrator account '{username}' was deleted successfully.",
    )

    return redirect("admin_accounts")


# ==========================================
# INSTRUCTORS
# ==========================================

@login_required
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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

            record.remarks = (
                    form.cleaned_data.get("remarks") or "Monitored"
            )

            record.other_remarks = (
                form.cleaned_data.get("other_remarks", "").strip()
                if record.remarks == "Other"
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
    remarks = str(
        getattr(record, "remarks", "") or ""
    ).strip()

    other = str(
        getattr(record, "other_remarks", "") or ""
    ).strip()

    missed = str(
        getattr(record, "missed_reason", "") or ""
    ).strip()

    # If "Other" was selected, display the user's actual explanation
    # instead of the literal word "Other".
    if remarks.lower() in {"other", "others"}:
        return other or "Other"

    # Keep the explanation for records marked Not Monitored.
    if remarks == "Not Monitored" and other:
        return f"{remarks} - {other}"

    # Keep the explanation for records marked Missed.
    if remarks == "Missed" and missed:
        return f"{remarks} - {missed}"

    # Handle any other record that has additional remarks.
    if other:
        return f"{remarks} - {other}" if remarks else other

    if missed:
        return f"{remarks} - {missed}" if remarks else missed

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


def _get_report_template(template_id=None, classification="ALL"):
    """
    Select the actual uploaded Excel template.

    Classification filters the report's records and heading.
    It does NOT select a different template based on its filename.
    """
    from pathlib import Path

    templates = ReportTemplate.objects.all()
    template = None

    if template_id:
        try:
            template = templates.get(pk=template_id)
        except (
            ReportTemplate.DoesNotExist,
            ValueError,
            TypeError,
        ):
            template = None

    # If no valid template was explicitly selected, use the active one.
    if template is None:
        template = (
            templates
            .filter(is_active=True)
            .exclude(file="")
            .first()
        )

    # Last fallback: use any uploaded template.
    if template is None:
        template = (
            templates
            .exclude(file="")
            .first()
        )

    if template is None:
        raise ValueError(
            "No uploaded Excel report template was found. "
            "Upload an Excel template and activate it in Report Templates."
        )

    if not template.file:
        raise ValueError(
            "The selected report template does not contain a file."
        )

    try:
        template_path = Path(template.file.path)
    except Exception as exc:
        raise ValueError(
            "The selected report template file could not be accessed."
        ) from exc

    if not template_path.exists():
        raise ValueError(
            "The selected report template file could not be found."
        )

    if template_path.suffix.lower() != ".xlsx":
        raise ValueError(
            "The report template must be an Excel .xlsx file."
        )

    return template, template_path


def _update_template_classification_heading(
    worksheet,
    header_rows,
    classification="ALL",
):
    """
    Update the existing classification heading in the uploaded
    Excel template according to the dean's selected classification.

    ALL:
        PART-TIME AND FULL-TIME INSTRUCTORS

    PART_TIME:
        PART-TIME INSTRUCTORS

    REGULAR:
        FULL-TIME INSTRUCTORS

    Preserves the existing heading cell, surrounding text, and
    formatting. Does not create a new title or hard-code the
    school's other heading lines.
    """

    import re

    # ---------------------------------------------------------
    # NORMALIZE THE SELECTED CLASSIFICATION
    # ---------------------------------------------------------

    classification = str(
        classification or "ALL"
    ).strip().upper().replace("-", "_").replace(" ", "_")

    if classification in {
        "PART_TIME",
        "PARTTIME",
    }:
        classification = "PART_TIME"

    elif classification in {
        "REGULAR",
        "FULL_TIME",
        "FULLTIME",
    }:
        classification = "REGULAR"

    else:
        classification = "ALL"

    # ---------------------------------------------------------
    # RECOGNIZE THE EXISTING HEADING
    # ---------------------------------------------------------

    category_pattern = (
        r"(?:PART[\s-]*TIME|FULL[\s-]*TIME|REGULAR)"
    )

    section_heading_pattern = re.compile(
        rf"(?P<categories>{category_pattern}"
        rf"(?:\s*(?:AND|&)\s*{category_pattern})*)"
        rf"\s+(?P<noun>INSTRUCTORS?|FACULTY|TEACHERS?)\b",
        re.IGNORECASE,
    )

    category_token_pattern = re.compile(
        category_pattern,
        re.IGNORECASE,
    )

    # ---------------------------------------------------------
    # MATCH THE TEMPLATE'S CAPITALIZATION STYLE
    # ---------------------------------------------------------

    def match_case(value, reference):
        if reference.isupper():
            return value.upper()

        if reference.islower():
            return value.lower()

        return value.title()

    # ---------------------------------------------------------
    # CREATE PART-TIME AND FULL-TIME LABELS
    # ---------------------------------------------------------

    def get_category_labels(category_text):
        tokens = category_token_pattern.findall(
            category_text or ""
        )

        part_token = next(
            (
                token
                for token in tokens
                if re.fullmatch(
                    r"PART[\s-]*TIME",
                    token,
                    re.IGNORECASE,
                )
            ),
            None,
        )

        full_token = next(
            (
                token
                for token in tokens
                if re.fullmatch(
                    r"FULL[\s-]*TIME",
                    token,
                    re.IGNORECASE,
                )
            ),
            None,
        )

        reference = (
            tokens[0]
            if tokens
            else "PART-TIME"
        )

        # Prefer the spelling already present in the template.
        # If no spelling exists, use the standard hyphenated labels.
        if part_token:
            part_label = part_token
        else:
            part_separator = (
                "-"
                if not tokens or "-" in reference
                else " "
            )

            part_label = match_case(
                f"PART{part_separator}TIME",
                reference,
            )

        if full_token:
            full_label = full_token
        else:
            full_separator = (
                "-"
                if (
                    not tokens
                    or "-" in reference
                    or part_token is None
                )
                else " "
            )

            full_label = match_case(
                f"FULL{full_separator}TIME",
                reference,
            )

        # Make both labels follow the same capitalization style.
        part_label = match_case(
            part_label,
            reference,
        )

        full_label = match_case(
            full_label,
            reference,
        )

        return part_label, full_label, reference

    # ---------------------------------------------------------
    # BUILD THE CORRECT HEADING
    # ---------------------------------------------------------

    def build_new_label(match):
        category_text = match.group("categories")
        noun = match.group("noun")

        part_label, full_label, reference = (
            get_category_labels(category_text)
        )

        if classification == "PART_TIME":
            new_categories = part_label

        elif classification == "REGULAR":
            new_categories = full_label

        else:
            # Standard combined heading when ALL is selected.
            if reference.isupper():
                conjunction = " AND "
            elif reference.islower():
                conjunction = " and "
            else:
                conjunction = " and "

            new_categories = (
                part_label
                + conjunction
                + full_label
            )

        return f"{new_categories} {noun}"

    # ---------------------------------------------------------
    # FIND THE EXISTING CLASSIFICATION HEADING
    # ---------------------------------------------------------

    if not header_rows:
        return

    table_header_row = min(header_rows)
    candidates = []

    for row in worksheet.iter_rows():
        for cell in row:
            if not isinstance(cell.value, str):
                continue

            match = section_heading_pattern.search(
                cell.value
            )

            if not match:
                continue

            # Prefer the nearest classification heading above
            # the report's column headings.
            if cell.row <= table_header_row:
                priority = 0
                distance = table_header_row - cell.row
            else:
                priority = 1
                distance = cell.row - table_header_row

            candidates.append(
                (
                    priority,
                    distance,
                    cell,
                    match,
                )
            )

    if not candidates:
        # Do not invent a classification title if the template
        # doesn't contain an existing, recognizable heading.
        return

    candidates.sort(
        key=lambda item: (
            item[0],
            item[1],
        )
    )

    _, _, target_cell, match = candidates[0]

    original_text = target_cell.value
    new_label = build_new_label(match)

    # Change only the classification phrase within the cell.
    # Keep the rest of the cell text unchanged.
    target_cell.value = (
        original_text[:match.start()]
        + new_label
        + original_text[match.end():]
    )

# ==========================================
# REPORT PRINT PAGE / WINDOWS PRINTING
# ==========================================

@login_required
@user_passes_test(is_admin_or_monitoring_staff, login_url="dashboard")
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

# =============================================================
# REPORT GENERATION — COMPLETE CORRECTED REPLACEMENT
# =============================================================

def _report_template_id(request):
    """Read the optional uploaded ReportTemplate primary key."""
    value = _request_value(request, "template_id", "")
    return str(value or "").strip() or None


def _checked_report_build_result(result, function_name):
    """Validate report builder results before unpacking them."""
    if not isinstance(result, (tuple, list)) or len(result) != 3:
        raise RuntimeError(
            f"{function_name} must return "
            "(output_path, temp_dir, template). "
            "Check for a missing return statement or duplicate function."
        )

    output_path, temp_dir, template = result

    if not output_path or not temp_dir:
        raise RuntimeError(
            f"{function_name} did not return both "
            "output_path and temp_dir."
        )

    return output_path, temp_dir, template


def _report_last_name_sort_key(name):
    """Sort instructors by surname, including 'Surname, Given' names."""
    name = " ".join(str(name or "").split())

    if not name:
        return ("", "", "")

    if "," in name:
        surname, given = name.split(",", 1)
        return (
            surname.strip().casefold(),
            given.strip().casefold(),
            name.casefold(),
        )

    parts = name.split()

    if len(parts) == 1:
        return (parts[0].casefold(), "", name.casefold())

    return (
        parts[-1].casefold(),
        " ".join(parts[:-1]).casefold(),
        name.casefold(),
    )


# =============================================================
# EXCEL REPORT BUILDER
# =============================================================

def _build_report_excel_file(
    records,
    week_start,
    week_end,
    template_id=None,
    classification="ALL",
    settings=None,
):
    """
    Generate an Excel monitoring report from an uploaded template.

    Returns:
        (output_xlsx, temp_dir, template)

    The selected uploaded Excel workbook is used as the layout base.
    Instructor names are merged vertically when an instructor has
    multiple monitoring records.
    """

    import os
    import re
    import shutil
    import tempfile
    from copy import copy
    from itertools import groupby

    from django.utils import timezone
    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter

    record_list = list(records or [])

    if not record_list:
        raise ValueError(
            "No monitoring records match the selected week, "
            "classification, search, and selected-record filters."
        )

    classification = (classification or "ALL").strip().upper()

    # Get the selected uploaded template, or automatically select one
    # based on classification when template_id is not provided.
    template, template_path = _get_report_template(
        template_id=template_id,
        classification=classification,
    )

    template_path = str(template_path)

    if os.path.splitext(template_path)[1].lower() != ".xlsx":
        raise ValueError(
            "The selected report template must be an .xlsx file."
        )

    if not os.path.isfile(template_path):
        raise FileNotFoundError(
            "The selected Excel report template was not found."
        )

    temp_dir = tempfile.mkdtemp(
        prefix="teacher_monitoring_report_"
    )

    output_xlsx = os.path.join(
        temp_dir,
        "Monitoring_Report.xlsx",
    )

    workbook = None

    try:
        # Work on a copy. Never modify the original uploaded template.
        shutil.copy2(template_path, output_xlsx)

        workbook = load_workbook(
            output_xlsx,
            data_only=False,
        )

        aliases = {
            "INSTRUCTOR": {
                "NAME",
                "INSTRUCTOR",
                "INSTRUCTOR NAME",
                "FACULTY NAME",
                "TEACHER NAME",
                "NAME OF INSTRUCTOR",
                "FACULTY",
            },
            "SUBJECT": {
                "SUBJECT",
                "SUBJECT CODE",
                "COURSE",
                "SUBJECT / COURSE",
            },
            "TIME": {
                "TIME",
                "CLASS TIME",
                "SCHEDULED TIME",
                "ORIGINAL TIME",
            },
            "DAY": {
                "DAY",
                "CLASS DAY",
            },
            "ROOM": {
                "ROOM",
                "ROOM USED",
                "ACTUAL ROOM",
                "ACTUAL ROOM USED",
            },
            "MONITORING_TIME": {
                "MONITORING TIME",
                "DATE & TIME",
                "DATE AND TIME",
                "DATE & TIME MONITORED",
                "ACTUAL TIME",
                "MONITORED AT",
            },
            "SIGNATURE": {
                "SIGNATURE",
                "INSTRUCTOR SIGNATURE",
            },
            "REMARKS": {
                "REMARKS",
                "REMARK",
                "STATUS",
            },
        }

        def normalize_heading(value):
            if value is None:
                return ""

            return re.sub(
                r"\s+",
                " ",
                str(value).replace("\n", " ").strip().upper(),
            )

        # The actual uploaded templates use a TWO-ROW table heading:
        # row 2 contains NAME, while row 3 contains SUBJECT, TIME, DAY,
        # SIGNATURE, and REMARKS. Some templates also repeat these labels
        # between instructor groups. Therefore, scan each possible NAME
        # row together with the next two rows instead of requiring every
        # heading to be on the same row.
        worksheet = None
        header_row = None
        header_map = {}
        header_rows = set()
        best_score = 0

        for sheet in workbook.worksheets:
            for candidate_row in range(1, sheet.max_row + 1):
                candidate_map = {}
                candidate_header_rows = set()

                # The instructor/name heading can be on its own row,
                # with the remaining column headings on the next row(s).
                scan_end = min(sheet.max_row, candidate_row + 2)

                for scan_row in range(candidate_row, scan_end + 1):
                    for cell in sheet[scan_row]:
                        heading = normalize_heading(cell.value)

                        if not heading:
                            continue

                        for field, choices in aliases.items():
                            if (
                                heading in choices
                                and field not in candidate_map
                            ):
                                candidate_map[field] = cell.column
                                candidate_header_rows.add(scan_row)

                score = len(candidate_map)

                if (
                    "INSTRUCTOR" in candidate_map
                    and score >= 3
                    and score > best_score
                ):
                    worksheet = sheet
                    header_map = candidate_map
                    header_rows = candidate_header_rows
                    # The data starts after the LAST row used by the
                    # detected column headings, not after the NAME row.
                    header_row = max(candidate_header_rows)
                    best_score = score

        if worksheet is None or header_row is None:
            # Include useful diagnostics: list the visible non-empty
            # cells from the first few rows of each sheet.
            sheet_details = []

            for sheet in workbook.worksheets:
                sample_rows = []

                for row in sheet.iter_rows(
                    min_row=1,
                    max_row=min(sheet.max_row, 6),
                ):
                    values = [
                        str(cell.value).replace("\\n", " ").strip()
                        for cell in row
                        if cell.value is not None
                    ]

                    if values:
                        sample_rows.append(
                            f"Row {row[0].row}: " + " | ".join(values[:10])
                        )

                sheet_details.append(
                    f"{sheet.title}: " + "; ".join(sample_rows)
                )

            raise ValueError(
                "Could not recognize the report table headings in the "
                "selected Excel template. The template may place NAME on "
                "one row and SUBJECT/TIME/DAY/SIGNATURE/REMARKS on the next. "
                "Detected worksheet contents: "
                + " || ".join(sheet_details)
            )

        # Update the classification heading before generating the exports.
        _update_template_classification_heading(
            worksheet,
            header_rows,
            classification=classification,
        )

        # Keep the title and headings exactly as they appear in the
        # uploaded Excel template. Do not hard-code title text here.
        data_start = header_row + 1
        name_column = header_map["INSTRUCTOR"]

        # Remove existing sample-data merges while preserving merges
        # located above the report's data rows.
        for merged_range in list(worksheet.merged_cells.ranges):
            if merged_range.max_row >= data_start:
                worksheet.unmerge_cells(str(merged_range))

        original_max_row = worksheet.max_row

        original_max_column = max(
            worksheet.max_column,
            max(header_map.values()),
        )

        # Capture the formatting of the template's first data row.
        source_styles = {}

        style_source_row = data_start

        source_height = worksheet.row_dimensions[
            style_source_row
        ].height

        for col in range(1, original_max_column + 1):
            cell = worksheet.cell(style_source_row, col)

            if isinstance(cell, MergedCell):
                continue

            source_styles[col] = {
                "style": copy(cell._style) if cell.has_style else None,
                "number_format": cell.number_format,
                "alignment": copy(cell.alignment),
                "protection": copy(cell.protection),
            }

        # ---------------------------------------------------------
        # EXTRACT VALUES FROM MONITORING RECORDS
        # ---------------------------------------------------------

        def extract_values(record):
            schedule = getattr(record, "schedule", None)

            instructor = (
                getattr(schedule, "instructor", None)
                if schedule else None
            )

            subject = (
                getattr(schedule, "subject", None)
                if schedule else None
            )

            room = (
                getattr(schedule, "room", None)
                if schedule else None
            )

            instructor_name = str(
                getattr(instructor, "name", "") or ""
            ).strip()

            subject_code = str(
                getattr(subject, "subject_code", "") or ""
            ).strip()

            subject_name = str(
                getattr(subject, "subject_name", "") or ""
            ).strip()

            if subject_code and subject_name:
                subject_text = (
                    f"{subject_code} - {subject_name}"
                )
            else:
                subject_text = subject_code or subject_name

            start_time = (
                getattr(schedule, "start_time", None)
                if schedule else None
            )

            end_time = (
                getattr(schedule, "end_time", None)
                if schedule else None
            )

            scheduled_time = ""

            if start_time and end_time:
                scheduled_time = (
                    f"{start_time.strftime('%I:%M %p')} - "
                    f"{end_time.strftime('%I:%M %p')}"
                )

            if schedule:
                try:
                    day_text = schedule.get_day_display()
                except Exception:
                    day_text = str(
                        getattr(schedule, "day", "") or ""
                    )
            else:
                day_text = ""

            assigned_room = str(
                getattr(room, "room_name", "") or ""
            ).strip()

            actual_room = str(
                getattr(record, "transfer_room", "") or ""
            ).strip()

            room_text = actual_room or assigned_room

            record_datetime = getattr(
                record,
                "record_datetime",
                None,
            )

            monitoring_time = ""

            if record_datetime:
                if timezone.is_aware(record_datetime):
                    record_datetime = timezone.localtime(
                        record_datetime
                    )

                monitoring_time = record_datetime.strftime(
                    "%B %d, %Y - %I:%M %p"
                )

            remarks = str(
                getattr(record, "remarks", "") or ""
            ).strip()

            other_remarks = str(
                getattr(record, "other_remarks", "") or ""
            ).strip()

            missed_reason = str(
                getattr(record, "missed_reason", "") or ""
            ).strip()

            if remarks == "Not Monitored" and other_remarks:
                remarks = f"{remarks} - {other_remarks}"

            elif remarks == "Missed" and missed_reason:
                remarks = f"{remarks} - {missed_reason}"

            elif other_remarks:
                remarks = (
                    f"{remarks} - {other_remarks}"
                    if remarks else other_remarks
                )

            elif missed_reason:
                remarks = (
                    f"{remarks} - {missed_reason}"
                    if remarks else missed_reason
                )

            return {
                "INSTRUCTOR": instructor_name,
                "SUBJECT": subject_text,
                "TIME": scheduled_time,
                "DAY": str(day_text or ""),
                "ROOM": room_text,
                "MONITORING_TIME": monitoring_time,
                "SIGNATURE": str(
                    getattr(record, "signature", "") or ""
                ).strip(),
                "REMARKS": remarks,
            }

        record_pairs = [
            (record, extract_values(record))
            for record in record_list
        ]

        # Alphabetize by instructor surname, then given name.
        record_pairs.sort(
            key=lambda pair: (
                _report_last_name_sort_key(
                    pair[1]["INSTRUCTOR"]
                ),
                str(
                    getattr(pair[0], "record_datetime", "") or ""
                ),
                getattr(pair[0], "pk", 0) or 0,
            )
        )

        total_records = len(record_pairs)

        clear_until = max(
            original_max_row,
            data_start + total_records - 1,
        )

        # Clear old sample values before inserting the current report.
        for row_number in range(data_start, clear_until + 1):
            for col in range(1, original_max_column + 1):
                cell = worksheet.cell(row_number, col)

                if not isinstance(cell, MergedCell):
                    cell.value = None

        # ---------------------------------------------------------
        # WRITE RECORDS AND GROUP EACH INSTRUCTOR'S ROWS
        # ---------------------------------------------------------

        row_groups = []
        current_row = data_start

        grouped_records = groupby(
            record_pairs,
            key=lambda pair: pair[1]["INSTRUCTOR"].casefold(),
        )

        for _, grouped_items in grouped_records:
            grouped_items = list(grouped_items)

            group_start = current_row
            group_end = group_start + len(grouped_items) - 1

            instructor_name = grouped_items[0][1]["INSTRUCTOR"]

            row_groups.append(
                (group_start, group_end, instructor_name)
            )

            for _record, values in grouped_items:
                if source_height is not None:
                    worksheet.row_dimensions[
                        current_row
                    ].height = source_height

                # Apply the uploaded template's row formatting.
                for col, info in source_styles.items():
                    cell = worksheet.cell(current_row, col)

                    if isinstance(cell, MergedCell):
                        continue

                    if info["style"] is not None:
                        cell._style = copy(info["style"])

                    cell.number_format = info["number_format"]
                    cell.alignment = copy(info["alignment"])
                    cell.protection = copy(info["protection"])

                # Populate the columns that actually exist in the template.
                for field, col in header_map.items():
                    if field == "INSTRUCTOR":
                        continue

                    worksheet.cell(
                        current_row,
                        col,
                    ).value = values.get(field, "")

                current_row += 1

        # One vertically merged NAME cell per instructor.
        for group_start, group_end, instructor_name in row_groups:
            if group_end > group_start:
                worksheet.merge_cells(
                    start_row=group_start,
                    start_column=name_column,
                    end_row=group_end,
                    end_column=name_column,
                )

            name_cell = worksheet.cell(
                group_start,
                name_column,
            )

            name_cell.value = instructor_name
            name_cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        # Replace supported week placeholders without changing other
        # headings, titles, or wording in the selected uploaded template.
        replacements = {
            "{{WEEK_START}}": week_start.strftime("%B %d, %Y"),
            "{{WEEK_END}}": week_end.strftime("%B %d, %Y"),
            "{{WEEK}}": (
                f"{week_start:%B %d, %Y} - "
                f"{week_end:%B %d, %Y}"
            ),
        }

        for sheet in workbook.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if (
                        isinstance(cell, MergedCell)
                        or not isinstance(cell.value, str)
                    ):
                        continue

                    new_value = cell.value

                    for placeholder, replacement in replacements.items():
                        new_value = new_value.replace(
                            placeholder,
                            replacement,
                        )

                    if new_value != cell.value:
                        cell.value = new_value

        # Configure the print area and page layout.
        final_last_row = max(
            header_row,
            data_start + total_records - 1,
        )

        # Only print the columns actually used by the report.
        # Ignore unused formatted columns outside the actual template.
        final_last_column = max(header_map.values())

        final_last_column = max(
            worksheet.max_column,
            max(header_map.values()),
        )

        worksheet.print_area = (
            f"A1:{get_column_letter(final_last_column)}{final_last_row}"
        )

        worksheet.print_title_rows = f"1:{header_row}"

        # Automatically scale the report to the selected paper size.
        worksheet.page_setup.scale = None
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = 0
        worksheet.sheet_properties.pageSetUpPr.fitToPage = True

        # Apply selected print settings when provided.
        settings = settings or {}

        paper_codes = {
            "A4": "9",
            "LETTER": "1",
            "LEGAL": "5",
            "FOLIO": "14",
        }

        requested_paper = str(
            settings.get("paper_size", "")
        ).upper()

        if requested_paper in paper_codes:
            worksheet.page_setup.paperSize = paper_codes[
                requested_paper
            ]

        requested_orientation = str(
            settings.get("orientation", "")
        ).lower()

        if requested_orientation in {"portrait", "landscape"}:
            worksheet.page_setup.orientation = requested_orientation

        margin = (
            0.30
            if str(settings.get("margins", "normal")).lower() == "narrow"
            else 0.50
        )

        if settings:
            worksheet.page_margins.left = margin
            worksheet.page_margins.right = margin
            worksheet.page_margins.top = margin
            worksheet.page_margins.bottom = margin

        # Make the report sheet active so DOCX conversion uses the same sheet.
        workbook.active = workbook.worksheets.index(worksheet)

        workbook.save(output_xlsx)
        workbook.close()
        workbook = None

        # IMPORTANT: always return all three values.
        return output_xlsx, temp_dir, template

    except Exception:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass

        shutil.rmtree(temp_dir, ignore_errors=True)
        raise


# =============================================================
# DOCX REPORT BUILDER
# =============================================================
def _build_report_docx_file(
    records,
    week_start,
    week_end,
    template_id=None,
    classification="ALL",
    settings=None,
):
    """Create an editable DOCX using the selected Excel template's layout."""

    import os
    import shutil
    from io import BytesIO
    from xml.etree import ElementTree as ET

    from openpyxl import load_workbook
    from openpyxl.styles.colors import COLOR_INDEX
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries

    from docx import Document
    from docx.enum.table import (
        WD_CELL_VERTICAL_ALIGNMENT,
        WD_ROW_HEIGHT_RULE,
        WD_TABLE_ALIGNMENT,
    )
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Inches, Pt, RGBColor

    # Generate the Excel report using the selected uploaded template.
    result = _build_report_excel_file(
        records,
        week_start,
        week_end,
        template_id=template_id,
        classification=classification,
        settings=settings,
    )

    excel_path, temp_dir, template = _checked_report_build_result(
        result,
        "_build_report_excel_file",
    )

    workbook = None

    try:
        workbook = load_workbook(
            excel_path,
            data_only=False,
        )

        worksheet = workbook.active
        settings = settings or {}

        # ---------------------------------------------------------
        # READ EXCEL THEME COLORS
        # ---------------------------------------------------------

        theme_rgb = {}

        try:
            theme_root = ET.fromstring(workbook.loaded_theme)

            namespace = {
                "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
            }

            scheme = theme_root.find(
                ".//a:clrScheme",
                namespace,
            )

            if scheme is not None:
                for index, theme_color in enumerate(list(scheme)):
                    children = list(theme_color)

                    if children:
                        raw = (
                            children[0].attrib.get("lastClr")
                            or children[0].attrib.get("val")
                        )

                        if raw and len(raw) >= 6:
                            theme_rgb[index] = raw[-6:].upper()

        except Exception:
            theme_rgb = {}

        def excel_color_to_hex(color):
            """Convert Excel RGB, theme, or indexed colors to HEX."""

            if color is None:
                return None

            raw = None

            if color.type == "rgb" and color.rgb:
                raw = str(color.rgb)[-6:]

            elif color.type == "theme" and color.theme is not None:
                raw = theme_rgb.get(int(color.theme))

            elif color.type == "indexed" and color.indexed is not None:
                try:
                    raw = str(
                        COLOR_INDEX[int(color.indexed)]
                    )[-6:]

                except (IndexError, TypeError, ValueError):
                    raw = None

            if not raw:
                return None

            raw = raw.upper()

            try:
                tint = float(
                    getattr(color, "tint", 0) or 0
                )
            except (TypeError, ValueError):
                tint = 0

            if tint:
                channels = []

                for offset in (0, 2, 4):
                    channel = int(
                        raw[offset:offset + 2],
                        16,
                    )

                    if tint < 0:
                        channel = round(
                            channel * (1 + tint)
                        )
                    else:
                        channel = round(
                            channel * (1 - tint)
                            + 255 * tint
                        )

                    channels.append(
                        max(0, min(255, channel))
                    )

                raw = "".join(
                    f"{value:02X}"
                    for value in channels
                )

            return raw

        # ---------------------------------------------------------
        # READ THE EXCEL PRINT AREA
        # ---------------------------------------------------------

        min_col, min_row = 1, 1
        max_col = worksheet.max_column
        max_row = worksheet.max_row

        if worksheet.print_area:
            try:
                area = str(
                    worksheet.print_area
                ).split(",")[0].strip()

                if "!" in area:
                    area = area.rsplit("!", 1)[1]

                (
                    min_col,
                    min_row,
                    max_col,
                    max_row,
                ) = range_boundaries(
                    area.replace("$", "")
                )

            except (ValueError, TypeError):
                min_col, min_row = 1, 1
                max_col = worksheet.max_column
                max_row = worksheet.max_row

        # ---------------------------------------------------------
        # DETECT THE LAST REAL REPORT COLUMN
        # ---------------------------------------------------------
        # Some templates contain extra formatted columns after the
        # real table. For example, the Regular template may extend
        # to column K although its actual report headings end at G.
        # These unused columns must not be included in the DOCX table.

        name_labels = {
            "NAME",
            "INSTRUCTOR",
            "INSTRUCTOR NAME",
            "FACULTY NAME",
        }

        report_header_labels = {
            "NAME",
            "INSTRUCTOR",
            "INSTRUCTOR NAME",
            "FACULTY NAME",
            "SUBJECT",
            "SUBJECT CODE",
            "TIME",
            "CLASS TIME",
            "ORIGINAL TIME",
            "DAY",
            "ROOM",
            "SIGNATURE",
            "REMARKS",
            "MONITORING TIME",
            "DATE & TIME",
            "DATE AND TIME",
            "ASSIGNED ROOM",
            "ACTUAL ROOM",
            "ACTUAL ROOM USED",
            "SCHEDULED CLASS TIME",
            "DATE MONITORED",
        }

        detected_header_row = None

        for candidate_row in range(
            min_row,
            min(max_row, min_row + 12) + 1,
        ):
            for candidate_col in range(
                1,
                worksheet.max_column + 1,
            ):
                raw_value = worksheet.cell(
                    candidate_row,
                    candidate_col,
                ).value

                normalized = " ".join(
                    str(raw_value or "").upper().split()
                )

                if normalized in name_labels:
                    detected_header_row = candidate_row
                    break

            if detected_header_row is not None:
                break

        if detected_header_row is not None:
            header_columns = []

            scan_end_row = min(
                max_row,
                detected_header_row + 2,
            )

            for scan_row in range(
                detected_header_row,
                scan_end_row + 1,
            ):
                for scan_col in range(
                    1,
                    worksheet.max_column + 1,
                ):
                    raw_value = worksheet.cell(
                        scan_row,
                        scan_col,
                    ).value

                    normalized = " ".join(
                        str(raw_value or "").upper().split()
                    )

                    if normalized in report_header_labels:
                        header_columns.append(scan_col)

                        # Include the complete width of a merged header.
                        for merged_range in worksheet.merged_cells.ranges:
                            if (
                                merged_range.min_row
                                <= scan_row
                                <= merged_range.max_row
                                and merged_range.min_col
                                <= scan_col
                                <= merged_range.max_col
                            ):
                                header_columns.append(
                                    merged_range.max_col
                                )

            if header_columns:
                max_col = max(header_columns)

        # ---------------------------------------------------------
        # DETECT TEMPLATE IMAGES
        # ---------------------------------------------------------

        excel_images = list(
            getattr(worksheet, "_images", []) or []
        )

        has_image_on_first_row = any(
            getattr(
                getattr(image, "anchor", None),
                "_from",
                None,
            ) is not None
            and image.anchor._from.row == 0
            for image in excel_images
        )

        first_row_has_text = any(
            worksheet.cell(1, col).value not in (None, "")
            for col in range(min_col, max_col + 1)
        )

        # Images anchored on an empty Excel row are handled in the
        # separate Word header block below.
        if (
            min_row <= 1
            and has_image_on_first_row
            and not first_row_has_text
        ):
            min_row = 2

        row_count = max(
            1,
            max_row - min_row + 1,
        )

        col_count = max(
            1,
            max_col - min_col + 1,
        )

        # ---------------------------------------------------------
        # SET PAPER SIZE, ORIENTATION, AND MARGINS
        # ---------------------------------------------------------

        paper_sizes = {
            "A4": (8.27, 11.69),
            "LETTER": (8.5, 11.0),
            "LEGAL": (8.5, 14.0),
            "FOLIO": (8.5, 13.0),
        }

        excel_paper_sizes = {
            "1": (8.5, 11.0),
            "2": (8.5, 11.0),
            "3": (11.0, 17.0),
            "4": (17.0, 11.0),
            "5": (8.5, 14.0),
            "6": (5.5, 8.5),
            "7": (7.25, 10.5),
            "8": (11.69, 16.54),
            "9": (8.27, 11.69),
            "10": (8.27, 11.69),
            "11": (5.83, 8.27),
            "12": (9.84, 13.90),
            "13": (7.17, 10.12),
            "14": (8.5, 13.0),
            "15": (8.5, 10.83),
            "16": (10.0, 14.0),
            "17": (11.0, 17.0),
        }

        requested_paper = str(
            settings.get("paper_size", "")
        ).upper()

        if requested_paper in paper_sizes:
            page_width, page_height = paper_sizes[
                requested_paper
            ]
        else:
            page_width, page_height = excel_paper_sizes.get(
                str(worksheet.page_setup.paperSize),
                (8.27, 11.69),
            )

        orientation = str(
            settings.get(
                "orientation",
                worksheet.page_setup.orientation or "portrait",
            )
        ).lower()

        # Normalize dimensions to the requested orientation.
        if orientation == "landscape":
            page_width, page_height = (
                max(page_width, page_height),
                min(page_width, page_height),
            )
        else:
            page_width, page_height = (
                min(page_width, page_height),
                max(page_width, page_height),
            )

        document = Document()
        section = document.sections[0]

        section.page_width = Inches(page_width)
        section.page_height = Inches(page_height)

        if settings:
            page_margin = (
                0.30
                if str(
                    settings.get("margins", "normal")
                ).lower() == "narrow"
                else 0.50
            )

            margin_left = page_margin
            margin_right = page_margin

            section.left_margin = Inches(page_margin)
            section.right_margin = Inches(page_margin)
            section.top_margin = Inches(page_margin)
            section.bottom_margin = Inches(page_margin)

        else:
            margin_left = float(
                worksheet.page_margins.left or 0.5
            )
            margin_right = float(
                worksheet.page_margins.right or 0.5
            )

            section.left_margin = Inches(margin_left)
            section.right_margin = Inches(margin_right)

            section.top_margin = Inches(
                float(worksheet.page_margins.top or 0.5)
            )
            section.bottom_margin = Inches(
                float(worksheet.page_margins.bottom or 0.5)
            )

        # Use the complete printable page width.
        table_width = max(
            1.0,
            page_width - margin_left - margin_right,
        )

        total_table_twips = int(
            round(table_width * 1440)
        )

        # ---------------------------------------------------------
        # CALCULATE COLUMN WIDTHS
        # ---------------------------------------------------------

        def excel_col_width(col_number):
            letter = get_column_letter(col_number)

            dimension = worksheet.column_dimensions[letter]

            if dimension.hidden:
                return 0.01

            return max(
                0.01,
                float(dimension.width or 8.43),
            )

        column_width_units = [
            excel_col_width(col)
            for col in range(min_col, max_col + 1)
        ]

        total_column_units = (
            sum(column_width_units)
            or float(col_count)
        )

        # Convert Excel's column proportions to Word twips.
        # The widths add up to the full printable page width.

        column_twips = []
        used_twips = 0

        for index, width_units in enumerate(column_width_units):
            if index == col_count - 1:
                width_twips = (
                    total_table_twips - used_twips
                )
            else:
                width_twips = int(
                    round(
                        total_table_twips
                        * width_units
                        / total_column_units
                    )
                )

                used_twips += width_twips

            column_twips.append(
                max(1, width_twips)
            )

        column_twips[-1] += (
            total_table_twips - sum(column_twips)
        )

        # ---------------------------------------------------------
        # WORD TABLE WIDTH HELPER
        # ---------------------------------------------------------

        def set_table_width(word_table, widths_twips):
            """Apply an exact width and fixed grid to a Word table."""

            properties = word_table._tbl.tblPr

            # Remove duplicate width definitions.
            for existing in list(
                properties.findall(qn("w:tblW"))
            ):
                properties.remove(existing)

            width_element = OxmlElement("w:tblW")
            width_element.set(
                qn("w:w"),
                str(sum(widths_twips)),
            )
            width_element.set(
                qn("w:type"),
                "dxa",
            )
            properties.append(width_element)

            # Use a fixed table layout.
            for existing in list(
                properties.findall(qn("w:tblLayout"))
            ):
                properties.remove(existing)

            layout = OxmlElement("w:tblLayout")
            layout.set(
                qn("w:type"),
                "fixed",
            )
            properties.append(layout)

            # Prevent an automatic Word table indent.
            indent = properties.find(qn("w:tblInd"))

            if indent is None:
                indent = OxmlElement("w:tblInd")
                properties.append(indent)

            indent.set(qn("w:w"), "0")
            indent.set(qn("w:type"), "dxa")

            # Set the actual Word table grid columns.
            grid_columns = word_table._tbl.tblGrid.gridCol_lst

            for index, width_value in enumerate(widths_twips):
                if index < len(grid_columns):
                    grid_columns[index].set(
                        qn("w:w"),
                        str(width_value),
                    )

        # ---------------------------------------------------------
        # RECREATE THE LETTERHEAD AND LOGOS
        # ---------------------------------------------------------

        banner_images = []
        logo_images = []

        for image in excel_images:
            anchor = getattr(image, "anchor", None)

            if anchor is not None and hasattr(anchor, "to"):
                banner_images.append(image)
            else:
                logo_images.append(image)

        if banner_images or logo_images:
            banner = (
                banner_images[0]
                if banner_images
                else None
            )

            banner_width = 0.0

            if banner is not None:
                anchor = banner.anchor

                start_col = max(
                    0,
                    int(anchor._from.col),
                )

                end_col = max(
                    start_col + 1,
                    int(anchor.to.col),
                )

                span_units = sum(
                    excel_col_width(col0 + 1)
                    for col0 in range(
                        start_col,
                        end_col,
                    )
                    if min_col <= col0 + 1 <= max_col
                )

                span_ratio = min(
                    0.92,
                    max(
                        0.30,
                        span_units / total_column_units,
                    ),
                )

                banner_width = (
                    table_width * span_ratio
                )

            # Preserve each logo's aspect ratio.
            logo_widths = []

            for image in logo_images:
                anchor = image.anchor
                ext = getattr(anchor, "ext", None)

                if (
                    ext is not None
                    and getattr(ext, "cx", None)
                ):
                    image_width = (
                        float(ext.cx) / 914400.0
                    )
                else:
                    image_width = (
                        float(image.width or 80) / 96.0
                    )

                logo_widths.append(
                    min(
                        max(0.20, image_width),
                        table_width * 0.14,
                    )
                )

            logo_total = sum(logo_widths)

            indent_left = (
                0.20
                if banner is not None
                else 0.0
            )

            gap = (
                0.10
                if banner is not None and logo_images
                else 0.0
            )

            if banner is not None:
                banner_cell_width = min(
                    table_width - logo_total - gap,
                    banner_width + indent_left + gap,
                )
            else:
                banner_cell_width = max(
                    0.25,
                    table_width - logo_total,
                )

            banner_cell_width = max(
                0.25,
                min(table_width, banner_cell_width),
            )

            if logo_images:
                first_header_width = int(
                    round(banner_cell_width * 1440)
                )

                header_widths = [
                    first_header_width,
                    total_table_twips - first_header_width,
                ]
            else:
                header_widths = [
                    total_table_twips
                ]

            header_table = document.add_table(
                rows=1,
                cols=len(header_widths),
            )

            header_table.alignment = WD_TABLE_ALIGNMENT.LEFT
            header_table.autofit = False

            set_table_width(
                header_table,
                header_widths,
            )

            for index, width_value in enumerate(header_widths):
                cell_width = Inches(
                    width_value / 1440.0
                )

                header_table.columns[index].width = cell_width
                header_table.rows[0].cells[index].width = cell_width

                header_table.rows[0].cells[
                    index
                ].vertical_alignment = (
                    WD_CELL_VERTICAL_ALIGNMENT.CENTER
                )

            first_cell = header_table.cell(0, 0)
            first_paragraph = first_cell.paragraphs[0]

            first_paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            first_paragraph.paragraph_format.space_before = Pt(0)
            first_paragraph.paragraph_format.space_after = Pt(0)

            if banner is not None:
                first_paragraph.paragraph_format.left_indent = Inches(
                    indent_left
                )

                max_banner_width = max(
                    0.25,
                    banner_cell_width
                    - indent_left
                    - gap
                    - 0.10,
                )

                first_paragraph.add_run().add_picture(
                    BytesIO(banner._data()),
                    width=Inches(
                        min(banner_width, max_banner_width)
                    ),
                )

            elif logo_images:
                first_paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                )

                first_paragraph.add_run().add_picture(
                    BytesIO(logo_images[0]._data()),
                    width=Inches(logo_widths[0]),
                )

                logo_images = logo_images[1:]
                logo_widths = logo_widths[1:]

            if len(header_widths) > 1 and logo_images:
                logo_cell = header_table.cell(0, 1)
                logo_paragraph = logo_cell.paragraphs[0]

                logo_paragraph.alignment = (
                    WD_ALIGN_PARAGRAPH.LEFT
                )
                logo_paragraph.paragraph_format.left_indent = Inches(0.05)
                logo_paragraph.paragraph_format.space_before = Pt(0)
                logo_paragraph.paragraph_format.space_after = Pt(0)

                available_logo_width = max(
                    0.2,
                    header_widths[1] / 1440.0 - 0.10,
                )

                for image, width_value in zip(
                    logo_images,
                    logo_widths,
                ):
                    logo_paragraph.add_run().add_picture(
                        BytesIO(image._data()),
                        width=Inches(
                            min(
                                width_value,
                                available_logo_width,
                            )
                        ),
                    )

        # ---------------------------------------------------------
        # CREATE THE REPORT TABLE AT THE FULL PAGE WIDTH
        # ---------------------------------------------------------

        table = document.add_table(
            rows=row_count,
            cols=col_count,
        )

        table.alignment = WD_TABLE_ALIGNMENT.LEFT
        table.autofit = False

        set_table_width(
            table,
            column_twips,
        )

        for index, width_twips in enumerate(column_twips):
            cell_width = Inches(
                width_twips / 1440.0
            )

            table.columns[index].width = cell_width

            for row in table.rows:
                row.cells[index].width = cell_width

        # ---------------------------------------------------------
        # CELL MARGINS
        # ---------------------------------------------------------

        def set_cell_margins(
            cell,
            top=20,
            start=35,
            bottom=20,
            end=35,
        ):
            tc_pr = cell._tc.get_or_add_tcPr()

            tc_mar = tc_pr.find(
                qn("w:tcMar")
            )

            if tc_mar is None:
                tc_mar = OxmlElement("w:tcMar")
                tc_pr.append(tc_mar)

            for edge, amount in (
                ("top", top),
                ("start", start),
                ("bottom", bottom),
                ("end", end),
            ):
                element = tc_mar.find(
                    qn(f"w:{edge}")
                )

                if element is None:
                    element = OxmlElement(
                        f"w:{edge}"
                    )
                    tc_mar.append(element)

                element.set(
                    qn("w:w"),
                    str(amount),
                )

                element.set(
                    qn("w:type"),
                    "dxa",
                )

        default_row_height = float(
            getattr(
                worksheet.sheet_format,
                "defaultRowHeight",
                None,
            ) or 15.0
        )

        # ---------------------------------------------------------
        # DETECT THE TABLE HEADER ROWS
        # ---------------------------------------------------------

        header_name_row = None
        header_end_row = None

        name_aliases = {
            "NAME",
            "INSTRUCTOR",
            "INSTRUCTOR NAME",
            "FACULTY NAME",
        }

        heading_aliases = {
            "SUBJECT",
            "TIME",
            "DAY",
            "ROOM",
            "SIGNATURE",
            "REMARKS",
            "MONITORING TIME",
            "DATE & TIME",
            "DATE AND TIME",
        }

        for candidate_row in range(
            min_row,
            min(max_row, min_row + 10) + 1,
        ):
            values = [
                " ".join(
                    str(
                        worksheet.cell(candidate_row, col).value
                        or ""
                    ).upper().split()
                )
                for col in range(
                    min_col,
                    max_col + 1,
                )
            ]

            if any(
                value in name_aliases
                for value in values
            ):
                header_name_row = candidate_row
                header_end_row = candidate_row

                for next_row in range(
                    candidate_row + 1,
                    min(max_row, candidate_row + 2) + 1,
                ):
                    next_values = [
                        " ".join(
                            str(
                                worksheet.cell(next_row, col).value
                                or ""
                            ).upper().split()
                        )
                        for col in range(
                            min_col,
                            max_col + 1,
                        )
                    ]

                    if any(
                        value in heading_aliases
                        for value in next_values
                    ):
                        header_end_row = next_row

                break

        # ---------------------------------------------------------
        # COPY EXCEL VALUES AND FORMATTING
        # ---------------------------------------------------------

        for excel_row in range(
            min_row,
            max_row + 1,
        ):
            word_row_index = excel_row - min_row

            excel_height = worksheet.row_dimensions[
                excel_row
            ].height

            table.rows[word_row_index].height = Pt(
                float(
                    excel_height or default_row_height
                )
            )

            table.rows[
                word_row_index
            ].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

            for excel_col in range(
                min_col,
                max_col + 1,
            ):
                source = worksheet.cell(
                    excel_row,
                    excel_col,
                )

                target = table.cell(
                    word_row_index,
                    excel_col - min_col,
                )

                target.text = (
                    ""
                    if source.value is None
                    else str(source.value)
                )

                target.vertical_alignment = (
                    WD_CELL_VERTICAL_ALIGNMENT.CENTER
                )

                set_cell_margins(target)

                paragraph = target.paragraphs[0]

                horizontal = source.alignment.horizontal

                paragraph.alignment = {
                    "center": WD_ALIGN_PARAGRAPH.CENTER,
                    "right": WD_ALIGN_PARAGRAPH.RIGHT,
                }.get(
                    horizontal,
                    WD_ALIGN_PARAGRAPH.LEFT,
                )

                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.keep_together = True

                run = (
                    paragraph.runs[0]
                    if paragraph.runs
                    else paragraph.add_run()
                )

                if source.font.name:
                    run.font.name = source.font.name

                if source.font.sz:
                    run.font.size = Pt(
                        source.font.sz
                    )

                run.bold = bool(source.font.bold)
                run.italic = bool(source.font.italic)
                run.underline = bool(source.font.underline)

                # Preserve font colors.
                if (
                    source.font.color
                    and source.font.color.type == "rgb"
                    and source.font.color.rgb
                ):
                    try:
                        run.font.color.rgb = (
                            RGBColor.from_string(
                                str(source.font.color.rgb)[-6:]
                            )
                        )
                    except ValueError:
                        pass

                # Preserve Excel fills, including green header cells.
                if (
                    source.fill
                    and source.fill.fill_type
                    and source.fill.fgColor
                ):
                    fill_color = excel_color_to_hex(
                        source.fill.fgColor
                    )

                    if fill_color:
                        tc_pr = target._tc.get_or_add_tcPr()

                        shading = OxmlElement("w:shd")
                        shading.set(
                            qn("w:fill"),
                            fill_color,
                        )

                        tc_pr.append(shading)

                # Preserve visible borders.
                tc_pr = target._tc.get_or_add_tcPr()

                borders = tc_pr.find(
                    qn("w:tcBorders")
                )

                if borders is None:
                    borders = OxmlElement("w:tcBorders")
                    tc_pr.append(borders)

                for edge_name in (
                    "top",
                    "left",
                    "bottom",
                    "right",
                ):
                    edge = getattr(
                        source.border,
                        edge_name,
                        None,
                    )

                    if edge is None or not edge.style:
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

                    edge_color = excel_color_to_hex(
                        edge.color
                    )

                    if edge_color:
                        border.set(
                            qn("w:color"),
                            edge_color,
                        )

                    borders.append(border)

        # ---------------------------------------------------------
        # RESTORE MERGED CELLS FROM EXCEL
        # ---------------------------------------------------------

        for merged in list(
            worksheet.merged_cells.ranges
        ):
            top_row = max(
                min_row,
                merged.min_row,
            )

            bottom_row = min(
                max_row,
                merged.max_row,
            )

            left_col = max(
                min_col,
                merged.min_col,
            )

            right_col = min(
                max_col,
                merged.max_col,
            )

            if (
                top_row > bottom_row
                or left_col > right_col
            ):
                continue

            if (
                top_row == bottom_row
                and left_col == right_col
            ):
                continue

            anchor_cell = table.cell(
                top_row - min_row,
                left_col - min_col,
            )

            anchor_cell.merge(
                table.cell(
                    bottom_row - min_row,
                    right_col - min_col,
                )
            )

            for extra_paragraph in list(
                anchor_cell.paragraphs[1:]
            ):
                if (
                    not extra_paragraph.text.strip()
                    and not extra_paragraph._p.xpath(
                        ".//w:drawing"
                    )
                ):
                    extra_paragraph._element.getparent().remove(
                        extra_paragraph._element
                    )

        # ---------------------------------------------------------
        # FINAL WIDTH ENFORCEMENT
        # ---------------------------------------------------------
        # Reapply the exact absolute table width after merged cells are
        # reconstructed. This prevents Word Online from shrinking the grid.

        set_table_width(
            table,
            column_twips,
        )

        # ---------------------------------------------------------
        # SAVE THE DOCX FILE
        # ---------------------------------------------------------

        output_path = os.path.join(
            temp_dir,
            "Monitoring_Report.docx",
        )

        document.save(output_path)

        workbook.close()
        workbook = None

        return output_path, temp_dir, template

    except Exception:
        if workbook is not None:
            try:
                workbook.close()
            except Exception:
                pass

        shutil.rmtree(
            temp_dir,
            ignore_errors=True,
        )

        raise

# =============================================================
# REPORT TEMPLATE EXTENSION
# =============================================================

def _get_template_extension(template):
    """Return the uploaded template's extension in lowercase."""
    import os

    return os.path.splitext(
        getattr(template.file, "name", "")
    )[1].lower()


# =============================================================
# BUILD REPORT FILE — COMPATIBILITY HELPER
# =============================================================

def _build_report_file(
    records,
    week_start,
    week_end,
    template_id=None,
    classification="ALL",
    settings=None,
):
    """Return the standard Excel build result."""
    return _build_report_excel_file(
        records,
        week_start,
        week_end,
        template_id=template_id,
        classification=classification,
        settings=settings,
    )


# =============================================================
# BUILD PDF BYTES FROM EXCEL
# =============================================================
def _build_report_pdf_bytes(
    records,
    week_start,
    week_end,
    template_id=None,
    classification="ALL",
    settings=None,
):
    """
    Generate PDF bytes from the selected uploaded Excel template.

    Preserves the template's title, logos, table formatting, and merged
    cells. Adjusts the report columns, wraps long text, centers the
    header images, and enlarges the report title to 12-point font.
    """

    import os
    import shutil
    import subprocess
    import textwrap
    from pathlib import Path
    from copy import copy

    from openpyxl import load_workbook
    from openpyxl.cell.cell import MergedCell
    from openpyxl.drawing.spreadsheet_drawing import (
        AnchorMarker,
        OneCellAnchor,
    )
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.utils.cell import range_boundaries
    from openpyxl.worksheet.properties import PageSetupProperties

    settings = settings or {}

    EMU_PER_PIXEL = 9525
    PIXELS_PER_INCH = 96

    # =========================================================
    # 1. GENERATE THE EXCEL REPORT FROM THE UPLOADED TEMPLATE
    # =========================================================

    build_result = _build_report_excel_file(
        records,
        week_start,
        week_end,
        template_id=template_id,
        classification=classification,
        settings=settings,
    )

    output_xlsx, temp_dir, _template = (
        _checked_report_build_result(
            build_result,
            "_build_report_excel_file",
        )
    )

    try:
        workbook = load_workbook(
            output_xlsx,
            data_only=False,
        )

        try:
            worksheet = workbook.active

            # -------------------------------------------------
            # NORMALIZE EXCEL HEADING TEXT
            # -------------------------------------------------

            def normalize_heading(value):
                if value is None:
                    return ""

                return " ".join(
                    str(value)
                    .replace("\n", " ")
                    .replace("\r", " ")
                    .replace("–", "-")
                    .replace("—", "-")
                    .upper()
                    .split()
                )

            # -------------------------------------------------
            # DETECT THE REPORT TABLE HEADINGS
            # -------------------------------------------------

            aliases = {
                "INSTRUCTOR": {
                    "NAME",
                    "INSTRUCTOR",
                    "INSTRUCTOR NAME",
                    "FACULTY NAME",
                    "TEACHER NAME",
                    "NAME OF INSTRUCTOR",
                    "FACULTY",
                },
                "SUBJECT": {
                    "SUBJECT",
                    "SUBJECT CODE",
                    "COURSE",
                    "SUBJECT / COURSE",
                },
                "TIME": {
                    "TIME",
                    "CLASS TIME",
                    "SCHEDULED TIME",
                    "ORIGINAL TIME",
                },
                "DAY": {
                    "DAY",
                    "CLASS DAY",
                },
                "ROOM": {
                    "ROOM",
                    "ROOM USED",
                    "ACTUAL ROOM",
                    "ACTUAL ROOM USED",
                },
                "MONITORING_TIME": {
                    "MONITORING TIME",
                    "DATE & TIME",
                    "DATE AND TIME",
                    "DATE & TIME MONITORED",
                    "ACTUAL TIME",
                    "MONITORED AT",
                },
                "SIGNATURE": {
                    "SIGNATURE",
                    "INSTRUCTOR SIGNATURE",
                },
                "REMARKS": {
                    "REMARKS",
                    "REMARK",
                    "STATUS",
                },
            }

            header_map = {}
            header_rows = set()
            best_score = 0

            for candidate_row in range(
                1,
                min(worksheet.max_row, 15) + 1,
            ):
                candidate_map = {}
                candidate_rows = set()

                scan_end = min(
                    worksheet.max_row,
                    candidate_row + 2,
                )

                for scan_row in range(
                    candidate_row,
                    scan_end + 1,
                ):
                    for cell in worksheet[scan_row]:
                        heading = normalize_heading(
                            cell.value
                        )

                        if not heading:
                            continue

                        for field, choices in aliases.items():
                            if (
                                heading in choices
                                and field not in candidate_map
                            ):
                                candidate_map[field] = cell.column
                                candidate_rows.add(scan_row)

                score = len(candidate_map)

                if (
                    "INSTRUCTOR" in candidate_map
                    and score >= 3
                    and score > best_score
                ):
                    header_map = candidate_map
                    header_rows = candidate_rows
                    best_score = score

            if (
                not header_map
                or "INSTRUCTOR" not in header_map
            ):
                raise ValueError(
                    "The report headings could not be identified "
                    "in the selected Excel template."
                )

            header_row = max(header_rows)
            data_start = header_row + 1

            # -------------------------------------------------
            # ENLARGE THE FOUR REPORT TITLE LINES TO 12 POINTS
            # -------------------------------------------------

            title_phrases = {
                "OFFICE OF THE DEAN, COLLEGE OF ARTS AND SCIENCES",
                "OFFICE OF THE DEAN",
                "MONITORING OF CLASSES",
                "PART-TIME INSTRUCTORS",
                "2ND SEMESTER, SCHOOL YEAR 2026-2027",
            }

            for row in worksheet.iter_rows(
                min_row=1,
                max_row=max(1, header_row - 1),
            ):
                for cell in row:
                    if not isinstance(cell.value, str):
                        continue

                    title_text = normalize_heading(
                        cell.value
                    )

                    if any(
                        phrase in title_text
                        for phrase in title_phrases
                    ):
                        title_font = copy(cell.font)
                        title_font.sz = 12
                        cell.font = title_font

            # -------------------------------------------------
            # DETERMINE THE ACTUAL LAST REPORT COLUMN
            # -------------------------------------------------

            final_last_column = max(
                header_map.values()
            )

            # -------------------------------------------------
            # DETERMINE THE LAST PRINTED ROW
            # -------------------------------------------------

            final_last_row = None

            if worksheet.print_area:
                try:
                    area = str(
                        worksheet.print_area
                    ).split(",")[0].strip()

                    if "!" in area:
                        area = area.rsplit("!", 1)[1]

                    (
                        _first_col,
                        _first_row,
                        _last_col,
                        final_last_row,
                    ) = range_boundaries(
                        area.replace("$", "")
                    )

                except (ValueError, TypeError):
                    final_last_row = None

            if not final_last_row:
                for row_number in range(
                    1,
                    worksheet.max_row + 1,
                ):
                    if any(
                        not isinstance(cell, MergedCell)
                        and cell.value not in (None, "")
                        for cell in worksheet[row_number]
                    ):
                        final_last_row = row_number

            final_last_row = max(
                header_row,
                final_last_row or worksheet.max_row,
            )

            # -------------------------------------------------
            # WIDEN THE TABLE COLUMNS
            # -------------------------------------------------

            preferred_widths = {
                "INSTRUCTOR": 24.0,
                "SUBJECT": 21.0,
                "TIME": 20.0,
                "DAY": 13.0,
                "ROOM": 16.0,
                "MONITORING_TIME": 23.0,
                "SIGNATURE": 16.0,
                "REMARKS": 31.0,
            }

            field_by_column = {
                column: field
                for field, column in header_map.items()
            }

            for column in range(
                1,
                final_last_column + 1,
            ):
                column_letter = get_column_letter(
                    column
                )

                field = field_by_column.get(column)

                if field:
                    worksheet.column_dimensions[
                        column_letter
                    ].width = preferred_widths.get(
                        field,
                        16.0,
                    )
                else:
                    # Keep intentional spacer columns narrow.
                    worksheet.column_dimensions[
                        column_letter
                    ].width = 3.5

            # -------------------------------------------------
            # PAPER SIZE, ORIENTATION, AND MARGINS
            # -------------------------------------------------

            paper_sizes = {
                "A4": (8.27, 11.69),
                "LETTER": (8.5, 11.0),
                "LEGAL": (8.5, 14.0),
                "FOLIO": (8.5, 13.0),
            }

            excel_paper_sizes = {
                "1": (8.5, 11.0),
                "2": (8.5, 11.0),
                "3": (11.0, 17.0),
                "4": (17.0, 11.0),
                "5": (8.5, 14.0),
                "6": (5.5, 8.5),
                "7": (7.25, 10.5),
                "8": (11.69, 16.54),
                "9": (8.27, 11.69),
                "10": (8.27, 11.69),
                "11": (5.83, 8.27),
                "12": (9.84, 13.90),
                "13": (7.17, 10.12),
                "14": (8.5, 13.0),
                "15": (8.5, 10.83),
                "16": (10.0, 14.0),
                "17": (11.0, 17.0),
            }

            requested_paper = str(
                settings.get("paper_size", "")
            ).upper()

            if requested_paper in paper_sizes:
                page_width, page_height = paper_sizes[
                    requested_paper
                ]
            else:
                page_width, page_height = (
                    excel_paper_sizes.get(
                        str(worksheet.page_setup.paperSize),
                        (8.27, 11.69),
                    )
                )

            orientation = str(
                settings.get(
                    "orientation",
                    worksheet.page_setup.orientation or "landscape",
                )
            ).lower()

            if orientation == "landscape":
                page_width, page_height = (
                    max(page_width, page_height),
                    min(page_width, page_height),
                )
                worksheet.page_setup.orientation = "landscape"
            else:
                page_width, page_height = (
                    min(page_width, page_height),
                    max(page_width, page_height),
                )
                worksheet.page_setup.orientation = "portrait"

            paper_codes = {
                "A4": "9",
                "LETTER": "1",
                "LEGAL": "5",
                "FOLIO": "14",
            }

            if requested_paper in paper_codes:
                worksheet.page_setup.paperSize = (
                    paper_codes[requested_paper]
                )

            margin_mode = str(
                settings.get("margins", "normal")
            ).lower()

            margin = (
                0.30
                if margin_mode == "narrow"
                else 0.50
            )

            worksheet.page_margins.left = margin
            worksheet.page_margins.right = margin
            worksheet.page_margins.top = margin
            worksheet.page_margins.bottom = margin

            printable_width = max(
                1.0,
                page_width - (margin * 2),
            )

            # -------------------------------------------------
            # WRAP THE TABLE HEADINGS
            # -------------------------------------------------

            for row_number in header_rows:
                has_long_heading = False

                for column in range(
                    1,
                    final_last_column + 1,
                ):
                    cell = worksheet.cell(
                        row_number,
                        column,
                    )

                    if isinstance(cell, MergedCell):
                        continue

                    if cell.value in (None, ""):
                        continue

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center",
                        wrap_text=True,
                        text_rotation=cell.alignment.textRotation,
                        indent=cell.alignment.indent,
                    )

                    if len(str(cell.value)) > 10:
                        has_long_heading = True

                old_height = (
                    worksheet.row_dimensions[
                        row_number
                    ].height or 15.0
                )

                worksheet.row_dimensions[
                    row_number
                ].height = max(
                    float(old_height),
                    30.0 if has_long_heading else 24.0,
                )

            # -------------------------------------------------
            # WRAP REMARKS AND PREVENT TEXT OVERLAPPING
            # -------------------------------------------------

            total_column_width_units = sum(
                float(
                    worksheet.column_dimensions[
                        get_column_letter(column)
                    ].width or 8.43
                )
                for column in range(
                    1,
                    final_last_column + 1,
                )
            )

            approximate_content_width = (
                total_column_width_units * 7.0 / 96.0
            )

            estimated_print_scale = min(
                1.0,
                printable_width
                / max(
                    1.0,
                    approximate_content_width,
                ),
            )

            for row_number in range(
                data_start,
                final_last_row + 1,
            ):
                max_lines = 1
                largest_font_size = 10.0

                for field, column in header_map.items():
                    cell = worksheet.cell(
                        row_number,
                        column,
                    )

                    if isinstance(cell, MergedCell):
                        continue

                    value = str(cell.value or "")
                    previous_alignment = cell.alignment

                    if field == "REMARKS":
                        horizontal = "left"
                        vertical = "top"
                    else:
                        horizontal = (
                            previous_alignment.horizontal
                            or "center"
                        )
                        vertical = "center"

                    cell.alignment = Alignment(
                        horizontal=horizontal,
                        vertical=vertical,
                        wrap_text=True,
                        text_rotation=previous_alignment.textRotation,
                        indent=previous_alignment.indent,
                        shrink_to_fit=False,
                    )

                    if cell.font and cell.font.sz:
                        largest_font_size = max(
                            largest_font_size,
                            float(cell.font.sz),
                        )

                    if not value:
                        continue

                    # Don't count merged instructor names as if they
                    # must fit inside a single short row.
                    if field == "INSTRUCTOR":
                        vertically_merged = any(
                            merged.min_col == column
                            and merged.max_col == column
                            and merged.min_row == row_number
                            and merged.max_row > row_number
                            for merged in worksheet.merged_cells.ranges
                        )

                        if vertically_merged:
                            continue

                    column_width = float(
                        worksheet.column_dimensions[
                            get_column_letter(column)
                        ].width or 8.43
                    )

                    approximate_characters_per_line = max(
                        5,
                        int(
                            column_width
                            * max(
                                0.35,
                                estimated_print_scale,
                            )
                            * 0.90
                        ),
                    )

                    required_lines = 0

                    for paragraph_text in value.splitlines() or [value]:
                        wrapped_text = textwrap.wrap(
                            paragraph_text,
                            width=approximate_characters_per_line,
                            break_long_words=True,
                            break_on_hyphens=True,
                        )

                        required_lines += max(
                            1,
                            len(wrapped_text),
                        )

                    max_lines = max(
                        max_lines,
                        required_lines,
                    )

                needed_height = (
                    max_lines * largest_font_size * 1.35 + 6
                )

                current_height = (
                    worksheet.row_dimensions[
                        row_number
                    ].height or 18.0
                )

                worksheet.row_dimensions[
                    row_number
                ].height = max(
                    float(current_height),
                    float(needed_height),
                )

            # -------------------------------------------------
            # KEEP TITLE MERGES INSIDE THE REPORT WIDTH
            # -------------------------------------------------

            for merged_range in list(
                worksheet.merged_cells.ranges
            ):
                if (
                    merged_range.min_row < data_start
                    and merged_range.min_col <= final_last_column
                    < merged_range.max_col
                ):
                    start_row = merged_range.min_row
                    end_row = merged_range.max_row
                    start_col = merged_range.min_col

                    worksheet.unmerge_cells(
                        str(merged_range)
                    )

                    worksheet.merge_cells(
                        start_row=start_row,
                        start_column=start_col,
                        end_row=end_row,
                        end_column=final_last_column,
                    )

            # -------------------------------------------------
            # CENTER THE UNIVERSITY HEADER AND COLLEGE LOGO
            # -------------------------------------------------

            excel_images = list(
                getattr(worksheet, "_images", []) or []
            )

            banner_images = []
            logo_images = []

            for image in excel_images:
                anchor = getattr(image, "anchor", None)

                if anchor is not None and hasattr(anchor, "to"):
                    banner_images.append(image)
                else:
                    logo_images.append(image)

            def column_width_pixels(column):
                width = float(
                    worksheet.column_dimensions[
                        get_column_letter(column)
                    ].width or 8.43
                )

                return max(
                    1,
                    int(round(width * 7 + 5)),
                )

            column_pixels = {
                column: column_width_pixels(column)
                for column in range(
                    1,
                    final_last_column + 1,
                )
            }

            total_grid_pixels = sum(
                column_pixels.values()
            )

            print_scale = min(
                1.0,
                printable_width
                * PIXELS_PER_INCH
                / max(1, total_grid_pixels),
            )

            # Center the entire header image group over the report.
            group_width_pixels = int(
                total_grid_pixels * 0.88
            )

            group_start_pixels = (
                total_grid_pixels - group_width_pixels
            ) / 2

            if banner_images and logo_images:
                banner_group_width = int(
                    group_width_pixels * 0.85
                )
                logo_group_width = int(
                    group_width_pixels * 0.10
                )
                group_gap = (
                    group_width_pixels
                    - banner_group_width
                    - logo_group_width
                )
            elif banner_images:
                banner_group_width = group_width_pixels
                logo_group_width = 0
                group_gap = 0
            else:
                banner_group_width = 0
                logo_group_width = group_width_pixels
                group_gap = 0

            def horizontal_anchor(position_pixels):
                remaining = max(
                    0.0,
                    float(position_pixels),
                )

                for column in range(
                    1,
                    final_last_column + 1,
                ):
                    width_pixels = column_pixels[column]

                    if remaining < width_pixels:
                        return AnchorMarker(
                            col=column - 1,
                            colOff=int(
                                remaining * EMU_PER_PIXEL
                            ),
                            row=0,
                            rowOff=0,
                        )

                    remaining -= width_pixels

                return AnchorMarker(
                    col=final_last_column - 1,
                    colOff=int(
                        remaining * EMU_PER_PIXEL
                    ),
                    row=0,
                    rowOff=0,
                )

            largest_image_height_pixels = 0.0

            for banner in banner_images:
                source_width = float(
                    banner.width or 900
                )
                source_height = float(
                    banner.height or 155
                )

                image_width = max(
                    1,
                    int(banner_group_width),
                )

                image_height = (
                    image_width
                    * source_height
                    / max(1.0, source_width)
                )

                banner.anchor = OneCellAnchor(
                    _from=horizontal_anchor(
                        group_start_pixels
                    ),
                    ext=XDRPositiveSize2D(
                        cx=int(
                            image_width * EMU_PER_PIXEL
                        ),
                        cy=int(
                            image_height * EMU_PER_PIXEL
                        ),
                    ),
                )

                largest_image_height_pixels = max(
                    largest_image_height_pixels,
                    image_height,
                )

            if logo_images:
                logo_width = max(
                    1,
                    int(
                        logo_group_width
                        / len(logo_images)
                    ),
                )

                logo_position = (
                    group_start_pixels
                    + banner_group_width
                    + group_gap
                )

                for logo in logo_images:
                    source_width = float(
                        logo.width or 186
                    )
                    source_height = float(
                        logo.height or 156
                    )

                    image_width = logo_width

                    image_height = (
                        image_width
                        * source_height
                        / max(1.0, source_width)
                    )

                    logo.anchor = OneCellAnchor(
                        _from=horizontal_anchor(
                            logo_position
                        ),
                        ext=XDRPositiveSize2D(
                            cx=int(
                                image_width * EMU_PER_PIXEL
                            ),
                            cy=int(
                                image_height * EMU_PER_PIXEL
                            ),
                        ),
                    )

                    largest_image_height_pixels = max(
                        largest_image_height_pixels,
                        image_height,
                    )

                    logo_position += image_width

            if largest_image_height_pixels > 0:
                required_header_height = (
                    largest_image_height_pixels
                    * 72
                    / PIXELS_PER_INCH
                    / max(0.01, print_scale)
                    + 2
                )

                current_header_height = (
                    worksheet.row_dimensions[1].height
                    or 15.0
                )

                worksheet.row_dimensions[1].height = max(
                    float(current_header_height),
                    float(required_header_height),
                )

            # -------------------------------------------------
            # FINAL PRINT SETTINGS
            # -------------------------------------------------

            worksheet.print_area = (
                f"A1:"
                f"{get_column_letter(final_last_column)}"
                f"{final_last_row}"
            )

            worksheet.page_setup.scale = None
            worksheet.page_setup.fitToWidth = 1
            worksheet.page_setup.fitToHeight = 0

            if worksheet.sheet_properties.pageSetUpPr is None:
                worksheet.sheet_properties.pageSetUpPr = (
                    PageSetupProperties(
                        fitToPage=True,
                    )
                )
            else:
                worksheet.sheet_properties.pageSetUpPr.fitToPage = True

            worksheet.sheet_properties.pageSetUpPr.autoPageBreaks = False

            # Save the temporary report, not the original uploaded template.
            workbook.save(output_xlsx)

        finally:
            workbook.close()

        # =====================================================
        # 2. CONVERT THE ADJUSTED EXCEL REPORT INTO PDF
        # =====================================================

        pdf_dir = os.path.join(
            temp_dir,
            "pdf",
        )

        os.makedirs(
            pdf_dir,
            exist_ok=True,
        )

        candidates = [
            os.environ.get("LIBREOFFICE_PATH", ""),
            shutil.which("libreoffice") or "",
            shutil.which("soffice") or "",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]

        office = next(
            (
                path
                for path in candidates
                if path and (
                    os.path.isfile(path)
                    or shutil.which(path)
                )
            ),
            None,
        )

        if not office:
            raise FileNotFoundError(
                "LibreOffice was not found. Install it on the computer "
                "running Django or set LIBREOFFICE_PATH to the "
                "soffice executable."
            )

        profile_dir = os.path.join(
            temp_dir,
            "libreoffice_profile",
        )

        os.makedirs(
            profile_dir,
            exist_ok=True,
        )

        command = [
            office,
            (
                f"-env:UserInstallation="
                f"{Path(profile_dir).resolve().as_uri()}"
            ),
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            pdf_dir,
            output_xlsx,
        ]

        try:
            conversion_result = subprocess.run(
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

        except subprocess.TimeoutExpired as exc:
            raise RuntimeError(
                "Excel-to-PDF conversion timed out."
            ) from exc

        pdf_path = os.path.join(
            pdf_dir,
            "Monitoring_Report.pdf",
        )

        if (
            conversion_result.returncode != 0
            or not os.path.isfile(pdf_path)
            or os.path.getsize(pdf_path) == 0
        ):
            raise RuntimeError(
                "LibreOffice could not create the PDF. "
                f"Output: {conversion_result.stdout.strip()} "
                f"Error: {conversion_result.stderr.strip()}"
            )

        with open(pdf_path, "rb") as pdf_file:
            pdf_bytes = pdf_file.read()

        if not pdf_bytes.startswith(b"%PDF"):
            raise RuntimeError(
                "LibreOffice produced an invalid PDF file."
            )

        return pdf_bytes

    finally:
        shutil.rmtree(
            temp_dir,
            ignore_errors=True,
        )

# =============================================================
# REPORTS PAGE
# =============================================================

@login_required
@user_passes_test(is_admin_or_monitoring_staff, login_url="dashboard")
def reports(request):
    """Display the selected week's monitoring records and report controls."""

    week_start, week_end, current_week_start = (
        _report_week_range(request)
    )

    all_records = MonitoringRecord.objects.only(
        "record_datetime"
    )

    week_starts = {current_week_start}

    for record in all_records:
        if not record.record_datetime:
            continue

        record_dt = record.record_datetime

        if timezone.is_aware(record_dt):
            record_dt = timezone.localtime(record_dt)

        record_date = record_dt.date()

        week_starts.add(
            record_date - timedelta(days=record_date.weekday())
        )

    weeks = []

    for start in sorted(week_starts, reverse=True):
        end = start + timedelta(days=6)

        weeks.append({
            "value": start.strftime("%Y-%m-%d"),
            "label": (
                f"{start:%B %d, %Y} - {end:%B %d, %Y}"
            ),
        })

    records, instructor_search = _report_records(
        request,
        week_start,
        week_end,
    )

    classification = _report_classification(request)

    templates = list(
        ReportTemplate.objects.all().order_by("name")
    )

    selected_template_id = _report_template_id(request) or ""

    instructors = list(Instructor.objects.all())

    instructors.sort(
        key=lambda instructor: _report_last_name_sort_key(
            instructor.name
        )
    )

    return render(
        request,
        "instructors/reports/list.html",
        {
            "records": records,
            "weeks": weeks,
            "selected_week": week_start.strftime("%Y-%m-%d"),
            "week_start": week_start,
            "week_end": week_end,
            "previous_week": week_start - timedelta(days=7),
            "next_week": week_start + timedelta(days=7),
            "current_week_start": current_week_start,
            "current_week_end": current_week_start + timedelta(days=6),
            "instructors": instructors,
            "instructor_search": instructor_search,
            "selected_classification": classification,
            "selected_template_id": selected_template_id,
            "report_settings": _report_settings(request),
            "report_templates": templates,
            "windows_printers": _windows_printers(),
        },
    )


# =============================================================
# PRINT REPORT
# =============================================================

@login_required
@user_passes_test(is_admin_or_monitoring_staff, login_url="dashboard")
def report_print(request):
    """Generate the selected report and send it to a Windows printer."""

    import os
    import shutil
    import subprocess
    import tempfile
    import threading
    from pathlib import Path

    if os.name != "nt":
        return JsonResponse(
            {
                "success": False,
                "message": (
                    "Direct printer access requires Django to run on Windows."
                ),
            },
            status=400,
        )

    week_start, week_end, _ = _report_week_range(request)
    classification = _report_classification(request)
    template_id = _report_template_id(request)

    records, _ = _report_records(
        request,
        week_start,
        week_end,
    )

    if not records:
        return JsonResponse(
            {
                "success": False,
                "message": (
                    "No monitoring records match your selected week and filters."
                ),
            },
            status=400,
        )

    printer_name = _request_value(
        request,
        "printer",
        "",
    ).strip()

    if not printer_name:
        return JsonResponse(
            {
                "success": False,
                "message": "Select a printer first.",
            },
            status=400,
        )

    valid_printers = {
        printer.get("name", "")
        for printer in _windows_printers()
    }

    if printer_name not in valid_printers:
        return JsonResponse(
            {
                "success": False,
                "message": (
                    f"Printer '{printer_name}' was not found "
                    "on the Django computer."
                ),
            },
            status=400,
        )

    temp_pdf_path = None
    keep_file_for_spooler = False

    def remove_temp_file(path):
        try:
            if path and os.path.exists(path):
                os.remove(path)
        except OSError:
            pass

    try:
        pdf_bytes = _build_report_pdf_bytes(
            records,
            week_start,
            week_end,
            template_id=template_id,
            classification=classification,
            settings=_report_settings(request),
        )

        file_descriptor, temp_pdf_path = tempfile.mkstemp(
            prefix="teacher_monitoring_print_",
            suffix=".pdf",
        )

        os.close(file_descriptor)

        with open(temp_pdf_path, "wb") as pdf_file:
            pdf_file.write(pdf_bytes)

        paths = [
            os.environ.get("LIBREOFFICE_PATH", ""),
            shutil.which("libreoffice") or "",
            shutil.which("soffice") or "",
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]

        office = next(
            (
                path
                for path in paths
                if path and (
                    os.path.isfile(path)
                    or shutil.which(path)
                )
            ),
            None,
        )

        if office:
            profile_dir = tempfile.mkdtemp(
                prefix="teacher_monitoring_printer_profile_"
            )

            try:
                command = [
                    office,
                    f"-env:UserInstallation={Path(profile_dir).resolve().as_uri()}",
                    "--headless",
                    "--pt",
                    printer_name,
                    temp_pdf_path,
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

                if result.returncode != 0:
                    detail = (
                        result.stderr.strip()
                        or result.stdout.strip()
                        or "Unknown printer error."
                    )

                    raise RuntimeError(detail)

            finally:
                shutil.rmtree(
                    profile_dir,
                    ignore_errors=True,
                )

        else:
            # Fallback to the installed Windows PDF application's PrintTo action.
            os.startfile(
                temp_pdf_path,
                "printto",
                printer_name,
            )

        # Give the operating system's print spooler time to read the PDF.
        threading.Timer(
            120,
            remove_temp_file,
            args=(temp_pdf_path,),
        ).start()

        keep_file_for_spooler = True

        return JsonResponse({
            "success": True,
            "message": f"The report was sent to '{printer_name}'.",
            "printer": printer_name,
        })

    except Exception as exc:
        return JsonResponse(
            {
                "success": False,
                "message": "The report could not be printed.",
                "detail": str(exc),
            },
            status=500,
        )

    finally:
        if temp_pdf_path and not keep_file_for_spooler:
            remove_temp_file(temp_pdf_path)


# =============================================================
# PDF DOWNLOAD AND PREVIEW
# =============================================================

@login_required
@user_passes_test(is_admin_or_monitoring_staff, login_url="dashboard")
def report_pdf(request):
    """
    Return the generated PDF inline for preview or as a download.

    ?preview=1 displays the PDF in the preview iframe.
    Without preview=1, the PDF downloads normally.
    """

    week_start, week_end, _ = _report_week_range(request)
    classification = _report_classification(request)
    template_id = _report_template_id(request)

    records, _ = _report_records(
        request,
        week_start,
        week_end,
    )

    try:
        pdf_bytes = _build_report_pdf_bytes(
            records,
            week_start,
            week_end,
            template_id=template_id,
            classification=classification,
            settings=_report_settings(request),
        )

        filename = (
            f"Monitoring_Report_{week_start:%Y-%m-%d}.pdf"
        )

        response = HttpResponse(
            pdf_bytes,
            content_type="application/pdf",
        )

        disposition = (
            "inline"
            if request.GET.get("preview") == "1"
            else "attachment"
        )

        response["Content-Disposition"] = (
            f'{disposition}; filename="{filename}"'
        )

        return response

    except Exception as exc:
        return HttpResponse(
            f"Report PDF could not be generated: {exc}",
            status=500,
            content_type="text/plain; charset=utf-8",
        )


# =============================================================
# DOCX DOWNLOAD
# =============================================================

@login_required
@user_passes_test(is_admin_or_monitoring_staff, login_url="dashboard")
def report_docx(request):
    """Generate and download the Word report."""

    import os
    import shutil

    week_start, week_end, _ = _report_week_range(request)
    classification = _report_classification(request)
    template_id = _report_template_id(request)

    records, _ = _report_records(
        request,
        week_start,
        week_end,
    )

    temp_dir = None

    try:
        result = _build_report_docx_file(
            records,
            week_start,
            week_end,
            template_id=template_id,
            classification=classification,
            settings=_report_settings(request),
        )

        output_path, temp_dir, _template = (
            _checked_report_build_result(
                result,
                "_build_report_docx_file",
            )
        )

        if not os.path.isfile(output_path):
            raise FileNotFoundError(
                "The DOCX builder returned a missing file path."
            )

        with open(output_path, "rb") as file:
            content = file.read()

        if not content:
            raise RuntimeError(
                "The generated Word report is empty."
            )

        response = HttpResponse(
            content,
            content_type=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
        )

        response["Content-Disposition"] = (
            f'attachment; filename="Monitoring_Report_{week_start:%Y-%m-%d}.docx"'
        )

        return response

    except Exception as exc:
        return HttpResponse(
            f"DOCX report could not be generated: {exc}",
            status=500,
            content_type="text/plain; charset=utf-8",
        )

    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)


# =============================================================
# EXCEL DOWNLOAD
# =============================================================

@login_required
@user_passes_test(is_admin_or_monitoring_staff, login_url="dashboard")
def report_excel(request):
    """Generate and download the Excel report."""

    import os
    import shutil

    week_start, week_end, _ = _report_week_range(request)
    classification = _report_classification(request)
    template_id = _report_template_id(request)

    records, _ = _report_records(
        request,
        week_start,
        week_end,
    )

    temp_dir = None

    try:
        result = _build_report_excel_file(
            records,
            week_start,
            week_end,
            template_id=template_id,
            classification=classification,
            settings=_report_settings(request),
        )

        output_path, temp_dir, _template = (
            _checked_report_build_result(
                result,
                "_build_report_excel_file",
            )
        )

        if not os.path.isfile(output_path):
            raise FileNotFoundError(
                "The Excel builder returned a missing file path."
            )

        with open(output_path, "rb") as file:
            content = file.read()

        if not content:
            raise RuntimeError(
                "The generated Excel report is empty."
            )

        response = HttpResponse(
            content,
            content_type=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

        response["Content-Disposition"] = (
            f'attachment; filename="Monitoring_Report_{week_start:%Y-%m-%d}.xlsx"'
        )

        return response

    except Exception as exc:
        return HttpResponse(
            f"Excel report could not be generated: {exc}",
            status=500,
            content_type="text/plain; charset=utf-8",
        )

    finally:
        if temp_dir:
            shutil.rmtree(temp_dir, ignore_errors=True)

# ==========================================
# REPORT TEMPLATES
# ==========================================


@login_required
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
def schedule_import_part_time(request):

    return _schedule_import(
        request,
        "PART_TIME"
    )


@login_required
@user_passes_test(is_system_admin, login_url="dashboard")
def schedule_import_regular(request):

    return _schedule_import(
        request,
        "REGULAR"
    )


@login_required
@user_passes_test(is_system_admin, login_url="dashboard")
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
@user_passes_test(is_system_admin, login_url="dashboard")
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

class RegisteredMainAdminPasswordResetForm(PasswordResetForm):

    def clean_email(self):
        email = self.cleaned_data["email"].strip()

        exists = SystemAdministrator.objects.filter(
            is_main_admin=True,
            user__is_active=True,
            user__email__iexact=email,
        ).exists()

        if not exists:
            raise ValidationError(
                "This email does not match the registered "
                "Main Administrator email. Please use your "
                "registered email address."
            )

        return email

    def get_users(self, email):
        admins = SystemAdministrator.objects.filter(
            is_main_admin=True,
            user__is_active=True,
            user__email__iexact=email,
        ).select_related("user")

        for admin in admins:
            user = admin.user

            if user.has_usable_password() and user.email:
                yield user


class RegisteredMainAdminPasswordResetView(PasswordResetView):
    form_class = RegisteredMainAdminPasswordResetForm
    template_name = "registration/password_reset_form.html"
    email_template_name = "registration/password_reset_email.html"
    subject_template_name = "registration/password_reset_subject.txt"
    success_url = reverse_lazy("password_reset_done")

# =========================================================
# EXPLICIT MAIN ADMINISTRATOR PASSWORD RESET
# =========================================================

# =========================================================
# STRICT MAIN ADMINISTRATOR PASSWORD RESET
# =========================================================

class MainAdminOnlyPasswordResetForm(PasswordResetForm):
    """
    Accept only the email registered to the active
    Main Administrator account.
    """


    def send_mail(
        self,
        subject_template_name,
        email_template_name,
        context,
        from_email,
        to_email,
        html_email_template_name=None,
    ):
        logger.warning(
            "PASSWORD RESET: send_mail() called."
        )

        try:
            result = super().send_mail(
                subject_template_name,
                email_template_name,
                context,
                from_email,
                to_email,
                html_email_template_name=html_email_template_name,
            )

            logger.warning(
                "PASSWORD RESET: send_mail() completed successfully."
            )

            return result

        except Exception:
            logger.exception(
                "PASSWORD RESET: send_mail() failed."
            )
            raise



    def clean_email(self):
        email = self.cleaned_data["email"].strip()

        exists = SystemAdministrator.objects.filter(
            is_main_admin=True,
            user__is_active=True,
            user__email__iexact=email,
        ).exists()

        if not exists:
            raise ValidationError(
                "Incorrect email address. Please enter the email "
                "registered to your Main Administrator account.",
                code="incorrect_admin_email",
            )

        return email

    def get_users(self, email):
        records = SystemAdministrator.objects.filter(
            is_main_admin=True,
            user__is_active=True,
            user__email__iexact=email,
        ).select_related("user")

        logger.warning(
            "PASSWORD RESET: matching Main Administrator records found: %s",
            records.count(),
        )

        for record in records:
            user = record.user

            eligible = (
                    user.is_active
                    and user.has_usable_password()
                    and bool(user.email)
                    and user.email.strip().casefold()
                    == email.strip().casefold()
            )

            logger.warning(
                "PASSWORD RESET: account eligible for email delivery: %s",
                eligible,
            )

            if eligible:
                yield user


def main_admin_password_reset(request):
    logger.warning(
        "PASSWORD RESET VIEW REACHED: method=%s path=%s",
        request.method,
        request.path,
    )

    form = MainAdminOnlyPasswordResetForm(
        request.POST or None
    )

    if request.method == "POST":
        logger.warning("PASSWORD RESET: POST received.")

        if form.is_valid():
            logger.warning(
                "PASSWORD RESET: form validation passed."
            )

            try:
                sent_count = form.save(
                    request=request,
                    use_https=request.is_secure(),
                    from_email=settings.DEFAULT_FROM_EMAIL,
                    email_template_name=(
                        "registration/password_reset_email.html"
                    ),
                    subject_template_name=(
                        "registration/password_reset_subject.txt"
                    ),
                )

                logger.warning(
                    "PASSWORD RESET: save() completed; "
                    "email backend returned %s.",
                    sent_count,
                )

            except Exception:
                logger.exception(
                    "PASSWORD RESET: email processing raised an exception."
                )
                raise

            return redirect("password_reset_done")

        logger.warning(
            "PASSWORD RESET: form validation failed. Errors: %s",
            form.errors.as_json(),
        )

    return render(
        request,
        "registration/password_reset_form.html",
        {"form": form},
    )
