from django.contrib import admin

from .models import (
    Instructor,
    Subject,
    Room,
    Schedule,
    MonitoringRecord,
    ReportTemplate,
)


@admin.register(Instructor)
class InstructorAdmin(admin.ModelAdmin):
    list_display = ("name", "department", "contact")
    search_fields = ("name", "department")


@admin.register(Subject)
class SubjectAdmin(admin.ModelAdmin):
    list_display = ("subject_code", "subject_name")
    search_fields = ("subject_code", "subject_name")


@admin.register(Room)
class RoomAdmin(admin.ModelAdmin):
    list_display = ("room_name",)
    search_fields = ("room_name",)


@admin.register(Schedule)
class ScheduleAdmin(admin.ModelAdmin):
    list_display = (
        "instructor",
        "subject",
        "room",
        "day",
        "start_time",
        "end_time",
    )

    list_filter = ("day", "room", "subject")
    search_fields = (
        "instructor__name",
        "subject__subject_code",
        "subject__subject_name",
        "room__room_name",
    )


@admin.register(MonitoringRecord)
class MonitoringRecordAdmin(admin.ModelAdmin):
    list_display = (
        "schedule",
        "signature",
        "remarks",
        "record_datetime",
    )

    list_filter = ("record_datetime",)
    search_fields = (
        "schedule__instructor__name",
        "schedule__subject__subject_code",
        "remarks",
    )

@admin.register(ReportTemplate)
class ReportTemplateAdmin(admin.ModelAdmin):
    list_display = (
        "name",
        "school_year",
        "is_active",
        "uploaded_at",
    )

    list_filter = (
        "school_year",
        "is_active",
    )

    search_fields = (
        "name",
        "school_year",
    )
