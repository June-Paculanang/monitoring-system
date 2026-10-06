from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User


class Instructor(models.Model):

    CLASSIFICATION_CHOICES = [
        ("PART_TIME", "Part-Time"),
        ("REGULAR", "Regular"),
    ]

    name = models.CharField(max_length=150)

    classification = models.CharField(
        max_length=20,
        choices=CLASSIFICATION_CHOICES,
        default="PART_TIME"
    )

    department = models.CharField(
        max_length=150,
        blank=True
    )

    contact = models.CharField(
        max_length=50,
        blank=True
    )

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name


class Subject(models.Model):
    subject_code = models.CharField(max_length=30)
    subject_name = models.CharField(
        max_length=150,
        blank=True
    )

    class Meta:
        ordering = ["subject_code"]

    def __str__(self):
        return f"{self.subject_code} - {self.subject_name}"


class Room(models.Model):
    room_name = models.CharField(max_length=50)

    class Meta:
        ordering = ["room_name"]

    def __str__(self):
        return self.room_name


class Schedule(models.Model):

    DAYS = [
        ("M", "Monday"),
        ("T", "Tuesday"),
        ("W", "Wednesday"),
        ("TH", "Thursday"),
        ("F", "Friday"),
        ("S", "Saturday"),
    ]

    instructor = models.ForeignKey(
        Instructor,
        on_delete=models.CASCADE,
        related_name="schedules"
    )

    subject = models.ForeignKey(
        Subject,
        on_delete=models.CASCADE,
        related_name="schedules"
    )

    room = models.ForeignKey(
        Room,
        on_delete=models.CASCADE,
        related_name="schedules"
    )

    day = models.CharField(
        max_length=5,
        choices=DAYS
    )

    start_time = models.TimeField()

    end_time = models.TimeField()

    def __str__(self):
        return f"{self.instructor} - {self.subject} - {self.day}"

class MonitoringRecord(models.Model):

    REMARK_CHOICES = [
        ("Monitored", "Monitored"),
        ("Missed", "Missed"),
        ("Late", "Late"),
        ("Not Monitored", "Not Monitored"),
    ]

    schedule = models.ForeignKey(
        Schedule,
        on_delete=models.CASCADE,
        related_name="monitoring_records"
    )

    signature = models.CharField(
        max_length=150,
        blank=True
    )

    transfer_room = models.CharField(
        max_length=50,
        blank=True
    )

    remarks = models.TextField(
        blank=True,
        default="Monitored"
    )

    other_remarks = models.TextField(
        blank=True
    )

    missed_reason = models.TextField(
        blank=True
    )

    record_datetime = models.DateTimeField(
        default=timezone.now
    )

    def __str__(self):
        return f"{self.schedule} - {self.record_datetime}"

class ReportTemplate(models.Model):
    name = models.CharField(
        max_length=150
    )

    school_year = models.CharField(
        max_length=20
    )

    file = models.FileField(
        upload_to="report_templates/"
    )

    is_active = models.BooleanField(
        default=False
    )

    uploaded_at = models.DateTimeField(
        auto_now_add=True
    )

    def __str__(self):
        return f"{self.name} - {self.school_year}"


class SystemAdministrator(models.Model):
    """
    Identifies the administrators of the Teacher Class Monitoring System.

    is_main_admin=True:
        The Main Administrator / Dean.
        Only this account can create and manage other Admin accounts.

    is_main_admin=False:
        A regular Admin account created by the Main Administrator.
        This account can use the system but cannot create or manage
        other Admin accounts.
    """

    user = models.OneToOneField(
        User,
        on_delete=models.CASCADE,
        related_name="system_administrator"
    )

    is_main_admin = models.BooleanField(
        default=False
    )

    def __str__(self):
        if self.is_main_admin:
            return f"{self.user.username} - Main Admin"
        return f"{self.user.username} - Admin"