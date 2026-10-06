from django import forms
from .models import Instructor
from datetime import datetime
from django.utils import timezone
from django.contrib.auth.models import User

from .models import (
    Instructor,
    Subject,
    Room,
    Schedule,
    MonitoringRecord,
    ReportTemplate,
)

class InstructorForm(forms.ModelForm):

    class Meta:

        model = Instructor

        fields = [
            "name",
            "classification",
            "department",
        ]

        widgets = {

            "name": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Enter instructor name"
                }
            ),

            "classification": forms.Select(
                attrs={
                    "class": "form-control"
                }
            ),

            "department": forms.TextInput(
                attrs={
                    "class": "form-control",
                    "placeholder": "Enter department"
                }
            ),
        }


class SubjectForm(forms.ModelForm):
    class Meta:
        model = Subject
        fields = ["subject_code", "subject_name"]
        labels = {
            "subject_code": "Subject Code",
            "subject_name": "Subject Name",
        }
        widgets = {
            "subject_code": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Example: GE 1",
            }),
            "subject_name": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Enter subject name",
            }),
        }


class RoomForm(forms.ModelForm):
    class Meta:
        model = Room
        fields = ["room_name"]
        labels = {"room_name": "Room"}
        widgets = {
            "room_name": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Example: IT 102",
            }),
        }


class ScheduleForm(forms.ModelForm):
    class Meta:
        model = Schedule
        fields = [
            "instructor",
            "subject",
            "room",
            "day",
            "start_time",
            "end_time",
        ]
        labels = {
            "instructor": "Instructor",
            "subject": "Subject",
            "room": "Room",
            "day": "Day",
            "start_time": "Start Time",
            "end_time": "End Time",
        }
        widgets = {
            "instructor": forms.Select(attrs={
                "class": "form-select searchable-select",
                "id": "instructor-select",
            }),
            "subject": forms.Select(attrs={
                "class": "form-select searchable-select",
                "id": "subject-select",
            }),
            "room": forms.Select(attrs={
                "class": "form-select searchable-select",
                "id": "room-select",
            }),
            "day": forms.Select(attrs={
                "class": "form-select searchable-select",
                "id": "day-select",
            }),
            "start_time": forms.TimeInput(attrs={
                "class": "form-control",
                "type": "time",
            }),
            "end_time": forms.TimeInput(attrs={
                "class": "form-control",
                "type": "time",
            }),
        }


class MonitoringRecordForm(forms.ModelForm):

    REMARKS_CHOICES = [
        ("Present / Class ongoing", "✅ Present / Class ongoing"),
        ("Present but late", "⚠️ Present but late"),
        ("ABSENT", "❌ ABSENT"),
        ("Room/Class transferred", "🔄 Room/Class transferred"),
        ("Official activity / approved absence", "📋 Official activity / approved absence"),
        ("No students present", "🚫 No students present"),
        ("Other", "Other"),
    ]

    remarks = forms.ChoiceField(
        choices=REMARKS_CHOICES,
        widget=forms.Select(attrs={
            "class": "form-select",
            "id": "remarks-select",
        }),
        label="Remarks",
    )

    class Meta:
        model = MonitoringRecord

        fields = [
            "transfer_room",
            "remarks",
            "other_remarks",
        ]

        labels = {
            "transfer_room": "Actual Room Used",
            "remarks": "Remarks",
            "other_remarks": "Other Remarks",
        }

        widgets = {
            "transfer_room": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Example: IT 103",
            }),

            "other_remarks": forms.Textarea(attrs={
                "class": "form-control",
                "id": "other-remarks",
                "rows": 3,
                "placeholder": "Specify additional remarks...",
            }),
        }

    def clean(self):
        cleaned_data = super().clean()

        remarks = cleaned_data.get("remarks")

        # Only keep Other Remarks when "Other" is selected.
        if remarks != "Other":
            cleaned_data["other_remarks"] = ""

        return cleaned_data


class MissedMonitoringForm(forms.ModelForm):
    class Meta:
        model = MonitoringRecord
        fields = ["missed_reason"]
        labels = {
            "missed_reason": "Reason for Missing the Monitoring",
        }
        widgets = {
            "missed_reason": forms.Textarea(attrs={
                "class": "form-control",
                "placeholder": "Enter the reason why the class was not monitored...",
                "rows": 4,
            }),
        }


class LateMonitoringRecordForm(forms.ModelForm):
    REMARKS_CHOICES = [
        ("Late", "Late"),
    ]

    monitoring_date = forms.DateField(
        label="Monitoring Date",
        widget=forms.DateInput(attrs={
            "class": "form-control",
            "type": "date",
        }),
    )

    monitoring_time = forms.TimeField(
        label="Monitoring Time",
        widget=forms.TimeInput(attrs={
            "class": "form-control",
            "type": "time",
        }),
    )

    remarks = forms.ChoiceField(
        choices=REMARKS_CHOICES,
        widget=forms.Select(attrs={
            "class": "form-select",
        }),
        label="Remarks",
    )

    class Meta:
        model = MonitoringRecord
        fields = [
            "monitoring_date",
            "monitoring_time",
            "transfer_room",
            "remarks",
            "other_remarks",
        ]
        labels = {
            "monitoring_date": "Monitoring Date",
            "monitoring_time": "Monitoring Time",
            "transfer_room": "Actual Room Used",
            "remarks": "Remarks",
            "other_remarks": "Other Remarks",
        }
        widgets = {
            "transfer_room": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Example: IT 103",
            }),
            "other_remarks": forms.Textarea(attrs={
                "class": "form-control",
                "rows": 3,
                "placeholder": "Specify additional remarks...",
            }),
        }

    def clean(self):
        cleaned_data = super().clean()
        cleaned_data["other_remarks"] = ""
        return cleaned_data

    def save(self, commit=True):
        instance = super().save(commit=False)
        monitoring_date = self.cleaned_data["monitoring_date"]
        monitoring_time = self.cleaned_data["monitoring_time"]
        naive_datetime = datetime.combine(
            monitoring_date,
            monitoring_time,
        )
        instance.record_datetime = timezone.make_aware(naive_datetime)
        instance.remarks = "Late"
        instance.other_remarks = ""
        instance.signature = ""

        if commit:
            instance.save()

        return instance


class ReportTemplateForm(forms.ModelForm):

    class Meta:

        model = ReportTemplate

        fields = [
            "name",
            "school_year",
            "file",
            "is_active",
        ]

        widgets = {

            "name": forms.TextInput(
                attrs={
                    "placeholder":
                        "Example: Part-Time Instructor Monitoring Form"
                }
            ),

            "school_year": forms.TextInput(
                attrs={
                    "placeholder":
                        "Example: 2026-2027"
                }
            ),

            "file": forms.FileInput(
                attrs={
                    "accept": ".xlsx,.xls"
                }
            ),

        }


    def clean_file(self):

        file = self.cleaned_data.get("file")

        if file:

            filename = file.name.lower()

            if not (
                filename.endswith(".xlsx")
                or filename.endswith(".xls")
            ):

                raise forms.ValidationError(
                    "Please upload a Microsoft Excel (.xlsx or .xls) file."
                )

        return file

class FirstAdminSetupForm(forms.ModelForm):
    password = forms.CharField(
        label="Password",
        widget=forms.PasswordInput(attrs={
            "class": "form-control",
            "placeholder": "Create a password",
        }),
    )
    confirm_password = forms.CharField(
        label="Confirm Password",
        widget=forms.PasswordInput(attrs={
            "class": "form-control",
            "placeholder": "Confirm your password",
        }),
    )

    class Meta:
        model = User
        fields = ["username"]
        labels = {"username": "Username"}
        widgets = {
            "username": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Create your username",
                "autocomplete": "username",
            }),
        }

    def clean_username(self):
        username = self.cleaned_data["username"].strip()
        if User.objects.filter(username__iexact=username).exists():
            raise forms.ValidationError("This username is already taken.")
        return username

    def clean(self):
        cleaned_data = super().clean()
        password = cleaned_data.get("password")
        confirm_password = cleaned_data.get("confirm_password")
        if password and confirm_password:
            if password != confirm_password:
                raise forms.ValidationError("The passwords do not match.")
            if len(password) < 8:
                raise forms.ValidationError(
                    "Password must contain at least 8 characters."
                )
        return cleaned_data

    def save(self, commit=True):
        user = super().save(commit=False)
        user.set_password(self.cleaned_data["password"])
        user.is_active = True
        user.is_staff = True
        if commit:
            user.save()
        return user


class AdminAccountForm(forms.ModelForm):
    password = forms.CharField(
        label="Password",
        widget=forms.PasswordInput(attrs={
            "class": "form-control",
            "placeholder": "Create a password",
            "autocomplete": "new-password",
        }),
    )
    confirm_password = forms.CharField(
        label="Confirm Password",
        widget=forms.PasswordInput(attrs={
            "class": "form-control",
            "placeholder": "Confirm the password",
            "autocomplete": "new-password",
        }),
    )

    class Meta:
        model = User
        fields = ["username"]
        labels = {"username": "Username"}
        widgets = {
            "username": forms.TextInput(attrs={
                "class": "form-control",
                "placeholder": "Enter admin username",
                "autocomplete": "off",
            }),
        }

    def clean_username(self):
        username = self.cleaned_data["username"].strip()
        if User.objects.filter(username__iexact=username).exists():
            raise forms.ValidationError("This username is already taken.")
        return username

    def clean(self):
        cleaned_data = super().clean()
        password = cleaned_data.get("password")
        confirm_password = cleaned_data.get("confirm_password")
        if password and confirm_password:
            if password != confirm_password:
                raise forms.ValidationError("The passwords do not match.")
            if len(password) < 8:
                raise forms.ValidationError(
                    "Password must contain at least 8 characters."
                )
        return cleaned_data

    def save(self, commit=True):
        user = super().save(commit=False)
        user.set_password(self.cleaned_data["password"])
        user.is_active = True
        user.is_staff = False
        user.is_superuser = False
        if commit:
            user.save()
        return user
