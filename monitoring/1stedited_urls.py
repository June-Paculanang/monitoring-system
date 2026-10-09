from django.urls import path
from . import views
from django.contrib.auth import views as auth_views


urlpatterns = [

    # Dashboard
    path(
        "",
        views.dashboard,
        name="dashboard"
    ),

    # Instructors
    path(
        "instructors/",
        views.instructor_list,
        name="instructor_list"
    ),

    path(
        "instructors/add/",
        views.instructor_create,
        name="instructor_create"
    ),

    path(
        "instructors/bulk-delete/",
        views.instructor_bulk_delete,
        name="instructor_bulk_delete"
    ),

    path(
        "instructors/<int:pk>/edit/",
        views.instructor_update,
        name="instructor_update"
    ),

    path(
        "instructors/<int:pk>/delete/",
        views.instructor_delete,
        name="instructor_delete"
    ),

    path(
        "instructors/<int:instructor_id>/schedule/",
        views.instructor_schedule,
        name="instructor_schedule"
    ),


    # Subjects
    path(
        "subjects/",
        views.subject_list,
        name="subject_list"
    ),

    path(
        "subjects/add/",
        views.subject_create,
        name="subject_create"
    ),

    path(
        "subjects/bulk-delete/",
        views.subject_bulk_delete,
        name="subject_bulk_delete"
    ),

    path(
        "subjects/<int:pk>/edit/",
        views.subject_update,
        name="subject_update"
    ),

    path(
        "subjects/<int:pk>/delete/",
        views.subject_delete,
        name="subject_delete"
    ),


    # Rooms
    path(
        "rooms/",
        views.room_list,
        name="room_list"
    ),

    path(
        "rooms/add/",
        views.room_create,
        name="room_create"
    ),

    path(
        "rooms/<int:pk>/edit/",
        views.room_update,
        name="room_update"
    ),

    path(
        "rooms/<int:pk>/delete/",
        views.room_delete,
        name="room_delete"
    ),

    path(
        "rooms/bulk-delete/",
        views.room_bulk_delete,
        name="room_bulk_delete"
    ),


    # Schedules
    path(
        "schedules/",
        views.schedule_list,
        name="schedule_list"
    ),

    path(
        "schedules/add/",
        views.schedule_create,
        name="schedule_create"
    ),

    path(
        "schedules/import/",
        views.schedule_import,
        name="schedule_import"
    ),

    path(
        "schedules/import/part-time/",
        views.schedule_import_part_time,
        name="schedule_import_part_time"
    ),

    path(
        "schedules/import/regular/",
        views.schedule_import_regular,
        name="schedule_import_regular"
    ),

    path(
        "schedules/import/confirm/",
        views.schedule_import_confirm,
        name="schedule_import_confirm"
    ),

    path(
        "schedules/bulk-delete/",
        views.schedule_bulk_delete,
        name="schedule_bulk_delete"
    ),

    path(
        "schedules/<int:pk>/edit/",
        views.schedule_update,
        name="schedule_update"
    ),

    path(
        "schedules/<int:pk>/delete/",
        views.schedule_delete,
        name="schedule_delete"
    ),


    # Monitoring
    path(
        "monitoring/",
        views.monitoring_list,
        name="monitoring_list"
    ),

    path(
        "schedules/<int:schedule_id>/monitor/",
        views.monitoring_create,
        name="monitoring_create"
    ),

    path(
        "schedules/<int:schedule_id>/monitor-late/",
        views.monitor_late,
        name="monitor_late"
    ),

    path(
        "schedules/<int:schedule_id>/monitor-missed/",
        views.monitoring_missed,
        name="monitoring_missed"
    ),

    path(
        "monitoring/<int:pk>/edit/",
        views.monitoring_update,
        name="monitoring_update"
    ),

    path(
        "monitoring/<int:pk>/delete/",
        views.monitoring_delete,
        name="monitoring_delete"
    ),

    path(
        "monitoring/bulk-delete/",
        views.monitoring_bulk_delete,
        name="monitoring_bulk_delete"
    ),


    # Reports
    path(
        "reports/",
        views.reports,
        name="reports"
    ),

    path(
        "reports/printers/",
        views.report_printers,
        name="report_printers"
    ),

    path(
        "reports/windows-printers/",
        views.report_printers,
        name="windows_printers"
    ),

    path(
        "reports/print/",
        views.report_print,
        name="report_print"
    ),

    path(
        "reports/template/",
        views.report_template_list,
        name="report_template_list"
    ),

    path(
        "reports/template/add/",
        views.report_template_create,
        name="report_template_create"
    ),

    path(
        "reports/template/<int:pk>/edit/",
        views.report_template_update,
        name="report_template_update"
    ),

    path(
        "reports/template/<int:pk>/delete/",
        views.report_template_delete,
        name="report_template_delete"
    ),

    path(
        "reports/template/bulk-delete/",
        views.report_template_bulk_delete,
        name="report_template_bulk_delete"
    ),

    path(
        "reports/pdf/",
        views.report_pdf,
        name="report_pdf"
    ),

    path(
        "reports/docx/",
        views.report_docx,
        name="report_docx"
    ),

    path(
        "reports/excel/",
        views.report_excel,
        name="report_excel"
    ),


    # System Administrator Setup
    path(
        "setup/",
        views.system_setup,
        name="system_setup"
    ),


    # Admin Account Management
    path(
        "admin-accounts/",
        views.admin_accounts,
        name="admin_accounts"
    ),

    path(
        "admin-accounts/add/",
        views.admin_account_create,
        name="admin_account_create"
    ),

    path(
        "admin-accounts/<int:user_id>/password/",
        views.admin_account_password,
        name="admin_account_password"
    ),

    path(
        "admin-accounts/<int:user_id>/deactivate/",
        views.admin_account_deactivate,
        name="admin_account_deactivate"
    ),

    path(
        "admin-accounts/<int:user_id>/delete/",
        views.admin_account_delete,
        name="admin_account_delete"
    ),


    # Password Reset
    path(
        "password-reset/",
        auth_views.PasswordResetView.as_view(
            template_name="registration/password_reset_form.html",
            email_template_name="registration/password_reset_email.html",
            subject_template_name="registration/password_reset_subject.txt",
            success_url="/password-reset/done/",
        ),
        name="password_reset"
    ),

    path(
        "password-reset/done/",
        auth_views.PasswordResetDoneView.as_view(
            template_name="registration/password_reset_done.html"
        ),
        name="password_reset_done"
    ),

    path(
        "reset/<uidb64>/<token>/",
        auth_views.PasswordResetConfirmView.as_view(
            template_name="registration/password_reset_confirm.html",
            success_url="/reset/done/",
        ),
        name="password_reset_confirm"
    ),

    path(
        "reset/done/",
        auth_views.PasswordResetCompleteView.as_view(
            template_name="registration/password_reset_complete.html"
        ),
        name="password_reset_complete"
    ),


    # Logout
    path(
        "logout/",
        auth_views.LogoutView.as_view(),
        name="logout"
    ),
]