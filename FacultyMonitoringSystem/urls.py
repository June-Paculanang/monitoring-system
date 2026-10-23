from django.contrib import admin
from django.urls import path, include
from django.contrib.auth import views as auth_views


urlpatterns = [
    # Django administration
    path("admin/", admin.site.urls),

    # Login
    path(
        "login/",
        auth_views.LoginView.as_view(
            template_name="registration/login.html"
        ),
        name="login",
    ),

    # All application routes, including logout and password reset
    path("", include("monitoring.urls")),
]