"""Корневые маршруты проекта."""
from django.conf import settings
from django.conf.urls.static import static
from django.contrib import admin
from django.contrib.auth import views as auth_views
from django.contrib.auth.decorators import login_not_required
from django.urls import include, path

urlpatterns = [
    path("admin/", admin.site.urls),
    # Страница входа освобождена от LoginRequiredMiddleware (иначе — цикл редиректов).
    path(
        "accounts/login/",
        login_not_required(
            auth_views.LoginView.as_view(template_name="registration/login.html")
        ),
        name="login",
    ),
    path("accounts/logout/", auth_views.LogoutView.as_view(), name="logout"),
    path("", include("attestation.urls")),
]

# В DEBUG отдаём медиа-файлы (загруженные архивы и сгенерированные документы)
if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)
