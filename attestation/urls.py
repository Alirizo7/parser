"""Маршруты приложения аттестации."""
from django.urls import path

from . import views

app_name = "attestation"

urlpatterns = [
    path("", views.dashboard, name="dashboard"),
    path("upload/", views.upload, name="upload"),
    path("batch/<int:pk>/", views.detail, name="detail"),
    path("batch/<int:pk>/status/", views.status, name="status"),
    path("batch/<int:pk>/cell/", views.edit_cell, name="edit_cell"),
    path("batch/<int:pk>/generate/", views.generate, name="generate"),
    path("batch/<int:pk>/download/<str:which>/", views.download, name="download"),
]
