from django.contrib import admin

from .models import Batch, SourceFile


class SourceFileInline(admin.TabularInline):
    model = SourceFile
    extra = 0
    fields = ("path", "kind", "converted_docx_path", "parsed")
    readonly_fields = fields


@admin.register(Batch)
class BatchAdmin(admin.ModelAdmin):
    list_display = ("id", "original_filename", "status", "workplaces_count", "created_at")
    list_filter = ("status",)
    search_fields = ("original_filename",)
    readonly_fields = ("created_at", "updated_at")
    inlines = [SourceFileInline]


@admin.register(SourceFile)
class SourceFileAdmin(admin.ModelAdmin):
    list_display = ("path", "kind", "batch", "parsed")
    list_filter = ("kind", "parsed")
    search_fields = ("path",)
