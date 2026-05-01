from django.contrib import admin

from .models import DailyReport


@admin.register(DailyReport)
class DailyReportAdmin(admin.ModelAdmin):
    list_display = ("date", "user", "department", "submitted_at")
    list_filter = ("date", "user__department")
    search_fields = ("user__username", "user__first_name", "user__last_name")
    date_hierarchy = "date"
    readonly_fields = ("submitted_at", "created_at")

    @admin.display(description="Department", ordering="user__department__name")
    def department(self, obj):
        return obj.user.department.name if obj.user.department else "—"
