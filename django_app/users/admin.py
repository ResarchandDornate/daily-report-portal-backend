from django.contrib import admin
from django.contrib.auth.admin import UserAdmin

from .models import User


@admin.register(User)
class CustomUserAdmin(UserAdmin):
    list_display = ("username", "email", "first_name", "last_name", "role", "department", "title", "organisation", "reporting_manager", "date_of_joining", "is_staff")
    list_filter = ("role", "department", "organisation", "is_staff", "is_active")
    search_fields = ("username", "email", "first_name", "last_name", "reporting_manager")

    fieldsets = UserAdmin.fieldsets + (
        ("Portal", {"fields": ("role", "department", "title", "contact_number", "organisation", "reporting_manager", "date_of_joining")}),
    )
    add_fieldsets = UserAdmin.add_fieldsets + (
        ("Portal", {"fields": ("role", "department", "title", "contact_number", "organisation", "reporting_manager", "date_of_joining")}),
    )
