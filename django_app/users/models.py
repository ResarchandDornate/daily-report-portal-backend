from django.contrib.auth.models import AbstractUser
from django.db import models


class User(AbstractUser):
    """Custom user — extends Django's AbstractUser with portal-specific fields."""

    class Role(models.TextChoices):
        HR = "hr", "HR"
        EMPLOYEE = "employee", "Employee"

    contact_number = models.CharField(max_length=20, blank=True, default="")
    role = models.CharField(max_length=16, choices=Role.choices, default=Role.EMPLOYEE)
    department = models.ForeignKey(
        "departments.Department",
        on_delete=models.SET_NULL,
        null=True,
        blank=True,
        related_name="employees",
    )
    title = models.CharField(max_length=64, blank=True, default="")

    # Imported from the HR roster spreadsheet
    organisation = models.CharField(max_length=64, blank=True, default="")
    reporting_manager = models.CharField(max_length=128, blank=True, default="")
    date_of_joining = models.DateField(null=True, blank=True)

    def save(self, *args, **kwargs):
        # Auto-promote: anyone in the HR department (slug="hrDept") becomes
        # HR role + superuser so they can manage the portal alongside admin.
        in_hr_dept = bool(self.department_id) and self.department and self.department.slug == "hrDept"
        if in_hr_dept:
            self.role = self.Role.HR
            self.is_superuser = True
            self.is_staff = True

        # Hard rule: role="hr" is reserved for superusers only.  Any other
        # attempt to set it (admin form, shell, REST endpoint) is silently
        # downgraded.
        if self.role == self.Role.HR and not self.is_superuser:
            self.role = self.Role.EMPLOYEE
        super().save(*args, **kwargs)

    def __str__(self):
        return self.get_full_name() or self.username
