from django.db import models


class Department(models.Model):
    """A company department. Each department has its own report column template."""

    slug = models.SlugField(unique=True, max_length=32)         # e.g. "sales", "insideSales"
    name = models.CharField(max_length=64)                      # e.g. "Sales"
    color = models.CharField(max_length=16, default="zinc")     # tailwind color name

    # JSON list of {key, label} entries. Mirrors the frontend's reportFields.
    # Example: [{"key": "meeting", "label": "Meeting"}, ...]
    report_fields = models.JSONField(default=list, blank=True)

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        ordering = ["name"]

    def __str__(self):
        return self.name
