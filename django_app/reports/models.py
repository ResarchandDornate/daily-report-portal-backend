from django.conf import settings
from django.db import models


class DailyReport(models.Model):
    """One report per user per day. The actual report fields live in `data`
    as JSON so each department can have its own column shape (Sales fields
    differ from Procurement fields, etc.).
    """

    user = models.ForeignKey(
        settings.AUTH_USER_MODEL,
        on_delete=models.CASCADE,
        related_name="daily_reports",
    )
    date = models.DateField(db_index=True)

    # Whatever fields the user's department template specifies, keyed by `key`.
    # Example for Sales: {"meeting": "...", "revenue": "...", "newCustomerOnboard": "..."}
    data = models.JSONField(default=dict, blank=True)

    submitted_at = models.DateTimeField(auto_now=True)
    created_at = models.DateTimeField(auto_now_add=True)

    class Meta:
        unique_together = [("user", "date")]
        ordering = ["-date", "user_id"]
        indexes = [models.Index(fields=["date", "user"])]

    def __str__(self):
        return f"{self.user} — {self.date}"
