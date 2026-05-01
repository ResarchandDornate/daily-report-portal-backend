from django.contrib import admin
from django.http import JsonResponse
from django.urls import path


def root(_request):
    return JsonResponse({
        "service": "Daily Report Portal — Django admin",
        "admin": "/admin/",
        "api": "Frontend should call FastAPI on http://localhost:8001",
    })


urlpatterns = [
    path("", root),
    path("admin/", admin.site.urls),
]
