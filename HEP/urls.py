from django.contrib import admin
from django.urls import path
from reports.views import upload_excel

urlpatterns = [
    path("admin/", admin.site.urls),
    path("upload/", upload_excel, name="upload_excel"),
    path("", upload_excel, name="upload_excel"),
]
