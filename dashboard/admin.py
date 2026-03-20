from django.contrib import admin
from .models import Host
from .models import Service
from .models import Process
# Register your models here.
@admin.register(Host)
class HostAdmin(admin.ModelAdmin):
    list_display = (
        "host_name",
        "host_availability",
        "cpu_usage",
        "memory_usage",
        "record_date",
    )

    list_filter = ("record_date",)
    search_fields = ("host_name",)

@admin.register(Service)
class ServiceAdmin(admin.ModelAdmin):
    list_display = (
        "service_name",
        "request_count",
        "response_time",
        "failure_rate",
        "record_date",
    )

    list_filter = ("record_date",)
    search_fields = ("service_name",)
    ordering = ("-record_date",)

@admin.register(Process)
class ProcessAdmin(admin.ModelAdmin):
    list_display = (
        "process_name",
        "availability",
        "cpu_usage",
        "memory_usage",
        "record_date",
    )

    list_filter = ("record_date",)
    search_fields = ("process_name",)
    ordering = ("-record_date",)