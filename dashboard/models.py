from django.db import models
from django.contrib.auth.models import User
from django.db import models

class ScheduledReport(models.Model):
    REPORT_CHOICES = [
        ("capacity", "Capacity Management"),
        ("problem", "Problem Analysis"),
        ("user", "User Management"),
        ("software", "Software Inventory"),
    ]

    report_type = models.CharField(max_length=50, choices=REPORT_CHOICES)

    # Dynatrace (for problem/software)
    tenant_url = models.CharField(max_length=255, blank=True, null=True)
    access_token = models.CharField(max_length=255, blank=True, null=True)

    # User Management credentials (NEW)
    account_uuid = models.CharField(max_length=255, blank=True, null=True)
    client_id = models.CharField(max_length=255, blank=True, null=True)
    client_secret = models.CharField(max_length=255, blank=True, null=True)

    management_zone = models.CharField(max_length=255, blank=True, null=True)
    timeframe = models.CharField(max_length=50)
    recipient_email = models.TextField()   # instead of EmailField
    report_format = models.CharField(max_length=10)
    recurrence = models.CharField(max_length=10)
    next_run = models.DateTimeField()

    status = models.CharField(max_length=20, default="Scheduled")
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.report_type} → {self.recipient_email} ({self.recurrence})"
    


class Host(models.Model):
    host_name = models.CharField(max_length=255)
    host_availability = models.FloatField()
    cpu_usage = models.FloatField()
    memory_usage = models.FloatField()
    record_date = models.DateField()

    class Meta:
        unique_together = ('host_name', 'record_date')  # Host



class Service(models.Model):
    service_name = models.CharField(max_length=255)
    request_count = models.FloatField()
    response_time = models.FloatField()
    failure_rate = models.FloatField()
    record_date = models.DateField()

    class Meta:
        unique_together = ('service_name', 'record_date')

    
class Process(models.Model):
    process_name = models.CharField(max_length=255)
    availability = models.FloatField()
    cpu_usage = models.FloatField()
    memory_usage = models.FloatField()
    record_date = models.DateField()

    class Meta:
        unique_together = ('process_name', 'record_date')  # Process

