# scheduler.py — lightweight runner for ScheduledReport (SQLite-friendly)
# -----------------------------------------------------------------------------
# - Uses APScheduler with an in-memory jobstore (avoids SQLite locking)
# - Polls ScheduledReport every minute and runs due jobs
# - Supports: Problem Analysis, SBOM, User Management
# - Supports MULTIPLE EMAIL RECIPIENTS (comma separated)
# -----------------------------------------------------------------------------

import logging
import time
from datetime import timedelta

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.jobstores.memory import MemoryJobStore
from apscheduler.executors.pool import ThreadPoolExecutor, ProcessPoolExecutor

from django.utils import timezone
from django.db.utils import OperationalError as DjangoOperationalError

from .models import ScheduledReport
from .views import (
    generate_problem_analysis_report,
    generate_sbom_report,
    generate_user_management_report,
    generate_capacity_management_report, 
)

logger = logging.getLogger(__name__)

_scheduler = None  # Prevent double creation


# ============================================================================
#  SAVE WITH RETRY (SQLite-safe)
# ============================================================================
def _save_with_retry(instance, max_retries: int = 5, delay_sec: float = 0.25):
    for attempt in range(1, max_retries + 1):
        try:
            instance.save()
            return
        except DjangoOperationalError as e:
            if "database is locked" in str(e).lower() and attempt < max_retries:
                time.sleep(delay_sec)
                continue
            raise


# ============================================================================
#  BUMP NEXT RUN DATE
# ============================================================================
def _bump_next_run(s: ScheduledReport, now_aware):
    r = (s.recurrence or "").lower()

    if r == "daily":
        s.next_run = now_aware + timedelta(days=1)
    elif r == "weekly":
        s.next_run = now_aware + timedelta(weeks=1)
    elif r == "monthly":
        s.next_run = now_aware + timedelta(days=30)
    else:
        s.next_run = now_aware + timedelta(days=1)


# ============================================================================
#  RUN DUE REPORTS
# ============================================================================
def run_due_reports():
    tz = timezone.get_fixed_timezone(330)  # Asia/Kolkata (+05:30)
    now_aware = timezone.now().astimezone(tz)

    logger.info("🕒 Checking for scheduled reports...")

    due = ScheduledReport.objects.filter(
        next_run__lte=timezone.now()
    ).order_by("next_run")

    if not due.exists():
        logger.info("✅ No pending reports at this time.")
        return

    # ----------------------------------------------------
    # PROCESS EACH SCHEDULED REPORT
    # ----------------------------------------------------
    for job in due:
        try:
            # MULTIPLE EMAILS SUPPORT
            raw_emails = (job.recipient_email or "").strip()
            recipients = [e.strip() for e in raw_emails.split(",") if e.strip()]

            if not recipients:
                logger.info(f"⚠ No valid recipients found for job #{job.id}")
                _bump_next_run(job, now_aware)
                _save_with_retry(job)
                continue

            logger.info(f"▶ Running: {job.report_type} for {recipients}")

            rtype = (job.report_type or "").strip().lower()
            fmt = (job.report_format or "").strip().lower()

            # --------------------------------------------------------
            # SEND REPORT TO EACH RECIPIENT
            # --------------------------------------------------------
            for email in recipients:
                try:
                    if rtype in ("problem", "problem analysis"):
                        generate_problem_analysis_report(
                            tenant_url=job.tenant_url,
                            access_token=job.access_token,
                            management_zone=job.management_zone,
                            timeframe=job.timeframe,
                            report_format=fmt,
                            email=email,
                        )

                    elif rtype in ("sbom", "software", "software inventory"):
                        generate_sbom_report(
                            tenant_url=job.tenant_url,
                            access_token=job.access_token,
                            management_zone=job.management_zone,
                            report_format=fmt,
                            email=email,
                        )

                    elif rtype in ("user", "user management"):
                        generate_user_management_report(
                            account_uuid=job.account_uuid,
                            client_id=job.client_id,
                            client_secret=job.client_secret,
                            timeframe=job.timeframe,
                            report_format=fmt,
                            email=email,
                        )

                    # --------------------------------------------------------
                    # 4. Capacity Management (Daily Activity)  ← NEW
                    # --------------------------------------------------------
                    elif rtype in ("capacity", "capacity management", "daily activity"):
                        generate_capacity_management_report(
                            tenant_url=job.tenant_url,
                            access_token=job.access_token,
                            management_zone=job.management_zone,
                            timeframe=job.timeframe,
                            report_format=fmt,
                            email=job.recipient_email,
                )

                    else:
                        logger.info(f"⚠ Skipping unsupported report type: {job.report_type}")
                        continue

                except Exception as inner_e:
                    logger.error(f"❌ Error sending report to {email}: {inner_e}")

            # --------------------------------------------------------
            # SUCCESS: update status + reschedule
            # --------------------------------------------------------
            job.status = "Completed"
            _bump_next_run(job, now_aware)
            _save_with_retry(job)

            logger.info(f"✅ Completed & Rescheduled → {job.next_run.astimezone(tz)}")

        except Exception as e:
            job.status = f"Failed: {e}"
            try:
                _save_with_retry(job)
            except Exception:
                logger.exception("Failed to save job status after exception.")

            logger.error(f"❌ Scheduler error for job #{job.id}: {e}")


# ============================================================================
#  START SCHEDULER
# ============================================================================
def start_scheduler():
    global _scheduler
    if _scheduler and _scheduler.running:
        return _scheduler

    logger.info("⚙️ Initializing APScheduler (Memory jobstore)...")

    _scheduler = BackgroundScheduler(
        jobstores={"default": MemoryJobStore()},
        executors={
            "default": ThreadPoolExecutor(10),
            "processpool": ProcessPoolExecutor(2),
        },
        job_defaults={"coalesce": False, "max_instances": 1},
        timezone="Asia/Kolkata",
    )

    _scheduler.add_job(
        run_due_reports,
        trigger="interval",
        minutes=1,
        id="run_due_reports",
        replace_existing=True,
        misfire_grace_time=60,
        jitter=5,
    )

    _scheduler.start()
    logger.info("🟢 APScheduler started — polling every 1 minute.")
    return _scheduler
