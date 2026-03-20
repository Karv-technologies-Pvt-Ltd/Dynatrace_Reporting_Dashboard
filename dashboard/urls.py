from django.contrib import admin
from django.urls import path
from dashboard import views


urlpatterns = [
    path('', views.index, name='index'),
    path('login', views.login_view, name="login"),
    path('signup', views.signup_view, name="signup"),
    path('logout', views.logout_view, name="logout"),
    path('daily-activity', views.daily_activity, name='daily_activity'),
    path('problem-data', views.problem_data, name='problem_data'),
    path('user-management', views.user_management, name='user_management'),
    path('sbom', views.sbom, name='sbom'),
    path('email_scheduler', views.email_scheduler, name='email_scheduler'),
    path('edit_schedule/<int:pk>/', views.edit_schedule, name='edit_schedule'),
    path('delete_schedule/<int:pk>/', views.delete_schedule, name='delete_schedule'),
    path("forgot-password", views.forgot_password, name="forgot_password"),
    path("verify-otp", views.verify_otp, name="verify_otp"),
    path("reset-password", views.reset_password, name="reset_password"),
    path("host-metrics/", views.HostLevelMetrics, name="host-metrics"),
    path("service-metrics/", views.ServiceLevelMetrics, name="service-metrics"),
    path("process-metrics/", views.ProcessLevelMetrics, name="process-metrics"),
    path("predictive-ui/", views.predictive_ui, name="predictive-ui"),
    path("generative-ui", views.generative_ui, name="generative-ui"),
    path("ai-query/", views.AIQueryRouter, name="ai-query"),
    path('ask-ai', views.ask_ai, name="ask-ai"),
    path("capacity-management/", views.capacity_management, name="capacity_management"),
    path("capacity-base/", views.capacity_base, name="capacity-base"),
    path("ai-search/", views.ai_search, name="ai_search"),



]

