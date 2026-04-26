from django.urls import path

from . import views

app_name = 'weekly'

urlpatterns = [
	path('', views.home, name='home'),
	path('login/', views.login_view, name='login'),
	path('logout/', views.logout_view, name='logout'),
	path('dashboard/', views.dashboard, name='dashboard'),
	path('dashboard/weekly-report/', views.weekly_report, name='weekly_report'),
	path('dashboard/monthly-report/', views.monthly_report, name='monthly_report'),
	path(
		'dashboard/employee-hour-contracts/',
		views.employee_hour_contracts,
		name='employee_hour_contracts',
	),
	path('dashboard/weekly-report/download/', views.download_weekly_excel, name='weekly_download'),
	path('api/health', views.health_api, name='health'),
]
