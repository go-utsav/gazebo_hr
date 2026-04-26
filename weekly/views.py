from django.contrib.auth import authenticate, login, logout
from django.contrib.auth.decorators import login_required
from django.http import HttpRequest, JsonResponse
from django.shortcuts import redirect, render
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.http import require_GET, require_http_methods


@require_GET
def home(request: HttpRequest):
	if request.user.is_authenticated:
		return redirect('weekly:dashboard')
	return render(
		request,
		'weekly/home.html',
		{
			'title': 'Gazebo',
			'page_heading': 'Payroll and reporting',
		},
	)


@require_http_methods(['GET', 'POST'])
def login_view(request: HttpRequest):
	if request.user.is_authenticated:
		return redirect('weekly:dashboard')
	error = ''
	if request.method == 'POST':
		username = request.POST.get('username', '').strip()
		password = request.POST.get('password', '')
		user = authenticate(request, username=username, password=password)
		if user is not None:
			login(request, user)
			next_url = (request.POST.get('next') or request.GET.get('next') or '').strip()
			if next_url and url_has_allowed_host_and_scheme(
				next_url,
				allowed_hosts={request.get_host()},
				require_https=request.is_secure(),
			):
				return redirect(next_url)
			return redirect('weekly:dashboard')
		error = 'Enter a valid username and password.'
	return render(
		request,
		'weekly/login.html',
		{
			'title': 'Sign in — Gazebo',
			'page_heading': 'Sign in',
			'error': error,
		},
	)


@require_http_methods(['GET', 'POST'])
def logout_view(request: HttpRequest):
	logout(request)
	return redirect('weekly:home')


@login_required
@require_GET
def dashboard(request: HttpRequest):
	return render(
		request,
		'weekly/dashboard.html',
		{
			'title': 'Dashboard — Gazebo',
			'page_heading': 'Dashboard',
		},
	)


@login_required
@require_GET
def weekly_report(request: HttpRequest):
	return render(
		request,
		'weekly/weekly_report.html',
		{
			'title': 'Weekly report — Gazebo',
			'page_heading': 'Weekly report',
		},
	)


@login_required
@require_GET
def monthly_report(request: HttpRequest):
	return render(
		request,
		'weekly/monthly_report.html',
		{
			'title': 'Monthly report — Gazebo',
			'page_heading': 'Monthly report',
		},
	)


@login_required
@require_GET
def employee_hour_contracts(request: HttpRequest):
	return render(
		request,
		'weekly/employee_hour_contracts.html',
		{
			'title': 'Employee hour contracts — Gazebo',
			'page_heading': 'Employee hour contracts',
		},
	)


@require_GET
def health_api(request: HttpRequest):
	return JsonResponse({'ok': True, 'app': 'weekly'})
