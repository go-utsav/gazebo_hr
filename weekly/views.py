from django.contrib.auth import authenticate, login, logout
from django.http import HttpRequest, HttpResponse, JsonResponse
from django.shortcuts import redirect, render
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.http import require_GET, require_http_methods
from django.contrib import messages

from .payroll_service import (
	AGENCY_CATEGORIES,
	PayrollResult,
	build_excel_bytes,
	calculate_payroll,
	parse_employee_hours,
	total_paid_hours_from_rows,
)


@require_GET
def home(request: HttpRequest):
	if request.user.is_authenticated:
		return redirect('weekly:weekly_report')
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


@require_http_methods(['GET', 'POST'])
def weekly_report(request: HttpRequest):
	result_data = request.session.get('weekly_last_result', {})
	preview_rows = result_data.get('rows', [])[:200]
	summary = result_data.get('summary', {})

	if request.method == 'POST':
		employee_file = request.FILES.get('employee_file')
		contracted_file = request.FILES.get('contracted_file')
		if not employee_file or not contracted_file:
			messages.error(request, 'Upload both files: employee hours and contracted hours.')
			return redirect('weekly:weekly_report')
		try:
			employee_rows = parse_employee_hours(employee_file)
			payroll_result = calculate_payroll(employee_rows, contracted_file)
		except Exception as exc:
			messages.error(request, f'Could not process files: {exc}')
			return redirect('weekly:weekly_report')

		request.session['weekly_last_result'] = {
			'rows': payroll_result.rows,
			'summary': {
				'total_rows': len(payroll_result.rows),
				'total_paid_hours': payroll_result.total_paid_hours,
				'agency_rows': len(payroll_result.agency_rows),
				'gazebo_rows': len(payroll_result.gazebo_rows),
			},
		}
		request.session.modified = True
		messages.success(request, f'Processed {len(payroll_result.rows)} employee rows.')
		return redirect('weekly:weekly_report')

	return render(
		request,
		'weekly/weekly_report.html',
		{
			'title': 'Weekly report — Gazebo',
			'page_heading': 'Weekly report',
			'preview_rows': preview_rows,
			'summary': summary,
		},
	)


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


@require_GET
def download_weekly_excel(request: HttpRequest):
	result_data = request.session.get('weekly_last_result', {})
	rows = result_data.get('rows', [])
	if not rows:
		messages.error(request, 'No processed data available. Upload files first.')
		return redirect('weekly:weekly_report')

	agency_rows = [r for r in rows if str(r.get('Category', '')).strip().upper() in AGENCY_CATEGORIES]
	gazebo_rows = [r for r in rows if str(r.get('Category', '')).strip().upper() not in AGENCY_CATEGORIES]
	payroll_result = PayrollResult(
		rows=rows,
		agency_rows=agency_rows,
		gazebo_rows=gazebo_rows,
		total_paid_hours=total_paid_hours_from_rows(rows),
	)
	file_bytes = build_excel_bytes(payroll_result)
	response = HttpResponse(
		file_bytes,
		content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	)
	response['Content-Disposition'] = 'attachment; filename="weekly_report_output.xlsx"'
	return response


