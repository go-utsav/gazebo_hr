# Gazebo — Django (Weekly payroll)

Location: `C:\Users\utsav\projects\gazebo\python\`

- **Project package:** `config` (settings, root URLs)
- **Single app:** `weekly` — page templates + JSON routes live in the same app

## Run (activate venv first)

```text
(venv) cd C:\Users\utsav\projects\gazebo\python
python manage.py migrate
python manage.py runserver
```

- **Home:** http://127.0.0.1:8000/
- **API:** http://127.0.0.1:8000/api/health
- **Admin:** http://127.0.0.1:8000/admin/ (create superuser first)

## Next steps

Move parsers and Excel export from the .NET app into `weekly` (services + views + forms + templates) using the migration plan in the ClockRite repo: `gazebo_clockrite_holiday_report_dot_net\docs\django_migration_business_and_audit_plan.md`.
