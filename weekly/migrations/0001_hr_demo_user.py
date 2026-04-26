from django.db import migrations


def create_hr_user(apps, schema_editor):
    User = apps.get_model('auth', 'User')
    if User.objects.filter(username='hr').exists():
        return
    User.objects.create_user(username='hr', password='ENTRYCL')


def noop_reverse(apps, schema_editor):
    pass


class Migration(migrations.Migration):
    initial = True

    dependencies = [
        ('auth', '0012_alter_user_first_name_max_length'),
    ]

    operations = [
        migrations.RunPython(create_hr_user, noop_reverse),
    ]
