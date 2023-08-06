import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'treningsmonitorering.settings')
django.setup()
from trening_db.models import Student


student = Student.objects.filter(name='Test')
for stud in student:
    stud.delete()