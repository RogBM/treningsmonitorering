import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'treningsmonitorering.settings')
django.setup()
from trening_db.models import Student, ResistanceExercise, OvulationRegistry, MenstruationRegistry


student = Student.objects.filter(name='Test')
for stud in student:
    pass


student2 = OvulationRegistry.objects.filter(ovulation__student__name="Test")
student3 = MenstruationRegistry.objects.filter(menstruation__student__name="Test")
for stud in student2:
    print(stud.date,
          stud.day,
          stud.ovulation_situation)

for stud in student3:
    print(stud.date,
          stud.day,
          stud.menstruation_situation,
          stud.numerical_menstruation_situation)

