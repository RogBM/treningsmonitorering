import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'treningsmonitorering.settings')
django.setup()
from trening_db.models import Student, ResistanceExercise, OvulationRegistry, MenstruationRegistry

intervals = {
            "0-10%": (0.0, 0.10),
            "11-20%": (0.11, 0.20),
            "21-30%": (0.21, 0.30),
            "31-40%": (0.31, 0.4),
            "41-50%": (0.41, 0.5),
            "51-60%": (0.51, 0.6),
            "61-70%": (0.61, 0.7),
            "71-80%": (0.71, 0.8),
            "81-90%": (0.81, 0.9),
            "91-100%": (0.91, 1.0),
            "101-110%": (1.01, 1.1),
            "111-120%": (1.11, 1.2),
            "121-130%": (1.21, 1.3)
        }
lower_range = None
ranges = []
oneRM = 82

for interval in intervals.items():
    upper_interval = interval[1][1] * oneRM
    lower_interval = interval[1][0] * oneRM

    if lower_range == None:
        upper_range = int(upper_interval * 100)
        lower_range = int(lower_interval * 100) + 1
        rangeee = range(lower_range, upper_range)
        lower_range = upper_range + 1

    elif lower_range:
        upper_range = int(upper_interval * 100)
        rangeee = range(lower_range, upper_range)
        lower_range = upper_range + 1

    print(rangeee)


example = 55
example2 = 0.0
print(example * example2)

student = Student.objects.all()
for stud in student:
    stud.delete()


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

