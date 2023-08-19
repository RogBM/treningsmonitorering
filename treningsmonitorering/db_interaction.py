import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'treningsmonitorering.settings')
django.setup()
import pandas as pd
from trening_db.models import ResistanceExercise, DailyTrainingVolume, WeeklyTrainingVolum, Student, TrainingRepMax, \
    DailyTrainingVolumeWeight, WeeklyTrainingVolumeWeight, MenstruationRegistry, OvulationRegistry
from PySide6.QtWidgets import QMessageBox


def validate_athlete(name, age):
    if not Student.objects.filter(name=name):
        try:
            student = Student(name=name, age=age)
            student.save()
            mode_of_exercise = ResistanceExercise(student=student)
            mode_of_exercise.save()
        except Exception as e:
            print(e)
    if Student.objects.filter(name=name):
        print('This student/athlete already exists')


def find_rm_data(exercise=None, name=None, to_date=None, from_date=None, reps=None):  # Må gjøre rm senere ettersom den krever at jeg forandrer på modellene og legger til reps
    '''

    Creates an Excel spreadsheet with the requested data through the kwargs.

    :param name: Provide the name of the student as a string
    :param exercise: Provide the exercise name as a string
    :param to_date: Provide the ending date as a string in the format Year-Month-Day
    :param from_date: Provide the starting date as a string in the format Year-Month-Day
    :return: name of Excel file created.
    '''

    if name and exercise and to_date and from_date and reps:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
            obj_ex = obj.filter(exercise=exercise)
            obj_ex_rep = obj_ex.filter(reps=reps)
            obj_to_date = obj_ex_rep.objects.exclude(pr_date__lt=to_date)
            obj_from_date = list(obj_ex_rep.objects.exclude(pr_date__gt=from_date))
            if obj_to_date and obj_from_date:
                dates = []
                for ob in obj_to_date:
                    if ob.pr_date in obj_from_date:
                        dates.append(ob)
                for ob in dates:
                    if ob.exercise == exercise:
                        re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all {exercise} PR's for {name}.xlsx", index=False)

        return print(f'''File: "all {exercise} PR's for {name}.xlsx" was created ''')

    elif name and exercise and to_date and from_date:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
            obj_ex = obj.filter(exercise=exercise)
            obj_to_date = obj_ex.objects.exclude(pr_date__lt=to_date)
            obj_from_date = list(obj_ex.objects.exclude(pr_date__gt=from_date))
            if obj_to_date and obj_from_date:
                dates = []
                for ob in obj_to_date:
                    if ob.pr_date in obj_from_date:
                        dates.append(ob)
                for ob in dates:
                    if ob.exercise == exercise:
                        re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all {exercise} PR's for {name}.xlsx", index=False)

        return print(f'''File: "all {exercise} PR's for {name}.xlsx" was created ''')

    elif name and exercise and reps:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
            obj_ex = obj.filter(exercise=exercise)
            obj_final = obj_ex.filter(reps=reps)
            if obj_final:
                for ob in obj_final:
                    re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])

        except Exception as e:
            error = QMessageBox()
            error.setWindowTitle('Feil')
            error.setText(e)
            error.exec()
            error.setIcon(QMessageBox.Information)

        if re_variables:
            df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
            df.to_excel(f"all {exercise} {reps}RM PR's for {name}.xlsx", index=False)
        else:
            error = QMessageBox()
            error.setWindowTitle('Feil')
            error.setText('Data does not exist')
            error.setIcon(QMessageBox.Information)
            error.exec()


    elif name and exercise and from_date:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
            obj_ex = obj.filter(exercise=exercise)
            obj_date = obj_ex.objects.exclude(pr_date__gt=from_date).order_by('id')
            if obj_date:
                for ob in obj_date:
                    if ob.exercise == exercise:
                        re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all {exercise} PR's for {name}.xlsx", index=False)

        return print(f'''File: "all {exercise} PR's for {name}.xlsx" was created ''')

    elif name and exercise and to_date:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
            obj_ex = obj.filter(exercise=exercise)
            obj_date = obj_ex.objects.exclude(pr_date__lt=to_date)
            if obj_date:
                for ob in obj_date:
                    if ob.exercise == exercise:
                        re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all {exercise} PR's for {name}.xlsx", index=False)

        return print(f'''File: "all {exercise} PR's for {name}.xlsx" was created ''')

    elif name and to_date:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
            obj_date = obj.objects.exclude(pr_date__lt=to_date)
            if obj_date:
                for ob in obj_date:
                    re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all PR's for {name}.xlsx", index=False)

        return print(f'''File: "all PR's for {name}.xlsx" was created ''')

    elif name and from_date:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
            obj_date = obj.objects.exclude(pr_date__gt=from_date)
            if obj_date:
                for ob in obj_date:
                    re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all PR's for {name}.xlsx", index=False)

        return print(f'''File: "all PR's for {name}.xlsx" was created ''')

    elif name and exercise:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name).order_by('id')
            if obj:
                for ob in obj:
                    if ob.exercise == exercise:
                        re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all {exercise} PR's for {name}.xlsx", index=False)

        return print(f'''File: "all {exercise}  PR's for {name}.xlsx" was created ''')

    elif name and reps:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name).order_by('id')
            obj_rep = obj.filter(reps=reps)
            if obj_rep:
                for ob in obj_rep:
                    re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])
        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all PR's for {name}.xlsx", index=False)

        return print(f'''File: "all PR's for {name}.xlsx" was created ''')

    elif name:
        re_variables = []
        try:
            obj = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name).order_by('id')
            if obj:
                for ob in obj:
                    re_variables.append([ob.exercise, ob.rm_value, ob.reps, ob.pr_date.strftime('%Y-%m-%d')])

        except Exception as e:
            print(e)

        df = pd.DataFrame(re_variables, columns=['Exercise', 'PR (kg)', 'RM', 'PR-date (y-m-d)'])
        df.to_excel(f"all PR's for {name}.xlsx", index=False)

        return print(f'''File: "all PR's for {name}.xlsx" was created ''')


def add_objects_to_database(name=None, age=None, exercise=None, pr=None, date=None, reps=None):
    '''

    Saves the provided exercises data to the database by either creating a new instance of student
    and its children instances or saves data to provided student given that the student exists in
    the database.

    :param name: Provide the name of the student as a string
    :param age: Provide the age of the student as an integer
    :param exercise: Provide the exercise name as a string
    :param pr: Provide the exercise PR as an integer
    :param date: Provide the exercise date as a string in the format Year-Month-Day
    :return: None.
    '''

    if not Student.objects.filter(name=name):
        try:
            student = Student(name=name, age=age)
            student.save()
            mode_of_exercise = ResistanceExercise(student=student)
            mode_of_exercise.save()
            pr = TrainingRepMax(mode_of_exercise=mode_of_exercise, exercise=exercise, rm_value=pr, reps=reps,
                                pr_date=date)
            pr.save()

        except Exception as e:
            print(e)

    elif Student.objects.filter(name=name):
        try:
            lst_id = list(ResistanceExercise.objects.filter(student__name=name).values_list('id', flat=True))
            ids = lst_id[0]
            instance = ResistanceExercise.objects.get(pk=ids)
            pr = TrainingRepMax(mode_of_exercise=instance, exercise=exercise, rm_value=pr, reps=reps, pr_date=date)
            pr.save()

        except Exception as e:
            print(e)


def update_rm(name, exercise, date, rmval, reps):
    stud_id_list = list(ResistanceExercise.objects.filter(student__name=name).values_list('id', flat=True))
    stud_id = stud_id_list[0]
    instance = ResistanceExercise.objects.get(pk=stud_id)
    name_pr = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name).order_by('id')
    name_ex_pr = name_pr.filter(exercise=exercise).order_by('id')
    prs = [pr.rm_value for pr in name_ex_pr]
    pr_reps = [iden.reps for iden in name_ex_pr]
    pr_reps_val = zip(prs, pr_reps)
    reps_pr = []

    for x in pr_reps_val:
        pr_current = x[0]
        rm = x[1]

        if reps == rm:
            reps_pr.append(pr_current)

    current_pr = 0
    try:
        current_pr = max(reps_pr)

    except ValueError:
        new_pr = TrainingRepMax(mode_of_exercise=instance, exercise=exercise, rm_value=rmval, pr_date=date, reps=reps)
        new_pr.save()
        return (f"You do not have a PR at this RM. A new {reps}RM of {rmval}kg's was created")

    if rmval == current_pr:
        new_pr = TrainingRepMax(mode_of_exercise=instance, exercise=exercise, rm_value=rmval, pr_date=date, reps=reps)
        new_pr.save()
        return f"You just matched your {exercise} {reps}RM"


    elif rmval > current_pr:
        new_pr = TrainingRepMax(mode_of_exercise=instance, exercise=exercise, rm_value=rmval, pr_date=date, reps=reps)
        new_pr.save()
        return f"You just PR'd your {exercise} {reps}RM by {rmval - current_pr}kg"


    else:
        return f"You were {current_pr - rmval}kg's behind your current {reps}RM PR"


def get_exercise_distribution(name=None, exercise=None, to_date=None, from_date=None, type=None):
    stud = DailyTrainingVolume.objects.filter(mode_of_exercise__student__name=name).order_by('id')
    if stud and type == 'dist reps':

        if name and exercise == 'All exercises':
            daily_re_variables = []
            weekly_re_variables = []
            exs = set(ex.exercise for ex in DailyTrainingVolume.objects.all())
            for ex in exs:
                obj = DailyTrainingVolume.objects.filter(mode_of_exercise__student__name=name, exercise=ex).order_by(
                    'id')
                obj2 = WeeklyTrainingVolum.objects.filter(mode_of_exercise__student__name=name, exercise=ex).order_by(
                    'id')
                if obj:
                    for ob in obj:
                        daily_re_variables.append([ob.date.strftime('%Y-%m-%d'),
                                                   ob.day,
                                                   ob.exercise,
                                                   ob.total_exercise_reps_volume,
                                                   ob.intensity_zone_1,
                                                   ob.intensity_zone_2,
                                                   ob.intensity_zone_3,
                                                   ob.intensity_zone_4,
                                                   ob.intensity_zone_5,
                                                   ob.intensity_zone_6,
                                                   ob.intensity_zone_7,
                                                   ob.intensity_zone_8,
                                                   ob.intensity_zone_9,
                                                   ob.intensity_zone_10,
                                                   ob.intensity_zone_11,
                                                   ob.intensity_zone_12,
                                                   ob.intensity_zone_13])
                if obj2:
                    for ob in obj2:
                        weekly_re_variables.append([ob.wknr,
                                                    ob.exercise,
                                                    ob.total_intensity_zone,
                                                    ob.intensity_zone_1,
                                                    ob.intensity_zone_2,
                                                    ob.intensity_zone_3,
                                                    ob.intensity_zone_4,
                                                    ob.intensity_zone_5,
                                                    ob.intensity_zone_6,
                                                    ob.intensity_zone_7,
                                                    ob.intensity_zone_8,
                                                    ob.intensity_zone_9,
                                                    ob.intensity_zone_10,
                                                    ob.intensity_zone_11,
                                                    ob.intensity_zone_12,
                                                    ob.intensity_zone_13])

            df_days = pd.DataFrame(daily_re_variables, columns=['Date',
                                                                'Day',
                                                                'Exercise',
                                                                'total reps',
                                                                '0-10%',
                                                                '11-20%',
                                                                '21-30%',
                                                                '31-40%',
                                                                '41-50%',
                                                                '51-60%',
                                                                '61-70%',
                                                                '71-80%',
                                                                '81-90%',
                                                                '91-100%',
                                                                '101-110%',
                                                                '111-120%',
                                                                '121-130%'])

            df_weeks = pd.DataFrame(weekly_re_variables, columns=['Week',
                                                                  'Exercise',
                                                                  'total reps',
                                                                  '0-10%',
                                                                  '11-20%',
                                                                  '21-30%',
                                                                  '31-40%',
                                                                  '41-50%',
                                                                  '51-60%',
                                                                  '61-70%',
                                                                  '71-80%',
                                                                  '81-90%',
                                                                  '91-100%',
                                                                  '101-110%',
                                                                  '111-120%',
                                                                  '121-130%'])
            # df3 = df1.sort_values(['Date'])
            # df4 = df2.sort_values(['Week'])
            df_days.to_excel(f"Daily reps distribution for {name}.xlsx", index=False)
            df_weeks.to_excel(f"Weekly reps distribution for {name}.xlsx", index=False)

            return None


        elif name and exercise:
            daily_re_variables = []
            weekly_re_variables = []
            obj = DailyTrainingVolume.objects.filter(mode_of_exercise__student__name=name, exercise=exercise).order_by(
                'id')
            obj2 = WeeklyTrainingVolum.objects.filter(mode_of_exercise__student__name=name, exercise=exercise).order_by(
                'id')
            if obj:
                for ob in obj:
                    daily_re_variables.append([ob.date.strftime('%Y-%m-%d'),
                                               ob.day,
                                               ob.exercise,
                                               ob.total_exercise_reps_volume,
                                               ob.intensity_zone_1,
                                               ob.intensity_zone_2,
                                               ob.intensity_zone_3,
                                               ob.intensity_zone_4,
                                               ob.intensity_zone_5,
                                               ob.intensity_zone_6,
                                               ob.intensity_zone_7,
                                               ob.intensity_zone_8,
                                               ob.intensity_zone_9,
                                               ob.intensity_zone_10,
                                               ob.intensity_zone_11,
                                               ob.intensity_zone_12,
                                               ob.intensity_zone_13])
                if obj2:
                    for ob in obj2:
                        weekly_re_variables.append([ob.wknr,
                                                    ob.exercise,
                                                    ob.total_intensity_zone,
                                                    ob.intensity_zone_1,
                                                    ob.intensity_zone_2,
                                                    ob.intensity_zone_3,
                                                    ob.intensity_zone_4,
                                                    ob.intensity_zone_5,
                                                    ob.intensity_zone_6,
                                                    ob.intensity_zone_7,
                                                    ob.intensity_zone_8,
                                                    ob.intensity_zone_9,
                                                    ob.intensity_zone_10,
                                                    ob.intensity_zone_11,
                                                    ob.intensity_zone_12,
                                                    ob.intensity_zone_13])

            df1 = pd.DataFrame(daily_re_variables, columns=['Date',
                                                            'Day',
                                                            'Exercise',
                                                            'total reps',
                                                            '0-10%',
                                                            '11-20%',
                                                            '21-30%',
                                                            '31-40%',
                                                            '41-50%',
                                                            '51-60%',
                                                            '61-70%',
                                                            '71-80%',
                                                            '81-90%',
                                                            '91-100%',
                                                            '101-110%',
                                                            '111-120%',
                                                            '121-130%'])

            df2 = pd.DataFrame(weekly_re_variables, columns=['Week',
                                                             'Exercise',
                                                             'total reps',
                                                             '0-10%',
                                                             '11-20%',
                                                             '21-30%',
                                                             '31-40%',
                                                             '41-50%',
                                                             '51-60%',
                                                             '61-70%',
                                                             '71-80%',
                                                             '81-90%',
                                                             '91-100%',
                                                             '101-110%',
                                                             '111-120%',
                                                             '121-130%'])
            # df = df1.sort_values(['Date'])
            # df3 = df2.sort_values(['Week'])

            df1.to_excel(f"Daily reps distribution for {name}.xlsx", index=False)
            df2.to_excel(f"Weekly reps distribution for {name}.xlsx", index=False)

            return None

        elif name:
            re_variables = []
            obj = list(DailyTrainingVolume.objects.filter(mode_of_exercise__student__name=name).order_by('id'))

            if obj:
                for ob in obj:
                    re_variables.append([ob.date.strftime('%Y-%m-%d'),
                                         ob.day,
                                         ob.exercise,
                                         ob.total_exercise_reps_volume,
                                         ob.intensity_zone_1,
                                         ob.intensity_zone_2,
                                         ob.intensity_zone_3,
                                         ob.intensity_zone_4,
                                         ob.intensity_zone_5,
                                         ob.intensity_zone_6,
                                         ob.intensity_zone_7,
                                         ob.intensity_zone_8,
                                         ob.intensity_zone_9,
                                         ob.intensity_zone_10,
                                         ob.intensity_zone_11,
                                         ob.intensity_zone_12,
                                         ob.intensity_zone_13])
            df = pd.DataFrame(re_variables, columns=['Date',
                                                     'Day',
                                                     'Exercise',
                                                     'total reps',
                                                     '0-10%',
                                                     '11-20%',
                                                     '21-30%',
                                                     '31-40%',
                                                     '41-50%',
                                                     '51-60%',
                                                     '61-70%',
                                                     '71-80%',
                                                     '81-90%',
                                                     '91-100%',
                                                     '101-110%',
                                                     '111-120%',
                                                     '121-130%'])
            df_srt_date = df.sort_values(['Date'])
            df_srt_date.to_excel(f"Daily reps distribution for {name}.xlsx", index=False)

            return None

    elif stud and type == 'dist weight':

        if name and exercise == 'All exercises':
            daily_re_variables = []
            weekly_re_variables = []
            exs = set(ex.exercise for ex in DailyTrainingVolumeWeight.objects.all())
            for ex in exs:
                obj = DailyTrainingVolumeWeight.objects.filter(mode_of_exercise__student__name=name,
                                                               exercise=ex).order_by('id')
                obj2 = WeeklyTrainingVolumeWeight.objects.filter(mode_of_exercise__student__name=name,
                                                                 exercise=ex).order_by('id')
                if obj:
                    for ob in obj:
                        daily_re_variables.append([ob.date.strftime('%Y-%m-%d'),
                                                   ob.day,
                                                   ob.exercise,
                                                   ob.total_exercise_weight_volume,
                                                   ob.intensity_zone_1,
                                                   ob.intensity_zone_2,
                                                   ob.intensity_zone_3,
                                                   ob.intensity_zone_4,
                                                   ob.intensity_zone_5,
                                                   ob.intensity_zone_6,
                                                   ob.intensity_zone_7,
                                                   ob.intensity_zone_8,
                                                   ob.intensity_zone_9,
                                                   ob.intensity_zone_10,
                                                   ob.intensity_zone_11,
                                                   ob.intensity_zone_12,
                                                   ob.intensity_zone_13])
                if obj2:
                    for ob in obj2:
                        weekly_re_variables.append([ob.wknr,
                                                    ob.exercise,
                                                    ob.total_exercise_weight_volume,
                                                    ob.intensity_zone_1,
                                                    ob.intensity_zone_2,
                                                    ob.intensity_zone_3,
                                                    ob.intensity_zone_4,
                                                    ob.intensity_zone_5,
                                                    ob.intensity_zone_6,
                                                    ob.intensity_zone_7,
                                                    ob.intensity_zone_8,
                                                    ob.intensity_zone_9,
                                                    ob.intensity_zone_10,
                                                    ob.intensity_zone_11,
                                                    ob.intensity_zone_12,
                                                    ob.intensity_zone_13])

            df_days = pd.DataFrame(daily_re_variables, columns=['Date',
                                                                'Day',
                                                                'Exercise',
                                                                'total reps',
                                                                '0-10%',
                                                                '11-20%',
                                                                '21-30%',
                                                                '31-40%',
                                                                '41-50%',
                                                                '51-60%',
                                                                '61-70%',
                                                                '71-80%',
                                                                '81-90%',
                                                                '91-100%',
                                                                '101-110%',
                                                                '111-120%',
                                                                '121-130%'])

            df_weeks = pd.DataFrame(weekly_re_variables, columns=['Week',
                                                                  'Exercise',
                                                                  'total reps',
                                                                  '0-10%',
                                                                  '11-20%',
                                                                  '21-30%',
                                                                  '31-40%',
                                                                  '41-50%',
                                                                  '51-60%',
                                                                  '61-70%',
                                                                  '71-80%',
                                                                  '81-90%',
                                                                  '91-100%',
                                                                  '101-110%',
                                                                  '111-120%',
                                                                  '121-130%'])
            # df3 = df1.sort_values(['Date'])
            # df4 = df2.sort_values(['Week'])
            df_days.to_excel(f"Daily weight distribution for {name}.xlsx", index=False)
            df_weeks.to_excel(f"Weekly weight distribution for {name}.xlsx", index=False)

            return None


        elif name and exercise:
            daily_re_variables = []
            weekly_re_variables = []
            obj = DailyTrainingVolumeWeight.objects.filter(mode_of_exercise__student__name=name,
                                                           exercise=exercise).order_by('id')
            obj2 = WeeklyTrainingVolumeWeight.objects.filter(mode_of_exercise__student__name=name,
                                                             exercise=exercise).order_by('id')
            if obj:
                for ob in obj:
                    daily_re_variables.append([ob.date.strftime('%Y-%m-%d'),
                                               ob.day,
                                               ob.exercise,
                                               ob.total_exercise_weight_volume,
                                               ob.intensity_zone_1,
                                               ob.intensity_zone_2,
                                               ob.intensity_zone_3,
                                               ob.intensity_zone_4,
                                               ob.intensity_zone_5,
                                               ob.intensity_zone_6,
                                               ob.intensity_zone_7,
                                               ob.intensity_zone_8,
                                               ob.intensity_zone_9,
                                               ob.intensity_zone_10,
                                               ob.intensity_zone_11,
                                               ob.intensity_zone_12,
                                               ob.intensity_zone_13])
                if obj2:
                    for ob in obj2:
                        weekly_re_variables.append([ob.wknr,
                                                    ob.exercise,
                                                    ob.total_exercise_weight_volume,
                                                    ob.intensity_zone_1,
                                                    ob.intensity_zone_2,
                                                    ob.intensity_zone_3,
                                                    ob.intensity_zone_4,
                                                    ob.intensity_zone_5,
                                                    ob.intensity_zone_6,
                                                    ob.intensity_zone_7,
                                                    ob.intensity_zone_8,
                                                    ob.intensity_zone_9,
                                                    ob.intensity_zone_10,
                                                    ob.intensity_zone_11,
                                                    ob.intensity_zone_12,
                                                    ob.intensity_zone_13])

            df1 = pd.DataFrame(daily_re_variables, columns=['Date',
                                                            'Day',
                                                            'Exercise',
                                                            'total weight',
                                                            '0-10%',
                                                            '11-20%',
                                                            '21-30%',
                                                            '31-40%',
                                                            '41-50%',
                                                            '51-60%',
                                                            '61-70%',
                                                            '71-80%',
                                                            '81-90%',
                                                            '91-100%',
                                                            '101-110%',
                                                            '111-120%',
                                                            '121-130%'])

            df2 = pd.DataFrame(weekly_re_variables, columns=['Week',
                                                             'Exercise',
                                                             'total weight',
                                                             '0-10%',
                                                             '11-20%',
                                                             '21-30%',
                                                             '31-40%',
                                                             '41-50%',
                                                             '51-60%',
                                                             '61-70%',
                                                             '71-80%',
                                                             '81-90%',
                                                             '91-100%',
                                                             '101-110%',
                                                             '111-120%',
                                                             '121-130%'])
            # df = df1.sort_values(['Date'])
            # df3 = df2.sort_values(['Week'])

            df1.to_excel(f"Daily weight distribution for {name}.xlsx", index=False)
            df2.to_excel(f"Weekly weight distribution for {name}.xlsx", index=False)

            return None


    else:
        print(f'{name} has no training data associated with him/her')

    return None


def retrieve_ovulation_and_menstruation_distribution(name=None):

    ovulation_information = []
    menstruation_information = []
    ovul_object = OvulationRegistry.objects.filter(ovulation__student__name=name).order_by('id')
    m_object = MenstruationRegistry.objects.filter(menstruation__student__name=name).order_by('id')

    for obj in ovul_object:
        ovulation_information.append([obj.date.strftime('%Y-%m-%d'),
                                     obj.day,
                                     obj.ovulation_situation])
    for obj in m_object:
        menstruation_information.append([obj.date.strftime('%Y-%m-%d'),
                                        obj.day,
                                        obj.menstruation_situation,
                                        obj.numerical_menstruation_situation])

    ovulation_ouput = pd.DataFrame(ovulation_information, columns=["Date",
                                                                   "Day",
                                                                   "Ovulation status"])

    menstruation_output = pd.DataFrame(menstruation_information, columns=["Date",
                                                                          "Day",
                                                                          "Menstruation status",
                                                                          "Numerical Menstruation Status"])

    ovulation_ouput.to_excel(f"Ovulation information for {name}.xlsx", index=False)
    menstruation_output.to_excel(f"Menstruation information for {name}.xlsx", index=False)


def Check_weekly_training_volume():
    e = WeeklyTrainingVolum.objects.all()
    for f in e:
        print(f.wknr, f.exercise, f.total_intensity_zone, f.intensity_zone_1,
              f.intensity_zone_2,
              f.intensity_zone_3,
              f.intensity_zone_4,
              f.intensity_zone_5,
              f.intensity_zone_6,
              f.intensity_zone_7,
              f.intensity_zone_8,
              f.intensity_zone_9,
              f.intensity_zone_10,
              f.intensity_zone_11,
              f.intensity_zone_12,
              f.intensity_zone_13,
              f.week)
    print(len(e))


def check_daily_training_volume():
    e = DailyTrainingVolume.objects.all()
    for f in e:
        print(f.date, f.day, f.exercise, f.total_exercise_volume, f.intensity_zone_1,
              f.intensity_zone_2,
              f.intensity_zone_3,
              f.intensity_zone_4,
              f.intensity_zone_5,
              f.intensity_zone_6,
              f.intensity_zone_7,
              f.intensity_zone_8,
              f.intensity_zone_9,
              f.intensity_zone_10,
              f.intensity_zone_11,
              f.intensity_zone_12,
              f.intensity_zone_13)
    print(len(e))


def check_training_rep_max(name):
    e = TrainingRepMax.objects.filter(mode_of_exercise__student__name=name)
    for f in e:
        print(f.exercise, f.rm_value, f.pr_date, f.reps)

