from django.db import models


class Student(models.Model):
    name = models.CharField(max_length=100)
    # injury_status = models.BooleanField()
    age = models.IntegerField()
    picture = models.ImageField()


class ResistanceExercise(models.Model):
    student = models.ForeignKey(Student, on_delete=models.CASCADE)


class TrainingRepMax(models.Model):
    mode_of_exercise = models.ForeignKey(ResistanceExercise, on_delete=models.CASCADE)
    exercise = models.CharField(max_length=100)
    rm_value = models.IntegerField()
    pr_date = models.DateField()
    reps = models.IntegerField(null=True)


class DailyTrainingVolumeWeight(models.Model):
    mode_of_exercise = models.ForeignKey(ResistanceExercise, on_delete=models.CASCADE)
    date = models.DateField(null=True)
    day = models.CharField(max_length=100)
    exercise = models.CharField(max_length=100)
    total_exercise_weight_volume = models.IntegerField()
    intensity_zone_1 = models.IntegerField()
    intensity_zone_2 = models.IntegerField()
    intensity_zone_3 = models.IntegerField()
    intensity_zone_4 = models.IntegerField()
    intensity_zone_5 = models.IntegerField()
    intensity_zone_6 = models.IntegerField()
    intensity_zone_7 = models.IntegerField()
    intensity_zone_8 = models.IntegerField()
    intensity_zone_9 = models.IntegerField()
    intensity_zone_10 = models.IntegerField()
    intensity_zone_11 = models.IntegerField()
    intensity_zone_12 = models.IntegerField()
    intensity_zone_13 = models.IntegerField()


class WeeklyTrainingVolumeWeight(models.Model):
    mode_of_exercise = models.ForeignKey(ResistanceExercise, on_delete=models.CASCADE)
    wknr = models.IntegerField(null=True)
    exercise = models.CharField(max_length=100)
    total_exercise_weight_volume = models.IntegerField()
    intensity_zone_1 = models.IntegerField()
    intensity_zone_2 = models.IntegerField()
    intensity_zone_3 = models.IntegerField()
    intensity_zone_4 = models.IntegerField()
    intensity_zone_5 = models.IntegerField()
    intensity_zone_6 = models.IntegerField()
    intensity_zone_7 = models.IntegerField()
    intensity_zone_8 = models.IntegerField()
    intensity_zone_9 = models.IntegerField()
    intensity_zone_10 = models.IntegerField()
    intensity_zone_11 = models.IntegerField()
    intensity_zone_12 = models.IntegerField()
    intensity_zone_13 = models.IntegerField()
    week = models.IntegerField()

class DailyTrainingVolume(models.Model):
    mode_of_exercise = models.ForeignKey(ResistanceExercise, on_delete=models.CASCADE)
    date = models.DateField(null=True)
    day = models.CharField(max_length=100)
    exercise = models.CharField(max_length=100)
    total_exercise_reps_volume = models.IntegerField()
    intensity_zone_1 = models.IntegerField()
    intensity_zone_2 = models.IntegerField()
    intensity_zone_3 = models.IntegerField()
    intensity_zone_4 = models.IntegerField()
    intensity_zone_5 = models.IntegerField()
    intensity_zone_6 = models.IntegerField()
    intensity_zone_7 = models.IntegerField()
    intensity_zone_8 = models.IntegerField()
    intensity_zone_9 = models.IntegerField()
    intensity_zone_10 = models.IntegerField()
    intensity_zone_11 = models.IntegerField()
    intensity_zone_12 = models.IntegerField()
    intensity_zone_13 = models.IntegerField()


class WeeklyTrainingVolum(models.Model):
    mode_of_exercise = models.ForeignKey(ResistanceExercise, on_delete=models.CASCADE)
    wknr = models.IntegerField(null=True)
    exercise = models.CharField(max_length=100)
    total_intensity_zone = models.IntegerField()
    intensity_zone_1 = models.IntegerField()
    intensity_zone_2 = models.IntegerField()
    intensity_zone_3 = models.IntegerField()
    intensity_zone_4 = models.IntegerField()
    intensity_zone_5 = models.IntegerField()
    intensity_zone_6 = models.IntegerField()
    intensity_zone_7 = models.IntegerField()
    intensity_zone_8 = models.IntegerField()
    intensity_zone_9 = models.IntegerField()
    intensity_zone_10 = models.IntegerField()
    intensity_zone_11 = models.IntegerField()
    intensity_zone_12 = models.IntegerField()
    intensity_zone_13 = models.IntegerField()
    week = models.IntegerField()


class DailyVariables(models.Model):
    mode_of_exercise = models.ForeignKey(ResistanceExercise, on_delete=(models.CASCADE))
    date = models.DateField(null=True)
    weight = models.IntegerField()
    injury_status = models.BooleanField()
    injury_type = models.CharField(max_length=100)
    sleep_quality = models.IntegerField()
    sleep_duration = models.IntegerField()
    rpe = models.IntegerField()
    prs = models.IntegerField()
    training_period = models.CharField(max_length=100)