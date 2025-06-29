from django.db import models
from django.contrib.auth.models import AbstractUser

class Group(models.Model):
    name = models.CharField(verbose_name='نام گروه', max_length=100)

    class Meta:
        verbose_name = 'گروه'
        verbose_name_plural = 'گروه ها'

    def __str__(self):
        return self.name

class CustomUser(AbstractUser):
    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='custom_user_group', verbose_name='گروه', null=True, blank=True)
    is_manager = models.BooleanField(verbose_name='دسترسی مسئول', default=True)

    class Meta:
        verbose_name = 'اکانت'
        verbose_name_plural = 'اکانت‌ ها'

    def __str__(self):
        return self.username

class Excel(models.Model):
    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='excel_group', verbose_name='گروه')
    file = models.FileField(verbose_name='فایل اکسل', upload_to='excels/', default='')

    class Meta:
        verbose_name = 'اکسل'
        verbose_name_plural = 'اکسل ها'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class Expertise(models.Model):
    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='expertise_group', verbose_name='گروه')
    name = models.CharField(verbose_name='نام تخصص', max_length=100)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = 'تخصص'
        verbose_name_plural = 'تخصص ها'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class Section(models.Model):
    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='section_group', verbose_name='گروه')
    name = models.CharField(verbose_name='نام بخش', max_length=100)
    sheet = models.CharField(verbose_name='برگ بخش', max_length=100, null=True, blank=True)
    expertises = models.ManyToManyField(Expertise, related_name='section_expertises', verbose_name='تخصص ها', blank=True)

    def __str__(self):
        return self.name
    
    class Meta:
        verbose_name = 'بخش'
        verbose_name_plural = 'بخش ها'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class Room(models.Model):
    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='room_group', verbose_name='گروه')
    name = models.CharField(verbose_name='نام اتاق عمل', max_length=100)
    sheet = models.CharField(verbose_name='برگ اتاق عمل', max_length=100, null=True, blank=True)
    expertises = models.ManyToManyField(Expertise, related_name='room_expertises', verbose_name='تخصص ها', blank=True)

    def __str__(self):
        return self.name
    
    class Meta:
        verbose_name = 'اتاق عمل'
        verbose_name_plural = 'اتاق های عمل'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class Doctor(models.Model):
    grade_choices = [
        ('1', 'عمومی'),
        ('2', 'متخصص'),
        ('3', 'فوق تخصص'),
    ]

    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='doctor_group', verbose_name='گروه')
    full_name = models.CharField(verbose_name='نام و نام خانوادگی', max_length=100)
    sections = models.ManyToManyField(Section, related_name='doctor_sections', verbose_name='بخش ها', blank=True)
    rooms = models.ManyToManyField(Room, related_name='doctor_rooms', verbose_name='اتاق های عمل', blank=True)
    expertises = models.ManyToManyField(Expertise, related_name='doctor_expertises', verbose_name='تخصص ها', blank=True)
    grade = models.CharField(verbose_name='درجه', max_length=1, choices=grade_choices, null=True, blank=True)
    personnel_code = models.CharField(verbose_name='کد پرسنلی', max_length=100, null=True, blank=True)

    def __str__(self):
        return self.full_name
    
    class Meta:
        verbose_name = 'پزشک'
        verbose_name_plural = 'پزشکان'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class Patient(models.Model):
    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='patient_group', verbose_name='گروه')
    full_name = models.CharField(verbose_name='نام و نام خانوادگی', max_length=500)
    sections = models.ManyToManyField(Section, related_name='patient_sections', verbose_name='بخش ها', blank=True)
    rooms = models.ManyToManyField(Room, related_name='patient_rooms', verbose_name='اتاق های عمل', blank=True)

    def __str__(self):
        return f'بیمار با نام {self.full_name}'
    
    class Meta:
        verbose_name = 'بیمار'
        verbose_name_plural = 'بیماران'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class SectionCase(models.Model):
    defect_sheet_choices = [
        ('1', 'برگ پذیرش خلاصه ترخیص'),
        ('2', 'برگ خلاصه پرونده'),
        ('3', 'برگ شرح حال'),
        ('4', 'برگ سیربیماری'),
        ('5', 'برگ مشاوره'),
        ('6', 'برگ مراقبت قبل از عمل'),
        ('7', 'برگ بیهوشی'),
        ('8', 'برگ شرح عمل'),
        ('9', 'برگ مراقبت بعد از عمل'),
        ('10', 'دستورات پزشک'),
        ('11', 'گزارش پرستار'),
        ('12', 'نمودار علائم حیاتی'),
        ('13', 'رضایت آگاهانه'),
        ('14', 'صورتحساب'),
        ('15', 'چک لیست'),
    ]

    defect_type_choices = [
        ('1', 'عدم درج مهر پزشک'),
        ('2', 'مهر مشاوره'),
        ('3', 'مهر tellorder'),
        ('4', 'فقدان برگ'),
        ('5', 'عدم تکمیل گزارش'),
        ('6', 'خط خوردگی'),
        ('7', 'عدم تکمیل سربرگ'),
        ('8', 'عدم تشخیص نویسی'),
        ('9', 'عدم اخذ رضایت'),
        ('10', 'عدم اثر انگشت و امضا'),
        ('11', 'عدم ثبت دقیق آدرس و تلفن بیمار'),
    ]

    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='section_case_group', verbose_name='گروه')
    insurance = models.CharField(verbose_name='بیمه', max_length=100, null=True, blank=True)
    discharge_date = models.CharField(verbose_name='تاریخ ترخیص', max_length=10, null=True, blank=True)
    section = models.ForeignKey(Section, on_delete=models.CASCADE, related_name='section_case', verbose_name='بخش')
    doctor = models.ForeignKey(Doctor, on_delete=models.CASCADE, related_name='doctor_case', verbose_name='پزشک')
    admission_date = models.CharField(verbose_name='تاریخ پذیرش', max_length=10)
    number = models.CharField(verbose_name='شماره پرونده', max_length=50)
    representative_doctor = models.ForeignKey(Doctor, on_delete=models.CASCADE, related_name='r_doctor_case', verbose_name='پزشک معرف')
    patient = models.ForeignKey(Patient, on_delete=models.CASCADE, related_name='patient_case', verbose_name='بیمار')
    delivery_date = models.CharField(verbose_name='تاریخ تحویل', max_length=10, null=True, blank=True)
    defect_sheet = models.CharField(verbose_name='برگ نقص', max_length=2, choices=defect_sheet_choices, null=True, blank=True)
    defect_type = models.CharField(verbose_name='نوع نقص', max_length=2, choices=defect_type_choices, null=True, blank=True)
    defect_sheet2 = models.CharField(verbose_name='2 برگ نقص', max_length=2, choices=defect_sheet_choices, null=True, blank=True)
    defect_type2 = models.CharField(verbose_name='2 نوع نقص', max_length=2, choices=defect_type_choices, null=True, blank=True)

    def __str__(self):
        return self.number
    
    class Meta:
        verbose_name = 'پرونده بخش'
        verbose_name_plural = 'پرونده های بخش'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class RoomCase(models.Model):
    operation_type_choices = [
        ('1', 'عمل کوچک'),
        ('2', 'عمل متوسط'),
        ('3', 'عمل بزرگ'),
    ]

    anesthesia_type_choices = [
        ('1', 'موضعی'),
        ('2', 'غیر موضعی'),
    ]

    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='room_case_group', verbose_name='گروه')
    hospitalization_date = models.CharField(verbose_name='تاریخ بستری', max_length=10)
    discharge_date = models.CharField(verbose_name='تاریخ ترخیص', max_length=10, null=True, blank=True)
    operation_date = models.CharField(verbose_name='تاریخ عمل', max_length=10)
    patient = models.ForeignKey(Patient, on_delete=models.CASCADE, related_name='patient_room_case', verbose_name='بیمار')
    number = models.CharField(verbose_name='شماره پرونده', max_length=50)
    room = models.ForeignKey(Room, on_delete=models.CASCADE, related_name='room_case', verbose_name='اتاق عمل')
    operation_type = models.CharField(verbose_name='نوع جراحی', max_length=1, choices=operation_type_choices, null=True, blank=True)
    k = models.IntegerField(verbose_name='کا')
    doctor = models.ForeignKey(Doctor, on_delete=models.CASCADE, related_name='doctor_room_case', verbose_name='جراح')
    anesthesia_type = models.CharField(verbose_name='نوع بیهوشی', max_length=1, choices=anesthesia_type_choices, null=True, blank=True)

    def __str__(self):
        return self.number
    
    class Meta:
        verbose_name = 'پرونده اتاق عمل'
        verbose_name_plural = 'پرونده های اتاق عمل'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)

class DC(models.Model):
    gender_choices = [
        ('1', 'مرد'),
        ('2', 'زن'),
    ]

    group = models.ForeignKey(Group, on_delete=models.CASCADE, related_name='dc_group', verbose_name='گروه')
    number = models.CharField(verbose_name='شماره پرونده', max_length=50)
    doctor = models.ForeignKey(Doctor, on_delete=models.CASCADE, related_name='doctor_dc', verbose_name='پزشک')
    cause_of_death = models.CharField(verbose_name='علت فوت', max_length=100, null=True, blank=True)
    location_of_death = models.CharField(verbose_name='بخش محل فوت', max_length=100, null=True, blank=True)
    hospitalization_section = models.ForeignKey(Section, on_delete=models.CASCADE, related_name='section_dc', verbose_name='بخش بستری')
    death_date = models.CharField(verbose_name='تاریخ فوت', max_length=10)
    admission_date = models.CharField(verbose_name='تاریخ پذیرش', max_length=10)
    age = models.CharField(verbose_name='سن بیمار', max_length=3)
    gender = models.CharField(verbose_name='جنسیت', max_length=1, choices=gender_choices)
    patient = models.ForeignKey(Patient, on_delete=models.CASCADE, related_name='patient_dc', verbose_name='بیمار')
    delivery_date = models.CharField(verbose_name='تاریخ تحویل', max_length=10, null=True, blank=True)

    def __str__(self):
        return self.number
    
    class Meta:
        verbose_name = 'پرونده فوت'
        verbose_name_plural = 'پرونده های فوت'
    
    def save(self, *args, **kwargs):
        if hasattr(self, '_user'):
            self.group = self._user.group
        super().save(*args, **kwargs)