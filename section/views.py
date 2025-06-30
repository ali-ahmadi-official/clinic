import pandas as pd
from functools import wraps
from collections import Counter
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse, reverse_lazy
from django.contrib.auth import authenticate, login
from django.contrib.auth.decorators import login_required
from django.contrib.auth.mixins import LoginRequiredMixin
from django.contrib.auth.forms import UserCreationForm
from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.db.models import Q
from django.views.generic import ListView, CreateView, DetailView, UpdateView, DeleteView
from .models import Excel, Expertise, Section, Room, Doctor, Patient, SectionCase, RoomCase, DC
from .forms import CustomUserCreationForm, LoginForm, ExcelForm, ExpertiseForm, SectionForm, RoomForm, DoctorForm, SectionCaseForm, ConfirmDeleteForm
from .mixins import ManagerRequiredMixin, UserIsOwnerMixin
from .jalali import Persian
from .dictionary import defect_sheet_map, defect_type_map, operation_type_dict, gender_dict

# در این ویو کامنت گذاری در توابع پیچیده تر انجام شده

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

# میکسین های ویو تابع بیس

def manager_required(view_func):
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.user.is_manager:
            raise PermissionDenied("دسترسی فقط برای مسئول درمانگاه مجاز است.")
        return view_func(request, *args, **kwargs)
    return _wrapped_view

def group_is_owner(model, lookup_field='pk', group_field='group'):
    def decorator(view_func):
        @wraps(view_func)
        def _wrapped_view(request, *args, **kwargs):
            obj = get_object_or_404(model, **{lookup_field: kwargs.get(lookup_field)})
            if getattr(obj, group_field) != request.user.group:
                raise PermissionDenied("شما اجازه دسترسی به این مورد را ندارید.")
            return view_func(request, *args, **kwargs)
        return _wrapped_view
    return decorator

class SignUpView(CreateView):
    form_class = CustomUserCreationForm
    template_name = 'signup.html'
    success_url = reverse_lazy('login')

def custom_login_view(request):
    form = LoginForm()

    if request.method == 'POST':
        form = LoginForm(request.POST)
        if form.is_valid():
            username = form.cleaned_data['username']
            password = form.cleaned_data['password']

            user = authenticate(request, username=username, password=password)
            if user:
                login(request, user)

                if getattr(user, 'is_manager', False):
                    return redirect('main')
                else:
                    return redirect('section_case_list')
            else:
                messages.error(request, 'نام کاربری یا رمز عبور اشتباه است.')

    return render(request, 'login.html', {'form': form})

@login_required
@manager_required
def main(request):
    if Excel.objects.filter(group=request.user.group):
        doctors_count = Doctor.objects.filter(group=request.user.group).count()
        patients_count = Patient.objects.filter(group=request.user.group).count()
        sections_count = Section.objects.filter(group=request.user.group).count()
        rooms_count = Room.objects.filter(group=request.user.group).count()
        section_cases_count = SectionCase.objects.filter(group=request.user.group).count()
        room_cases = RoomCase.objects.filter(group=request.user.group)
        bigroom_cases = RoomCase.objects.filter(group=request.user.group, operation_type='3')
        mediumroom_cases = RoomCase.objects.filter(group=request.user.group, operation_type='2')
        smallroom_cases = RoomCase.objects.filter(group=request.user.group, operation_type='1')
        room_cases_count = room_cases.count()
        cases_count = section_cases_count + room_cases_count
        defect_section_cases_count = SectionCase.objects.filter(
            Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
        ).filter(group=request.user.group).count()

        sections = Section.objects.filter(group=request.user.group).order_by('-id')[:3]
        rooms = Room.objects.filter(group=request.user.group).order_by('-id')[:3]

        defect_counts = {}
        defect_type_counts = {}
        doctor_room_list = []
        doctor_bigroom_list = []
        doctor_mediumroom_list = []
        doctor_smallroom_list = []

        # تعیین لیست پزشکان پرونده های اتاق عمل
        for room_case in room_cases:
            doctor_name = room_case.doctor.full_name
            doctor_room_list.append(doctor_name)
        
        # تعیین لیست پزشکان پرونده های اتاق عمل بزرگ
        for bigroom_case in bigroom_cases:
            doctor_name = bigroom_case.doctor.full_name
            doctor_bigroom_list.append(doctor_name)
        
        # تعیین لیست پزشکان پرونده های اتاق عمل متوسط
        for mediumroom_case in mediumroom_cases:
            doctor_name = mediumroom_case.doctor.full_name
            doctor_mediumroom_list.append(doctor_name)
        
        # تعیین لیست پزشکان پرونده های اتاق عمل کوچک
        for smallroom_case in smallroom_cases:
            doctor_name = smallroom_case.doctor.full_name
            doctor_smallroom_list.append(doctor_name)
        
        # تعیین پزشک با بیشترین عمل در انواع سایز
        most_doctor_room_list = Counter(doctor_room_list).most_common(1)
        most_doctor_bigroom_list = Counter(doctor_bigroom_list).most_common(1)
        most_doctor_mediumroom_list = Counter(doctor_mediumroom_list).most_common(1)
        most_doctor_smallroom_list = Counter(doctor_smallroom_list).most_common(1)

        # پراکندگی اوراق نقص در بین پرونده های نقص خورده
        for code, name in defect_sheet_choices:
            count_sheet1 = SectionCase.objects.filter(group=request.user.group, defect_sheet=code).count()
            count_sheet2 = SectionCase.objects.filter(group=request.user.group, defect_sheet2=code).count()
            defect_counts[name] = count_sheet1 + count_sheet2
        
        # پراکندگی انواع نقص در بین پرونده های نقص خورده
        for code, name in defect_type_choices:
            count_sheet1 = SectionCase.objects.filter(group=request.user.group, defect_type=code).count()
            count_sheet2 = SectionCase.objects.filter(group=request.user.group, defect_type2=code).count()
            defect_type_counts[name] = count_sheet1 + count_sheet2

        context = {
            'doctors_count': doctors_count,
            'patients_count': patients_count,
            'sections_count': sections_count,
            'rooms_count': rooms_count,
            'cases_count': cases_count,
            'section_cases_count': section_cases_count,
            'defect_section_cases_count': defect_section_cases_count,
            'room_cases_count': room_cases_count,
            'sections': sections,
            'rooms': rooms,
            'defect_counts': defect_counts,
            'defect_type_counts': defect_type_counts,
            'most_doctor_room_list': most_doctor_room_list,
            'most_doctor_bigroom_list': most_doctor_bigroom_list,
            'most_doctor_mediumroom_list': most_doctor_mediumroom_list,
            'most_doctor_smallroom_list': most_doctor_smallroom_list,
        }

        return render(request, 'main.html', context=context)
    else:
        if request.method == 'POST':
            form = ExcelForm(request.POST, request.FILES)
            if form.is_valid():
                excel_instance = form.save(commit=False)
                excel_instance.group = request.user.group
                excel_instance.save()

                file_path = excel_instance.file.path
                excel_file = pd.ExcelFile(file_path)

                # دریافت نام همه شیت‌ها
                sheet_names = excel_file.sheet_names

                # جدا کردن شیت‌ها بر اساس کلمات کلیدی در اسم‌شون
                section_sheets = [name for name in sheet_names if 'section' in name.lower()]
                room_sheets = [name for name in sheet_names if 'room' in name.lower()]
                DC_sheets = [name for name in sheet_names if 'dc' in name.lower()]

                section_values = {}
                room_values = {}
                doctor_data = {}
                patient_data = {}

                for sheet in section_sheets:
                    try:
                        # پاک سازی و اصلاح جداول شیت
                        df = pd.read_excel(file_path, sheet_name=sheet)
                        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                        # ردیفی که ستون 9 آن دارای حرف P است به عنوان پرونده درنظر گرفته میشه
                        if df.shape[1] >= 9:
                            df = df[df.iloc[:, 8].astype(str).str.contains("P", na=False)]

                        # استخراج نام بخش ها از ستون سوم هر شیت
                        if df.shape[1] >= 3:
                            third_col_values = df.iloc[:, 2].dropna().astype(str).str.strip()
                            for value in third_col_values:
                                if value in section_values:
                                    section_values[value].add(sheet)
                                else:
                                    section_values[value] = {sheet}
                        
                        # استخراج نام پزشکان از ستون های 4 و 8
                        for index, row in df.iterrows():
                            if df.shape[1] >= 4 and df.shape[1] >= 9:
                                doctor_name_4 = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else None
                                doctor_name_8 = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else None
                                related_value = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else None

                            for doctor_name in [doctor_name_4, doctor_name_8]:
                                if doctor_name:
                                    if doctor_name in doctor_data:
                                        doctor_data[doctor_name].add(related_value)
                                    else:
                                        doctor_data[doctor_name] = {related_value}
                        
                        # استخراج نام بیماران از ستون نهم
                        for index, row in df.iterrows():
                            if df.shape[1] >= 9:
                                col9_value = str(row.iloc[8]).strip() if pd.notna(row.iloc[8]) else None
                                related_value = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else None

                                if col9_value:
                                    if col9_value in patient_data:
                                        patient_data[col9_value].add(related_value)
                                    else:
                                        patient_data[col9_value] = {related_value}

                    except Exception as e:
                        print(f"خطا در شیت '{sheet}': {e}")

                # اعمال کار های مشابه در شیت های اتاق عمل
                # نام اتاق از ستون 7
                # نام پزشک از ستون 10
                # نام بیمار از ترکیب ستون های نام بیمار و شناسه آن
                for sheet in room_sheets:
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet)
                        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                        if df.shape[1] >= 7:
                            seventh_col_values = df.iloc[:, 6].dropna().astype(str).str.strip()
                            for value in seventh_col_values:
                                if value in room_values:
                                    room_values[value].add(sheet)
                                else:
                                    room_values[value] = {sheet}
                        
                        for index, row in df.iterrows():
                            if df.shape[1] >= 10:
                                col10_value = str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else None
                                related_value = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else None

                            if col10_value:
                                if col10_value in doctor_data:
                                    doctor_data[col10_value].add(related_value)
                                else:
                                    doctor_data[col10_value] = {related_value}

                        if 'شناسه بیمار' in df.columns and 'نام بیمار' in df.columns:
                            combined_values = df[['شناسه بیمار', 'نام بیمار', df.columns[2]]].dropna().astype(str).apply(
                                lambda row: row['شناسه بیمار'].strip() + ' ' + row['نام بیمار'].strip(), axis=1)
    
                            related_values = df.iloc[:, 6].dropna().astype(str).str.strip()

                            for index, value in enumerate(combined_values):
                                related_value = related_values.iloc[index] if index < len(related_values) else None
                                if value and related_value:
                                    if value in patient_data:
                                        patient_data[value].add(related_value)
                                    else:
                                        patient_data[value] = {related_value}
                    except Exception as e:
                        print(f"خطا در شیت '{sheet}': {e}")

                # اتصال اتاق های به دست آمده به نام شیت آنها
                for value, sheets in room_values.items():
                    for sheet in sheets:
                        room, created = Room.objects.get_or_create(group=request.user.group, name=value, sheet='')
                
                # اعمال کار های مشابه در شیت های فوت
                # نام بخش از ستون 5
                # نام پزشک از ستون 2
                # نام بیمار از ترکیب ستون 12
                for DC_sheet in DC_sheets:
                    try:
                        df = pd.read_excel(file_path, sheet_name=DC_sheet)
                        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                        # ردیفی که ستون 1 آن دارای حرف U است به عنوان پرونده درنظر گرفته میشه
                        if df.shape[1] >= 1:
                            df = df[df.iloc[:, 0].astype(str).str.contains("U", na=False)]
                        
                        if df.shape[1] >= 5:
                            third_col_values = df.iloc[:, 4].dropna().astype(str).str.strip()
                            for value in third_col_values:
                                if value in section_values:
                                    section_values[value].add(DC_sheet)
                                else:
                                    section_values[value] = {DC_sheet}
                        
                        for index, row in df.iterrows():
                            if df.shape[1] >= 2:
                                col10_value = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else None
                                related_value = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else None

                            if col10_value:
                                if col10_value in doctor_data:
                                    doctor_data[col10_value].add(related_value)
                                else:
                                    doctor_data[col10_value] = {related_value}
                        
                        for index, row in df.iterrows():
                            if df.shape[1] >= 12:
                                col9_value = str(row.iloc[11]).strip() if pd.notna(row.iloc[11]) else None
                                related_value = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else None

                                if col9_value:
                                    if col9_value in patient_data:
                                        patient_data[col9_value].add(related_value)
                                    else:
                                        patient_data[col9_value] = {related_value}
                    except Exception as e:
                        print(f"خطا در شیت '{DC_sheet}': {e}")
                
                # اتصال بخش های به دست آمده به نام شیت آنها
                for value, sheets in section_values.items():
                    for sheet in sheets:
                        section, created = Section.objects.get_or_create(group=request.user.group, name=value, sheet='')
                
                # اتصال پزشکان به بخش ها و اتاق های آنها
                for full_name, sheets in doctor_data.items():
                    doctor, created = Doctor.objects.get_or_create(group=request.user.group, full_name=full_name)

                    for sheet in sheets:
                        section = Section.objects.filter(group=request.user.group, name=sheet).first()
                        if section:
                            doctor.sections.add(section)
                        
                        room = Room.objects.filter(group=request.user.group, name=sheet).first()
                        if room:
                            doctor.rooms.add(room)

                        doctor.save()
                
                # اتصال بیماران به بخش ها و اتاق های آنها
                for full_name, sheets in patient_data.items():
                    patient, created = Patient.objects.get_or_create(group=request.user.group, full_name=full_name)

                    for sheet in sheets:
                        section = Section.objects.filter(group=request.user.group, name=sheet).first()
                        if section:
                            patient.sections.add(section)

                        room = Room.objects.filter(group=request.user.group, name=sheet).first()
                        if room:
                            patient.rooms.add(room)

                        patient.save()
                
                # برداشت ردیف های شیت های بخش به عنوان پرونده بخش
                for sheet in section_sheets:
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet)
                        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                        if df.shape[1] >= 9:
                            df = df[df.iloc[:, 8].astype(str).str.contains("P", na=False)]

                        if df.shape[1] >= 14:
                            for i in range(len(df)):
                                try:
                                    insurance = str(df.iloc[i, 0]).strip()
                                    discharge_date = str(df.iloc[i, 1]).strip()
                                    section_name = str(df.iloc[i, 2]).strip()
                                    doctor_name = str(df.iloc[i, 3]).strip()
                                    admission_date = str(df.iloc[i, 4]).strip()
                                    number = str(df.iloc[i, 6]).strip()
                                    rep_doctor_name = str(df.iloc[i, 7]).strip()
                                    patient_name = str(df.iloc[i, 8]).strip()
                                    delivery_date = str(df.iloc[i, 9]).strip()
                                    defect_sheet = str(df.iloc[i, 10]).strip()
                                    defect_type = str(df.iloc[i, 11]).strip()
                                    defect_sheet2 = str(df.iloc[i, 12]).strip()
                                    defect_type2 = str(df.iloc[i, 13]).strip()

                                    section = Section.objects.filter(group=request.user.group, name=section_name).first()
                                    doctor = Doctor.objects.filter(group=request.user.group, full_name=doctor_name).first()
                                    rep_doctor = Doctor.objects.filter(group=request.user.group, full_name=rep_doctor_name).first()
                                    patient = Patient.objects.filter(group=request.user.group, full_name=patient_name).first()
                                    defect_sheet = defect_sheet_map.get(defect_sheet, None)
                                    defect_type = defect_type_map.get(defect_type, None)
                                    defect_sheet2 = defect_sheet_map.get(defect_sheet2, None)
                                    defect_type2 = defect_type_map.get(defect_type2, None)

                                    SectionCase.objects.create(
                                        group=request.user.group,
                                        insurance=insurance,
                                        discharge_date=discharge_date,
                                        section=section,
                                        doctor=doctor,
                                        admission_date=admission_date,
                                        number=number,
                                        representative_doctor=rep_doctor,
                                        patient=patient,
                                        delivery_date=delivery_date,
                                        defect_sheet=defect_sheet,
                                        defect_type=defect_type,
                                        defect_sheet2=defect_sheet2,
                                        defect_type2=defect_type2
                                    )
                                except Exception as row_error:
                                    print(f" خطا در ردیف {i} از شیت {sheet}: {row_error}")
                    except Exception as e:
                        print(f" خطا در شیت '{sheet}': {e}")

                # برداشت ردیف های شیت های اتاق عمل به عنوان پرونده اتاق عمل
                for sheet in room_sheets:
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet)
                        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                        if df.shape[1] >= 11:
                            for i in range(len(df)):
                                try:
                                    hospitalization_date = str(df.iloc[i, 0]).strip()
                                    discharge_date = str(df.iloc[i, 1]).strip()
                                    operation_date = str(df.iloc[i, 2]).strip()
                                    patient_name = str(df.iloc[i, 3]).strip()
                                    number = str(df.iloc[i, 5]).strip()
                                    room_name = str(df.iloc[i, 6]).strip()
                                    operation_type = str(df.iloc[i, 7]).strip()
                                    k = str(df.iloc[i, 8]).strip()
                                    doctor_name = str(df.iloc[i, 9]).strip()
                                    anesthesia_type = str(df.iloc[i, 10]).strip()

                                    patient = Patient.objects.filter(group=request.user.group, full_name__icontains=patient_name).first()
                                    room = Room.objects.filter(group=request.user.group, name=room_name).first()
                                    operation_type = operation_type_dict.get(operation_type, None)
                                    doctor = Doctor.objects.filter(group=request.user.group, full_name=doctor_name).first()

                                    RoomCase.objects.create(
                                        group=request.user.group,
                                        hospitalization_date=hospitalization_date,
                                        discharge_date=discharge_date,
                                        operation_date=operation_date,
                                        patient=patient,
                                        number=number,
                                        room=room,
                                        operation_type=operation_type,
                                        k=k,
                                        doctor=doctor,
                                        anesthesia_type=anesthesia_type,
                                    )
                                except Exception as row_error:
                                    print(f" خطا در ردیف {i} از شیت {sheet}: {row_error}")
                    except Exception as e:
                        print(f" خطا در شیت '{sheet}': {e}")
                
                # برداشت ردیف های شیت فوت به عنوان پرونده فوت
                for DC_sheet in DC_sheets:
                    try:
                        df = pd.read_excel(file_path, sheet_name=DC_sheet)
                        df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                        if df.shape[1] >= 1:
                            df = df[df.iloc[:, 0].astype(str).str.contains("U", na=False)]
                        
                        if df.shape[1] >= 14:
                            for i in range(len(df)):
                                try:
                                    number = str(df.iloc[i, 0]).strip()
                                    doctor_full_name = str(df.iloc[i, 1]).strip()
                                    cause_of_death = str(df.iloc[i, 2]).strip()
                                    location_of_death = str(df.iloc[i, 3]).strip()
                                    hospitalization_section = str(df.iloc[i, 4]).strip()
                                    death_date = str(df.iloc[i, 5]).strip()
                                    admission_date = str(df.iloc[i, 6]).strip()
                                    age = str(df.iloc[i, 9]).strip()
                                    gender = str(df.iloc[i, 10]).strip()
                                    patient_name = str(df.iloc[i, 11]).strip()
                                    delivery_date = str(df.iloc[i, 13]).strip()

                                    doctor = Doctor.objects.filter(group=request.user.group, full_name=doctor_full_name).first()
                                    section = Section.objects.filter(group=request.user.group, name=hospitalization_section).first()
                                    gender = gender_dict.get(gender, None)
                                    patient = Patient.objects.filter(group=request.user.group, full_name=patient_name).first()

                                    DC.objects.create(
                                        group=request.user.group,
                                        number=number,
                                        doctor=doctor,
                                        cause_of_death=cause_of_death,
                                        location_of_death=location_of_death,
                                        hospitalization_section=section,
                                        death_date=death_date,
                                        admission_date=admission_date,
                                        age=age,
                                        gender=gender,
                                        patient=patient,
                                        delivery_date=delivery_date,
                                    )
                                except Exception as row_error:
                                    print(f" خطا در ردیف {i} از شیت {sheet}: {row_error}")
                    except Exception as e:
                        print(f"خطا در شیت '{sheet}': {e}")
                return redirect('main')
        else:
            form = ExcelForm()
        
        context = {
            'form': form,
        }

        return render(request, 'create_excel.html', context=context)

@login_required
@manager_required
@group_is_owner(Section, lookup_field='pk', group_field='group')
def section_detail(request, pk):
    section = get_object_or_404(Section, pk=pk)
    doctors = section.doctor_sections.all()
    doctors_count = doctors.count()
    patients_count = section.patient_sections.all().count()

    section_cases = SectionCase.objects.filter(group=request.user.group, section=section)
    dc_section_cases = DC.objects.filter(group=request.user.group, hospitalization_section=section)
    not_arrived_cases = SectionCase.objects.filter(group=request.user.group, section=section, delivery_date=None)

    defect_cases = SectionCase.objects.filter(
        group=request.user.group,
        section=section
    ).filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
    )

    social_security_cases = SectionCase.objects.filter(
        group=request.user.group,
        section=section
    ).filter(
        Q(insurance__icontains="تامین اجتماعی")
    )

    filtered_section_cases = []
    filtered_dc_section_cases = []
    filtered_not_arrived_cases = []
    filtered_defect_cases = []
    filtered_social_security_cases = []
    arrive_daies = []
    stay_daies = []
    defect_counts = {}
    defect_type_counts = {}
    doctor_cases = {}
    doctor_defects = {}

    if request.GET.get("start") and request.GET.get("end"):
        try:
            start = Persian(request.GET["start"]).gregorian_datetime()
            end = Persian(request.GET["end"]).gregorian_datetime()

            if start > end:
                start, end = end, start
            
            # استخراج پرونده های بخش در بازه زمانی
            for section_case in section_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(section_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_section_cases.append(section_case)
                    except:
                        continue

            # استخراج پرونده های فوت در بازه زمانی
            for dc_section_case in dc_section_cases:
                if isinstance(dc_section_case.admission_date, str):
                    try:
                        case_date = Persian(dc_section_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_dc_section_cases.append(dc_section_case)
                    except:
                        continue

            # استخراج پرونده های نرسیده بخش در بازه زمانی
            for not_arrived_case in not_arrived_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(not_arrived_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_not_arrived_cases.append(not_arrived_case)
                    except:
                        continue

            # استخراج پرونده های ناقص بخش در بازه زمانی
            for defect_case in defect_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(defect_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_defect_cases.append(defect_case)
                    except:
                        continue
            
            # استخراج پرونده های تامین اجتماعی بخش در بازه زمانی
            for social_security_case in social_security_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(social_security_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_social_security_cases.append(social_security_case)
                    except:
                        continue
            
            # استخراج میانگین رسیدن پرونده های بازه زمانی بخش
            for section_case in filtered_section_cases:
                d_date = section_case.discharge_date
                dl_date = section_case.delivery_date

                if isinstance(d_date, str) and isinstance(dl_date, str):
                    try:
                        discharge_date = Persian(d_date).gregorian_datetime()
                        delivery_date = Persian(dl_date).gregorian_datetime()

                        arrive_day = (delivery_date - discharge_date).days
                        arrive_daies.append(arrive_day)
                    except Exception as e:
                        continue
            
            # استخراج میانگین اقامت بیماران در پرونده های موجود در بازه زمانی بخش
            for section_case in filtered_section_cases:
                d_date = section_case.discharge_date
                ad_date = section_case.admission_date

                if isinstance(d_date, str) and isinstance(ad_date, str):
                    try:
                        discharge_date = Persian(d_date).gregorian_datetime()
                        admission_date = Persian(ad_date).gregorian_datetime()

                        stay_day = (discharge_date - admission_date).days
                        stay_daies.append(stay_day)
                    except Exception as e:
                        continue
            
            # بررسی پراکندگی اوراق نقص پرونده های بازه زمانی بخش
            for code, name in defect_sheet_choices:
                count_sheet = 0
                for section_case in filtered_section_cases:
                    if section_case.defect_sheet==code:
                        count_sheet += 1
                    
                    if section_case.defect_sheet2==code:
                        count_sheet += 1
                defect_counts[name] = count_sheet
            
            # بررسی پراکندگی انواع نقص پرونده های بازه زمانی بخش
            for code, name in defect_type_choices:
                count_sheet = 0
                for section_case in filtered_section_cases:
                    if section_case.defect_type==code:
                        count_sheet += 1
                    
                    if section_case.defect_type2==code:
                        count_sheet += 1
                defect_type_counts[name] = count_sheet
            
            # بررسی تعداد پرونده های پزشکان بخش در بازه زمانی
            # بررسی تعداد پرونده های نقص خورده پزشکان بخش در بازه زمانی
            for doctor in doctors:
                in_time_doctor_cases = SectionCase.objects.filter(group=request.user.group, section=section, doctor=doctor)
                in_time_doctor_defects = SectionCase.objects.filter(
                    group=request.user.group,
                    section=section,
                    doctor=doctor
                ).filter(
                    Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
                )

                num_doctor_cases = 0
                for in_time_doctor_case in in_time_doctor_cases:
                    if isinstance(in_time_doctor_case.admission_date, str):
                        try:
                            case_date = Persian(in_time_doctor_case.admission_date).gregorian_datetime()
                            if start <= case_date <= end:
                                num_doctor_cases += 1
                        except:
                            continue
                doctor_cases[doctor.full_name] = num_doctor_cases

                num_doctor_defects = 0
                for in_time_doctor_defect in in_time_doctor_defects:
                    if isinstance(in_time_doctor_defect.admission_date, str):
                        try:
                            case_date = Persian(in_time_doctor_defect.admission_date).gregorian_datetime()
                            if start <= case_date <= end:
                                num_doctor_defects += 1
                        except:
                            continue
                doctor_defects[doctor.full_name] = num_doctor_defects
        except Exception as e:
            print("Error converting dates:", e)
    else:
        # اگر بازه انتخاب نشود کل پرونده های بخش بررسی می شود
        filtered_section_cases = section_cases
        filtered_dc_section_cases = dc_section_cases
        filtered_not_arrived_cases = not_arrived_cases
        filtered_defect_cases = defect_cases
        filtered_social_security_cases = social_security_cases

        # استخراج میانگین رسیدن پرونده های بخش
        for section_case in section_cases:
            d_date = section_case.discharge_date
            dl_date = section_case.delivery_date

            if isinstance(d_date, str) and isinstance(dl_date, str):
                try:
                    discharge_date = Persian(d_date).gregorian_datetime()
                    delivery_date = Persian(dl_date).gregorian_datetime()

                    arrive_day = (delivery_date - discharge_date).days
                    arrive_daies.append(arrive_day)
                except Exception as e:
                    continue
        
        # استخراج میانگین اقامت بیماران در پرونده های موجود در بخش
        for section_case in section_cases:
            d_date = section_case.discharge_date
            ad_date = section_case.admission_date

            if isinstance(d_date, str) and isinstance(ad_date, str):
                try:
                    discharge_date = Persian(d_date).gregorian_datetime()
                    admission_date = Persian(ad_date).gregorian_datetime()

                    stay_day = (discharge_date - admission_date).days
                    stay_daies.append(stay_day)
                except Exception as e:
                    continue
        
        # بررسی پراکندگی اوراق نقص پرونده های بخش
        for code, name in defect_sheet_choices:
            count_sheet1 = SectionCase.objects.filter(group=request.user.group, section=section, defect_sheet=code).count()
            count_sheet2 = SectionCase.objects.filter(group=request.user.group, section=section, defect_sheet2=code).count()
            defect_counts[name] = count_sheet1 + count_sheet2

        # بررسی پراکندگی انواع نقص پرونده های بخش
        for code, name in defect_type_choices:
                count_sheet1 = SectionCase.objects.filter(group=request.user.group, section=section, defect_type=code).count()
                count_sheet2 = SectionCase.objects.filter(group=request.user.group, section=section, defect_type2=code).count()
                defect_type_counts[name] = count_sheet1 + count_sheet2

        # برداشت تعداد پرونده های پزشکان در این بخش
        for doctor in doctors:
            doctor_cases[doctor.full_name] = SectionCase.objects.filter(group=request.user.group, section=section, doctor=doctor).count()
            doctor_defects[doctor.full_name] = SectionCase.objects.filter(
                group=request.user.group,
                section=section,
                doctor=doctor
            ).filter(
                Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
            ).count()

    # بررسی شمار پرونده ها یا میانگین مدت زمان ها
    filtered_section_cases_count = len(filtered_section_cases)
    filtered_dc_section_cases_count = len(filtered_dc_section_cases)
    filtered_not_arrived_cases_count = len(filtered_not_arrived_cases)
    filtered_defect_cases_count = len(filtered_defect_cases)
    filtered_social_security_cases_count = len(filtered_social_security_cases)

    if arrive_daies:
        average_arrive_daies = round(sum(arrive_daies) / len(arrive_daies), 0)
    else:
        average_arrive_daies = '0'
    
    if stay_daies:
        average_stay_daies = round(sum(stay_daies) / len(stay_daies), 0)
    else:
        average_stay_daies = '0'

    context = {
        'section': section,
        'doctors_count': doctors_count,
        'patients_count': patients_count,
        'filtered_section_cases': filtered_section_cases,
        'filtered_dc_section_cases_count': filtered_dc_section_cases_count,
        'filtered_not_arrived_cases': filtered_not_arrived_cases,
        'filtered_section_cases_count': filtered_section_cases_count,
        'filtered_not_arrived_cases_count': filtered_not_arrived_cases_count,
        'filtered_defect_cases_count': filtered_defect_cases_count,
        'filtered_social_security_cases_count': filtered_social_security_cases_count,
        'average_arrive_daies': average_arrive_daies,
        'average_stay_daies': average_stay_daies,
        'defect_counts': defect_counts,
        'defect_type_counts': defect_type_counts,
        'doctor_cases': doctor_cases,
        'doctor_defects': doctor_defects,
    }

    return render(request, 'section_detail.html', context)

@login_required
@manager_required
@group_is_owner(Room, lookup_field='pk', group_field='group')
def room_detail(request, pk):
    room = get_object_or_404(Room, pk=pk)
    doctors = room.doctor_rooms.all()
    doctors_count = doctors.count()
    patients_count = room.patient_rooms.all().count()
    room_cases = RoomCase.objects.filter(group=request.user.group, room=room)
    big_room_cases = RoomCase.objects.filter(group=request.user.group, room=room, operation_type='3')
    medium_room_cases = RoomCase.objects.filter(group=request.user.group, room=room, operation_type='2')
    small_room_cases = RoomCase.objects.filter(group=request.user.group, room=room, operation_type='1')

    filtered_room_cases = []
    filtered_big_room_cases = []
    filtered_medium_room_cases = []
    filtered_small_room_cases = []
    doctor_cases = {}

    if request.GET.get("start") and request.GET.get("end"):
        try:
            start = Persian(request.GET["start"]).gregorian_datetime()
            end = Persian(request.GET["end"]).gregorian_datetime()

            if start > end:
                start, end = end, start

            # استخراج پرونده های اتاق عمل در بازه زمانی
            for room_case in room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_room_cases.append(room_case)
                    except:
                        continue
            
            # استخراج پرونده های عمل بزرگ اتاق عمل در بازه زمانی
            for room_case in big_room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_big_room_cases.append(room_case)
                    except:
                        continue
            
            # استخراج پرونده های عمل متوسط اتاق عمل در بازه زمانی
            for room_case in medium_room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_medium_room_cases.append(room_case)
                    except:
                        continue
            
            # استخراج پرونده های عمل کوچک اتاق عمل در بازه زمانی
            for room_case in small_room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_small_room_cases.append(room_case)
                    except:
                        continue
            
            # بررسی تعداد پرونده های پزشکان اتاق عمل در بازه زمانی
            for doctor in doctors:
                in_time_doctor_cases = RoomCase.objects.filter(group=request.user.group, room=room, doctor=doctor)

                num_doctor_cases = 0
                for in_time_doctor_case in in_time_doctor_cases:
                    if isinstance(in_time_doctor_case.operation_date, str):
                        try:
                            case_date = Persian(in_time_doctor_case.operation_date).gregorian_datetime()
                            if start <= case_date <= end:
                                num_doctor_cases += 1
                        except:
                            continue
                doctor_cases[doctor.full_name] = num_doctor_cases
        except Exception as e:
            print("Error converting dates:", e)
    else:
        filtered_room_cases = room_cases
        filtered_big_room_cases = big_room_cases
        filtered_medium_room_cases = medium_room_cases
        filtered_small_room_cases = small_room_cases

        # برداشت تعداد پرونده های پزشکان در این اتاق عمل
        for doctor in doctors:
            doctor_cases[doctor.full_name] = RoomCase.objects.filter(group=request.user.group, room=room, doctor=doctor).count()
    
    filtered_room_cases_count = len(filtered_room_cases)
    filtered_big_room_cases_count = len(filtered_big_room_cases)
    filtered_medium_room_cases_count = len(filtered_medium_room_cases)
    filtered_small_room_cases_count = len(filtered_small_room_cases)

    context = {
        'room': room,
        'doctors_count': doctors_count,
        'patients_count': patients_count,
        'filtered_room_cases_count': filtered_room_cases_count,
        'filtered_big_room_cases_count': filtered_big_room_cases_count,
        'filtered_medium_room_cases_count': filtered_medium_room_cases_count,
        'filtered_small_room_cases_count': filtered_small_room_cases_count,
        'doctor_cases': doctor_cases,
    }
    
    return render(request, 'room_detail.html', context)

@login_required
@manager_required
@group_is_owner(Doctor, lookup_field='pk', group_field='group')
def doctor_detail(request, pk):
    doctor = get_object_or_404(Doctor, pk=pk)
    section_cases = SectionCase.objects.filter(group=request.user.group, doctor=doctor)
    dc_section_cases = DC.objects.filter(group=request.user.group, doctor=doctor)
    room_cases = RoomCase.objects.filter(group=request.user.group, doctor=doctor)
    big_room_cases = RoomCase.objects.filter(group=request.user.group, doctor=doctor, operation_type='3')
    medium_room_cases = RoomCase.objects.filter(group=request.user.group, doctor=doctor, operation_type='2')
    small_room_cases = RoomCase.objects.filter(group=request.user.group, doctor=doctor, operation_type='1')

    not_arrived_cases = SectionCase.objects.filter(group=request.user.group, doctor=doctor, delivery_date=None)

    defect_cases = SectionCase.objects.filter(
        group=request.user.group,
        doctor=doctor
    ).filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
    )

    all_defect_cases = SectionCase.objects.filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
    ).filter(group=request.user.group)

    social_security_cases = SectionCase.objects.filter(
        group=request.user.group,
        doctor=doctor
    ).filter(
        Q(insurance__icontains="تامین اجتماعی")
    )

    patients = []
    filtered_section_cases = []
    filtered_dc_section_cases = []
    filtered_not_arrived_cases = []
    filtered_defect_cases = []
    filtered_all_defect_cases = []
    filtered_social_security_cases = []
    filtered_room_cases = []
    filtered_big_room_cases = []
    filtered_medium_room_cases = []
    filtered_small_room_cases = []
    arrive_daies = []
    stay_daies = []
    defect_counts = {}
    defect_type_counts = {}

    for section_case in section_cases:
        patient = section_case.patient
        if patient in patients:
            pass
        else:
            patients.append(patient)
    
    for room_case in room_cases:
        patient = room_case.patient
        if patient in patients:
            pass
        else:
            patients.append(patient)
    
    patients_count = len(patients)

    if request.GET.get("start") and request.GET.get("end"):
        try:
            start = Persian(request.GET["start"]).gregorian_datetime()
            end = Persian(request.GET["end"]).gregorian_datetime()

            if start > end:
                start, end = end, start

            # استخراج پرونده های پزشک در بازه زمانی
            for section_case in section_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(section_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_section_cases.append(section_case)
                    except:
                        continue
            
            # استخراج پرونده های فوت در بازه زمانی
            for section_case in dc_section_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(section_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_dc_section_cases.append(section_case)
                    except:
                        continue
            
            # استخراج پرونده های نرسیده پزشک در بازه زمانی
            for not_arrived_case in not_arrived_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(not_arrived_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_not_arrived_cases.append(not_arrived_case)
                    except:
                        continue

            # استخراج پرونده های ناقص پزشک در بازه زمانی
            for defect_case in defect_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(defect_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_defect_cases.append(defect_case)
                    except:
                        continue
            
            # استخراج همه پرونده های ناقص پزشک در بازه زمانی
            for all_defect_case in all_defect_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(defect_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_all_defect_cases.append(all_defect_case)
                    except:
                        continue
            
            # استخراج پرونده های تامین اجتماعی پزشک در بازه زمانی
            for social_security_case in social_security_cases:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(social_security_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_social_security_cases.append(social_security_case)
                    except:
                        continue
            
            # استخراج پرونده های اتاق عمل در بازه زمانی
            for room_case in room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_room_cases.append(room_case)
                    except:
                        continue
            
            # استخراج پرونده های عمل بزرگ اتاق عمل در بازه زمانی
            for room_case in big_room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_big_room_cases.append(room_case)
                    except:
                        continue
            
            # استخراج پرونده های عمل متوسط اتاق عمل در بازه زمانی
            for room_case in medium_room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_medium_room_cases.append(room_case)
                    except:
                        continue
            
            # استخراج پرونده های عمل کوچک اتاق عمل در بازه زمانی
            for room_case in small_room_cases:
                if isinstance(room_case.operation_date, str):
                    try:
                        case_date = Persian(room_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_small_room_cases.append(room_case)
                    except:
                        continue

            # استخراج میانگین رسیدن پرونده های بازه زمانی دکتر
            for section_case in filtered_section_cases:
                d_date = section_case.discharge_date
                dl_date = section_case.delivery_date

                if isinstance(d_date, str) and isinstance(dl_date, str):
                    try:
                        discharge_date = Persian(d_date).gregorian_datetime()
                        delivery_date = Persian(dl_date).gregorian_datetime()

                        arrive_day = (delivery_date - discharge_date).days
                        arrive_daies.append(arrive_day)
                    except Exception as e:
                        continue
            
            # استخراج میانگین اقامت بیماران در پرونده های موجود در بازه زمانی دکتر
            for section_case in filtered_section_cases:
                d_date = section_case.discharge_date
                ad_date = section_case.admission_date

                if isinstance(d_date, str) and isinstance(ad_date, str):
                    try:
                        discharge_date = Persian(d_date).gregorian_datetime()
                        admission_date = Persian(ad_date).gregorian_datetime()

                        stay_day = (discharge_date - admission_date).days
                        stay_daies.append(stay_day)
                    except Exception as e:
                        continue

            # بررسی پراکندگی انواع نقص پرونده های بازه زمانی دکتر
            for code, name in defect_sheet_choices:
                count_sheet = 0
                for section_case in filtered_section_cases:
                    if section_case.defect_sheet==code:
                        count_sheet += 1
                    
                    if section_case.defect_sheet2==code:
                        count_sheet += 1
                defect_counts[name] = count_sheet
            
            # بررسی پراکندگی انواع نقص پرونده های بازه زمانی دکتر
            for code, name in defect_type_choices:
                count_sheet = 0
                for section_case in filtered_section_cases:
                    if section_case.defect_type==code:
                        count_sheet += 1
                    
                    if section_case.defect_type2==code:
                        count_sheet += 1
                defect_type_counts[name] = count_sheet
        except Exception as e:
            print("Error converting dates:", e)
    else:
        # اگر بازه انتخاب نشود کل پرونده های پزشک بررسی می شود
        filtered_section_cases = section_cases
        filtered_dc_section_cases = dc_section_cases
        filtered_not_arrived_cases = not_arrived_cases
        filtered_defect_cases = defect_cases
        filtered_all_defect_cases = all_defect_cases
        filtered_social_security_cases = social_security_cases
        filtered_room_cases = room_cases
        filtered_big_room_cases = big_room_cases
        filtered_medium_room_cases = medium_room_cases
        filtered_small_room_cases = small_room_cases

        # استخراج میانگین رسیدن پرونده های پزشک
        for section_case in section_cases:
            d_date = section_case.discharge_date
            dl_date = section_case.delivery_date

            if isinstance(d_date, str) and isinstance(dl_date, str):
                try:
                    discharge_date = Persian(d_date).gregorian_datetime()
                    delivery_date = Persian(dl_date).gregorian_datetime()

                    arrive_day = (delivery_date - discharge_date).days
                    arrive_daies.append(arrive_day)
                except Exception as e:
                    continue
        
        # استخراج میانگین اقامت بیماران در پرونده های موجود در پزشک
        for section_case in section_cases:
            d_date = section_case.discharge_date
            ad_date = section_case.admission_date

            if isinstance(d_date, str) and isinstance(ad_date, str):
                try:
                    discharge_date = Persian(d_date).gregorian_datetime()
                    admission_date = Persian(ad_date).gregorian_datetime()

                    stay_day = (discharge_date - admission_date).days
                    stay_daies.append(stay_day)
                except Exception as e:
                    continue
        
        # بررسی پراکندگی اوراق نقص پرونده های دکتر
        for code, name in defect_sheet_choices:
            count_sheet1 = SectionCase.objects.filter(group=request.user.group, doctor=doctor, defect_sheet=code).count()
            count_sheet2 = SectionCase.objects.filter(group=request.user.group, doctor=doctor, defect_sheet2=code).count()
            defect_counts[name] = count_sheet1 + count_sheet2

        # بررسی پراکندگی انواع نقص پرونده های دکتر
        for code, name in defect_type_choices:
                count_sheet1 = SectionCase.objects.filter(group=request.user.group, doctor=doctor, defect_type=code).count()
                count_sheet2 = SectionCase.objects.filter(group=request.user.group, doctor=doctor, defect_type2=code).count()
                defect_type_counts[name] = count_sheet1 + count_sheet2
    
    # بررسی شمار پرونده ها یا میانگین مدت زمان ها
    filtered_section_cases_count = len(filtered_section_cases)
    filtered_dc_section_cases_count = len(filtered_dc_section_cases)
    filtered_not_arrived_cases_count = len(filtered_not_arrived_cases)
    filtered_defect_cases_count = len(filtered_defect_cases)
    filtered_social_security_cases_count = len(filtered_social_security_cases)
    filtered_room_cases_count = len(filtered_room_cases)
    filtered_big_room_cases_count = len(filtered_big_room_cases)
    filtered_medium_room_cases_count = len(filtered_medium_room_cases)
    filtered_small_room_cases_count = len(filtered_small_room_cases)

    if arrive_daies:
        average_arrive_daies = round(sum(arrive_daies) / len(arrive_daies), 0)
    else:
        average_arrive_daies = '0'
    
    if stay_daies:
        average_stay_daies = round(sum(stay_daies) / len(stay_daies), 0)
    else:
        average_stay_daies = '0'
    
    if filtered_all_defect_cases:
        percent_defect_cases = (len(filtered_defect_cases) * 100) // len(filtered_all_defect_cases)
    else:
        percent_defect_cases = '0'

    context = {
        'doctor': doctor,
        'patients_count': patients_count,
        'filtered_section_cases': filtered_section_cases,
        'filtered_dc_section_cases_count': filtered_dc_section_cases_count,
        'filtered_not_arrived_cases': filtered_not_arrived_cases,
        'filtered_section_cases_count': filtered_section_cases_count,
        'filtered_not_arrived_cases_count': filtered_not_arrived_cases_count,
        'filtered_defect_cases_count': filtered_defect_cases_count,
        'filtered_social_security_cases_count': filtered_social_security_cases_count,
        'filtered_room_cases_count': filtered_room_cases_count,
        'filtered_big_room_cases_count': filtered_big_room_cases_count,
        'filtered_medium_room_cases_count': filtered_medium_room_cases_count,
        'filtered_small_room_cases_count': filtered_small_room_cases_count,
        'average_arrive_daies': average_arrive_daies,
        'average_stay_daies': average_stay_daies,
        'defect_counts': defect_counts,
        'defect_type_counts': defect_type_counts,
        'percent_defect_cases': percent_defect_cases,
    }

    return render(request, 'doctor_detail.html', context)

@login_required
@manager_required
@group_is_owner(Patient, lookup_field='pk', group_field='group')
def patient_detail(request, pk):
    patient = get_object_or_404(Patient, pk=pk)
    section_cases = SectionCase.objects.filter(group=request.user.group, patient=patient)
    room_cases = RoomCase.objects.filter(group=request.user.group, patient=patient)
    dc_cases = DC.objects.filter(group=request.user.group, patient=patient)
    doctors = []

    for section_case in section_cases:
        doctor = section_case.doctor
        if doctor in doctors:
            pass
        else:
            doctors.append(doctor)
    
    for room_case in room_cases:
        doctor = room_case.doctor
        if doctor in doctors:
            pass
        else:
            doctors.append(doctor)
    
    for dc_case in dc_cases:
        doctor = dc_case.doctor
        if doctor in doctors:
            pass
        else:
            doctors.append(doctor)

    context = {
        'patient': patient,
        'section_cases': section_cases,
        'room_cases': room_cases,
        'dc_cases': dc_cases,
        'doctors': doctors,
    }

    return render(request, 'patient_detail.html', context=context)

class SectionListView(LoginRequiredMixin, ManagerRequiredMixin, ListView):
    model = Section
    template_name = 'section_list.html'
    context_object_name = 'sections'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()

        search_query = self.request.GET.get('q')
        id = self.request.GET.get('id')
        expertise_filter = self.request.GET.getlist('expertise')

        if search_query:
            queryset = queryset.filter(Q(name__icontains=search_query))
        if id:
            queryset = queryset.filter(id=id)
        if expertise_filter:
            queryset = queryset.filter(expertises__in=expertise_filter)

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        context['selected_expertises'] = [int(e) for e in self.request.GET.getlist('expertise')]
        return context

class SectionCreateView(LoginRequiredMixin, ManagerRequiredMixin, CreateView):
    model = Section
    form_class = SectionForm
    template_name = 'section_create.html'

    def get_success_url(self):
        return reverse('section_detail', kwargs={'pk': self.object.id})

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        return context
    
    def form_valid(self, form):
        form.instance.group = self.request.user.group
        return super().form_valid(form)

class SectionUpdateView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, UpdateView):
    model = Section
    form_class = SectionForm
    template_name = 'section_update.html'

    def get_success_url(self):
        return reverse('section_detail', kwargs={'pk': self.object.id})

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)

        return context

class SectionDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, DeleteView):
    model = Section
    template_name = 'section_confirm_delete.html'
    success_url = reverse_lazy('section_list')

class RoomListView(LoginRequiredMixin, ManagerRequiredMixin, ListView):
    model = Room
    template_name = 'room_list.html'
    context_object_name = 'rooms'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        search_query = self.request.GET.get('q')
        id = self.request.GET.get('id')
        expertise_filter = self.request.GET.getlist('expertise')

        if search_query:
            queryset = queryset.filter(Q(name__icontains=search_query))
        if id:
            queryset = queryset.filter(id=id)
        if expertise_filter:
            queryset = queryset.filter(expertises__in=expertise_filter)

        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        context['selected_expertises'] = [int(e) for e in self.request.GET.getlist('expertise')]
        return context

class RoomCreateView(LoginRequiredMixin, ManagerRequiredMixin, CreateView):
    model = Room
    form_class = RoomForm
    template_name = 'room_create.html'

    def get_success_url(self):
        return reverse('room_detail', kwargs={'pk': self.object.id})

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        return context
    
    def form_valid(self, form):
        form.instance.group = self.request.user.group
        return super().form_valid(form)

class RoomUpdateView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, UpdateView):
    model = Room
    form_class = RoomForm
    template_name = 'room_update.html'

    def get_success_url(self):
        return reverse('room_detail', kwargs={'pk': self.object.id})

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        return context

class RoomDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, DeleteView):
    model = Room
    template_name = 'room_confirm_delete.html'
    success_url = reverse_lazy('room_list')

class ExpertiseListView(LoginRequiredMixin, ManagerRequiredMixin, ListView):
    model = Expertise
    template_name = 'expertise_list.html'
    context_object_name = 'expertises'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        return queryset

class ExpertiseCreateView(LoginRequiredMixin, ManagerRequiredMixin, CreateView):
    model = Expertise
    form_class = ExpertiseForm
    template_name = 'expertise_create.html'

    def get_success_url(self):
        return reverse('expertise_list')
    
    def form_valid(self, form):
        form.instance.group = self.request.user.group
        return super().form_valid(form)

class ExpertiseUpdateView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, UpdateView):
    model = Expertise
    form_class = ExpertiseForm
    template_name = 'expertise_update.html'

    def get_success_url(self):
        return reverse('expertise_list')

class ExpertiseDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, DeleteView):
    model = Expertise
    template_name = 'expertise_confirm_delete.html'
    success_url = reverse_lazy('expertise_list')

class DoctorListView(LoginRequiredMixin, ManagerRequiredMixin, ListView):
    model = Doctor
    template_name = 'doctor_list.html'
    context_object_name = 'doctors'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        search_query = self.request.GET.get('q')
        personnel_code = self.request.GET.get('personnel_code')
        expertise_filter = self.request.GET.getlist('expertise')

        if search_query:
            queryset = queryset.filter(Q(full_name__icontains=search_query))
        if personnel_code:
            queryset = queryset.filter(personnel_code=personnel_code)
        if expertise_filter:
            queryset = queryset.filter(expertises__in=expertise_filter)
        
        return queryset
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        context['selected_expertises'] = [int(e) for e in self.request.GET.getlist('expertise')]
        return context

class DoctorCreateView(LoginRequiredMixin, ManagerRequiredMixin, CreateView):
    model = Doctor
    form_class = DoctorForm
    template_name = 'doctor_create.html'

    def get_success_url(self):
        return reverse('doctor_detail', kwargs={'pk': self.object.id})

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['section_list'] = Section.objects.filter(group=self.request.user.group)
        context['room_list'] = Room.objects.filter(group=self.request.user.group)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        return context
    
    def form_valid(self, form):
        form.instance.group = self.request.user.group
        return super().form_valid(form)

class DoctorUpdateView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, UpdateView):
    model = Doctor
    form_class = DoctorForm
    template_name = 'doctor_update.html'

    def get_success_url(self):
        return reverse('doctor_detail', kwargs={'pk': self.object.id})

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['section_list'] = Section.objects.filter(group=self.request.user.group)
        context['room_list'] = Room.objects.filter(group=self.request.user.group)
        context['expertise_list'] = Expertise.objects.filter(group=self.request.user.group)
        return context

class DoctorDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, DeleteView):
    model = Doctor
    template_name = 'doctor_confirm_delete.html'
    success_url = reverse_lazy('doctor_list')

class PatientListView(LoginRequiredMixin, ManagerRequiredMixin, ListView):
    model = Patient
    template_name = 'patient_list.html'
    context_object_name = 'patients'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        search_query = self.request.GET.get('q')
        section_filter = self.request.GET.get('section')
        room_filter = self.request.GET.get('room')

        if search_query:
            queryset = queryset.filter(Q(full_name__icontains=search_query))
        if section_filter:
            queryset = queryset.filter(sections__in=section_filter)
        if room_filter:
            queryset = queryset.filter(rooms__in=room_filter)
        
        return queryset

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        selected_section = self.request.GET.get('section')
        selected_room = self.request.GET.get('room')
        context['section_list'] = Section.objects.filter(group=self.request.user.group)
        context['room_list'] = Room.objects.filter(group=self.request.user.group)
        context['selected_section'] = int(selected_section) if selected_section and selected_section.isdigit() else None
        context['selected_room'] = int(selected_room) if selected_room and selected_room.isdigit() else None
        return context

class PatientDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin ,DeleteView):
    model = Patient
    template_name = 'patient_confirm_delete.html'
    success_url = reverse_lazy('patient_list')

class SectionCaseListView(LoginRequiredMixin, ListView):
    model = SectionCase
    template_name = 'section_case_list.html'
    context_object_name = 'section_cases'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        number = self.request.GET.get('number')
        search_query_admission_date = self.request.GET.get('admission_date')
        search_query_doctor = self.request.GET.get('doctor')
        search_query_section = self.request.GET.get('section')

        if number:
            queryset = queryset.filter(Q(number__icontains=number))
        if search_query_admission_date:
            queryset = queryset.filter(admission_date=search_query_admission_date)
        if search_query_doctor:
            queryset = queryset.filter(doctor=search_query_doctor)
        if search_query_section:
            queryset = queryset.filter(section=search_query_section)
        
        return queryset
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        selected_section = self.request.GET.get('section')
        selected_doctor = self.request.GET.get('doctor')
        context['section_list'] = Section.objects.filter(group=self.request.user.group)
        context['doctor_list'] = Doctor.objects.filter(group=self.request.user.group)
        context['selected_section'] = int(selected_section) if selected_section and selected_section.isdigit() else None
        context['selected_doctor'] = int(selected_doctor) if selected_doctor and selected_doctor.isdigit() else None
        return context

@login_required
@group_is_owner(SectionCase, lookup_field='pk', group_field='group')
def section_case_detail(request, pk):
    section_case = get_object_or_404(SectionCase, pk=pk)
    d_date = section_case.discharge_date
    dl_date = section_case.delivery_date
    ad_date = section_case.admission_date
    arrive_day = '----'
    stay_day = '----'

    if isinstance(d_date, str) and isinstance(dl_date, str):
        try:
            discharge_date = Persian(d_date).gregorian_datetime()
            delivery_date = Persian(dl_date).gregorian_datetime()

            arrive_day = (delivery_date - discharge_date).days
        except:
            pass
        
    if isinstance(d_date, str) and isinstance(ad_date, str):
        try:
            discharge_date = Persian(d_date).gregorian_datetime()
            admission_date = Persian(ad_date).gregorian_datetime()

            stay_day = (discharge_date - admission_date).days
        except:
            pass
    
    context = {
        'section_case': section_case,
        'arrive_day': arrive_day,
        'stay_day': stay_day,
    }

    return render(request, 'section_case_detail.html', context=context)

class SectionCaseUpdateView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, UpdateView):
    model = SectionCase
    form_class = SectionCaseForm
    template_name = 'section_case_update.html'

    def get_success_url(self):
        return reverse('section_case_detail', kwargs={'pk': self.object.id})

    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        context['defect_sheet_choices'] = defect_sheet_choices
        context['defect_type_choices'] = defect_type_choices
        return context

class SectionCaseDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, DeleteView):
    model = SectionCase
    template_name = 'section_case_confirm_delete.html'
    success_url = reverse_lazy('section_case_list')

class RoomCaseListView(LoginRequiredMixin, ListView):
    model = RoomCase
    template_name = 'room_case_list.html'
    context_object_name = 'room_cases'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        number = self.request.GET.get('number')
        search_query_admission_date = self.request.GET.get('operation_date')
        search_query_doctor = self.request.GET.get('doctor')
        search_query_section = self.request.GET.get('room')

        if number:
            queryset = queryset.filter(Q(number__icontains=number))
        if search_query_admission_date:
            queryset = queryset.filter(admission_date=search_query_admission_date)
        if search_query_doctor:
            queryset = queryset.filter(doctor=search_query_doctor)
        if search_query_section:
            queryset = queryset.filter(room=search_query_section)
        
        return queryset
    
    def get_context_data(self, **kwargs):
        context = super().get_context_data(**kwargs)
        selected_room = self.request.GET.get('room')
        selected_doctor = self.request.GET.get('doctor')
        context['room_list'] = Room.objects.filter(group=self.request.user.group)
        context['doctor_list'] = Doctor.objects.filter(group=self.request.user.group)
        context['selected_room'] = int(selected_room) if selected_room and selected_room.isdigit() else None
        context['selected_doctor'] = int(selected_doctor) if selected_doctor and selected_doctor.isdigit() else None
        return context
    
class RoomCaseDetailView(LoginRequiredMixin, UserIsOwnerMixin, DetailView):
    model = RoomCase
    template_name = 'room_case_detail.html'
    context_object_name = 'room_case'

class RoomCaseDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, DeleteView):
    model = RoomCase
    template_name = 'room_case_confirm_delete.html'
    success_url = reverse_lazy('room_case_list')

class DCListView(LoginRequiredMixin, ListView):
    model = DC
    template_name = 'dc_list.html'
    context_object_name = 'dc_cases'
    ordering = ['-id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        return queryset

class DCDetailView(LoginRequiredMixin, UserIsOwnerMixin, DetailView):
    model = DC
    template_name = 'dc_detail.html'
    context_object_name = 'dc'

class DCDeleteView(LoginRequiredMixin, ManagerRequiredMixin, UserIsOwnerMixin, DeleteView):
    model = DC
    template_name = 'dc_confirm_delete.html'
    success_url = reverse_lazy('dc_list')

@login_required
@manager_required
def add_section_case(request):
    if request.method == 'POST':
        form = ExcelForm(request.POST, request.FILES)
        if form.is_valid():
            excel_instance = form.save(commit=False)
            excel_instance.group = request.user.group
            excel_instance.save()

            file_path = excel_instance.file.path
            excel_file = pd.ExcelFile(file_path)

            section_sheets = excel_file.sheet_names[:1]
            section_values = {}
            doctor_data = {}
            patient_data = {}

            for sheet in section_sheets:
                try:
                    # پاک سازی و اصلاح جداول شیت
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                    # ردیفی که ستون 9 آن دارای حرف P است به عنوان پرونده درنظر گرفته میشه
                    if df.shape[1] >= 9:
                        df = df[df.iloc[:, 8].astype(str).str.contains("P", na=False)]
                    
                    # استخراج نام بخش ها از ستون سوم هر شیت
                    if df.shape[1] >= 3:
                        third_col_values = df.iloc[:, 2].dropna().astype(str).str.strip()
                        for value in third_col_values:
                            if value in section_values:
                                section_values[value].add(sheet)
                            else:
                                section_values[value] = {sheet}
                        
                    # استخراج نام پزشکان از ستون های 4 و 8
                    for index, row in df.iterrows():
                        if df.shape[1] >= 4 and df.shape[1] >= 9:
                            doctor_name_4 = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else None
                            doctor_name_8 = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else None
                            related_value = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else None

                        for doctor_name in [doctor_name_4, doctor_name_8]:
                            if doctor_name:
                                if doctor_name in doctor_data:
                                    doctor_data[doctor_name].add(related_value)
                                else:
                                    doctor_data[doctor_name] = {related_value}
                        
                    # استخراج نام بیماران از ستون نهم
                    for index, row in df.iterrows():
                        if df.shape[1] >= 9:
                            col9_value = str(row.iloc[8]).strip() if pd.notna(row.iloc[8]) else None
                            related_value = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else None

                            if col9_value:
                                if col9_value in patient_data:
                                    patient_data[col9_value].add(related_value)
                                else:
                                    patient_data[col9_value] = {related_value}
                except Exception as e:
                    print(f"خطا در شیت '{sheet}': {e}")
            
            # اتصال بخش های به دست آمده به نام شیت آنها
            for value, sheets in section_values.items():
                for sheet in sheets:
                    section, created = Section.objects.get_or_create(group=request.user.group, name=value, sheet='')

            # اتصال پزشکان به بخش ها و اتاق های آنها
            for full_name, sheets in doctor_data.items():
                doctor, created = Doctor.objects.get_or_create(group=request.user.group, full_name=full_name)

                for sheet in sheets:
                    section = Section.objects.filter(group=request.user.group, name=sheet).first()
                    if section not in doctor.sections.all():
                        doctor.sections.add(section)
                        doctor.save()

            # اتصال بیماران به بخش ها و اتاق های آنها
            for full_name, sheets in patient_data.items():
                patient, created = Patient.objects.get_or_create(group=request.user.group, full_name=full_name)

                for sheet in sheets:
                    section = Section.objects.filter(group=request.user.group, name=sheet).first()
                    if section not in doctor.sections.all():
                        patient.sections.add(section)
                        patient.save()
            
            for sheet in section_sheets:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                    if df.shape[1] >= 9:
                        df = df[df.iloc[:, 8].astype(str).str.contains("P", na=False)]

                    if df.shape[1] >= 14:
                        for i in range(len(df)):
                            try:
                                insurance = str(df.iloc[i, 0]).strip()
                                discharge_date = str(df.iloc[i, 1]).strip()
                                section_name = str(df.iloc[i, 2]).strip()
                                doctor_name = str(df.iloc[i, 3]).strip()
                                admission_date = str(df.iloc[i, 4]).strip()
                                number = str(df.iloc[i, 6]).strip()
                                rep_doctor_name = str(df.iloc[i, 7]).strip()
                                patient_name = str(df.iloc[i, 8]).strip()
                                delivery_date = str(df.iloc[i, 9]).strip()
                                defect_sheet = str(df.iloc[i, 10]).strip()
                                defect_type = str(df.iloc[i, 11]).strip()
                                defect_sheet2 = str(df.iloc[i, 12]).strip()
                                defect_type2 = str(df.iloc[i, 13]).strip()

                                section = Section.objects.filter(group=request.user.group, name=section_name).first()
                                doctor = Doctor.objects.filter(group=request.user.group, full_name=doctor_name).first()
                                rep_doctor = Doctor.objects.filter(group=request.user.group, full_name=rep_doctor_name).first()
                                patient = Patient.objects.filter(group=request.user.group, full_name=patient_name).first()
                                defect_sheet = defect_sheet_map.get(defect_sheet, None)
                                defect_type = defect_type_map.get(defect_type, None)
                                defect_sheet2 = defect_sheet_map.get(defect_sheet2, None)
                                defect_type2 = defect_type_map.get(defect_type2, None)

                                SectionCase.objects.create(
                                    group=request.user.group,
                                    insurance=insurance,
                                    discharge_date=discharge_date,
                                    section=section,
                                    doctor=doctor,
                                    admission_date=admission_date,
                                    number=number,
                                    representative_doctor=rep_doctor,
                                    patient=patient,
                                    delivery_date=delivery_date,
                                    defect_sheet=defect_sheet,
                                    defect_type=defect_type,
                                    defect_sheet2=defect_sheet2,
                                    defect_type2=defect_type2
                                )
                            except Exception as row_error:
                                print(f" خطا در ردیف {i} از شیت {sheet}: {row_error}")
                except Exception as e:
                    print(f" خطا در شیت '{sheet}': {e}")

            return redirect('main')
    else:
        form = ExcelForm()
    
    context = {
        'form': form,
    }

    return render(request, 'add_section_case.html', context=context)

@login_required
@manager_required
def add_room_case(request):
    if request.method == 'POST':
        form = ExcelForm(request.POST, request.FILES)
        if form.is_valid():
            excel_instance = form.save(commit=False)
            excel_instance.group = request.user.group
            excel_instance.save()

            file_path = excel_instance.file.path
            excel_file = pd.ExcelFile(file_path)

            room_sheets = excel_file.sheet_names[:1]
            room_values = {}
            doctor_data = {}
            patient_data = {}

            # اعمال کار های مشابه در شیت های اتاق عمل
            # نام اتاق از ستون 7
            # نام پزشک از ستون 10
            # نام بیمار از ترکیب ستون های نام بیمار و شناسه آن
            for sheet in room_sheets:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                    if df.shape[1] >= 7:
                        seventh_col_values = df.iloc[:, 6].dropna().astype(str).str.strip()
                        for value in seventh_col_values:
                            if value in room_values:
                                room_values[value].add(sheet)
                            else:
                                room_values[value] = {sheet}
                        
                    for index, row in df.iterrows():
                        if df.shape[1] >= 10:
                            col10_value = str(row.iloc[9]).strip() if pd.notna(row.iloc[9]) else None
                            related_value = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else None

                        if col10_value:
                            if col10_value in doctor_data:
                                doctor_data[col10_value].add(related_value)
                            else:
                                doctor_data[col10_value] = {related_value}

                    if 'شناسه بیمار' in df.columns and 'نام بیمار' in df.columns:
                        combined_values = df[['شناسه بیمار', 'نام بیمار', df.columns[2]]].dropna().astype(str).apply(
                            lambda row: row['شناسه بیمار'].strip() + ' ' + row['نام بیمار'].strip(), axis=1)
    
                        related_values = df.iloc[:, 6].dropna().astype(str).str.strip()

                        for index, value in enumerate(combined_values):
                            related_value = related_values.iloc[index] if index < len(related_values) else None
                            if value and related_value:
                                if value in patient_data:
                                    patient_data[value].add(related_value)
                                else:
                                    patient_data[value] = {related_value}
                except Exception as e:
                    print(f"خطا در شیت '{sheet}': {e}")
            
            # اتصال اتاق های به دست آمده به نام شیت آنها
            for value, sheets in room_values.items():
                for sheet in sheets:
                    room, created = Room.objects.get_or_create(group=request.user.group, name=value, sheet='')
            
            # اتصال پزشکان به بخش ها و اتاق های آنها
            for full_name, sheets in doctor_data.items():
                doctor, created = Doctor.objects.get_or_create(group=request.user.group, full_name=full_name)

                for sheet in sheets:
                    room = Room.objects.filter(group=request.user.group, name=sheet).first()
                    if room not in doctor.rooms.all():
                        doctor.rooms.add(room)
                        doctor.save()

            # اتصال بیماران به بخش ها و اتاق های آنها
            for full_name, sheets in patient_data.items():
                patient, created = Patient.objects.get_or_create(group=request.user.group, full_name=full_name)

                for sheet in sheets:
                    room = Room.objects.filter(group=request.user.group, name=sheet).first()
                    if room not in patient.rooms.all():
                        patient.rooms.add(room)
                        patient.save()
            
            # برداشت ردیف های شیت های اتاق عمل به عنوان پرونده اتاق عمل
            for sheet in room_sheets:
                try:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]

                    if df.shape[1] >= 11:
                        for i in range(len(df)):
                            try:
                                hospitalization_date = str(df.iloc[i, 0]).strip()
                                discharge_date = str(df.iloc[i, 1]).strip()
                                operation_date = str(df.iloc[i, 2]).strip()
                                patient_name = str(df.iloc[i, 3]).strip()
                                number = str(df.iloc[i, 5]).strip()
                                room_name = str(df.iloc[i, 6]).strip()
                                operation_type = str(df.iloc[i, 7]).strip()
                                k = str(df.iloc[i, 8]).strip()
                                doctor_name = str(df.iloc[i, 9]).strip()
                                anesthesia_type = str(df.iloc[i, 10]).strip()

                                patient = Patient.objects.filter(group=request.user.group, full_name__icontains=patient_name).first()
                                room = Room.objects.filter(group=request.user.group, name=room_name).first()
                                operation_type = operation_type_dict.get(operation_type, None)
                                doctor = Doctor.objects.filter(group=request.user.group, full_name=doctor_name).first()

                                RoomCase.objects.create(
                                    group=request.user.group,
                                    hospitalization_date=hospitalization_date,
                                    discharge_date=discharge_date,
                                    operation_date=operation_date,
                                    patient=patient,
                                    number=number,
                                    room=room,
                                    operation_type=operation_type,
                                    k=k,
                                    doctor=doctor,
                                    anesthesia_type=anesthesia_type,
                                )
                            except Exception as row_error:
                                print(f" خطا در ردیف {i} از شیت {sheet}: {row_error}")
                except Exception as e:
                    print(f" خطا در شیت '{sheet}': {e}")
            
            return redirect('main')
    else:
        form = ExcelForm()
    
    context = {
        'form': form,
    }

    return render(request, 'add_room_case.html', context=context)

@login_required
@manager_required
def all_delete(request):
    form = ConfirmDeleteForm(request.POST or None)

    if request.method == 'POST':
        if form.is_valid() and form.cleaned_data['confirm']:
            group = request.user.group

            # حذف داده‌های مرتبط با گروه
            models = [Excel, Expertise, Section, Room, Doctor, Patient, SectionCase, RoomCase, DC]
            for model in models:
                model.objects.filter(group=group).delete()

            return redirect('main')

    return render(request, 'all_delete_confirm.html', {'form': form})
