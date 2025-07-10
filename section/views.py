import pandas as pd
from functools import wraps
from collections import Counter, defaultdict
from django.shortcuts import render, redirect, get_object_or_404
from django.urls import reverse, reverse_lazy
from django.contrib.auth import authenticate, login
from django.contrib.auth.decorators import login_required
from django.contrib.auth.mixins import LoginRequiredMixin
from django.contrib import messages
from django.core.exceptions import PermissionDenied
from django.core.paginator import Paginator
from django.db.models import Q
from django.views.generic import ListView, CreateView, DetailView, UpdateView, DeleteView
from .models import Excel, Expertise, Section, Room, Doctor, Patient, SectionCase, RoomCase, DC
from .forms import (
    CustomUserCreationForm, LoginForm, 
    ExcelForm, ExpertiseForm, SectionForm, RoomForm, DoctorForm, SectionCaseForm, ConfirmDeleteForm,
    MultiSectionForm, MultiRoomForm, MultiDoctorForm
)
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
    group = request.user.group

    if Excel.objects.filter(group=group).exists():
        # آمار اولیه
        doctors_count = Doctor.objects.filter(group=group).count()
        patients_count = Patient.objects.filter(group=group).count()
        sections = Section.objects.filter(group=group)
        rooms = Room.objects.filter(group=group)
        sections_count = sections.count()
        rooms_count = rooms.count()

        section_cases = SectionCase.objects.filter(group=group)
        section_cases_count = section_cases.count()
        room_cases = RoomCase.objects.filter(group=group)
        room_cases_count = room_cases.count()
        cases_count = section_cases_count + room_cases_count

        defect_cases = section_cases.filter(
            Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False) |
            Q(defect_sheet3__isnull=False) | Q(defect_sheet4__isnull=False) |
            Q(defect_sheet5__isnull=False) | Q(defect_sheet6__isnull=False) |
            Q(defect_sheet7__isnull=False) | Q(defect_sheet8__isnull=False) |
            Q(defect_sheet9__isnull=False) | Q(defect_sheet10__isnull=False)
        )
        defect_section_cases_count = defect_cases.count()

        # آمار بیمه‌ها
        def count_insurance(keyword):
            return section_cases.filter(insurance__icontains=keyword).count()

        social_security_cases = count_insurance("تامین اجتماعی")
        medical_services_cases = count_insurance("خدمات درمانی")
        armed_forces_cases = count_insurance("نیرو های مسلح")
        free_cases = count_insurance("آزاد")

        # آخرین اتاق‌ها و بخش‌ها
        sections_recent = sections.order_by('id')[:3]
        rooms_recent = rooms.order_by('id')[:3]

        # تحلیل پزشکان اتاق عمل
        doctor_room_list = []
        doctor_bigroom_list = []
        doctor_mediumroom_list = []
        doctor_smallroom_list = []

        for case in room_cases.select_related('doctor'):
            full_name = case.doctor.full_name
            doctor_room_list.append(full_name)
            if case.operation_type == '1':
                doctor_smallroom_list.append(full_name)
            elif case.operation_type == '2':
                doctor_mediumroom_list.append(full_name)
            elif case.operation_type == '3':
                doctor_bigroom_list.append(full_name)

        def most_common_or_none(lst):
            return Counter(lst).most_common(1)

        # آماده‌سازی داده نقص
        defect_counts = {}
        defect_percents = {}
        defect_type_counts = {}
        defect_type_percents = {}

        filtered_section_cases = []

        if request.GET.get("start") and request.GET.get("end"):
            try:
                start = Persian(request.GET["start"]).gregorian_datetime()
                end = Persian(request.GET["end"]).gregorian_datetime()
                if start > end:
                    start, end = end, start

                for case in section_cases:
                    if isinstance(case.admission_date, str):
                        try:
                            case_date = Persian(case.admission_date).gregorian_datetime()
                            if start <= case_date <= end:
                                filtered_section_cases.append(case)
                        except:
                            continue

                # شمارش نقص‌ها در بازه زمانی
                def count_defects(
                        cases, 
                        field1, field2, field3, field4, field5, field6, field7, field8, field9, field10,
                        choices
                ):
                    counts, percents = {}, {}
                    for code, name in choices:
                        count = sum(1 for c in cases if getattr(c, field1) == code or getattr(c, field2) == code or getattr(c, field3) == code or getattr(c, field4) == code or getattr(c, field5) == code or getattr(c, field6) == code or getattr(c, field7) == code or getattr(c, field8) == code or getattr(c, field9) == code or getattr(c, field10) == code)
                        total = sum(1 for c in cases if getattr(c, field1) or getattr(c, field2))
                        counts[name] = count
                        percents[name] = round((count * 100 / total), 0) if total else 0
                    return counts, percents
                
                def count_multiselect_defects(cases, fields, choices):
                    counts, percents = {}, {}
                    total = sum(
                        1 for c in cases if any(getattr(c, f, None) for f in fields)
                    )

                    for code, name in choices:
                        count = sum(
                            1 for c in cases for f in fields
                            if code in (getattr(c, f, []) or [])
                        )
                        counts[name] = count
                        percents[name] = round((count * 100 / total), 0) if total else 0

                    return counts, percents

                defect_counts, defect_percents = count_defects(
                    filtered_section_cases, 
                    'defect_sheet', 'defect_sheet2', 'defect_sheet3', 'defect_sheet4', 'defect_sheet5', 'defect_sheet6', 'defect_sheet7', 'defect_sheet8', 'defect_sheet9', 'defect_sheet10', 
                    defect_sheet_choices
                )

                fields = [
                    'defect_type', 'defect_type2', 'defect_type3',
                    'defect_type4', 'defect_type5', 'defect_type6',
                    'defect_type7', 'defect_type8', 'defect_type9', 'defect_type10'
                ]

                defect_type_counts, defect_type_percents = count_multiselect_defects(
                    filtered_section_cases,
                    fields,
                    defect_type_choices
                )

            except:
                pass
        else:
            def count_global_defects(
                    field1, field2, field3, field4, field5, field6, field7, field8, field9, field10,
                    choices, filter_field
            ):
                counts, percents = {}, {}
                total = defect_cases.count()
                for code, name in choices:
                    count = section_cases.filter(**{field1: code}).count() + \
                            section_cases.filter(**{field2: code}).count() + \
                            section_cases.filter(**{field3: code}).count() + \
                            section_cases.filter(**{field4: code}).count() + \
                            section_cases.filter(**{field5: code}).count() + \
                            section_cases.filter(**{field6: code}).count() + \
                            section_cases.filter(**{field7: code}).count() + \
                            section_cases.filter(**{field8: code}).count() + \
                            section_cases.filter(**{field9: code}).count() + \
                            section_cases.filter(**{field10: code}).count()
                    counts[name] = count
                    percents[name] = round((count * 100 / total), 0) if total else 0
                return counts, percents
            
            def count_multiselect_defects(cases, fields, choices):
                counts, percents = {}, {}
                total = sum(
                    1 for c in cases if any(getattr(c, f, None) for f in fields)
                )

                for code, name in choices:
                    count = sum(
                        1 for c in cases for f in fields
                        if code in (getattr(c, f, []) or [])
                    )
                    counts[name] = count
                    percents[name] = round((count * 100 / total), 0) if total else 0

                return counts, percents

            defect_counts, defect_percents = count_global_defects(
                'defect_sheet', 'defect_sheet2', 'defect_sheet3', 'defect_sheet4', 'defect_sheet5', 'defect_sheet6', 'defect_sheet7', 'defect_sheet8', 'defect_sheet9', 'defect_sheet10', 
                defect_sheet_choices, 'defect_sheet'
            )

            fields = [
                'defect_type', 'defect_type2', 'defect_type3',
                'defect_type4', 'defect_type5', 'defect_type6',
                'defect_type7', 'defect_type8', 'defect_type9', 'defect_type10'
            ]

            defect_type_counts, defect_type_percents = count_multiselect_defects(
                section_cases,
                fields,
                defect_type_choices
            )

        context = {
            'doctors_count': doctors_count,
            'patients_count': patients_count,
            'sections_count': sections_count,
            'rooms_count': rooms_count,
            'cases_count': cases_count,
            'section_cases_count': section_cases_count,
            'defect_section_cases_count': defect_section_cases_count,
            'social_security_cases': social_security_cases,
            'medical_services_cases': medical_services_cases,
            'armed_forces_cases': armed_forces_cases,
            'free_cases': free_cases,
            'room_cases_count': room_cases_count,
            'sections': sections_recent,
            'rooms': rooms_recent,
            'defect_counts': defect_counts,
            'defect_type_counts': defect_type_counts,
            'defect_percents': defect_percents,
            'defect_type_percents': defect_type_percents,
            'most_doctor_room_list': most_common_or_none(doctor_room_list),
            'most_doctor_bigroom_list': most_common_or_none(doctor_bigroom_list),
            'most_doctor_mediumroom_list': most_common_or_none(doctor_mediumroom_list),
            'most_doctor_smallroom_list': most_common_or_none(doctor_smallroom_list),
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
                                    defect_type_raw = df.iloc[i, 11]
                                    defect_sheet2 = str(df.iloc[i, 12]).strip()
                                    defect_type2_raw = df.iloc[i, 13]

                                    section = Section.objects.filter(group=request.user.group, name=section_name).first()
                                    doctor = Doctor.objects.filter(group=request.user.group, full_name=doctor_name).first()
                                    rep_doctor = Doctor.objects.filter(group=request.user.group, full_name=rep_doctor_name).first()
                                    patient = Patient.objects.filter(group=request.user.group, full_name=patient_name).first()
                                    defect_sheet = defect_sheet_map.get(defect_sheet, None)
                                    defect_sheet2 = defect_sheet_map.get(defect_sheet2, None)
                                    defect_type_list = defect_type_raw if isinstance(defect_type_raw, list) else [defect_type_raw]
                                    defect_type2_list = defect_type2_raw if isinstance(defect_type2_raw, list) else [defect_type2_raw]

                                    defect_type_clean = [
                                        defect_type_map.get(str(dt).strip(), None)
                                        for dt in defect_type_list if dt is not None
                                ]

                                    defect_type2_clean = [
                                        defect_type_map.get(str(dt).strip(), None)
                                        for dt in defect_type2_list if dt is not None
                                    ]

                                    if defect_type_clean:
                                        defect_type_summary = ', '.join([str(dt) for dt in defect_type_clean if dt is not None])
                                    else:
                                        defect_type_summary = ''

                                    if defect_type2_clean:
                                        defect_type2_summary = ', '.join([str(dt) for dt in defect_type2_clean if dt is not None])
                                    else:
                                        defect_type2_summary = ''

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
                                        defect_type=defect_type_summary,
                                        defect_sheet2=defect_sheet2,
                                        defect_type2=defect_type2_summary
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
    user_group = request.user.group
    section_cases_all = SectionCase.objects.filter(group=user_group, section=section)
    dc_section_cases_all = DC.objects.filter(group=user_group, hospitalization_section=section)
    doctors = section.doctor_sections.all()

    def to_date(date_str):
        try:
            return Persian(date_str).gregorian_datetime() if isinstance(date_str, str) else None
        except:
            return None

    start, end = None, None
    if request.GET.get("start") and request.GET.get("end"):
        start = to_date(request.GET["start"])
        end = to_date(request.GET["end"])
        if start and end and start > end:
            start, end = end, start

    def in_range(date):
        return start <= date <= end if start and end and date else not start and not end

    # فیلتر کردن بر اساس تاریخ پذیرش
    section_cases = [sc for sc in section_cases_all if in_range(to_date(sc.admission_date))]
    dc_section_cases = [dc for dc in dc_section_cases_all if in_range(to_date(dc.admission_date))]
    not_arrived_cases = [
        sc for sc in section_cases_all
        if sc.delivery_date == 'nan' and in_range(to_date(sc.admission_date))
    ]

    defect_sheet_fields = ['defect_sheet'] + [f'defect_sheet{i}' for i in range(2, 11)]
    defect_cases = [
        sc for sc in section_cases
        if any(getattr(sc, field) for field in defect_sheet_fields)
    ]

    # فیلتر بر اساس بیمه
    def filter_insurance(keyword):
        return [sc for sc in section_cases if keyword in (sc.insurance or '')]

    insurance_filters = {
        'filtered_social_security_cases': 'تامین اجتماعی',
        'filtered_medical_services_cases': 'خدمات درمانی',
        'filtered_armed_forces_cases': 'نیرو های مسلح',
        'filtered_free_cases': 'آزاد'
    }
    insurance_results = {key: filter_insurance(val) for key, val in insurance_filters.items()}

    # آمار پزشکان
    doctor_cases = defaultdict(int)
    doctor_defects = defaultdict(int)
    for doctor in doctors:
        for sc in section_cases_all:
            if sc.doctor == doctor and in_range(to_date(sc.admission_date)):
                doctor_cases[doctor.full_name] += 1
                if sc.defect_sheet or sc.defect_sheet2:
                    doctor_defects[doctor.full_name] += 1

    # محاسبه متوسط اقامت و تحویل
    arrive_days, stay_days = [], []
    for sc in section_cases:
        discharge = to_date(sc.discharge_date)
        delivery = to_date(sc.delivery_date)
        admission = to_date(sc.admission_date)
        if discharge and delivery:
            arrive_days.append((delivery - discharge).days)
        if discharge and admission:
            stay_days.append((discharge - admission).days)

    avg_arrive = round(sum(arrive_days) / len(arrive_days), 0) if arrive_days else '0'
    avg_stay = round(sum(stay_days) / len(stay_days), 0) if stay_days else '0'

    # آمار نقص‌ها
    defect_sheet_fields = ['defect_sheet'] + [f'defect_sheet{i}' for i in range(2, 11)]
    defect_type_fields = ['defect_type'] + [f'defect_type{i}' for i in range(2, 11)]

    defect_counts = {
        name: sum([
            1 for sc in section_cases
            if any(getattr(sc, field) == code for field in defect_sheet_fields)
        ]) for code, name in defect_sheet_choices
    }

    defect_type_counts = {
        name: sum([
            1 for sc in section_cases
            if any(
                getattr(sc, field) and code in getattr(sc, field)
                for field in defect_type_fields
            )
        ]) for code, name in defect_type_choices
    }

    # آمار سن و جنسیت فوتی‌ها
    age_counts = {'less_20': 0, 'more_20_less_40': 0, 'more_40_less_60': 0, 'more_60_less_80': 0, 'more_80': 0}
    gender_counts = {'men': 0, 'women': 0}
    for dc in dc_section_cases:
        age = int(''.join(filter(str.isdigit, dc.age or '0')))
        if age < 20:
            age_counts['less_20'] += 1
        elif age < 40:
            age_counts['more_20_less_40'] += 1
        elif age < 60:
            age_counts['more_40_less_60'] += 1
        elif age < 80:
            age_counts['more_60_less_80'] += 1
        else:
            age_counts['more_80'] += 1
        gender_counts['men' if dc.gender == '1' else 'women'] += 1

    # Pagination helper
    def paginate(request, objects_list, per_page=100, name='page'):
        paginator = Paginator(objects_list, per_page)
        page_number = request.GET.get(name)
        return paginator.get_page(page_number)

    context = {
        'section': section,
        'doctors_count': len(set(sc.doctor for sc in section_cases if sc.doctor)),
        'patients_count': len(set(sc.patient for sc in section_cases if sc.patient)),
        'filtered_section_cases': paginate(request, section_cases, name='sc_page'),
        'filtered_section_cases_count': len(section_cases),
        'filtered_dc_section_cases': paginate(request, dc_section_cases, name='dc_page'),
        'filtered_dc_section_cases_count': len(dc_section_cases),
        'filtered_not_arrived_cases': paginate(request, not_arrived_cases, name='nc_page'),
        'filtered_not_arrived_cases_count': len(not_arrived_cases),
        'filtered_defect_cases': paginate(request, defect_cases, name='defc_page'),
        'filtered_defect_cases_count': len(defect_cases),
        **{key: val for key, val in insurance_results.items()},
        **{f'{key}_count': len(val) for key, val in insurance_results.items()},
        'average_arrive_daies': avg_arrive,
        'average_stay_daies': avg_stay,
        'defect_counts': defect_counts,
        'defect_type_counts': defect_type_counts,
        'doctor_cases': dict(doctor_cases),
        'doctor_defects': dict(doctor_defects),
        'age_counts': age_counts,
        'gender_counts': gender_counts,
    }
    return render(request, 'section_detail.html', context)

@login_required
@manager_required
@group_is_owner(Room, lookup_field='pk', group_field='group')
def room_detail(request, pk):
    room = get_object_or_404(Room, pk=pk)
    group = request.user.group

    room_cases = RoomCase.objects.filter(group=group, room=room)
    doctors = room.doctor_rooms.all()
    patients = set(room_cases.values_list('patient', flat=True))
    doctors_count = doctors.count()
    patients_count = len(patients)

    def to_gregorian(date_str):
        try:
            return Persian(date_str).gregorian_datetime()
        except:
            return None

    def filter_by_date(queryset):
        return [
            case for case in queryset
            if isinstance(case.operation_date, str)
            and (dt := to_gregorian(case.operation_date))
            and start <= dt <= end
        ]

    doctor_cases = {}
    filtered_room_cases = list(room_cases)
    filtered_big_room_cases = room_cases.filter(operation_type='3')
    filtered_medium_room_cases = room_cases.filter(operation_type='2')
    filtered_small_room_cases = room_cases.filter(operation_type='1')

    if request.GET.get("start") and request.GET.get("end"):
        try:
            start = to_gregorian(request.GET["start"])
            end = to_gregorian(request.GET["end"])
            if start and end and start > end:
                start, end = end, start

            filtered_room_cases = filter_by_date(room_cases)
            filtered_big_room_cases = filter_by_date(filtered_big_room_cases)
            filtered_medium_room_cases = filter_by_date(filtered_medium_room_cases)
            filtered_small_room_cases = filter_by_date(filtered_small_room_cases)

            doctors_count = len(set(c.doctor for c in filtered_room_cases))
            patients_count = len(set(c.patient for c in filtered_room_cases))

            for doctor in doctors:
                doctor_cases[doctor.full_name] = len([
                    c for c in filtered_room_cases if c.doctor == doctor
                ])
        except:
            pass
    else:
        for doctor in doctors:
            doctor_cases[doctor.full_name] = room_cases.filter(doctor=doctor).count()
    
    # Pagination helper
    def paginate(request, objects_list, per_page=100, name='page'):
        paginator = Paginator(objects_list, per_page)
        page_number = request.GET.get(name)
        return paginator.get_page(page_number)

    context = {
        'room': room,
        'doctors_count': doctors_count,
        'patients_count': patients_count,
        'filtered_room_cases': paginate(request, filtered_room_cases, name='rc'),
        'filtered_room_cases_count': len(filtered_room_cases),
        'filtered_big_room_cases_count': len(filtered_big_room_cases),
        'filtered_medium_room_cases_count': len(filtered_medium_room_cases),
        'filtered_small_room_cases_count': len(filtered_small_room_cases),
        'doctor_cases': doctor_cases,
    }

    return render(request, 'room_detail.html', context)

@login_required
@manager_required
@group_is_owner(Doctor, lookup_field='pk', group_field='group')
def doctor_detail(request, pk):
    doctor = get_object_or_404(Doctor, pk=pk)
    group = request.user.group

    def to_gregorian(date_str):
        try:
            return Persian(date_str).gregorian_datetime()
        except:
            return None

    def filter_by_date(queryset, date_field):
        result = []
        for obj in queryset:
            date_value = getattr(obj, date_field)
            if isinstance(date_value, str):
                case_date = to_gregorian(date_value)
                if case_date and start <= case_date <= end:
                    result.append(obj)
        return result

    section_cases = SectionCase.objects.filter(group=group, doctor=doctor)
    dc_section_cases = DC.objects.filter(group=group, doctor=doctor)
    room_cases = RoomCase.objects.filter(group=group, doctor=doctor)
    big_room_cases = room_cases.filter(operation_type='3')
    medium_room_cases = room_cases.filter(operation_type='2')
    small_room_cases = room_cases.filter(operation_type='1')

    not_arrived_cases = section_cases.filter(delivery_date='nan')

    defect_cases = section_cases.filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
    )
    all_defect_cases = SectionCase.objects.filter(group=group).filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False)
    )

    insurance_filter = lambda s: section_cases.filter(insurance__icontains=s)
    social_security_cases = insurance_filter("تامین اجتماعی")
    medical_services_cases = insurance_filter("خدمات درمانی")
    armed_forces_cases = insurance_filter("نیرو های مسلح")
    free_cases = insurance_filter("آزاد")

    patients = set(section_cases.values_list('patient', flat=True)) | set(room_cases.values_list('patient', flat=True))
    patients_count = len(patients)

    # فیلتر زمانی
    filtered_section_cases = section_cases
    filtered_dc_section_cases = dc_section_cases
    filtered_not_arrived_cases = not_arrived_cases
    filtered_defect_cases = defect_cases
    filtered_all_defect_cases = all_defect_cases
    filtered_social_security_cases = social_security_cases
    filtered_medical_services_cases = medical_services_cases
    filtered_armed_forces_cases = armed_forces_cases
    filtered_free_cases = free_cases
    filtered_room_cases = room_cases
    filtered_big_room_cases = big_room_cases
    filtered_medium_room_cases = medium_room_cases
    filtered_small_room_cases = small_room_cases

    if request.GET.get("start") and request.GET.get("end"):
        try:
            start = to_gregorian(request.GET["start"])
            end = to_gregorian(request.GET["end"])

            if start and end and start > end:
                start, end = end, start

            filtered_section_cases = filter_by_date(section_cases, 'admission_date')
            filtered_dc_section_cases = filter_by_date(dc_section_cases, 'admission_date')
            filtered_not_arrived_cases = filter_by_date(not_arrived_cases, 'admission_date')
            filtered_defect_cases = filter_by_date(defect_cases, 'admission_date')
            filtered_all_defect_cases = filter_by_date(all_defect_cases, 'admission_date')
            filtered_social_security_cases = filter_by_date(social_security_cases, 'admission_date')
            filtered_medical_services_cases = filter_by_date(medical_services_cases, 'admission_date')
            filtered_armed_forces_cases = filter_by_date(armed_forces_cases, 'admission_date')
            filtered_free_cases = filter_by_date(free_cases, 'admission_date')
            filtered_room_cases = filter_by_date(room_cases, 'operation_date')
            filtered_big_room_cases = filter_by_date(big_room_cases, 'operation_date')
            filtered_medium_room_cases = filter_by_date(medium_room_cases, 'operation_date')
            filtered_small_room_cases = filter_by_date(small_room_cases, 'operation_date')

            f_patients = set(c.patient for c in filtered_section_cases + filtered_room_cases)
            patients_count = len(f_patients)

        except Exception as e:
            print("Error filtering by date range:", e)

    def calc_average_days(cases, start_field, end_field):
        days = []
        for case in cases:
            start_date = getattr(case, start_field)
            end_date = getattr(case, end_field)
            if isinstance(start_date, str) and isinstance(end_date, str):
                start_dt = to_gregorian(start_date)
                end_dt = to_gregorian(end_date)
                if start_dt and end_dt:
                    days.append((end_dt - start_dt).days)
        return round(sum(days) / len(days), 0) if days else '0'

    average_arrive_daies = calc_average_days(filtered_section_cases, 'discharge_date', 'delivery_date')
    average_stay_daies = calc_average_days(filtered_section_cases, 'admission_date', 'discharge_date')

    # درصد نقص
    percent_defect_cases = (
        (len(filtered_defect_cases) * 100) // len(filtered_all_defect_cases)
        if filtered_all_defect_cases else '0'
    )

    # پراکندگی نقص
    defect_sheet_fields = ['defect_sheet'] + [f'defect_sheet{i}' for i in range(2, 11)]
    defect_type_fields = ['defect_type'] + [f'defect_type{i}' for i in range(2, 11)]

    defect_counts = {
        name: sum([
            1 for sc in section_cases
            if any(getattr(sc, field) == code for field in defect_sheet_fields)
        ]) for code, name in defect_sheet_choices
    }

    defect_type_counts = {
        name: sum([
            1 for sc in section_cases
            if any(
                getattr(sc, field) and code in getattr(sc, field)
                for field in defect_type_fields
            )
        ]) for code, name in defect_type_choices
    }

    # آمار فوت‌شدگان
    age_counts = {'less_20': 0, 'more_20_less_40': 0, 'more_40_less_60': 0, 'more_60_less_80': 0, 'more_80': 0}
    gender_counts = {'men': 0, 'women': 0}

    for case in filtered_dc_section_cases:
        age = int(''.join(filter(str.isdigit, case.age or '0')))
        gender = case.gender

        if age < 20:
            age_counts['less_20'] += 1
        elif age < 40:
            age_counts['more_20_less_40'] += 1
        elif age < 60:
            age_counts['more_40_less_60'] += 1
        elif age < 80:
            age_counts['more_60_less_80'] += 1
        else:
            age_counts['more_80'] += 1

        if gender == '1':
            gender_counts['men'] += 1
        else:
            gender_counts['women'] += 1
    
    # Pagination helper
    def paginate(request, objects_list, per_page=100, name='page'):
        paginator = Paginator(objects_list, per_page)
        page_number = request.GET.get(name)
        return paginator.get_page(page_number)

    context = {
        'doctor': doctor,
        'patients_count': patients_count,
        'filtered_section_cases': paginate(request, section_cases, name='sc_page'),
        'filtered_dc_section_cases': paginate(request, dc_section_cases, name='dc_page'),
        'filtered_not_arrived_cases': paginate(request, not_arrived_cases, name='nc_page'),
        'filtered_defect_cases': paginate(request, defect_cases, name='defc_page'),
        'filtered_dc_section_cases_count': len(filtered_dc_section_cases),
        'filtered_section_cases_count': len(filtered_section_cases),
        'filtered_not_arrived_cases_count': len(filtered_not_arrived_cases),
        'filtered_defect_cases_count': len(filtered_defect_cases),
        'filtered_social_security_cases_count': len(filtered_social_security_cases),
        'filtered_medical_services_cases_count': len(filtered_medical_services_cases),
        'filtered_armed_forces_cases_count': len(filtered_armed_forces_cases),
        'filtered_free_cases_count': len(filtered_free_cases),
        'filtered_room_cases_count': len(filtered_room_cases),
        'filtered_big_room_cases_count': len(filtered_big_room_cases),
        'filtered_medium_room_cases_count': len(filtered_medium_room_cases),
        'filtered_small_room_cases_count': len(filtered_small_room_cases),
        'average_arrive_daies': average_arrive_daies,
        'average_stay_daies': average_stay_daies,
        'defect_counts': defect_counts,
        'defect_type_counts': defect_type_counts,
        'percent_defect_cases': percent_defect_cases,
        'age_counts': age_counts,
        'gender_counts': gender_counts,
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

@login_required
@manager_required
def dc_all_detail(request):
    dc_cases = DC.objects.filter(group=request.user.group)
    dc_doctors = []
    dc_patients = []

    def to_gregorian(date_str):
        try:
            return Persian(date_str).gregorian_datetime()
        except:
            return None
    
    def filter_by_date(queryset, date_field):
        result = []
        for obj in queryset:
            date_value = getattr(obj, date_field)
            if isinstance(date_value, str):
                case_date = to_gregorian(date_value)
                if case_date and start <= case_date <= end:
                    result.append(obj)
        return result

    if request.GET.get("start") and request.GET.get("end"):
        try:
            start = to_gregorian(request.GET["start"])
            end = to_gregorian(request.GET["end"])

            if start and end and start > end:
                start, end = end, start
            
            dc_cases = filter_by_date(dc_cases, 'admission_date')

        except:
            pass

    def calc_average_days(cases, start_field, end_field):
        days = []
        for case in cases:
            start_date = getattr(case, start_field)
            end_date = getattr(case, end_field)
            if isinstance(start_date, str) and isinstance(end_date, str):
                start_dt = to_gregorian(start_date)
                end_dt = to_gregorian(end_date)
                if start_dt and end_dt:
                    days.append((end_dt - start_dt).days)
        return round(sum(days) / len(days), 0) if days else '0'

    for dc_case in dc_cases:
        dc_doctor = dc_case.doctor
        dc_patient = dc_case.patient

        if dc_doctor not in dc_doctors:
            dc_doctors.append(dc_doctor)

        if dc_patient not in dc_patients:
            dc_patients.append(dc_patient)

    average_arrive_daies = calc_average_days(dc_cases, 'death_date', 'delivery_date')
    average_stay_daies = calc_average_days(dc_cases, 'admission_date', 'death_date')

    # آمار فوت‌شدگان
    age_counts = {'less_20': 0, 'more_20_less_40': 0, 'more_40_less_60': 0, 'more_60_less_80': 0, 'more_80': 0}
    gender_counts = {'men': 0, 'women': 0}

    for case in dc_cases:
        age = int(''.join(filter(str.isdigit, case.age or '0')))
        gender = case.gender

        if age < 20:
            age_counts['less_20'] += 1
        elif age < 40:
            age_counts['more_20_less_40'] += 1
        elif age < 60:
            age_counts['more_40_less_60'] += 1
        elif age < 80:
            age_counts['more_60_less_80'] += 1
        else:
            age_counts['more_80'] += 1

        if gender == '1':
            gender_counts['men'] += 1
        else:
            gender_counts['women'] += 1
    
    # Pagination helper
    def paginate(request, objects_list, per_page=100, name='page'):
        paginator = Paginator(objects_list, per_page)
        page_number = request.GET.get(name)
        return paginator.get_page(page_number)

    context = {
        'filtered_dc_section_cases': paginate(request, dc_cases, name='sc_page'),
        'dc_cases_count': len(dc_cases),
        'dc_doctors_count': len(dc_doctors),
        'dc_patients_count': len(dc_patients),
        'average_arrive_daies': average_arrive_daies,
        'average_stay_daies': average_stay_daies,
        'age_counts': age_counts,
        'gender_counts': gender_counts,
    }

    return render(request, 'dc_all_detail.html', context)

class SectionListView(LoginRequiredMixin, ManagerRequiredMixin, ListView):
    model = Section
    template_name = 'section_list.html'
    context_object_name = 'sections'
    ordering = ['id']
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
    ordering = ['id']
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
    ordering = ['id']
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
    ordering = ['id']
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
    ordering = ['id']
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
    ordering = ['id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        start = self.request.GET.get('start')
        end = self.request.GET.get('end')
        number = self.request.GET.get('number')
        search_query_admission_date = self.request.GET.get('admission_date')
        search_query_doctor = self.request.GET.get('doctor')
        search_query_section = self.request.GET.get('section')
        search_query_patient = self.request.GET.get('patient')

        if start and end:
            start = Persian(start).gregorian_datetime()
            end = Persian(end).gregorian_datetime()

            filtered_queryset = []
            for section_case in queryset:
                if isinstance(section_case.admission_date, str):
                    try:
                        case_date = Persian(section_case.admission_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_queryset.append(section_case)
                    except:
                        continue

            queryset = filtered_queryset
        if number:
            queryset = queryset.filter(Q(number__icontains=number))
        if search_query_admission_date:
            queryset = queryset.filter(admission_date=search_query_admission_date)
        if search_query_doctor:
            queryset = queryset.filter(doctor=search_query_doctor)
        if search_query_section:
            queryset = queryset.filter(section=search_query_section)
        if search_query_patient:
            queryset = queryset.filter(patient__full_name=search_query_patient)
        
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
    ordering = ['id']
    paginate_by = 100

    def get_queryset(self):
        user_group = self.request.user.group
        queryset = super().get_queryset().filter(group=user_group).distinct()
        start = self.request.GET.get('start')
        end = self.request.GET.get('end')
        number = self.request.GET.get('number')
        search_query_admission_date = self.request.GET.get('operation_date')
        search_query_doctor = self.request.GET.get('doctor')
        search_query_section = self.request.GET.get('room')

        if start and end:
            start = Persian(start).gregorian_datetime()
            end = Persian(end).gregorian_datetime()

            filtered_queryset = []
            for section_case in queryset:
                if isinstance(section_case.operation_date, str):
                    try:
                        case_date = Persian(section_case.operation_date).gregorian_datetime()
                        if start <= case_date <= end:
                            filtered_queryset.append(section_case)
                    except:
                        continue

            queryset = filtered_queryset
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
    ordering = ['id']
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
                    if section not in patient.sections.all():
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
                                defect_type_raw = df.iloc[i, 11]
                                defect_sheet2 = str(df.iloc[i, 12]).strip()
                                defect_type2_raw = df.iloc[i, 13]

                                section = Section.objects.filter(group=request.user.group, name=section_name).first()
                                doctor = Doctor.objects.filter(group=request.user.group, full_name=doctor_name).first()
                                rep_doctor = Doctor.objects.filter(group=request.user.group, full_name=rep_doctor_name).first()
                                patient = Patient.objects.filter(group=request.user.group, full_name=patient_name).first()
                                defect_sheet = defect_sheet_map.get(defect_sheet, None)
                                defect_sheet2 = defect_sheet_map.get(defect_sheet2, None)
                                defect_type_list = defect_type_raw if isinstance(defect_type_raw, list) else [defect_type_raw]
                                defect_type2_list = defect_type2_raw if isinstance(defect_type2_raw, list) else [defect_type2_raw]

                                defect_type_clean = [
                                    defect_type_map.get(str(dt).strip(), None)
                                    for dt in defect_type_list if dt is not None
                                ]

                                defect_type2_clean = [
                                    defect_type_map.get(str(dt).strip(), None)
                                    for dt in defect_type2_list if dt is not None
                                ]

                                if defect_type_clean:
                                    defect_type_summary = ', '.join([str(dt) for dt in defect_type_clean if dt is not None])
                                else:
                                    defect_type_summary = ''

                                if defect_type2_clean:
                                    defect_type2_summary = ', '.join([str(dt) for dt in defect_type2_clean if dt is not None])
                                else:
                                    defect_type2_summary = ''

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
                                    defect_type=defect_type_summary,
                                    defect_sheet2=defect_sheet2,
                                    defect_type2=defect_type2_summary
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

def analyze_section(section, group, start=None, end=None):
    # اگه start و end وجود دارن، به میلادی تبدیل کن
    try:
        start_date = Persian(start).gregorian_datetime() if start else None
        end_date = Persian(end).gregorian_datetime() if end else None
    except:
        start_date = end_date = None

    def in_range(date_str):
        try:
            date = Persian(date_str).gregorian_datetime()
            if start_date and end_date:
                return start_date <= date <= end_date
            return True
        except:
            return False

    all_cases = SectionCase.objects.filter(group=group, section=section)
    dc_cases = DC.objects.filter(group=group, hospitalization_section=section)
    not_arrived = all_cases.filter(delivery_date='nan')
    defect_cases = all_cases.filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False) |
        Q(defect_sheet3__isnull=False) | Q(defect_sheet4__isnull=False) |
        Q(defect_sheet5__isnull=False) | Q(defect_sheet6__isnull=False) |
        Q(defect_sheet7__isnull=False) | Q(defect_sheet8__isnull=False) |
        Q(defect_sheet9__isnull=False) | Q(defect_sheet10__isnull=False)
    )

    insurance_filters = {
        'social_security': Q(insurance__icontains="تامین اجتماعی"),
        'medical_services': Q(insurance__icontains="خدمات درمانی"),
        'armed_forces': Q(insurance__icontains="نیرو های مسلح"),
        'free': Q(insurance__icontains="آزاد"),
    }

    filtered = lambda qs: [obj for obj in qs if in_range(obj.admission_date)] if start_date and end_date else list(qs)

    f_cases = filtered(all_cases)
    f_dc = filtered(dc_cases)
    f_not_arrived = filtered(not_arrived)
    f_defects = filtered(defect_cases)

    insurance_counts = {
        key: len(filtered(all_cases.filter(q)))
        for key, q in insurance_filters.items()
    }

    doctors = section.doctor_sections.all()
    f_doctors = {c.doctor for c in f_cases if c.doctor}
    f_patients = {c.patient for c in f_cases if c.patient}
    doctors_count = len(f_doctors)
    patients_count = len(f_patients)

    arrive_days = []
    stay_days = []
    for case in f_cases:
        try:
            if case.discharge_date and case.delivery_date:
                d = Persian(case.discharge_date).gregorian_datetime()
                dl = Persian(case.delivery_date).gregorian_datetime()
                arrive_days.append((dl - d).days)
            if case.admission_date and case.discharge_date:
                a = Persian(case.admission_date).gregorian_datetime()
                d = Persian(case.discharge_date).gregorian_datetime()
                stay_days.append((d - a).days)
        except:
            continue

    average_arrive = round(sum(arrive_days) / len(arrive_days), 0) if arrive_days else 0
    average_stay = round(sum(stay_days) / len(stay_days), 0) if stay_days else 0

    # آمار نقص‌ها
    defect_sheet_fields = ['defect_sheet'] + [f'defect_sheet{i}' for i in range(2, 11)]
    defect_type_fields = ['defect_type'] + [f'defect_type{i}' for i in range(2, 11)]

    defect_counts = {
        name: sum([
            1 for sc in f_cases
            if any(getattr(sc, field) == code for field in defect_sheet_fields)
        ]) for code, name in defect_sheet_choices
    }

    defect_type_counts = {
        name: sum([
            1 for sc in f_cases
            if any(
                getattr(sc, field) and code in getattr(sc, field)
                for field in defect_type_fields
            )
        ]) for code, name in defect_type_choices
    }

    doctor_cases = {}
    doctor_defects = {}
    for doc in doctors:
        cases = SectionCase.objects.filter(group=group, section=section, doctor=doc)
        cases = filtered(cases)
        doctor_cases[doc.full_name] = len(cases)

        defects = [c for c in cases if c.defect_sheet or c.defect_sheet2]
        doctor_defects[doc.full_name] = len(defects)

    age_buckets = {'less_20': 0, 'more_20_less_40': 0, 'more_40_less_60': 0, 'more_60_less_80': 0, 'more_80': 0}
    gender_counts = {'men': 0, 'women': 0}
    for c in f_dc:
        try:
            age = int(''.join(filter(str.isdigit, c.age or '0')))
            if age < 20:
                age_buckets['less_20'] += 1
            elif age < 40:
                age_buckets['more_20_less_40'] += 1
            elif age < 60:
                age_buckets['more_40_less_60'] += 1
            elif age < 80:
                age_buckets['more_60_less_80'] += 1
            else:
                age_buckets['more_80'] += 1

            if c.gender == '1':
                gender_counts['men'] += 1
            else:
                gender_counts['women'] += 1
        except:
            continue

    return {
        'section': section,
        'doctors_count': doctors_count,
        'patients_count': patients_count,
        'filtered_section_cases_count': len(f_cases),
        'filtered_dc_section_cases_count': len(f_dc),
        'filtered_not_arrived_cases_count': len(f_not_arrived),
        'filtered_defect_cases_count': len(f_defects),
        'filtered_social_security_cases_count': insurance_counts['social_security'],
        'filtered_medical_services_cases_count': insurance_counts['medical_services'],
        'filtered_armed_forces_cases_count': insurance_counts['armed_forces'],
        'filtered_free_cases_count': insurance_counts['free'],
        'average_arrive_daies': average_arrive,
        'average_stay_daies': average_stay,
        'defect_counts': defect_counts,
        'defect_type_counts': defect_type_counts,
        'doctor_cases': doctor_cases,
        'doctor_defects': doctor_defects,
        'age_counts': age_buckets,
        'gender_counts': gender_counts,
    }

def analyze_room(room, group, start=None, end=None):
    try:
        start_date = Persian(start).gregorian_datetime() if start else None
        end_date = Persian(end).gregorian_datetime() if end else None
    except:
        start_date = end_date = None

    def in_range(date_str):
        try:
            date = Persian(date_str).gregorian_datetime()
            if start_date and end_date:
                return start_date <= date <= end_date
            return True
        except:
            return False

    all_cases = RoomCase.objects.filter(group=group, room=room)
    big_cases = all_cases.filter(operation_type='3')
    medium_cases = all_cases.filter(operation_type='2')
    small_cases = all_cases.filter(operation_type='1')

    filtered = lambda qs: [obj for obj in qs if in_range(obj.operation_date)] if start_date and end_date else list(qs)

    f_cases = filtered(all_cases)
    f_big = filtered(big_cases)
    f_medium = filtered(medium_cases)
    f_small = filtered(small_cases)

    f_doctors = {c.doctor for c in f_cases if c.doctor}
    f_patients = {c.patient for c in f_cases if c.patient}

    doctor_cases = {}
    for doctor in room.doctor_rooms.all():
        cases = RoomCase.objects.filter(group=group, room=room, doctor=doctor)
        cases = filtered(cases)
        doctor_cases[doctor.full_name] = len(cases)

    return {
        'room': room,
        'doctors_count': len(f_doctors),
        'patients_count': len(f_patients),
        'filtered_room_cases_count': len(f_cases),
        'filtered_big_room_cases_count': len(f_big),
        'filtered_medium_room_cases_count': len(f_medium),
        'filtered_small_room_cases_count': len(f_small),
        'doctor_cases': doctor_cases,
    }

def analyze_doctor(doctor, group, start=None, end=None):
    try:
        start = Persian(start).gregorian_datetime() if start else None
        end = Persian(end).gregorian_datetime() if end else None
    except:
        start = end = None

    def to_gregorian(date_str):
        try:
            return Persian(date_str).gregorian_datetime()
        except:
            return None

    def in_range(date_str):
        date = to_gregorian(date_str)
        if date and start and end:
            return start <= date <= end
        return True

    def filter_by_date(qs, date_field):
        return [obj for obj in qs if in_range(getattr(obj, date_field))]

    def calc_average_days(cases, start_field, end_field):
        days = []
        for case in cases:
            start_date = getattr(case, start_field)
            end_date = getattr(case, end_field)
            if isinstance(start_date, str) and isinstance(end_date, str):
                start_dt = to_gregorian(start_date)
                end_dt = to_gregorian(end_date)
                if start_dt and end_dt:
                    days.append((end_dt - start_dt).days)
        return round(sum(days) / len(days), 0) if days else 0

    section_cases = SectionCase.objects.filter(group=group, doctor=doctor)
    dc_section_cases = DC.objects.filter(group=group, doctor=doctor)
    room_cases = RoomCase.objects.filter(group=group, doctor=doctor)
    big_room_cases = room_cases.filter(operation_type='3')
    medium_room_cases = room_cases.filter(operation_type='2')
    small_room_cases = room_cases.filter(operation_type='1')

    not_arrived_cases = section_cases.filter(delivery_date='nan')

    defect_cases = section_cases.filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False) |
        Q(defect_sheet3__isnull=False) | Q(defect_sheet4__isnull=False) |
        Q(defect_sheet5__isnull=False) | Q(defect_sheet6__isnull=False) |
        Q(defect_sheet7__isnull=False) | Q(defect_sheet8__isnull=False) |
        Q(defect_sheet9__isnull=False) | Q(defect_sheet10__isnull=False)
    )
    all_defect_cases = SectionCase.objects.filter(group=group).filter(
        Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False) |
        Q(defect_sheet3__isnull=False) | Q(defect_sheet4__isnull=False) |
        Q(defect_sheet5__isnull=False) | Q(defect_sheet6__isnull=False) |
        Q(defect_sheet7__isnull=False) | Q(defect_sheet8__isnull=False) |
        Q(defect_sheet9__isnull=False) | Q(defect_sheet10__isnull=False)
    )

    insurance_filter = lambda s: section_cases.filter(insurance__icontains=s)
    social_security_cases = insurance_filter("تامین اجتماعی")
    medical_services_cases = insurance_filter("خدمات درمانی")
    armed_forces_cases = insurance_filter("نیرو های مسلح")
    free_cases = insurance_filter("آزاد")

    # زمان‌بندی
    if start and end and start > end:
        start, end = end, start

    filtered_section_cases = filter_by_date(section_cases, 'admission_date')
    filtered_dc_section_cases = filter_by_date(dc_section_cases, 'admission_date')
    filtered_not_arrived_cases = filter_by_date(not_arrived_cases, 'admission_date')
    filtered_defect_cases = filter_by_date(defect_cases, 'admission_date')
    filtered_all_defect_cases = filter_by_date(all_defect_cases, 'admission_date')
    filtered_social_security_cases = filter_by_date(social_security_cases, 'admission_date')
    filtered_medical_services_cases = filter_by_date(medical_services_cases, 'admission_date')
    filtered_armed_forces_cases = filter_by_date(armed_forces_cases, 'admission_date')
    filtered_free_cases = filter_by_date(free_cases, 'admission_date')
    filtered_room_cases = filter_by_date(room_cases, 'operation_date')
    filtered_big_room_cases = filter_by_date(big_room_cases, 'operation_date')
    filtered_medium_room_cases = filter_by_date(medium_room_cases, 'operation_date')
    filtered_small_room_cases = filter_by_date(small_room_cases, 'operation_date')

    f_patients = set(c.patient for c in filtered_section_cases + filtered_room_cases)
    patients_count = len(f_patients)

    # نقص
    percent_defect_cases = (
        (len(filtered_defect_cases) * 100) // len(filtered_all_defect_cases)
        if filtered_all_defect_cases else 0
    )

    # پراکندگی نقص
    defect_sheet_fields = ['defect_sheet'] + [f'defect_sheet{i}' for i in range(2, 11)]
    defect_type_fields = ['defect_type'] + [f'defect_type{i}' for i in range(2, 11)]

    defect_counts = {
        name: sum([
            1 for sc in section_cases
            if any(getattr(sc, field) == code for field in defect_sheet_fields)
        ]) for code, name in defect_sheet_choices
    }

    defect_type_counts = {
        name: sum([
            1 for sc in section_cases
            if any(
                getattr(sc, field) and code in getattr(sc, field)
                for field in defect_type_fields
            )
        ]) for code, name in defect_type_choices
    }

    # فوت‌شدگان
    age_counts = {'less_20': 0, 'more_20_less_40': 0, 'more_40_less_60': 0, 'more_60_less_80': 0, 'more_80': 0}
    gender_counts = {'men': 0, 'women': 0}

    for case in filtered_dc_section_cases:
        age = int(''.join(filter(str.isdigit, case.age or '0')))
        gender = case.gender
        if age < 20:
            age_counts['less_20'] += 1
        elif age < 40:
            age_counts['more_20_less_40'] += 1
        elif age < 60:
            age_counts['more_40_less_60'] += 1
        elif age < 80:
            age_counts['more_60_less_80'] += 1
        else:
            age_counts['more_80'] += 1
        if gender == '1':
            gender_counts['men'] += 1
        else:
            gender_counts['women'] += 1

    return {
        'doctor': doctor,
        'patients_count': patients_count,
        'filtered_section_cases_count': len(filtered_section_cases),
        'filtered_dc_section_cases_count': len(filtered_dc_section_cases),
        'filtered_not_arrived_cases_count': len(filtered_not_arrived_cases),
        'filtered_defect_cases_count': len(filtered_defect_cases),
        'filtered_social_security_cases_count': len(filtered_social_security_cases),
        'filtered_medical_services_cases_count': len(filtered_medical_services_cases),
        'filtered_armed_forces_cases_count': len(filtered_armed_forces_cases),
        'filtered_free_cases_count': len(filtered_free_cases),
        'filtered_room_cases_count': len(filtered_room_cases),
        'filtered_big_room_cases_count': len(filtered_big_room_cases),
        'filtered_medium_room_cases_count': len(filtered_medium_room_cases),
        'filtered_small_room_cases_count': len(filtered_small_room_cases),
        'average_arrive_days': calc_average_days(filtered_section_cases, 'discharge_date', 'delivery_date'),
        'average_stay_days': calc_average_days(filtered_section_cases, 'admission_date', 'discharge_date'),
        'defect_counts': defect_counts,
        'defect_type_counts': defect_type_counts,
        'percent_defect_cases': percent_defect_cases,
        'age_counts': age_counts,
        'gender_counts': gender_counts,
    }

@login_required
@manager_required
def multi_section_analysis(request):
    if request.method == 'POST':
        form = MultiSectionForm(request.POST)
        if form.is_valid():
            sections = form.cleaned_data['sections']
            start = form.cleaned_data['start']
            end = form.cleaned_data['end']
            
            results = {}
            for section in sections:
                data = analyze_section(section, request.user.group, start, end)
                results[section.name] = data
            
            context = {
                'form': form,
                'results': results,
            }
            return render(request, 'multi_section_results.html', context)
    else:
        form = MultiSectionForm()

    return render(request, 'multi_section_form.html', {'form': form})

@login_required
@manager_required
def multi_room_analysis(request):
    if request.method == 'POST':
        form = MultiRoomForm(request.POST)
        if form.is_valid():
            rooms = form.cleaned_data['rooms']
            start = form.cleaned_data['start']
            end = form.cleaned_data['end']
            
            results = {}
            for room in rooms:
                data = analyze_room(room, request.user.group, start, end)
                results[room.name] = data
            
            context = {
                'form': form,
                'results': results,
            }
            return render(request, 'multi_room_results.html', context)
    else:
        form = MultiRoomForm()

    return render(request, 'multi_room_form.html', {'form': form})

@login_required
@manager_required
def multi_doctor_analysis(request):
    if request.method == 'POST':
        form = MultiDoctorForm(request.POST)
        if form.is_valid():
            doctors = form.cleaned_data['doctors']
            start = form.cleaned_data['start']
            end = form.cleaned_data['end']
            
            results = {}
            for doctor in doctors:
                data = analyze_doctor(doctor, request.user.group, start, end)
                results[doctor.full_name] = data
            
            context = {
                'form': form,
                'results': results,
            }
            return render(request, 'multi_doctor_results.html', context)
    else:
        form = MultiDoctorForm()

    return render(request, 'multi_doctor_form.html', {'form': form})

@login_required
@manager_required
def analyze_defect(request):
    group = request.user.group

    if Excel.objects.filter(group=group).exists():
        section_cases = SectionCase.objects.filter(group=group)
        defect_cases = section_cases.filter(
            Q(defect_sheet__isnull=False) | Q(defect_sheet2__isnull=False) |
            Q(defect_sheet3__isnull=False) | Q(defect_sheet4__isnull=False) |
            Q(defect_sheet5__isnull=False) | Q(defect_sheet6__isnull=False) |
            Q(defect_sheet7__isnull=False) | Q(defect_sheet8__isnull=False) |
            Q(defect_sheet9__isnull=False) | Q(defect_sheet10__isnull=False)
        )

        # آماده‌سازی داده نقص
        defect_counts = {}
        defect_percents = {}
        defect_type_counts = {}
        defect_type_percents = {}

        filtered_section_cases = []

        if request.GET.get("start") and request.GET.get("end"):
            try:
                start = Persian(request.GET["start"]).gregorian_datetime()
                end = Persian(request.GET["end"]).gregorian_datetime()
                if start > end:
                    start, end = end, start

                for case in section_cases:
                    if isinstance(case.admission_date, str):
                        try:
                            case_date = Persian(case.admission_date).gregorian_datetime()
                            if start <= case_date <= end:
                                filtered_section_cases.append(case)
                        except:
                            continue

                # شمارش نقص‌ها در بازه زمانی
                def count_defects(
                        cases, 
                        field1, field2, field3, field4, field5, field6, field7, field8, field9, field10,
                        choices
                ):
                    counts, percents = {}, {}
                    for code, name in choices:
                        count = sum(1 for c in cases if getattr(c, field1) == code or getattr(c, field2) == code or getattr(c, field3) == code or getattr(c, field4) == code or getattr(c, field5) == code or getattr(c, field6) == code or getattr(c, field7) == code or getattr(c, field8) == code or getattr(c, field9) == code or getattr(c, field10) == code)
                        total = sum(1 for c in cases if getattr(c, field1) or getattr(c, field2))
                        counts[name] = count
                        percents[name] = round((count * 100 / total), 0) if total else 0
                    return counts, percents
                
                def count_multiselect_defects(cases, fields, choices):
                    counts, percents = {}, {}
                    total = sum(
                        1 for c in cases if any(getattr(c, f, None) for f in fields)
                    )

                    for code, name in choices:
                        count = sum(
                            1 for c in cases for f in fields
                            if code in (getattr(c, f, []) or [])
                        )
                        counts[name] = count
                        percents[name] = round((count * 100 / total), 0) if total else 0

                    return counts, percents

                defect_counts, defect_percents = count_defects(
                    filtered_section_cases, 
                    'defect_sheet', 'defect_sheet2', 'defect_sheet3', 'defect_sheet4', 'defect_sheet5', 'defect_sheet6', 'defect_sheet7', 'defect_sheet8', 'defect_sheet9', 'defect_sheet10', 
                    defect_sheet_choices
                )

                fields = [
                    'defect_type', 'defect_type2', 'defect_type3',
                    'defect_type4', 'defect_type5', 'defect_type6',
                    'defect_type7', 'defect_type8', 'defect_type9', 'defect_type10'
                ]

                defect_type_counts, defect_type_percents = count_multiselect_defects(
                    filtered_section_cases,
                    fields,
                    defect_type_choices
                )

            except:
                pass
        else:
            def count_global_defects(
                    field1, field2, field3, field4, field5, field6, field7, field8, field9, field10,
                    choices, filter_field
            ):
                counts, percents = {}, {}
                total = defect_cases.count()
                for code, name in choices:
                    count = section_cases.filter(**{field1: code}).count() + \
                            section_cases.filter(**{field2: code}).count() + \
                            section_cases.filter(**{field3: code}).count() + \
                            section_cases.filter(**{field4: code}).count() + \
                            section_cases.filter(**{field5: code}).count() + \
                            section_cases.filter(**{field6: code}).count() + \
                            section_cases.filter(**{field7: code}).count() + \
                            section_cases.filter(**{field8: code}).count() + \
                            section_cases.filter(**{field9: code}).count() + \
                            section_cases.filter(**{field10: code}).count()
                    counts[name] = count
                    percents[name] = round((count * 100 / total), 0) if total else 0
                return counts, percents
            
            def count_multiselect_defects(cases, fields, choices):
                counts, percents = {}, {}
                total = sum(
                    1 for c in cases if any(getattr(c, f, None) for f in fields)
                )

                for code, name in choices:
                    count = sum(
                        1 for c in cases for f in fields
                        if code in (getattr(c, f, []) or [])
                    )
                    counts[name] = count
                    percents[name] = round((count * 100 / total), 0) if total else 0

                return counts, percents

            defect_counts, defect_percents = count_global_defects(
                'defect_sheet', 'defect_sheet2', 'defect_sheet3', 'defect_sheet4', 'defect_sheet5', 'defect_sheet6', 'defect_sheet7', 'defect_sheet8', 'defect_sheet9', 'defect_sheet10', 
                defect_sheet_choices, 'defect_sheet'
            )

            fields = [
                'defect_type', 'defect_type2', 'defect_type3',
                'defect_type4', 'defect_type5', 'defect_type6',
                'defect_type7', 'defect_type8', 'defect_type9', 'defect_type10'
            ]

            defect_type_counts, defect_type_percents = count_multiselect_defects(
                section_cases,
                fields,
                defect_type_choices
            )

        context = {
            'defect_counts': defect_counts,
            'defect_type_counts': defect_type_counts,
            'defect_percents': defect_percents,
            'defect_type_percents': defect_type_percents,
        }

        return render(request, 'analyze_defect.html', context=context)
    else:
        return render(request, 'main.html', context=context)