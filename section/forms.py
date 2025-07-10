from django import forms
from django.contrib.auth import get_user_model
from django.contrib.auth.forms import UserCreationForm
from .models import Excel, Expertise, Section, Room, Doctor, SectionCase

CustomUser = get_user_model()

class CustomUserCreationForm(UserCreationForm):
    username = forms.CharField(label='نام کاربری', max_length=100, widget=forms.TextInput(attrs={'class': 'form-control'}))
    password1 = forms.CharField(label='رمز عبور', widget=forms.PasswordInput(attrs={'class': 'form-control'}))
    password2 = forms.CharField(label='تکرار رمز عبور', widget=forms.PasswordInput(attrs={'class': 'form-control'}))

    class Meta:
        model = CustomUser
        fields = ('username', 'password1', 'password2')

class LoginForm(forms.Form):
    username = forms.CharField(label='نام کاربری')
    password = forms.CharField(widget=forms.PasswordInput, label='رمز عبور')

class ExcelForm(forms.ModelForm):
    class Meta:
        model = Excel
        exclude = ('group',)

class ExpertiseForm(forms.ModelForm):
    class Meta:
        model = Expertise
        exclude = ('group',)

class SectionForm(forms.ModelForm):
    class Meta:
        model = Section
        exclude = ('group',)
        widgets = {
            'expertises': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
        }

class RoomForm(forms.ModelForm):
    class Meta:
        model = Room
        exclude = ('group',)
        widgets = {
            'expertises': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
        }

class DoctorForm(forms.ModelForm):
    class Meta:
        model = Doctor
        exclude = ('group',)
        widgets = {
            'grade': forms.Select(attrs={'class': 'form-control'}),
            'sections': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'rooms': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'expertises': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
        }

class SectionCaseForm(forms.ModelForm):
    class Meta:
        model = SectionCase
        fields = [
            'defect_sheet', 'defect_type', 'defect_sheet2', 'defect_type2', 'defect_sheet3', 'defect_type3',
            'defect_sheet4', 'defect_type4', 'defect_sheet5', 'defect_type5', 'defect_sheet6', 'defect_type6',
            'defect_sheet7', 'defect_type7', 'defect_sheet8', 'defect_type8', 'defect_sheet9', 'defect_type9',
            'defect_sheet10', 'defect_type10',
        ]
        widgets = {
            'defect_sheet': forms.Select(attrs={'class': 'form-control'}),
            'defect_type': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet2': forms.Select(attrs={'class': 'form-control'}),
            'defect_type2': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet3': forms.Select(attrs={'class': 'form-control'}),
            'defect_type3': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet4': forms.Select(attrs={'class': 'form-control'}),
            'defect_type4': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet5': forms.Select(attrs={'class': 'form-control'}),
            'defect_type5': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet6': forms.Select(attrs={'class': 'form-control'}),
            'defect_type6': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet7': forms.Select(attrs={'class': 'form-control'}),
            'defect_type7': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet8': forms.Select(attrs={'class': 'form-control'}),
            'defect_type8': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet9': forms.Select(attrs={'class': 'form-control'}),
            'defect_type9': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
            'defect_sheet10': forms.Select(attrs={'class': 'form-control'}),
            'defect_type10': forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
        }

class ConfirmDeleteForm(forms.Form):
    confirm = forms.BooleanField(label="تایید نهایی")

class MultiSectionForm(forms.Form):
    sections = forms.ModelMultipleChoiceField(
        queryset=Section.objects.all(),
        widget=forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
        required=True,
        label="بخش‌ها"
    )
    start = forms.CharField(label="تاریخ شروع", required=False)
    end = forms.CharField(label="تاریخ پایان", required=False)

class MultiRoomForm(forms.Form):
    rooms = forms.ModelMultipleChoiceField(
        queryset=Room.objects.all(),
        widget=forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
        required=True,
        label="اتاق عمل ها"
    )
    start = forms.CharField(label="تاریخ شروع", required=False)
    end = forms.CharField(label="تاریخ پایان", required=False)

class MultiDoctorForm(forms.Form):
    doctors = forms.ModelMultipleChoiceField(
        queryset=Doctor.objects.all(),
        widget=forms.CheckboxSelectMultiple(attrs={'class': 'checkbox-multiple'}),
        required=True,
        label="پزشکان"
    )
    start = forms.CharField(label="تاریخ شروع", required=False)
    end = forms.CharField(label="تاریخ پایان", required=False)
