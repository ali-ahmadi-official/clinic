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
            'expertises': forms.SelectMultiple(attrs={'class': 'form-control'}),
        }

class RoomForm(forms.ModelForm):
    class Meta:
        model = Room
        exclude = ('group',)
        widgets = {
            'expertises': forms.SelectMultiple(attrs={'class': 'form-control'}),
        }

class DoctorForm(forms.ModelForm):
    class Meta:
        model = Doctor
        exclude = ('group',)
        widgets = {
            'grade': forms.Select(attrs={'class': 'form-control'}),
            'sections': forms.SelectMultiple(attrs={'class': 'form-control'}),
            'rooms': forms.SelectMultiple(attrs={'class': 'form-control'}),
            'expertises': forms.SelectMultiple(attrs={'class': 'form-control'}),
        }

class SectionCaseForm(forms.ModelForm):
    class Meta:
        model = SectionCase
        fields = ['defect_sheet', 'defect_type', 'defect_sheet2', 'defect_type2']
        widgets = {
            'defect_sheet': forms.Select(attrs={'class': 'form-control'}),
            'defect_type': forms.Select(attrs={'class': 'form-control'}),
            'defect_sheet2': forms.Select(attrs={'class': 'form-control'}),
            'defect_type2': forms.Select(attrs={'class': 'form-control'}),
        }

class ConfirmDeleteForm(forms.Form):
    confirm = forms.BooleanField(label="تایید نهایی")