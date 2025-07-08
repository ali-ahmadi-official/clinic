from django.urls import path
from django.contrib.auth.views import LogoutView

from .views import (
    main, SignUpView, custom_login_view,
    SectionListView, section_detail, SectionCreateView, SectionUpdateView, SectionDeleteView,
    RoomListView, room_detail, RoomCreateView, RoomUpdateView, RoomDeleteView,
    ExpertiseCreateView, ExpertiseListView, ExpertiseUpdateView, ExpertiseDeleteView,
    DoctorListView, doctor_detail, DoctorCreateView, DoctorUpdateView, DoctorDeleteView,
    PatientListView, patient_detail, PatientDeleteView,
    SectionCaseListView, add_section_case, section_case_detail, SectionCaseUpdateView, SectionCaseDeleteView,
    RoomCaseListView, add_room_case, RoomCaseDetailView, RoomCaseDeleteView,
    DCListView, DCDetailView, DCDeleteView, dc_all_detail,
    all_delete,
    multi_section_analysis, multi_room_analysis, multi_doctor_analysis
)

urlpatterns = [
    path('', main, name='main'),

    path('signup/', SignUpView.as_view(), name='signup'),
    path('login/', custom_login_view, name='login'),
    path('logout/', LogoutView.as_view(), name='logout'),

    path('sections/', SectionListView.as_view(), name='section_list'),
    path('sections/add/', SectionCreateView.as_view(), name='section_create'),
    path('sections/<int:pk>/', section_detail, name='section_detail'),
    path('sections/<int:pk>/update/', SectionUpdateView.as_view(), name='section_update'),
    path('sections/<int:pk>/delete/', SectionDeleteView.as_view(), name='section_delete'),

    path('rooms/', RoomListView.as_view(), name='room_list'),
    path('rooms/add/', RoomCreateView.as_view(), name='room_create'),
    path('rooms/<int:pk>/', room_detail, name='room_detail'),
    path('rooms/<int:pk>/update/', RoomUpdateView.as_view(), name='room_update'),
    path('rooms/<int:pk>/delete/', RoomDeleteView.as_view(), name='room_delete'),

    path('expertises/', ExpertiseListView.as_view(), name='expertise_list'),
    path('expertises/add/', ExpertiseCreateView.as_view(), name='expertise_create'),
    path('expertises/<int:pk>/update/', ExpertiseUpdateView.as_view(), name='expertise_update'),
    path('expertises/<int:pk>/delete/', ExpertiseDeleteView.as_view(), name='expertise_delete'),

    path('doctors/', DoctorListView.as_view(), name='doctor_list'),
    path('doctors/add/', DoctorCreateView.as_view(), name='doctor_create'),
    path('doctors/<int:pk>/', doctor_detail, name='doctor_detail'),
    path('doctors/<int:pk>/update/', DoctorUpdateView.as_view(), name='doctor_update'),
    path('doctors/<int:pk>/delete/', DoctorDeleteView.as_view(), name='doctor_delete'),

    path('patients/', PatientListView.as_view(), name='patient_list'),
    path('patients/<int:pk>/', patient_detail, name='patient_detail'),
    path('patients/<int:pk>/delete/', PatientDeleteView.as_view(), name='patient_delete'),

    path('section-cases/', SectionCaseListView.as_view(), name='section_case_list'),
    path('section-cases/add/', add_section_case, name='add_section_case'),
    path('section-cases/<int:pk>/', section_case_detail, name='section_case_detail'),
    path('section-cases/<int:pk>/update/', SectionCaseUpdateView.as_view(), name='section_case_update'),
    path('section-cases/<int:pk>/delete/', SectionCaseDeleteView.as_view(), name='section_case_delete'),

    path('room-cases/', RoomCaseListView.as_view(), name='room_case_list'),
    path('room-cases/add/', add_room_case, name='add_room_case'),
    path('room-cases/<int:pk>/', RoomCaseDetailView.as_view(), name='room_case_detail'),
    path('room-cases/<int:pk>/delete/', RoomCaseDeleteView.as_view(), name='room_case_delete'),

    path('dcs/', DCListView.as_view(), name='dc_list'),
    path('dcs/<int:pk>/', DCDetailView.as_view(), name='dc_detail'),
    path('dcs/<int:pk>/delete/', DCDeleteView.as_view(), name='dc_delete'),
    path('dcs/all-detail/', dc_all_detail, name='dc_all_detail'),

    path('all-delete/', all_delete, name='all_delete'),

    path('analyze/section/', multi_section_analysis, name='analyze_section'),
    path('analyze/room/', multi_room_analysis, name='analyze_room'),
    path('analyze/doctor/', multi_doctor_analysis, name='analyze_doctor'),
]