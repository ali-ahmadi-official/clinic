from django.contrib import admin
from .models import Group, CustomUser

admin.site.register(Group)
admin.site.register(CustomUser)