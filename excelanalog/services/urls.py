from django.contrib import admin
from django.urls import path
from .views import *


urlpatterns = [
    path('admin/', admin.site.urls),
    path('k/', index, name='index'),
    path('', base, name='base'),
    path('checklist/<int:pk>/', checklist_detail, name='checklist_detail'),
    path('reestr/<int:pk>/', edit_reestr, name='edit_reestr'),
    path('kk/', indexx, name='indexx'),
    path('kkk/', svod, name='svod'),
    path('kkkk/', svod2, name='svod2'),
    path('checklist/<int:pk>/download/', download_excel, name='download_excel'),
    path('checklist1/<int:pk>/download/', download_excel1, name='download_excel1'),
    path('upload/', upload_file, name='upload')


]
