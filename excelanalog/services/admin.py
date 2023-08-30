from django.contrib import admin
from .models import *
from import_export.admin import ImportExportModelAdmin

class TablesAdmin(ImportExportModelAdmin, admin.ModelAdmin):
    list_display = ('column1', 'column2', 'column3', 'column4', 'column5', 'column6', 'column7', 'column8', 'column9', 'column10', 'column11', 'column12', 'column13', 'column14', 'column15')

# Register your models here.
admin.site.register(Columns, TablesAdmin )
admin.site.register(CheckList)
admin.site.register(Reestr)