from django.http import HttpResponse
from django.shortcuts import render, get_object_or_404, redirect
from openpyxl.workbook import Workbook

from .forms import CheckListForm, ReestrForm
from .models import Columns, CheckList, Reestr


# Create your views here.
def index(request):
    return render(request, 'base.html')



def base(request):
    if request.method == 'POST':
        # Получаем данные из формы, введенные в админке
        column1_data = request.POST.get('column1')
        column2_data = request.POST.get('column2')
        column3_data = request.POST.get('column3')
        column4_data = request.POST.get('column4')
        column5_data = request.POST.get('column5')
        column6_data = request.POST.get('column6')
        column7_data = request.POST.get('column7')
        column8_data = request.POST.get('column8')
        column9_data = request.POST.get('column9')

        # Создаем новый объект модели Columns с введенными данными
        new_column = Columns(column1=column1_data, column2=column2_data, column3=column3_data,
                             column4=column4_data, column5=column5_data, column6=column6_data,
                             column7=column7_data, column8=column8_data, column9=column9_data)
        new_column.save()

    # Получаем все объекты модели Columns
    columns = Columns.objects.all()
    return render(request, 'tables.html', {'columns': columns})



def checklist_detail(request, pk):
    checklist = get_object_or_404(CheckList, pk=pk)
    columns = CheckList.objects.all()
    if request.method == 'POST':
        form = CheckListForm(request.POST, instance=checklist)
        if form.is_valid():
            form.save()
            return redirect('checklist_detail', pk=checklist.pk)
    else:
        form = CheckListForm(instance=checklist)
    # Создание нового файла Excel
    workbook = Workbook()
    worksheet = workbook.active
    # Запись заголовков таблицы
    headers = ['номер п/п', 'Код КП(общий)', 'Код КП(промежуточный)', 'Наименование ИП', 'Описание КП', 'Переодичность проведения',
               'Способ подсчета результаты проведения КП', 'Подразделение, ответственное за проведение контрольной процедуры',
               'Исполнитель КП', 'Количество выполненых КП', 'Количество выявленных ошибок', 'сведения об объекте контроля']
    for col_num, header in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.value = header
    # Запись данных из базы данных в таблицу
    for row_num, column in enumerate(columns, 2):
        worksheet.cell(row=row_num, column=1).value = column.number
        worksheet.cell(row=row_num, column=2).value = column.cod_kp_overall
        worksheet.cell(row=row_num, column=3).value = column.cod_kp_intervall
        worksheet.cell(row=row_num, column=4).value = column.name_ip
        worksheet.cell(row=row_num, column=5).value = column.description_ip
        worksheet.cell(row=row_num, column=6).value = column.pereodiction_carriage
        worksheet.cell(row=row_num, column=7).value = column.counting_abillity
        worksheet.cell(row=row_num, column=8).value = column.responsible_group
        worksheet.cell(row=row_num, column=9).value = column.perforemr_kp
        worksheet.cell(row=row_num, column=10).value = column.number_complete
        worksheet.cell(row=row_num, column=11).value = column.number_mistakes
        worksheet.cell(row=row_num, column=12).value = column.data_object
    # Сохранение файла Excel



    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=checklist.xlsx'
    workbook.save(response)
    return response





def edit_reestr(request, pk):
    reestr = get_object_or_404(Reestr, pk=pk)
    columns = Reestr.objects.all()  # Получение всех объектов Column
    if request.method == 'POST':
        form = ReestrForm(request.POST, instance=reestr)
        if form.is_valid():
            form.save()
            return redirect('edit_reestr', pk=reestr.pk)
    else:
        form = ReestrForm(instance=reestr)
    return render(request, 'reestr.html', {'reestr': reestr,'form': form, 'columns': columns })



def indexx(request):
    if request.method == 'POST':
        num_rows = int(request.POST['num_rows'])
        columns = []

        for i in range(num_rows):
            column = {
                'num': request.POST.get(f'num_{i}'),
                'cod_kp_inter': request.POST.get(f'cod_kp_inter_{i}'),
                'chek_num': request.POST.get(f'chek_num_{i}'),
                'obj_control': request.POST.get(f'obj_control_{i}'),
                'date_document': request.POST.get(f'date_document_{i}'),
                'num_document': request.POST.get(f'num_document_{i}'),
                'colvo_doc': request.POST.get(f'colvo_doc_{i}'),
                'colvo_errors': request.POST.get(f'colvo_errors_{i}'),
                'notes': request.POST.get(f'notes_{i}')
            }
            columns.append(column)

        # Создаем новый Excel-файл и заполняем его данными из таблицы
        workbook = Workbook()
        sheet = workbook.active

        # Записываем заголовки столбцов
        headers = ['Номер', 'Код КП интерфейса', 'Номер чека', 'Объект контроля',
                   'Дата документа', 'Номер документа', 'Количество документов',
                   'Количество ошибок', 'Примечания']
        for col_num, header in enumerate(headers, 1):
            sheet.cell(row=1, column=col_num).value = header

        # Записываем данные из таблицы
        for row_num, column in enumerate(columns, 2):
            sheet.cell(row=row_num, column=1).value = column['num']
            sheet.cell(row=row_num, column=2).value = column['cod_kp_inter']
            sheet.cell(row=row_num, column=3).value = column['chek_num']
            sheet.cell(row=row_num, column=4).value = column['obj_control']
            sheet.cell(row=row_num, column=5).value = column['date_document']
            sheet.cell(row=row_num, column=6).value = column['num_document']
            sheet.cell(row=row_num, column=7).value = column['colvo_doc']
            sheet.cell(row=row_num, column=8).value = column['colvo_errors']
            sheet.cell(row=row_num, column=9).value = column['notes']

        # Сохраняем Excel-файл
        workbook.save('реестр.xlsx')

        return render(request, 'index.html')

    return render(request, 'index.html')


def download_excel(request, pk):
    checklist = get_object_or_404(CheckList, pk=pk)

    workbook = Workbook()
    worksheet = workbook.active

    # Запись заголовков таблицы
    headers = ['номер п/п', 'Код КП(общий)', 'Код КП(промежуточный)', 'Наименование ИП', 'Описание КП',
               'Переодичность проведения',
               'Способ подсчета результаты проведения КП',
               'Подразделение, ответственное за проведение контрольной процедуры',
               'Исполнитель КП', 'Количество выполненых КП', 'Количество выявленных ошибок',
               'сведения об объекте контроля']
    for col_num, header in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.value = header

    # Запись данных из базы данных в таблицу
    worksheet.cell(row=2, column=1).value = checklist.number
    worksheet.cell(row=2, column=2).value = checklist.cod_kp_overall
    worksheet.cell(row=2, column=3).value = checklist.cod_kp_intervall
    worksheet.cell(row=2, column=4).value = checklist.name_ip
    worksheet.cell(row=2, column=5).value = checklist.description_ip
    worksheet.cell(row=2, column=6).value = checklist.pereodiction_carriage
    worksheet.cell(row=2, column=7).value = checklist.counting_abillity
    worksheet.cell(row=2, column=8).value = checklist.responsible_group
    worksheet.cell(row=2, column=9).value = checklist.perforemr_kp
    worksheet.cell(row=2, column=10).value = checklist.number_complete
    worksheet.cell(row=2, column=11).value = checklist.number_mistakes
    worksheet.cell(row=2, column=12).value = checklist.data_object

    worksheet.cell(row=15, column=2).value = "Должность:"
    worksheet.cell(row=15, column=5).value = "   Подпись:"
    worksheet.cell(row=15, column=7).value = "           ФИО:"

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=checklist.xlsx'
    workbook.save(response)

    return response



def download_excel1(request, pk):
    checklist = get_object_or_404(Reestr, pk=pk)

    workbook = Workbook()
    worksheet = workbook.active

    # Запись заголовков таблицы
    headers = ['номер п/п', 'Код КП(промежуточный)', 'номер чек листа', 'Обьект контроля', 'Дата документа',
               'Номер документа',
               'Количество документов/операций',
               'Количество ошибок/нарушений',
               'Примечаний',
               ]
    for col_num, header in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.value = header

    # Запись данных из базы данных в таблицу
    worksheet.cell(row=2, column=1).value = checklist.num
    worksheet.cell(row=2, column=2).value = checklist.cod_kp_inter
    worksheet.cell(row=2, column=3).value = checklist.chek_num
    worksheet.cell(row=2, column=4).value = checklist.obj_control
    worksheet.cell(row=2, column=5).value = checklist.date_document
    worksheet.cell(row=2, column=6).value = checklist.num_document
    worksheet.cell(row=2, column=7).value = checklist.colvo_doc
    worksheet.cell(row=2, column=8).value = checklist.colvo_errors
    worksheet.cell(row=2, column=9).value = checklist.notes


    worksheet.cell(row=15, column=2).value = "Должность:"
    worksheet.cell(row=15, column=5).value = "   Подпись:"


    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=checklist.xlsx'
    workbook.save(response)

    return response