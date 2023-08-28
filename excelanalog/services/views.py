import openpyxl
from django.http import HttpResponse
from django.shortcuts import render, get_object_or_404, redirect
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook

from .forms import CheckListForm, ReestrForm
from .models import Columns, CheckList, Reestr
from openpyxl.styles import Font, PatternFill, Alignment



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



def download_excel(request, pk):
    from openpyxl import Workbook
    from openpyxl.styles import Border, Side
    from django.shortcuts import get_object_or_404
    from django.http import HttpResponse
    from openpyxl.styles import Alignment
    from openpyxl import Workbook

    checklist = get_object_or_404(CheckList, pk=pk)

    workbook = Workbook()
    worksheet = workbook.active



    # Сохранение файла
    workbook.save('example.xlsx')

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

    thin_border = Border(top=Side(border_style="thin"))

    # Запись данных из базы данных в таблицу

    worksheet.cell(row=2, column=1).value = checklist.number
    worksheet.merge_cells('A1')
    worksheet['A1'] = 'номер п/п'
    worksheet['A1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=2).value = checklist.number
    worksheet.merge_cells('B1')
    worksheet['B1'] = 'Код КП(общий)'
    worksheet['B1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=3).value = checklist.number
    worksheet.merge_cells('C1')
    worksheet['C1'] = 'Код КП(общий)'
    worksheet['C1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=4).value = checklist.cod_kp_intervall
    worksheet.merge_cells('D1')
    worksheet['D1'] = 'Код КП(промеж.)'
    worksheet['D1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)


    worksheet.cell(row=2, column=5).value = checklist.name_ip
    worksheet.merge_cells('E2')
    worksheet['E1'] = 'Наименование ИП'
    worksheet['E1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)


    worksheet.cell(row=2, column=6).value = checklist.description_ip
    worksheet.merge_cells('F1')
    worksheet['F1'] = 'Описание КП'
    worksheet['F1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=7).value = checklist.pereodiction_carriage
    worksheet.merge_cells('G1')
    worksheet['G1'] = 'Переодичность проведения'
    worksheet['G1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=8).value = checklist.counting_abillity
    worksheet.merge_cells('H1')
    worksheet['H1'] = 'Способ подсчета результаты проведения КП'
    worksheet['H1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=9).value = checklist.responsible_group
    worksheet.merge_cells('I1')
    worksheet['I1'] = 'Подразделение, ответственное за проведение контрольной процедуры'
    worksheet['I1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=10).value = checklist.perforemr_kp
    worksheet.merge_cells('J1')
    worksheet['J1'] = 'Исполнитель КП'
    worksheet['J1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=11).value = checklist.number_complete
    worksheet.merge_cells('K1')
    worksheet['K1'] = 'Количество выполненных КП'
    worksheet['K1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=12).value = checklist.number_mistakes
    worksheet.merge_cells('L1')
    worksheet['L1'] = 'Количество выявленных ошибок'
    worksheet['L1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)

    worksheet.cell(row=2, column=13).value = checklist.data_object
    worksheet.merge_cells('M1')
    worksheet['M1'] = 'сведения об объекте контроля'
    worksheet['M1'].alignment = openpyxl.styles.Alignment(horizontal='center', vertical='center', wrap_text=True)


    # Создание стиля границы
    border_style = Border(left=Side(border_style="thin", color="000000"),
                          right=Side(border_style="thin", color="000000"),
                          top=Side(border_style="thin", color="000000"),
                          bottom=Side(border_style="thin", color="000000")
                          )

    # Автоматическое расширение столбцов
    for column in worksheet.columns:
        max_length = 15
        column_letter = get_column_letter(column[1].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)

            except:
                pass
        adjusted_width = 20
        worksheet.column_dimensions[column_letter].width = adjusted_width



    for row in worksheet.rows:
        max_length = 15
        for cell in row:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_height = 400
        alignment = Alignment(horizontal='centerContinuous', vertical='center', wrap_text=True)

        for cell in worksheet[row[1].column]:
            worksheet.row_dimensions[cell.row].height = adjusted_height
            cell.alignment = alignment

    # Сохранение файла
    workbook.save('example.xlsx')



    # Сохранение файла


    # Применение стиля границы к ячейкам
    for i in range(1, 10):
        for column in worksheet.iter_cols(min_row=1, max_row=2, min_col=i, max_col=i + 4):
            for cell in column:
                cell.border = border_style

    worksheet.cell(row=15, column=2).value = "Должность:"
    worksheet.cell(row=15, column=5).value = "   Подпись:"
    worksheet.cell(row=15, column=7).value = "           ФИО:"

    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=checklist.xlsx'
    workbook.save(response)

    return response






def download_excel1(request, pk):
    from openpyxl import Workbook
    from openpyxl.styles import Border, Side
    from django.shortcuts import get_object_or_404
    from django.http import HttpResponse
    checklist = get_object_or_404(Reestr, pk=pk)

    workbook = Workbook()
    worksheet = workbook.active

    # Запись заголовков таблицы
    headers = ['номер п/п', 'Код КП(промежуточный)', 'Исполнитель ИП', 'номер чек листа', 'Обьект контроля', 'Дата документа',
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
    worksheet.merge_cells('B1')
    worksheet['B1'] = 'Код КП(промежуточный)'
    cell.alignment = Alignment(wrap_text=True)
    worksheet.cell(row=2, column=3).value = checklist.performer_ip
    worksheet.cell(row=2, column=4).value = checklist.chek_num
    worksheet.cell(row=2, column=5).value = checklist.obj_control
    worksheet.cell(row=2, column=6).value = checklist.date_document
    worksheet.cell(row=2, column=7).value = checklist.num_document
    worksheet.cell(row=2, column=8).value = checklist.colvo_doc
    worksheet.cell(row=2, column=9).value = checklist.colvo_errors
    worksheet.cell(row=2, column=10).value = checklist.notes

    total_errors = 0  # Инициализация переменной для суммирования ошибок

    border_style = Border(left=Side(border_style="thin", color="000000"),
                          right=Side(border_style="thin", color="000000"),
                          top=Side(border_style="thin", color="000000"),
                          bottom=Side(border_style="thin", color="000000")
                          )

    # Автоматическое расширение столбцов
    for column in worksheet.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column_letter].width = adjusted_width

    # Автоматическое расширение строк


    # Применение стиля границы к ячейкам
    for i in range(1, 8):
        for column in worksheet.iter_cols(min_row=1, max_row=25, min_col=i, max_col=i + 3):
            for cell in column:
                cell.border = border_style




    # Установка значений для "Должность" и "Подпись"
    worksheet.cell(row=28, column=2).value = "_________________________"
    worksheet.cell(row=29, column=2).value = "               (должность)"
    worksheet.cell(row=28, column=6).value = "_________________"
    worksheet.cell(row=29, column=6).value = "         (подпись)"


    response = HttpResponse(content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response['Content-Disposition'] = 'attachment; filename=checklist.xlsx'
    workbook.save(response)

    return response


