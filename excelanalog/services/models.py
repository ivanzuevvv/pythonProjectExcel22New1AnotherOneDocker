from django.db import models

# Create your models here.
class Columns(models.Model):
    column1 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-1")
    column2 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-2")
    column3 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-3")
    column4 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-4")
    column5 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-5")
    column6 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-6")
    column7 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-7")
    column8 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-8")
    column9 = models.CharField(max_length=500, blank=True, null=True, verbose_name="Столб-9")



class CheckList(models.Model):
    number = models.CharField(max_length=500, blank=True, null=True, verbose_name="номер п/п")
    cod_kp_overall = models.CharField(max_length=500, blank=True, null=True, verbose_name="Код КП(общий)")
    cod_kp_intervall = models.CharField(max_length=500, blank=True, null=True, verbose_name="Код КП(промежуточный)")
    name_ip = models.CharField(max_length=500, blank=True, null=True, verbose_name="Наименование ИП")
    description_ip = models.CharField(max_length=500, blank=True, null=True, verbose_name="Описание КП")
    pereodiction_carriage = models.CharField(max_length=500, blank=True, null=True, verbose_name="Переодичность проведения")
    counting_abillity = models.CharField(max_length=500, blank=True, null=True, verbose_name="Способ подсчета результаты проведения КП")
    responsible_group = models.CharField(max_length=500, blank=True, null=True, verbose_name="Подразделение, ответственное за проведение контрольной процедуры")
    perforemr_kp = models.CharField(max_length=500, blank=True, null=True, verbose_name="Исполнитель КП")
    number_complete = models.CharField(max_length=500, blank=True, null=True, verbose_name="Количество выполненых КП")
    number_mistakes = models.CharField(max_length=500, blank=True, null=True, verbose_name="Количество выявленных ошибок")
    data_object = models.CharField(max_length=500, blank=True, null=True, verbose_name="сведения об объекте контроля")
    cheklist = models.ForeignKey(Columns, on_delete=models.CASCADE)


class Reestr(models.Model):
    num = models.CharField(max_length=500, blank=True, null=True, verbose_name="номер п/п")
    cod_kp_inter = models.CharField(max_length=500, blank=True, null=True, verbose_name="Код КП(промежуточный)")
    performer_ip = models.CharField(max_length=500, blank=True, null=True, verbose_name="Исполнитель ИП")

    chek_num = models.CharField(max_length=500, blank=True, null=True, verbose_name="номер чек листа")
    obj_control = models.CharField(max_length=500, blank=True, null=True, verbose_name="Обьект контроля")
    date_document = models.CharField(max_length=500, blank=True, null=True, verbose_name="Дата документа")
    num_document = models.CharField(max_length=500, blank=True, null=True, verbose_name="Номер документа")
    colvo_doc = models.CharField(max_length=500, blank=True, null=True, verbose_name="Количество документов/операций")
    colvo_errors = models.CharField(max_length=500, blank=True, null=True, verbose_name="Количество ошибок/нарушений")
    notes = models.CharField(max_length=500, blank=True, null=True, verbose_name="Примечаний")
    cheklist1 = models.ForeignKey(CheckList, on_delete=models.CASCADE)


