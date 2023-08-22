from django import forms
from .models import CheckList, Reestr


class CheckListForm(forms.ModelForm):
    class Meta:
        model = CheckList
        fields = ('responsible_group', 'perforemr_kp')


class ReestrForm(forms.ModelForm):
    class Meta:
        model = Reestr
        fields = ['date_document', 'num_document', 'colvo_doc', 'colvo_errors', 'notes']
