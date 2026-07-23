# Пять Excel-протоколов лабораторных замеров (output_excel_1..5).

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('attestation', '0003_batch_output_6_4'),
    ]

    operations = [
        migrations.AddField(
            model_name='batch',
            name='output_excel_1',
            field=models.CharField(blank=True, default='', max_length=512, verbose_name='Протокол 1 (вредные вещества)'),
        ),
        migrations.AddField(
            model_name='batch',
            name='output_excel_2',
            field=models.CharField(blank=True, default='', max_length=512, verbose_name='Протокол 2 (физические факторы)'),
        ),
        migrations.AddField(
            model_name='batch',
            name='output_excel_3',
            field=models.CharField(blank=True, default='', max_length=512, verbose_name='Протокол 3 (микроклимат)'),
        ),
        migrations.AddField(
            model_name='batch',
            name='output_excel_4',
            field=models.CharField(blank=True, default='', max_length=512, verbose_name='Протокол 4 (освещённость)'),
        ),
        migrations.AddField(
            model_name='batch',
            name='output_excel_5',
            field=models.CharField(blank=True, default='', max_length=512, verbose_name='Протокол 5 (ЭМИ/магнитные поля)'),
        ),
    ]
