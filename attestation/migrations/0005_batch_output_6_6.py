from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("attestation", "0004_batch_excel_protocols"),
    ]

    operations = [
        migrations.AddField(
            model_name="batch",
            name="output_6_6",
            field=models.CharField(
                blank=True,
                default="",
                max_length=512,
                verbose_name="Документ 6_6",
            ),
        ),
    ]
