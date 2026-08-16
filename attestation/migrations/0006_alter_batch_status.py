from django.db import migrations, models


class Migration(migrations.Migration):
    dependencies = [
        ("attestation", "0005_batch_output_6_6"),
    ]

    operations = [
        migrations.AlterField(
            model_name="batch",
            name="status",
            field=models.CharField(
                choices=[
                    ("uploaded", "Загружен"),
                    ("processing", "Обрабатывается"),
                    ("extracted", "Извлечён"),
                    ("done", "Готов"),
                    ("deleting", "Удаляется"),
                    ("failed", "Ошибка"),
                ],
                default="uploaded",
                max_length=16,
                verbose_name="Статус",
            ),
        ),
    ]
