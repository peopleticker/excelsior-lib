from rest_framework import serializers
from rest_framework.exceptions import ValidationError


class AnyField(serializers.Field):
    """
    Any Value.
    """

    def to_representation(self, obj):
        return obj

    def to_internal_value(self, data):
        return data


class ExcelRowSerializer(serializers.Serializer):
    row = serializers.IntegerField(min_value=0)
    height = serializers.IntegerField(min_value=0)
    format = serializers.CharField(max_length=255, required=False, allow_null=True)
    options = serializers.DictField(required=False, default={})


class ExcelColumnSerializer(serializers.Serializer):
    first_col = serializers.IntegerField(min_value=0)
    last_col = serializers.IntegerField(min_value=0)
    width = serializers.IntegerField(min_value=0)
    format = serializers.CharField(max_length=255, required=False, allow_null=True)
    options = serializers.DictField(required=False, default={})


class CellBaseSerializer(serializers.Serializer):
    row = serializers.IntegerField(min_value=0)
    col = serializers.IntegerField(min_value=0)
    format = serializers.CharField(max_length=255, required=False, allow_null=True)

    class Meta:
        abstract = True


class ExcelWriteSerializer(CellBaseSerializer):
    value = AnyField(required=False, allow_null=True)


class AutofilterSerializer(serializers.Serializer):
    first_row = serializers.IntegerField(min_value=0)
    first_col = serializers.IntegerField(min_value=0)
    last_row = serializers.IntegerField(min_value=0)
    last_col = serializers.IntegerField(min_value=0)


class FormulaSerializer(CellBaseSerializer):
    formula = serializers.CharField(max_length=1024)
    default_value = serializers.IntegerField(required=False)


class ImageSerializer(CellBaseSerializer):
    url = serializers.CharField(max_length=1024)
    options = serializers.DictField(required=False, default={})


class HyperlinkSerializer(CellBaseSerializer):
    url = serializers.CharField(max_length=255)
    label = serializers.CharField(max_length=255, required=False, allow_null=True)
    tip = serializers.CharField(max_length=255, required=False, allow_null=True)


class TableSerializer(serializers.Serializer):
    first_row = serializers.IntegerField(min_value=0)
    first_col = serializers.IntegerField(min_value=0)
    last_row = serializers.IntegerField(min_value=0)
    last_col = serializers.IntegerField(min_value=0)
    options = serializers.DictField(required=False, default={})

class MergedCellsSerializer(serializers.Serializer):
    first_row = serializers.IntegerField(min_value=0)
    first_col = serializers.IntegerField(min_value=0)
    last_row = serializers.IntegerField(min_value=0)
    last_col = serializers.IntegerField(min_value=0)
    data = AnyField(required=False, allow_null=True)
    format = serializers.CharField(max_length=255, required=False, allow_null=True)


class WorksheetSerializer(serializers.Serializer):
    label = serializers.CharField(max_length=255)
    cells = ExcelWriteSerializer(many=True)

    columns = ExcelColumnSerializer(many=True, required=False, allow_null=True)
    rows = ExcelRowSerializer(many=True, required=False, allow_null=True)
    formulas = FormulaSerializer(many=True, required=False, allow_null=True)
    images = ImageSerializer(many=True, required=False, allow_null=True)
    autofilters = AutofilterSerializer(many=True, required=False, allow_null=True)
    hyperlinks = HyperlinkSerializer(many=True, required=False, allow_null=True)
    tables = TableSerializer(many=True, required=False, allow_null=True)
    merged_cells = MergedCellsSerializer(many=True, required=False, allow_null=True)


class ExcelExportSerializer(serializers.Serializer):
    filename = serializers.CharField(max_length=255)
    worksheets = WorksheetSerializer(many=True)
    formats = serializers.DictField(required=False, allow_null=True)

    def validate(self, attrs):
        format_keys = attrs.get('formats')
        if format_keys:
            format_keys = format_keys.keys()

        for worksheet in attrs['worksheets']:
            for cell in worksheet['cells']:
                if 'format' in cell:
                    if format_keys is None:
                        raise ValidationError('formats key not in object!')

                    if cell['format'] not in format_keys:
                        raise ValidationError('Cell format key %s not in formats!' % cell['format'])

        return attrs
