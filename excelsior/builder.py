import six
from six.moves.urllib.request import urlopen

empty_dict = {}
empty_list = []

def create_sheet(wb, schema):
    """
        This func builds real workisheets using actual exporting library
            and provided exporting schema.
    """

    added_formats = {}
    for format in schema['formats']:
        added_formats[format] = wb.add_format(schema['formats'][format])

    for sheet in schema['worksheets']:
        worksheet = wb.add_worksheet(sheet['label'])
        if 'columns' in sheet:
            for column in sheet['columns']:
                worksheet.set_column(
                    column['first_col'],
                    column['last_col'],
                    column.get('width', None),
                    added_formats.get(column.get('format', None), None),
                    column.get('options', empty_dict)
                )

        if 'rows' in sheet:
            for row in sheet['rows']:
                worksheet.set_row(
                    row['row'],
                    row.get('height', None),
                    added_formats.get(row.get('format', None), None),
                    row.get('options', empty_dict)
                )

        for cell in sheet['cells']:
            worksheet.write(
                cell['row'],
                cell['col'],
                cell['value'],
                added_formats.get(cell.get('format', None), None)
            )

        if 'formulas' in sheet:
            for formula in sheet['formulas']:
                worksheet.write_formula(
                    formula['row'],
                    formula['col'],
                    formula['formula'],
                    added_formats.get(formula.get('format', None), None),
                    formula.get('default_value', 0)
                )

        if 'images' in sheet:
            for img in sheet['images']:
                url = img['url']
                options = img.get('options', empty_dict)
                image_data = options.get('image_data', None)

                if image_data is not None:
                    image_data = six.BytesIO(str(image_data.decode('base64')))
                elif image_data is None and url:
                    image_data = six.BytesIO(urlopen(url).read())

                options['image_data'] = image_data

                worksheet.insert_image(
                    img['row'],
                    img['col'],
                    url,
                    options
                )

        if 'hyperlinks' in sheet:
            for hyperlink in sheet['hyperlinks']:
                worksheet.write_url(
                    hyperlink['row'],
                    hyperlink['col'],
                    hyperlink['url'],
                    added_formats.get(hyperlink.get('format', None), None),
                    hyperlink.get('label', None),
                    hyperlink.get('tip', None)
                )

        if 'tables' in sheet:
            for table in sheet['tables']:
                options = table.get('options', empty_dict)
                columns = options.get('columns', empty_list)

                for column in columns:
                    format = column.get('format', None)
                    if format is not None:
                        column['format'] = added_formats.get(format, None)

                worksheet.add_table(
                    table['first_row'],
                    table['first_col'],
                    table['last_row'],
                    table['last_col'],
                    options
                )

        if 'merged_cells' in sheet:
            for merged in sheet['merged_cells']:
                worksheet.merge_range(
                    merged['first_row'],
                    merged['first_col'],
                    merged['last_row'],
                    merged['last_col'],
                    merged['data'],
                    added_formats.get(merged.get('format', None), None),
                )

        if 'autofilter' in sheet:
            atf = sheet['autofilter']
            worksheet.autofilter(
                atf['first_row'],
                atf['first_col'],
                atf['last_row'],
                atf['last_col']
            )

        if 'autofilters' in sheet:
            autofilters_list = sheet.get('autofilters') or []
            for autofilter in autofilters_list:
                worksheet.autofilter(
                    autofilter['first_row'],
                    autofilter['first_col'],
                    autofilter['last_row'],
                    autofilter['last_col']
                )

    return wb
