import six

"""
    Schema API layer. Use it in the real exports implementations
        to build proper export schema.
    This Impl ported from the reactr-frontend repo,
    See file reactr-frontend://src/shared/utils/excelsior.js
"""


__all__ = [
    'ExportError', 'Worksheet', 'WorkbookBuilder'
]


class ExportError(Exception):
    pass


class Worksheet(object):
    def __init__(self, workbook, label):
        self.workbook = workbook
        self.label = label
        self.cells = []
        self.columns = []
        self.rows = []
        self.tables = []
        self.merged_cells = []
        self.hyperlinks = []
        self.formulas = []
        self.images = []
        self.autofilters = []
        self.frozen_panes = []
        self.current_row_idx = 0

    def write_empty_rows(self, rows_to_write = 1):
        self.current_row_idx += rows_to_write
        return self

    def set_column_width(self, column_idx, width, format=None, options=None):
        column = {
          'first_col': column_idx,
          'last_col': column_idx,
          'width': width
        }

        if format is not None:
            if not self.workbook.has_format(format):
                raise ExportError(
                    'Column %d refers to non-existent format %s' \
                         % (column_idx, format)
                )
            column['format'] = format

        if options is not None:
            column['options'] = options

        self.columns.append(column)

        return self

    def set_row_height(self, row_idx, height, format=None, options=None):
        row = {
          'row': row_idx,
          'height': height,
        }

        if format is not None:
            if not self.workbook.has_format(format):
                raise ExportError(
                    'Row %d refers to non-existent format %s' \
                         % (row_idx, format)
                )
            row['format'] = format

        if options is not None:
            row['options'] = options

        self.rows.append(row)

        return self

    def write_row(self, *columns):
        """
            Writes a row

            Example:
                write_row(
                    {'v': 'Foo Bar'},
                    None,
                    {'f': 'bolder', 'v': 'I am bolder'},
                    'Also can be a string'
                )

        """
        current_column_idx = 0
        current_row_idx = self.current_row_idx
        self.current_row_idx+=1

        for column in columns:
            if column is None:
                current_column_idx+=1
                continue

            format = None
            value = None

            if isinstance(column, six.string_types):
                value = column
            elif isinstance(column, dict):
                format = column.get('f')
                value = column.get('v')
            else:
                raise ExportError('Invalid column of type %s provided' % type(column))

            if format is not None:
                if not self.workbook.has_format(format):
                    raise ExportError(
                        'Row %d, Column %d refers to non-existent format %s' \
                             % (current_row_idx, current_column_idx, format)
                    )

            if value is None:
                raise ExportError(
                    """Row %d, Column %d has no value.
                        Pass null instead of an object.""" \
                    % (current_row_idx, current_column_idx)
                )

            cell = {
                'row': current_row_idx,
                'col': current_column_idx,
                'value': value
            }
            current_column_idx += 1

            if format is not None :
                cell['format'] = format

            self.cells.append(cell)

        return self

    def write_cell(self, row_idx, col_idx, value, format=None):
        cell = {
            'row': row_idx,
            'col': col_idx,
            'value': value
        }

        if format is not None:
            if not self.workbook.has_format(format):
                raise ExportError(
                    'Row %d, Column %d refers to non-existent format %s' \
                         % (row_idx, col_idx, format)
                )
            cell['format'] = format

        self.cells.append(cell)

        return self

    def write_merged_cell(self, first_row_idx, first_col_idx, last_row_idx, last_col_idx,
                        data=None, format=None):
        cell = {
            'first_row': first_row_idx,
            'first_col': first_col_idx,
            'last_row': last_row_idx,
            'last_col': last_col_idx,
            'data': data
        }

        if format is not None:
            if not self.workbook.has_format(format):
                raise ExportError(
                    'Merged cell (from row %d, column %d, to row %d, column %d)' +
                    ' refers to non-existent format %s' \
                    % (first_row_idx, first_col_idx, last_row_idx, last_col_idx, format)
                )
            cell['format'] = format

        self.merged_cells.append(cell)

        return self

    def write_hyperlink(self, row_idx, col_idx, url,
                        format=None, label=None, tip=None):
        hyperlink = {
            'row': row_idx,
            'col': col_idx,
            'url': url
        }

        if format is not None:
            if not self.workbook.has_format(format):
                raise ExportError(
                    'Row %d, Column %d refers to non-existent format %s' \
                         % (row_idx, col_idx, format)
                )
            hyperlink['format'] = format
        if label is not None:
            hyperlink['label'] = label
        if tip is not None:
            hyperlink['tip'] = tip

        self.hyperlinks.append(hyperlink)

        return self

    def write_formula(self, row_idx, col_idx, formula,
                      format=None, default_value=None):
        formula = {
            'row': row_idx,
            'col': col_idx,
            'formula': formula
        }

        if format is not None:
            if not self.workbook.has_format(format):
                raise ExportError(
                    'Row %d, Column %d refers to non-existent format %s' \
                         % (row_idx, col_idx, format)
                )
            formula['format'] = format
        if default_value is not None:
            formula['default_value'] = default_value

        self.formulas.append(formula)

    def write_image(self, row_idx, col_idx, url, options={}):
        self.images.append({
            'row': row_idx,
            'col': col_idx,
            'url': url,
            'options': options
        })

        return self

    def write_table(self, first_row_idx, first_col_idx, last_row_idx, last_col_idx,
                    options={}):
        self.tables.append({
            'first_row': first_row_idx,
            'first_col': first_col_idx,
            'last_row': last_row_idx,
            'last_col': last_col_idx,
            'options': options
        })

        return self

    def add_autofilter(self, first_row_idx, first_col_idx, last_row_idx, last_col_idx):
        self.autofilters.append({
            'first_row': first_row_idx,
            'first_col': first_col_idx,
            'last_row': last_row_idx,
            'last_col': last_col_idx,
        })

    def add_frozen_pane(self, row_idx, col_idx, top_row=None, left_col=None):
        self.frozen_panes.append({
            'row': row_idx,
            'col': col_idx,
            'top_row': top_row,
            'left_col': left_col,
        })

    def as_dict(self):
        schema = {
            'cells': self.cells,
            'label': self.label
        }

        if len(self.columns) > 0:
            schema['columns'] = self.columns

        if len(self.rows) > 0:
            schema['rows'] = self.rows

        if len(self.hyperlinks) > 0:
            schema['hyperlinks'] = self.hyperlinks

        if len(self.formulas) > 0:
            schema['formulas'] = self.formulas

        if len(self.images) > 0:
            schema['images'] = self.images

        if len(self.tables) > 0:
            schema['tables'] = self.tables

        if len(self.merged_cells) > 0:
            schema['merged_cells'] = self.merged_cells

        if len(self.autofilters) > 0:
            schema['autofilters'] = self.autofilters

        if len(self.frozen_panes) > 0:
            schema['frozen_panes'] = self.frozen_panes        

        return schema


class WorkbookBuilder(object):
    def __init__(self, filename):
        self.filename = filename
        self.formats = {}
        self.worksheets = []

    def add_format(self, format_key, format_spec):
        if format_key in self.formats:
            raise ExportError(
                'A format with the key %s already exists.' % format_key
            )

        self.formats[format_key] = format_spec

        return self

    def has_format(self, format_key):
        return format_key in self.formats

    def add_worksheet(self, worksheet_label):

        """ Add a worksheet to the workbook. """

        worksheet = Worksheet(self, worksheet_label)
        self.worksheets.append(worksheet)

        return worksheet

    def as_dict(self):
        if len(self.worksheets) == 0:
            raise ExportError('Trying to serialize an empty workbook.')

        schema = {
            'filename': self.filename,
            'worksheets': map(lambda w: w.as_dict(), self.worksheets)
        }

        if len(self.formats.keys()) > 0:
            schema['formats'] = self.formats

        return schema
