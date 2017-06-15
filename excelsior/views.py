import json
import urllib
import xlsxwriter
from StringIO import StringIO
from django.http import FileResponse
from rest_framework import status
from rest_framework.views import APIView
from rest_framework.parsers import JSONParser, FormParser, MultiPartParser

# Import urlopen() for either Python 2 or 3.
try:
    from urllib.request import urlopen
except ImportError:
    from urllib2 import urlopen

from serializers import ExcelExportSerializer
from builder import create_sheet


def schema_to_file_object(schema):
    file_object = StringIO()
    workbook = xlsxwriter.Workbook(file_object, {'in_memory': True})
    workbook = create_sheet(workbook, schema)
    workbook.close()
    file_object.seek(0)

    return file_object


class ExportToExcelMixin(object):
    """
        This Mixin will help you make file object from your exporting schema
            and stream that file object back to the client
    """
    def schema_to_file_object(self, schema):
        return schema_to_file_object(schema)

    def stream_as_file(self, schema, **extra_params):
        file_object = self.schema_to_file_object(schema)
        response = FileResponse(file_object, status=status.HTTP_200_OK)
        filename = schema['filename']
        # NOTE: Gory details http://greenbytes.de/tech/tc2231/
        filename_fallback = "filename*=UTF-8''{}".format(urllib.quote(filename.encode('utf-8')))
        response['Content-Disposition'] = 'attachment; filename="{}"; {}'.format(filename, filename_fallback)
        response['Content-Type'] = 'application/vnd.ms-excel'

        cookie = (extra_params.get('cookie') or
                        self.request.data.get('cookie'))
        if cookie:
            response.set_cookie(cookie, 'true')

        return response


class ExportToExcelView(ExportToExcelMixin, APIView):
    """
        Basic exporting view which takes already prepared worksheet config
            from request body and turns it into the .xlsx file.
        Same does Excelsior API, so you can use this view in case you don't
            want to spent any time on API calls.
    """

    parser_classes = (JSONParser, MultiPartParser, FormParser)
    form_media_types = (FormParser.media_type, MultiPartParser.media_type)

    def post(self, request, *args, **kwargs):
        # Notice: We allow submitting both types - ajax and form based
        schema = request.data['data']
        if (request.content_type in self.form_media_types
                and isinstance(schema, basestring)):
            schema = json.loads(schema)

        serializer = ExcelExportSerializer(data=schema)
        serializer.is_valid(raise_exception=True)

        return self.stream_as_file(serializer.data)
