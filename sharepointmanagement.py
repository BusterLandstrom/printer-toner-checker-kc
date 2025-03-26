import codecs
from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.files.file import File
from chardet import detect
from io import BytesIO
from os import path

# Simple SharePoint functions
class SharePointHandler():
    '''
        These functions dont have explanations because
        the function names and variables should be easy to understand.

        Only thing is that the sp functions only do stuff via the SharePoint API and nothing else.
    '''

    def upload_item(self, function_context, local_file_path, rel_sp_path):
        target_folder = function_context.web.get_folder_by_server_relative_path(rel_sp_path)
        filename = path.basename(local_file_path)
        with open(local_file_path, 'rb') as content_file:
            file_content = content_file.read()
            target_file = target_folder.upload_file(filename, file_content).execute_query()

    def upload_sp_item(self, function_context, sp_file, rel_sp_path, name):
        target_folder = function_context.web.get_folder_by_server_relative_path(rel_sp_path)
        file_content = BytesIO(sp_file.buffer.read())
        target_file = target_folder.upload_file(name, file_content).execute_query()

    def download_item(self, function_context, sp_file_path, local_file_path):
        response = File.open_binary(function_context, sp_file_path)
        with open(local_file_path, "wb") as local_file:
            local_file.write(response.content)

    def read_from_sp_file(self, function_context, sp_file_path):
        response = File.open_binary(function_context, sp_file_path)
        response_file = BytesIO(response.content)
        encoding = detect(response_file.read())['encoding']
        response_file.seek(0)
        contentout = codecs.iterdecode(response_file, encoding)
        return contentout

    def write_from_sp_file(self, sp_file, local_file_path):
        with open(local_file_path, "wb") as local_file:
            local_file.write(sp_file)

    def fop_exist(self, function_context, sp_path, error_path):
        try:
            return function_context.web.get_file_by_server_relative_url(sp_path).get().execute_query()
        except ClientRequestException as e:
            if e.response.status_code == 404:
                return None
            else:
                raise ValueError(e.response.text)
# END