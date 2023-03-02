from urllib import response
import environ
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
import datetime
import sys, os
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

env = environ.Env()
environ.Env.read_env()

USERNAME = env('sharepoint_email')
PASSWORD = env('sharepoint_password')
SHAREPOINT_SITE = env('sharepoint_url_site')
SHAREPOINT_SITE_NAME = env('sharepoint_site_name')
SHAREPOINT_DOC = env('sharepoint_doc_library')

class SharePoint:
    def _auth(self):
        conn = ClientContext(SHAREPOINT_SITE).with_credentials(
            UserCredential(
                USERNAME,
                PASSWORD
            )
        )
        return conn

    def _get_files_list(self, folder_name):
        print("365_1", folder_name)
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        print("==target_folder_url==", target_folder_url)

        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        print("==root_folder.files==", root_folder.files)
        return root_folder.files
    
    def get_folder_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        root_folder.expand(["Folders"]).get().execute_query()
        return root_folder.folders

    def download_file(self, file_name, folder_name):
        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}/{file_name}'
        file = File.open_binary(conn, file_url)
        return file.content
    
    def download_latest_file(self, folder_name):
        date_format = "%Y-%m-%dT%H:%M:%SZ"
        files_list = self._get_files_list(folder_name)
        file_dict = {}
        for file in files_list:
            dt_obj = datetime.datetime.strptime(file.time_last_modified, date_format)
            file_dict[file.name] = dt_obj
        # sort dict object to get the latest file
        file_dict_sorted = {key:value for key, value in sorted(file_dict.items(), key=lambda item:item[1], reverse=True)}    
        latest_file_name = next(iter(file_dict_sorted))
        content = self.download_file(latest_file_name, folder_name)
        return latest_file_name, content
        

    def upload_file(self, file_name, folder_name, content):
        conn = self._auth()
        
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.upload_file(file_name, content).execute_query()

        return response
    
    def upload_file_in_chunks(self, file_path, folder_name, chunk_size, chunk_uploaded=None, **kwargs):
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.files.create_upload_session(
            source_path=file_path,
            chunk_size=chunk_size,
            chunk_uploaded=chunk_uploaded,
            **kwargs
        ).execute_query()
        return response
    
    def get_list(self, list_name):
        conn = self._auth()
        target_list = conn.web.lists.get_by_title(list_name)
        items = target_list.items.get().execute_query()
        return items
        
    def get_file_properties_from_folder(self, folder_name):
        files_list = self._get_files_list(folder_name)
        properties_list = []
        for file in files_list:
            file_dict = {
                'file_id': file.unique_id,
                'file_name': file.name,
                'major_version': file.major_version,
                'minor_version': file.minor_version,
                'file_size': file.length,
                'time_created': file.time_created,
                'time_last_modified': file.time_last_modified
            }
            properties_list.append(file_dict)
            file_dict = {}
        return properties_list
    

    


    def upload_to_sharepoint(self,files, folder_name):
        file_path = files
        upload_file_lst = os.listdir(file_path)
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC}/{folder_name}'
        print("target_folder_url==", target_folder_url)
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)

        no_of_upload_files = len(upload_file_lst)

        with tqdm(total=no_of_upload_files, desc="Uploading..", initial=0, unit_scale=True, colour='green') as pbar:
            with ThreadPoolExecutor(max_workers=5) as executor:
                for file_name in upload_file_lst:
                    with open(os.path.join(file_path, file_name), 'rb') as f:
                        file_content = f.read()
                        futures = executor.submit(target_folder.upload_file, file_name, file_content)
                        if as_completed(futures):
                            futures.result().execute_query()
                            pbar.update(1)
        return True
    

    def create_self_folder_upload_folder_to_sharepoint(self,files, folder_name,local_path):
        file_path = files
        upload_file_lst = os.listdir(file_path)
        conn = self._auth()
        local_folder = folder_name.split("/")
        f_path = "Shared Documents"
        for folder in local_folder:
            if len(folder) >= 1:
                f_path += f'/{folder}'
                archive_folder = conn.web.folders.add(f_path)
                conn.execute_query()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{f_path}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        no_of_upload_files = len(upload_file_lst)
        with tqdm(total=no_of_upload_files, desc="Uploading..", initial=0, unit_scale=True, colour='green') as pbar:
            with ThreadPoolExecutor(max_workers=5) as executor:
                for file_name in upload_file_lst:
                    with open(os.path.join(file_path, file_name), 'rb') as f:
                        file_content = f.read()
                        futures = executor.submit(target_folder.upload_file, file_name, file_content)
                        if as_completed(futures):
                            futures.result().execute_query()
                            pbar.update(1)
        return True
    
    def upload_folder_to_sharepoint(self,files, local_file_path):
        file_path = files
        upload_file_lst = os.listdir(file_path)
        conn = self._auth()
        local_folder = local_file_path.split("/")
        f_path = "Shared Documents"
        for folder in local_folder:
            if len(folder) >= 1:
                f_path += f'/{folder}'
                archive_folder = conn.web.folders.add(f_path)
                conn.execute_query()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{f_path}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        no_of_upload_files = len(upload_file_lst)
        with tqdm(total=no_of_upload_files, desc="Uploading..", initial=0, unit_scale=True, colour='green') as pbar:
            with ThreadPoolExecutor(max_workers=5) as executor:
                for file_name in upload_file_lst:
                    with open(os.path.join(file_path, file_name), 'rb') as f:
                        file_content = f.read()
                        futures = executor.submit(target_folder.upload_file, file_name, file_content)
                        if as_completed(futures):
                            futures.result().execute_query()
                            pbar.update(1)
        return True
