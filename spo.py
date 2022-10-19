import requests

class SharePointHelper():
    _tenant = None
    _client_id = None
    _client_secret = None
    _tenant_id = None
    _base_url = None
    _oauth2_token = None
    _token_url = "https://accounts.accesscontrol.windows.net"
    _default_principal_audience_id = "00000003–0000–0ff1-ce00–000000000000"
    
    def __init__(self, tenant, _client_id, _client_secret, _tenant_id, init_oauth2_token = True):
        self._tenant = tenant
        self._client_id = _client_id
        self._client_secret = _client_secret
        self._tenant_id = _tenant_id
        self._base_url = f"https://{tenant}.sharepoint.com"
        self._base = f"{tenant}.sharepoint.com"
        self._oauth2_token = self.get_oauth2_token() if init_oauth2_token else None
    
    def get_oauth2_token(self):
        spo_token_url = f"https://accounts.accesscontrol.windows.net/{self._tenant_id}/tokens/oauth/2"
        default_principal_id = "00000003-0000-0ff1-ce00-000000000000"
        resource = f"{default_principal_id}/{self._tenant}@{self._tenant_id}"
        client_id = f"{self._client_id}@{self._tenant_id}"
        headers = {
            'Content-Type' : 'application/x-www-form-urlencoded'
        }
        body = {
            "grant_type" : "client_credentials",
            "client_id" : client_id,
            "client_secret" :self._client_secret,
            "resource" : resource
        }
        token_data = requests.post(spo_token_url, headers=headers,data=body)
        return token_data.json()["access_token"]
    
    def upload_file(self,site,spo_path,file_path, file_name = None):
        data = open(file_path).read()
        spo_endpoint = f"{self._base_url}/sites/{site}/_api/web/getFolderByServerRelativeUrl('{spo_path}')/files/add(url='{file_name}', overwrite=True)"
        self.send_request(spo_endpoint, "POST",data, None, self._oauth2_token)
        #
    def send_request(self, url, method='GET',body=None,headers=None, token=None):
        if token == None and self._oauth2_token is None:
            raise Exception("Please Provide a valid token")
        else:
            headers={
            'Authorization': 'Bearer ' + self._oauth2_token if self._oauth2_token is not None else token,
            'Accept': 'application/json',
            'Content-Type': 'application/json'
            }
        try:
            response = requests.request(method, url, headers=headers, data=body)
            print(headers)
            print(url)
           # response = requests.post(url, data=body, headers=headers)
            print(response.json())
            return response.json()
        except:
            print("hallo")
          #  raise Exception("Could not send request. Please check your Input Parameters")

    