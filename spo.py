from email import header
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
    
    def __init__(self, tenant, _client_id, _client_secret, _tenant_id):
        self._tenant = tenant
        self._client_id = _client_id
        self._client_secret = _client_secret
        self._tenant_id = _tenant_id
        self._base_url = f"https://www.{self._tenant}.sharepoint.com/"
        self._base = f"{tenant}.sharepoint.com"
    
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
        
            
            
        
    def get_spo_oauth_token(self,client_id, client_secret,app_identifier, spo_base_url):
        spo_token_url = "https://accounts.accesscontrol.windows.net/"+app_identifier+"/tokens/oauth/2"
        default_principal_id = "00000003-0000-0ff1-ce00-000000000000"
        resource = default_principal_id + "/" + spo_base_url + "@" + app_identifier
        client_id = client_id + "@" + app_identifier
        headers = {
            'Content-Type' : 'application/x-www-form-urlencoded'
        }
        body = {
            "grant_type" : "client_credentials",
            "client_id" : client_id,
            "client_secret" : client_secret,
            "resource" : resource
        }
        print(body)
        token_data = requests.post(spo_token_url, headers=headers,data=body)
        return token_data.json()["access_token"]
    