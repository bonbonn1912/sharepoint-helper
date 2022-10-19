from email import header
import requests

class SharePointHelper():
    _tenant = None
    _client_id = None
    _client_secret = None
    _tenant_id = None
    _base_url = None
    _token_url = "https://accounts.accesscontrol.windows.net/"
    _default_principal_audience_id = "00000003–0000–0ff1-ce00–000000000000"
    
    def __init__(self, tenant, _client_id, _client_secret, _tenant_id):
        self._tenant = tenant
        self._client_id = _client_id
        self._client_secret = _client_secret
        self._tenant_id = _tenant_id
        self._base_url = f"https://www.{self._tenant}.sharepoint.com/"
        self._base = f"{self._tenant}.sharepoint.com"
    
    def get_oauth2_token(self):
        token_url = f"{self._token_url}/{self._tenant_id}/tokens/oauth/2"
        resource = f"{self._default_principal_audience_id}/{self._base}@{self._tenant_id}"
        client_id = f"{self._client_id}@{self._tenant_id}"
        token = None
        headers = {
            "Content-Type": "application/x-www-form-urlencoded"
        }
        body = {
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": self._client_secret,
            "resource": resource
        }
        try:
            data = requests.post(token_url, data=body, headers=headers)
            token = data.json()
            return token
        except:
            raise Exception("Could not retrieve token")
        
            
            
        
        
    